package pptx

import (
	"fmt"
	"sort"
)

// SlideData はスライドのパース結果
type SlideData struct {
	Number   int
	Title    string
	Hidden   bool
	HasNotes bool
	Shapes   []Shape
	Notes    []Paragraph
}

// Info は SlideData からヘッダ情報を返す
func (sd *SlideData) Info() SlideInfo {
	return SlideInfo{
		Slide:    sd.Number,
		Title:    sd.Title,
		HasNotes: sd.HasNotes,
		Hidden:   sd.Hidden,
	}
}

// LoadSlide は指定スライドの内容をパースする
func (f *File) LoadSlide(slideNum int, includeNotes bool) (*SlideData, error) {
	idx := slideNum - 1
	sld, err := f.loadSlideXML(idx)
	if err != nil {
		return nil, err
	}
	if sld == nil {
		return nil, fmt.Errorf("スライド %d が見つかりません", slideNum)
	}

	entry := f.slideEntries[idx]

	// スライドのリレーション（画像・コネクタ用）
	rels, err := loadRelsTyped(f, relsPathFor(entry.Path))
	if err != nil {
		return nil, fmt.Errorf("スライド %d のリレーション読み込みに失敗: %w", slideNum, err)
	}
	slideRels := relsToMap(rels)

	sd := &SlideData{
		Number:   slideNum,
		Title:    extractTitle(sld.CSld.SpTree.Children),
		Hidden:   sld.Show == "0", // OOXML: show属性のデフォルトは"1"（表示）、省略時は空文字列（表示扱い）
		HasNotes: f.hasNotes(idx),
	}

	// レイアウト/マスターの継承データを取得
	ic := f.getInheritCache()
	var layout *layoutData
	var master *masterData
	layoutPath := resolveLayoutPath(f, entry.Path)
	if layoutPath != "" {
		layout = ic.getLayout(f, layoutPath)
		if layout != nil && layout.masterPath != "" {
			master = ic.getMaster(f, layout.masterPath)
		}
	}
	if layout == nil {
		layout = &layoutData{placeholders: make(map[phKey]*placeholderDef)}
	}
	if master == nil {
		master = &masterData{placeholders: make(map[phKey]*placeholderDef)}
	}

	// 図形をパース
	ctx := &parseContext{
		f:         f,
		slideRels: slideRels,
		slidePath: entry.Path,
		pptxIDMap: make(map[int]int),
		layout:    layout,
		master:    master,
	}

	sd.Shapes = ctx.parseSpTree(sld.CSld.SpTree.Children)

	// コネクタの from/to を解決
	ctx.resolveConnectors(sd.Shapes)

	// ノート
	if includeNotes {
		sd.Notes = f.loadNotesParagraphs(idx)
	}

	return sd, nil
}

// parseContext はスライドパース中のコンテキスト
type parseContext struct {
	f         *File
	slideRels map[string]string
	slidePath string
	nextID    int         // 連番ID
	pptxIDMap map[int]int // PowerPoint図形ID → 連番ID
	nextZ     int         // z-order カウンタ
	layout    *layoutData // レイアウトデータ（継承解決用）
	master    *masterData // マスターデータ（継承解決用）
}

// newTextOnlyContext はテキスト解析専用の parseContext を生成する。
// ノート等、テキストのみを扱う場面で使用する。
// ID割り当てや画像抽出は行わないため、それらのフィールドは初期値のまま。
func newTextOnlyContext(f *File) *parseContext {
	return &parseContext{
		f:         f,
		pptxIDMap: make(map[int]int),
	}
}

// newChildContext はグループの子要素パース用のサブコンテキストを生成する。
// pptxIDMap は参照共有: 子で登録したIDが親の resolveConnectors からも参照可能。
// カウンタ類は値コピーされ、パース後に syncFromChild で同期する。
func (ctx *parseContext) newChildContext() *parseContext {
	return &parseContext{
		f:         ctx.f,
		slideRels: ctx.slideRels,
		slidePath: ctx.slidePath,
		nextID:    ctx.nextID,
		nextZ:     ctx.nextZ,
		pptxIDMap: ctx.pptxIDMap,
		layout:    ctx.layout,
		master:    ctx.master,
	}
}

// syncFromChild は子コンテキストのカウンタを親に同期する
func (ctx *parseContext) syncFromChild(child *parseContext) {
	ctx.nextID = child.nextID
	ctx.nextZ = child.nextZ
}

func (ctx *parseContext) allocID(pptxID int) int {
	ctx.nextID++
	id := ctx.nextID
	if pptxID != 0 {
		ctx.pptxIDMap[pptxID] = id
	}
	return id
}

func (ctx *parseContext) allocZ() int {
	z := ctx.nextZ
	ctx.nextZ++
	return z
}

// shapeItem はソート用の中間構造
type shapeItem struct {
	order      int
	shape      Shape
	isPH       bool
	phPriority int
}

// parseSpTree は子要素をXML出現順にパースする
func (ctx *parseContext) parseSpTree(children []xmlSpTreeChild) []Shape {
	items := make([]shapeItem, 0)

	for order, child := range children {
		var s *Shape
		var isPH bool
		var priority int

		switch {
		case child.Sp != nil:
			s = ctx.parseSp(*child.Sp)
			if s != nil {
				ph := child.Sp.NvSpPr.NvPr.Ph
				isPH = ph != nil
				priority = phPriority(ph)
			}
		case child.CxnSp != nil:
			s = ctx.parseCxnSp(*child.CxnSp)
		case child.Pic != nil:
			s = ctx.parsePic(*child.Pic)
		case child.GrpSp != nil:
			s = ctx.parseGrpSp(*child.GrpSp)
		case child.GraphicFrame != nil:
			s = ctx.parseGraphicFrame(*child.GraphicFrame)
		}

		if s == nil {
			continue
		}
		// z-order はXML出現順（PowerPointの描画順）を反映する
		s.Z = ctx.allocZ()
		items = append(items, shapeItem{order: order, shape: *s, isPH: isPH, phPriority: priority})
	}

	// ソート: プレースホルダー（優先度順）→ 非プレースホルダー（出現順）
	// 出力順序はプレースホルダー優先だが、z-order は元のXML出現順を保持する
	sortShapeItems(items)

	shapes := make([]Shape, 0, len(items))
	for _, item := range items {
		shapes = append(shapes, item.shape)
	}

	return shapes
}

// sortShapeItems はプレースホルダー優先でソートする
func sortShapeItems(items []shapeItem) {
	sort.SliceStable(items, func(i, j int) bool {
		return lessShapeItem(items[i], items[j])
	})
}

func lessShapeItem(a, b shapeItem) bool {
	if a.isPH != b.isPH {
		return a.isPH
	}
	if a.isPH && b.isPH {
		if a.phPriority != b.phPriority {
			return a.phPriority < b.phPriority
		}
	}
	return a.order < b.order
}

func phPriority(ph *xmlPh) int {
	if ph == nil {
		return 99
	}
	switch ph.Type {
	case "title", "ctrTitle":
		return 0
	case "subTitle":
		return 1
	case "body":
		return 2
	default:
		return 3
	}
}


