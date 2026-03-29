package pptx

import (
	"encoding/xml"
	"fmt"
	"sort"
)

// SlideData はスライドのパース結果
type SlideData struct {
	Number int
	Title  string
	Shapes []Shape
	Notes  []Paragraph
}

// LoadSlide は指定スライドの内容をパースする
func (f *File) LoadSlide(slideNum int, includeNotes bool, extractDir string) (*SlideData, error) {
	idx := slideNum - 1
	if idx < 0 || idx >= len(f.slideEntries) {
		return nil, fmt.Errorf("スライド番号 %d は範囲外です（1〜%d）", slideNum, len(f.slideEntries))
	}

	entry := f.slideEntries[idx]

	data, err := readZipFile(f.zi, entry.Path)
	if err != nil {
		return nil, fmt.Errorf("スライド %d の読み込みに失敗: %w", slideNum, err)
	}
	if data == nil {
		return nil, fmt.Errorf("スライド %d が見つかりません", slideNum)
	}

	var sld xmlSlide
	if err := xml.Unmarshal(data, &sld); err != nil {
		return nil, fmt.Errorf("スライド %d のパースに失敗: %w", slideNum, err)
	}

	// スライドのリレーション（画像・コネクタ用）
	slideRels, err := loadRels(f, slideRelsPath(entry.Path))
	if err != nil {
		return nil, fmt.Errorf("スライド %d のリレーション読み込みに失敗: %w", slideNum, err)
	}

	sd := &SlideData{
		Number: slideNum,
		Title:  extractTitle(sld.CSld.SpTree.Children),
	}

	// 図形をパース
	ctx := &parseContext{
		f:          f,
		slideRels:  slideRels,
		slidePath:  entry.Path,
		extractDir: extractDir,
		pptxIDMap:  make(map[int]int),
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
	f          *File
	slideRels  map[string]string
	slidePath  string
	extractDir string
	nextID     int         // 連番ID
	pptxIDMap  map[int]int // PowerPoint図形ID → 連番ID
	nextZ int // z-order カウンタ
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
		f:          ctx.f,
		slideRels:  ctx.slideRels,
		slidePath:  ctx.slidePath,
		extractDir: ctx.extractDir,
		nextID:    ctx.nextID,
		nextZ:     ctx.nextZ,
		pptxIDMap: ctx.pptxIDMap,
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

// parseSp は通常の図形をパースする
func (ctx *parseContext) parseSp(sp xmlSp) *Shape {
	ph := sp.NvSpPr.NvPr.Ph

	// テキスト・塗りつぶし・枠線のいずれもない図形はスキップ（プレースホルダー含む）
	hasText := hasTextContent(sp.TxBody)
	hasFill := sp.SpPr.SolidFill != nil
	hasLine := sp.SpPr.Ln != nil && sp.SpPr.Ln.NoFill == nil
	if !hasText && !hasFill && !hasLine {
		return nil
	}

	s := &Shape{
		ID: ctx.allocID(sp.NvSpPr.CNvPr.ID),
	}

	// 図形種別
	if sp.SpPr.PrstGeom != nil {
		s.Type = sp.SpPr.PrstGeom.Prst
	} else if sp.SpPr.CustGeom != nil {
		s.Type = "customShape"
	} else {
		s.Type = "rect" // デフォルト
	}

	// 名前とプレースホルダー
	if ph != nil {
		s.Placeholder = ph.Type
		if s.Placeholder == "" {
			s.Placeholder = "body" // type未指定のプレースホルダーはbody扱い
		}
	} else {
		s.Name = sp.NvSpPr.CNvPr.Name
	}

	// 位置
	s.Pos = xfrmToPosition(sp.SpPr.Xfrm)

	// 回転・反転
	if sp.SpPr.Xfrm != nil {
		s.Rotation = float64(sp.SpPr.Xfrm.Rot) / 60000.0
		s.Flip = xfrmFlip(sp.SpPr.Xfrm)
	}

	// 塗りつぶし
	s.Fill = ctx.resolveSolidFillColor(sp.SpPr.SolidFill)

	// 枠線
	s.Line = ctx.resolveLine(sp.SpPr.Ln)

	// 吹き出しポインタ
	s.CalloutPointer = resolveCalloutPointer(sp.SpPr.PrstGeom, s.Pos)

	// テキスト
	if sp.TxBody != nil {
		s.Paragraphs = ctx.parseParagraphs(sp.TxBody.Ps)
		s.Alignment = ctx.extractShapeLevelAlignment(sp.TxBody)
	}

	return s
}


// parsePic は画像をパースする
func (ctx *parseContext) parsePic(pic xmlPic) *Shape {
	s := &Shape{
		ID:   ctx.allocID(pic.NvPicPr.CNvPr.ID),
		Type: "picture",
		Name: pic.NvPicPr.CNvPr.Name,
	}

	// 代替テキスト
	s.AltText = pic.NvPicPr.CNvPr.Descr

	// 位置
	s.Pos = xfrmToPosition(pic.SpPr.Xfrm)

	// 画像の抽出（extractDir が指定されている場合のみ）
	if ctx.extractDir != "" && pic.BlipFill.Blip.Embed != "" {
		s.ImagePath = ctx.extractImage(pic.BlipFill.Blip.Embed)
	}

	return s
}

// parseGrpSp はグループをパースする
func (ctx *parseContext) parseGrpSp(grp xmlGrpSp) *Shape {
	s := &Shape{
		ID:   ctx.allocID(grp.NvGrpSpPr.CNvPr.ID),
		Type: "group",
		Name: grp.NvGrpSpPr.CNvPr.Name,
	}

	// グループの位置
	if grp.GrpSpPr.Xfrm != nil {
		s.Pos = &Position{
			X: grp.GrpSpPr.Xfrm.Off.X,
			Y: grp.GrpSpPr.Xfrm.Off.Y,
			W: grp.GrpSpPr.Xfrm.Ext.Cx,
			H: grp.GrpSpPr.Xfrm.Ext.Cy,
		}
	}

	// 子要素のパース
	childCtx := ctx.newChildContext()
	s.Children = childCtx.parseSpTree(grp.Children)
	ctx.syncFromChild(childCtx)

	if len(s.Children) == 0 {
		return nil
	}

	return s
}


// loadNotesParagraphs はスライドのノートの段落を取得する。
// ノートの読み込み・パース失敗時はnilを返す（スライド処理は継続する）。
func (f *File) loadNotesParagraphs(slideIdx int) []Paragraph {
	txBody := f.findNotesBody(slideIdx)
	if txBody == nil {
		return nil
	}
	ctx := newTextOnlyContext(f)
	paras := ctx.parseParagraphs(txBody.Ps)
	if len(paras) > 0 {
		return paras
	}
	return nil
}
