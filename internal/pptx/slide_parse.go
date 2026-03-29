package pptx

import (
	"encoding/xml"
	"fmt"
	"sort"
	"strings"
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
	nextZ      int         // z-order カウンタ
	imageCount int         // 画像カウンタ
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
		nextID:     ctx.nextID,
		nextZ:      ctx.nextZ,
		pptxIDMap:  ctx.pptxIDMap,
		imageCount: ctx.imageCount,
	}
}

// syncFromChild は子コンテキストのカウンタを親に同期する
func (ctx *parseContext) syncFromChild(child *parseContext) {
	ctx.nextID = child.nextID
	ctx.nextZ = child.nextZ
	ctx.imageCount = child.imageCount
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
	s.Position = xfrmToPosition(sp.SpPr.Xfrm)

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
	s.CalloutPointer = resolveCalloutPointer(sp.SpPr.PrstGeom, s.Position)

	// テキスト
	if sp.TxBody != nil {
		s.Paragraphs = ctx.parseParagraphs(sp.TxBody.Ps)
		s.Alignment = ctx.extractShapeLevelAlignment(sp.TxBody)
	}

	return s
}

// parseCxnSp はコネクタをパースする
func (ctx *parseContext) parseCxnSp(cxn xmlCxnSp) *Shape {
	s := &Shape{
		ID:   ctx.allocID(cxn.NvCxnSpPr.CNvPr.ID),
		Type: "connector",
		Name: cxn.NvCxnSpPr.CNvPr.Name,
	}

	// コネクタ形状
	if cxn.SpPr.PrstGeom != nil {
		s.ConnectorType = cxn.SpPr.PrstGeom.Prst
	}

	// 位置
	s.Position = xfrmToPosition(cxn.SpPr.Xfrm)

	// 枠線
	s.Line = ctx.resolveLine(cxn.SpPr.Ln)

	// 矢印
	s.Arrow = resolveArrow(cxn.SpPr.Ln)

	// 接続情報（PowerPoint ID。後で解決する）
	if cxn.NvCxnSpPr.CNvCxnSpPr.StCxn != nil {
		s.From = -cxn.NvCxnSpPr.CNvCxnSpPr.StCxn.ID // 負値で未解決マーク
	}
	if cxn.NvCxnSpPr.CNvCxnSpPr.EndCxn != nil {
		s.To = -cxn.NvCxnSpPr.CNvCxnSpPr.EndCxn.ID
	}

	// テキスト
	if cxn.TxBody != nil {
		paras := ctx.parseParagraphs(cxn.TxBody.Ps)
		if len(paras) > 0 {
			var texts []string
			for _, p := range paras {
				texts = append(texts, p.Text)
			}
			s.Label = strings.Join(texts, "\n")
		}
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
	s.Position = xfrmToPosition(pic.SpPr.Xfrm)

	// 画像の抽出（extractDir が指定されている場合のみ）
	if ctx.extractDir != "" && pic.BlipFill.Blip.Embed != "" {
		s.Image = ctx.extractImage(pic.BlipFill.Blip.Embed, s.Position)
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
		s.Position = &Position{
			X:  grp.GrpSpPr.Xfrm.Off.X,
			Y:  grp.GrpSpPr.Xfrm.Off.Y,
			Cx: grp.GrpSpPr.Xfrm.Ext.Cx,
			Cy: grp.GrpSpPr.Xfrm.Ext.Cy,
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

// parseGraphicFrame はテーブル等のgraphicFrameをパースする
func (ctx *parseContext) parseGraphicFrame(gf xmlGraphicFrame) *Shape {
	tbl := gf.Graphic.GraphicData.Tbl
	if tbl == nil {
		return nil // テーブル以外のgraphicFrameはスキップ
	}

	s := &Shape{
		ID:   ctx.allocID(gf.NvGraphicFramePr.CNvPr.ID),
		Type: "table",
		Name: gf.NvGraphicFramePr.CNvPr.Name,
	}

	// 位置
	s.Position = xfrmToPosition(gf.Xfrm)

	// テーブルデータ（被結合セルは null）
	cols := len(tbl.TblGrid.GridCols)
	var rows [][]*string

	// rowSpan による被結合セルを後のパスで null にするための記録
	type rowSpanArea struct {
		row, col, rowSpan, colSpan int
	}
	var rowSpans []rowSpanArea

	for _, tr := range tbl.Trs {
		row := make([]*string, cols)
		colIdx := 0
		for _, tc := range tr.Tcs {
			if colIdx >= cols {
				break
			}
			if tc.VMerge != "1" && tc.HMerge != "1" {
				text := extractTextFromTxBody(tc.TxBody)
				row[colIdx] = &text
			}
			span := tc.GridSpan
			if span < 1 {
				span = 1
			}
			if tc.RowSpan > 1 {
				rowSpans = append(rowSpans, rowSpanArea{
					row: len(rows), col: colIdx, rowSpan: tc.RowSpan, colSpan: span,
				})
			}
			colIdx += span
		}
		rows = append(rows, row)
	}

	// rowSpan による被結合セルを null にする
	// （標準XMLでは vMerge で既に null だが、vMerge 省略時のフォールバック）
	for _, rs := range rowSpans {
		for r := rs.row + 1; r < rs.row+rs.rowSpan && r < len(rows); r++ {
			for c := rs.col; c < rs.col+rs.colSpan && c < cols && c < len(rows[r]); c++ {
				rows[r][c] = nil
			}
		}
	}

	s.Table = &TableData{
		Cols: cols,
		Rows: rows,
	}

	return s
}

// resolveConnectors はコネクタの from/to を PowerPoint ID から連番IDに変換する
func (ctx *parseContext) resolveConnectors(shapes []Shape) {
	for i := range shapes {
		if shapes[i].Type == "connector" {
			if shapes[i].From < 0 {
				pptxID := -shapes[i].From
				if resolved, ok := ctx.pptxIDMap[pptxID]; ok {
					shapes[i].From = resolved
				} else {
					shapes[i].From = 0
				}
			}
			if shapes[i].To < 0 {
				pptxID := -shapes[i].To
				if resolved, ok := ctx.pptxIDMap[pptxID]; ok {
					shapes[i].To = resolved
				} else {
					shapes[i].To = 0
				}
			}
		}
		if shapes[i].Type == "group" {
			ctx.resolveConnectors(shapes[i].Children)
		}
	}
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
