package pptx

import (
	"encoding/xml"
	"fmt"
	"sort"
	"strconv"
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

	// スライドのリレーション（画像用）
	slideRels := loadRels(f, slideRelsPath(entry.Path))

	sd := &SlideData{
		Number: slideNum,
		Title:  extractTitle(sld.CSld.SpTree),
	}

	// 図形をパース
	ctx := &parseContext{
		f:          f,
		slideRels:  slideRels,
		slidePath:  entry.Path,
		extractDir: extractDir,
		pptxIDMap:  make(map[int]int),
	}

	sd.Shapes = ctx.parseSpTree(sld.CSld.SpTree)

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

// parseSpTree は spTree 内の全要素をパースする
func (ctx *parseContext) parseSpTree(spTree xmlSpTree) []Shape {
	items := make([]shapeItem, 0)
	order := 0

	// 通常の図形
	for _, sp := range spTree.Shapes {
		s := ctx.parseSp(sp)
		if s == nil {
			continue
		}
		ph := sp.NvSpPr.NvPr.Ph
		isPH := ph != nil
		priority := phPriority(ph)
		items = append(items, shapeItem{order: order, shape: *s, isPH: isPH, phPriority: priority})
		order++
	}

	// コネクタ
	for _, cxn := range spTree.Connectors {
		s := ctx.parseCxnSp(cxn)
		if s == nil {
			continue
		}
		items = append(items, shapeItem{order: order, shape: *s})
		order++
	}

	// 画像
	for _, pic := range spTree.Pictures {
		s := ctx.parsePic(pic)
		if s == nil {
			continue
		}
		items = append(items, shapeItem{order: order, shape: *s})
		order++
	}

	// グループ
	for _, grp := range spTree.GroupShapes {
		s := ctx.parseGrpSp(grp)
		if s == nil {
			continue
		}
		items = append(items, shapeItem{order: order, shape: *s})
		order++
	}

	// テーブル（graphicFrame）
	for _, gf := range spTree.GraphicFrames {
		s := ctx.parseGraphicFrame(gf)
		if s == nil {
			continue
		}
		items = append(items, shapeItem{order: order, shape: *s})
		order++
	}

	// ソート: プレースホルダー（優先度順）→ 非プレースホルダー（出現順）
	sortShapeItems(items)

	// z-order とIDを割り当て
	shapes := make([]Shape, 0, len(items))
	for _, item := range items {
		item.shape.Z = ctx.allocZ()
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

	// テキストが空でプレースホルダーでもない図形はスキップ
	hasText := hasTextContent(sp.TxBody)
	hasFill := sp.SpPr.SolidFill != nil
	hasLine := sp.SpPr.Ln != nil && sp.SpPr.Ln.NoFill == nil
	if !hasText && ph == nil && !hasFill && !hasLine {
		return nil
	}
	// 空のプレースホルダーもスキップ
	if !hasText && ph != nil && !hasFill && !hasLine {
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

	// 子要素のパース（サブコンテキストで z を0からリセット）
	childCtx := &parseContext{
		f:          ctx.f,
		slideRels:  ctx.slideRels,
		slidePath:  ctx.slidePath,
		extractDir: ctx.extractDir,
		nextID:     ctx.nextID,
		pptxIDMap:  ctx.pptxIDMap,
		imageCount: ctx.imageCount,
	}

	childTree := xmlSpTree{
		Shapes:        grp.Shapes,
		GroupShapes:   grp.GroupShapes,
		Connectors:    grp.Connectors,
		Pictures:      grp.Pictures,
		GraphicFrames: grp.GraphicFrames,
	}
	s.Children = childCtx.parseSpTree(childTree)

	// カウンタを同期
	ctx.nextID = childCtx.nextID
	ctx.imageCount = childCtx.imageCount

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

	// テーブルデータ（結合セルはスキップ）
	cols := len(tbl.TblGrid.GridCols)
	var rows [][]string
	for _, tr := range tbl.Trs {
		row := make([]string, 0, cols)
		for _, tc := range tr.Tcs {
			if tc.VMerge == "1" || tc.HMerge == "1" {
				continue // 結合で吸収されたセルをスキップ
			}
			text := extractTextFromTxBody(tc.TxBody)
			row = append(row, text)
		}
		rows = append(rows, row)
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

// ---------- ヘルパー ----------

func xfrmToPosition(xfrm *xmlXfrm) *Position {
	if xfrm == nil {
		return nil
	}
	return &Position{
		X:  xfrm.Off.X,
		Y:  xfrm.Off.Y,
		Cx: xfrm.Ext.Cx,
		Cy: xfrm.Ext.Cy,
	}
}

func xfrmFlip(xfrm *xmlXfrm) string {
	if xfrm == nil {
		return ""
	}
	h := xfrm.FlipH
	v := xfrm.FlipV
	if h && v {
		return "hv"
	}
	if h {
		return "h"
	}
	if v {
		return "v"
	}
	return ""
}

func hasTextContent(txBody *xmlTxBody) bool {
	if txBody == nil {
		return false
	}
	for _, p := range txBody.Ps {
		if extractParagraphText(p) != "" {
			return true
		}
	}
	return false
}

// calloutDefaults は吹き出し図形のデフォルトadj値（OOXML仕様準拠）
// adj1=水平方向オフセット, adj2=垂直方向オフセット（100000分率）
var calloutDefaults = map[string][2]int64{
	"wedgeRectCallout":      {-20833, 62500},
	"wedgeRoundRectCallout": {-20833, 62500},
	"wedgeEllipseCallout":   {-20833, 62500},
	"cloudCallout":          {-20833, 62500},
	"borderCallout1":        {18750, -8333},
	"borderCallout2":        {18750, -8333},
	"borderCallout3":        {18750, -8333},
}

// resolveCalloutPointer は吹き出し図形のポインタ位置を計算する
func resolveCalloutPointer(geom *xmlPrstGeom, pos *Position) *Point {
	if geom == nil || pos == nil {
		return nil
	}

	defaults, isCallout := calloutDefaults[geom.Prst]
	if !isCallout {
		return nil
	}

	a1, a2 := defaults[0], defaults[1]

	if geom.AvLst != nil {
		for _, gd := range geom.AvLst.Gd {
			val := parseGdVal(gd.Fmla)
			if val == nil {
				continue
			}
			switch gd.Name {
			case "adj1":
				a1 = *val
			case "adj2":
				a2 = *val
			}
		}
	}

	px := pos.X + pos.Cx/2 + a1*pos.Cx/100000
	py := pos.Y + pos.Cy/2 + a2*pos.Cy/100000

	return &Point{X: px, Y: py}
}

// parseGdVal は "val N" 形式の数式から値を取得する
func parseGdVal(fmla string) *int64 {
	if !strings.HasPrefix(fmla, "val ") {
		return nil
	}
	v, err := strconv.ParseInt(strings.TrimPrefix(fmla, "val "), 10, 64)
	if err != nil {
		return nil
	}
	return &v
}

// resolveArrow はコネクタの矢印情報を解決する
func resolveArrow(ln *xmlLn) string {
	if ln == nil {
		return ""
	}
	result := resolveArrowType(ln.HeadEnd, ln.TailEnd)
	if result == "none" {
		return ""
	}
	return result
}

// loadNotesParagraphs はスライドのノートの段落を取得する
// ノートの読み込み・パース失敗時はnilを返す（スライド処理は継続する）
func (f *File) loadNotesParagraphs(slideIdx int) []Paragraph {
	notesPath := f.notesPath(slideIdx)
	if notesPath == "" {
		return nil
	}

	data, err := readZipFile(f.zi, notesPath)
	if err != nil || data == nil {
		return nil
	}

	var notes xmlNotes
	if err := xml.Unmarshal(data, &notes); err != nil {
		return nil
	}

	for _, sp := range notes.CSld.SpTree.Shapes {
		ph := sp.NvSpPr.NvPr.Ph
		if ph == nil || ph.Type != "body" {
			continue
		}
		if sp.TxBody == nil {
			continue
		}
		ctx := newTextOnlyContext(f)
		paras := ctx.parseParagraphs(sp.TxBody.Ps)
		if len(paras) > 0 {
			return paras
		}
	}
	return nil
}
