package pptx

import "strings"

// parseCxnSp はコネクタをパースする
func (ctx *parseContext) parseCxnSp(cxn xmlCxnSp) *Shape {
	if cxn.NvCxnSpPr.CNvPr.Hidden {
		return nil
	}

	s := &Shape{
		ID:   ctx.allocID(cxn.NvCxnSpPr.CNvPr.ID),
		Type: "connector",
		Name: cxn.NvCxnSpPr.CNvPr.Name,
	}

	// コネクタ形状
	if cxn.SpPr.PrstGeom != nil {
		s.ConnectorType = cxn.SpPr.PrstGeom.Prst
		// 調整値
		if cxn.SpPr.PrstGeom.AvLst != nil {
			s.Adj = parseAdjValues(cxn.SpPr.PrstGeom.AvLst)
		}
	}

	// 位置
	s.Pos = xfrmToPosition(cxn.SpPr.Xfrm)

	// フリップ（始点・終点の算出に必要）
	flip := xfrmFlip(cxn.SpPr.Xfrm)

	// 始点・終点座標
	if s.Pos != nil {
		s.Start, s.End = connectorEndpoints(s.Pos, flip)
	}

	// 枠線
	s.Line = ctx.resolveLine(cxn.SpPr.Ln)

	// 矢印
	s.Arrow = resolveArrow(cxn.SpPr.Ln)

	// 接続情報（PowerPoint ID。後で解決する）
	if cxn.NvCxnSpPr.CNvCxnSpPr.StCxn != nil {
		s.From = -cxn.NvCxnSpPr.CNvCxnSpPr.StCxn.ID // 負値で未解決マーク
		idx := cxn.NvCxnSpPr.CNvCxnSpPr.StCxn.Idx
		s.FromIdx = &idx
	}
	if cxn.NvCxnSpPr.CNvCxnSpPr.EndCxn != nil {
		s.To = -cxn.NvCxnSpPr.CNvCxnSpPr.EndCxn.ID
		idx := cxn.NvCxnSpPr.CNvCxnSpPr.EndCxn.Idx
		s.ToIdx = &idx
	}

	// テキスト
	if cxn.TxBody != nil {
		paras := ctx.parseParagraphs(cxn.TxBody.Ps, nil)
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

// connectorEndpoints は pos と flip からコネクタの始点・終点を算出する
func connectorEndpoints(pos *Position, flip string) (*Point, *Point) {
	x1, y1 := pos.X, pos.Y
	x2, y2 := pos.X+pos.W, pos.Y+pos.H
	switch flip {
	case "h":
		x1, x2 = x2, x1
	case "v":
		y1, y2 = y2, y1
	case "hv":
		x1, x2 = x2, x1
		y1, y2 = y2, y1
	}
	return &Point{X: x1, Y: y1}, &Point{X: x2, Y: y2}
}

// parseAdjValues は avLst から調整値を取得する
func parseAdjValues(avLst *xmlAvLst) map[string]int {
	if len(avLst.Gd) == 0 {
		return nil
	}
	adj := make(map[string]int, len(avLst.Gd))
	for _, gd := range avLst.Gd {
		if gd.Name != "" {
			if v := parseGdVal(gd.Fmla); v != nil {
				adj[gd.Name] = int(*v)
			}
		}
	}
	if len(adj) == 0 {
		return nil
	}
	return adj
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
					shapes[i].FromIdx = nil
				}
			}
			if shapes[i].To < 0 {
				pptxID := -shapes[i].To
				if resolved, ok := ctx.pptxIDMap[pptxID]; ok {
					shapes[i].To = resolved
				} else {
					shapes[i].To = 0
					shapes[i].ToIdx = nil
				}
			}
		}
		if shapes[i].Type == "group" {
			ctx.resolveConnectors(shapes[i].Children)
		}
	}
}
