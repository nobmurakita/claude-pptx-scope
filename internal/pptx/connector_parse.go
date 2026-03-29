package pptx

import "strings"

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
