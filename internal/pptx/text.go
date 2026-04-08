package pptx

import "strings"

// extractTitle は子要素からタイトルテキストを取得する
func extractTitle(children []xmlSpTreeChild) string {
	for _, child := range children {
		if child.Sp == nil {
			continue
		}
		ph := child.Sp.NvSpPr.NvPr.Ph
		if ph == nil {
			continue
		}
		if ph.Type == "title" || ph.Type == "ctrTitle" {
			return extractTextFromTxBody(child.Sp.TxBody)
		}
	}
	return ""
}

// extractTextFromTxBody は txBody から全テキストを結合して返す
func extractTextFromTxBody(txBody *xmlTxBody) string {
	if txBody == nil {
		return ""
	}
	var parts []string
	for _, p := range txBody.Ps {
		text := extractParagraphText(p)
		if text != "" {
			parts = append(parts, text)
		}
	}
	return strings.Join(parts, " ")
}

// extractParagraphText は段落からプレーンテキストを結合する。
// a:br は改行、テキストランとフィールドはXML出現順に結合する。
func extractParagraphText(p xmlP) string {
	var sb strings.Builder
	for _, elem := range p.Elements {
		switch {
		case elem.R != nil:
			sb.WriteString(elem.R.T)
		case elem.Br:
			sb.WriteByte('\n')
		case elem.Fld != nil:
			sb.WriteString(elem.Fld.T)
		}
	}
	return sb.String()
}
