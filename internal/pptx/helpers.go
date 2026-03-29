package pptx

// xfrmToPosition は xfrm 要素を Position に変換する
func xfrmToPosition(xfrm *xmlXfrm) *Position {
	if xfrm == nil {
		return nil
	}
	return &Position{
		X: xfrm.Off.X,
		Y: xfrm.Off.Y,
		W: xfrm.Ext.Cx,
		H: xfrm.Ext.Cy,
	}
}

// xfrmFlip は xfrm の反転情報を文字列に変換する
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

// hasTextContent は txBody にテキストが含まれるかを判定する
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
