package pptx

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

// hasArrowHead は矢印があるかを判定する
func hasArrowHead(le *xmlLineEnd) bool {
	if le == nil {
		return false
	}
	return le.Type != "" && le.Type != "none"
}

// resolveArrowType は headEnd/tailEnd から矢印種別を判定する
func resolveArrowType(head, tail *xmlLineEnd) string {
	hasHead := hasArrowHead(head)
	hasTail := hasArrowHead(tail)
	if hasHead && hasTail {
		return "both"
	}
	if hasHead {
		return "start"
	}
	if hasTail {
		return "end"
	}
	return "none"
}
