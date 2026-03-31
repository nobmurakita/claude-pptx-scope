package pptx

// resolveSolidFillColor は solidFill から色文字列を返す
func (ctx *parseContext) resolveSolidFillColor(fill *xmlSolidFill) string {
	if fill == nil {
		return ""
	}

	if ctx.f == nil {
		return ""
	}

	if fill.SrgbClr != nil {
		color := normalizeHexColor(fill.SrgbClr.Val)
		return applyColorTransforms(color, fill.SrgbClr.Transforms)
	}

	if fill.SchemeClr != nil {
		tc := ctx.f.getTheme()
		if tc == nil {
			return ""
		}
		idx := schemeClrToThemeIndex(fill.SchemeClr.Val)
		if idx < 0 {
			return ""
		}
		base := tc.Get(idx)
		if base == "" {
			return ""
		}
		return applyColorTransforms(base, fill.SchemeClr.Transforms)
	}

	return ""
}

// applyColorTransforms は色変換をXML出現順に適用する
func applyColorTransforms(color string, transforms []colorTransform) string {
	if color == "" {
		return ""
	}
	for _, t := range transforms {
		switch t.Op {
		case "tint":
			color = applyTint(color, float64(t.Val)/100000.0)
		case "shade":
			color = applyTint(color, -(1.0-float64(t.Val)/100000.0))
		case "lumMod":
			color = applyLuminance(color, float64(t.Val)/100000.0, 0)
		case "lumOff":
			color = applyLuminance(color, 1.0, float64(t.Val)/100000.0)
		}
	}
	return color
}

// resolveGradFillColor は gradFill から代表色（最初のストップカラー）を返す
func (ctx *parseContext) resolveGradFillColor(fill *xmlGradFill) string {
	if fill == nil || len(fill.GsLst) == 0 {
		return ""
	}
	return ctx.resolveSolidFillColor(&fill.GsLst[0].SolidFill)
}

// resolveFillColor は図形の塗りつぶし色を解決する（solidFill 優先、なければ gradFill の代表色）
func (ctx *parseContext) resolveFillColor(spPr *xmlSpPr) string {
	if spPr.SolidFill != nil {
		return ctx.resolveSolidFillColor(spPr.SolidFill)
	}
	return ctx.resolveGradFillColor(spPr.GradFill)
}

// resolveLine は ln 要素から枠線情報を解決する
func (ctx *parseContext) resolveLine(ln *xmlLn) *LineStyle {
	if ln == nil {
		return nil
	}
	if ln.NoFill != nil {
		return nil
	}

	ls := &LineStyle{}

	// 色
	if ln.SolidFill != nil {
		ls.Color = ctx.resolveSolidFillColor(ln.SolidFill)
	}

	// スタイル
	if ln.PrstDash != nil {
		ls.Style = ln.PrstDash.Val
	} else if ln.SolidFill != nil {
		ls.Style = "solid"
	}

	// 幅（EMU）
	if ln.W > 0 {
		ls.Width = int64(ln.W)
	}

	if ls.Color == "" && ls.Style == "" && ls.Width == 0 {
		return nil
	}

	return ls
}

// schemeClrToThemeIndex はスキームカラー名をテーマインデックスに変換する
func schemeClrToThemeIndex(val string) int {
	switch val {
	case "bg1", "lt1":
		return 0
	case "tx1", "dk1":
		return 1
	case "bg2", "lt2":
		return 2
	case "tx2", "dk2":
		return 3
	case "accent1":
		return 4
	case "accent2":
		return 5
	case "accent3":
		return 6
	case "accent4":
		return 7
	case "accent5":
		return 8
	case "accent6":
		return 9
	case "hlink":
		return 10
	case "folHlink":
		return 11
	default:
		// "phClr" 等のプレースホルダーカラーは解決できない
		return -1
	}
}
