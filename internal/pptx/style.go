package pptx


// resolveSolidFillColor は solidFill から色文字列を返す
func (ctx *parseContext) resolveSolidFillColor(fill *xmlSolidFill) string {
	if fill == nil {
		return ""
	}

	tc := ctx.f.getTheme()

	if fill.SrgbClr != nil {
		color := normalizeHexColor(fill.SrgbClr.Val)
		color = applyColorTransforms(color, fill.SrgbClr.LumMod, fill.SrgbClr.LumOff, fill.SrgbClr.Tint, fill.SrgbClr.Shade)
		return color
	}

	if fill.SchemeClr != nil {
		idx := schemeClrToThemeIndex(fill.SchemeClr.Val)
		if idx < 0 {
			return ""
		}
		base := tc.Get(idx)
		if base == "" {
			return ""
		}
		base = applyColorTransforms(base, fill.SchemeClr.LumMod, fill.SchemeClr.LumOff, fill.SchemeClr.Tint, fill.SchemeClr.Shade)
		return base
	}

	return ""
}

// applyColorTransforms は色変換を適用する
func applyColorTransforms(color string, lumMod, lumOff, tint, shade *xmlPercentage) string {
	if color == "" {
		return ""
	}
	if tint != nil {
		t := float64(tint.Val) / 100000.0
		color = applyTint(color, t)
	}
	if shade != nil {
		t := -(1.0 - float64(shade.Val)/100000.0)
		color = applyTint(color, t)
	}
	if lumMod != nil || lumOff != nil {
		mod := 1.0
		off := 0.0
		if lumMod != nil {
			mod = float64(lumMod.Val) / 100000.0
		}
		if lumOff != nil {
			off = float64(lumOff.Val) / 100000.0
		}
		color = applyLuminance(color, mod, off)
	}
	return color
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

	// 幅（EMU → ポイント）
	if ln.W > 0 {
		ls.Width = float64(ln.W) / 12700.0
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

