package pptx

import (
	"fmt"
	"strings"
)

// parseParagraphs は段落の配列をパースする。
// inherited はプレースホルダーの継承スタイル（非プレースホルダーの場合はnil）。
func (ctx *parseContext) parseParagraphs(ps []xmlP, inherited *inheritedStyle) []Paragraph {
	var result []Paragraph

	// 自動番号のカウンタ管理
	autoNumCounters := make(map[int]int) // level → カウンタ
	lastAutoNumLevel := -1

	for _, p := range ps {
		text := extractParagraphText(p)
		if text == "" {
			// 空の段落で自動番号をリセット
			autoNumCounters = make(map[int]int)
			lastAutoNumLevel = -1
			continue
		}

		para := Paragraph{Text: text}

		// レベル
		level := 0
		if p.PPr != nil {
			level = p.PPr.Lvl
		}
		if level > 0 {
			para.Level = level
		}

		// 箇条書き（スライド上で明示指定されているか、継承から取得）
		buChar, buAutoNum, buNone := resolveBullet(p.PPr, level, inherited)
		if buChar != nil {
			para.Bullet = buChar.Char
			autoNumCounters = make(map[int]int)
			lastAutoNumLevel = -1
		} else if buAutoNum != nil {
			// レベルが変わるか、前が自動番号でなければカウンタリセット
			if level != lastAutoNumLevel {
				autoNumCounters[level] = 0
			}
			startAt := buAutoNum.StartAt
			if startAt == 0 {
				startAt = 1
			}
			autoNumCounters[level]++
			num := startAt + autoNumCounters[level] - 1
			para.Bullet = formatAutoNum(buAutoNum.Type, num)
			lastAutoNumLevel = level
		} else if buNone {
			// 明示的に箇条書きなし → リセット
			autoNumCounters = make(map[int]int)
			lastAutoNumLevel = -1
		} else {
			// 箇条書き指定なし（PPr なし含む）→ 自動番号をリセット
			autoNumCounters = make(map[int]int)
			lastAutoNumLevel = -1
		}

		// 配置
		if p.PPr != nil && p.PPr.Algn != "" {
			para.Alignment = &Alignment{Horizontal: mapAlignment(p.PPr.Algn)}
		}

		// フォント情報・リッチテキスト
		para.Font, para.RichText = ctx.parseRunStyles(p, level, inherited)

		result = append(result, para)
	}

	return result
}

// resolveBullet は段落の箇条書きプロパティを解決する。
// スライド上で明示的に指定されていれば優先し、なければ継承チェーンから取得する。
func resolveBullet(ppr *xmlPPr, level int, inherited *inheritedStyle) (buChar *xmlBuChar, buAutoNum *xmlBuAutoNum, buNone bool) {
	// スライド上の段落プロパティを優先
	if ppr != nil {
		if ppr.BuChar != nil {
			return ppr.BuChar, nil, false
		}
		if ppr.BuAutoNum != nil {
			return nil, ppr.BuAutoNum, false
		}
		if ppr.BuNone != nil {
			return nil, nil, true
		}
	}

	// 継承チェーンから取得（空の lvlPPr をスキップして辿る）
	if inherited != nil {
		for _, ls := range inherited.lstStyles {
			if ppr := ls.GetLevel(level); ppr != nil {
				if ppr.BuChar != nil {
					return ppr.BuChar, nil, false
				}
				if ppr.BuAutoNum != nil {
					return nil, ppr.BuAutoNum, false
				}
				if ppr.BuNone != nil {
					return nil, nil, true
				}
			}
		}
	}

	return nil, nil, false
}

// parseRunStyles は段落のランからフォント情報とリッチテキストを抽出する。
// level は段落のインデントレベル、inherited は継承スタイル（フォント補完用）。
func (ctx *parseContext) parseRunStyles(p xmlP, level int, inherited *inheritedStyle) (*FontStyle, []RichTextRun) {
	runs := p.Rs
	if len(runs) == 0 {
		// ランがなくても継承からフォント情報を取得できる場合がある
		font := ctx.applyInheritedFont(nil, level, inherited)
		if font != nil && isEmptyFont(font) {
			font = nil
		}
		return font, nil
	}

	// 単一ランの場合はフォント情報のみ
	if len(runs) == 1 {
		font := ctx.rprToFont(runs[0].RPr)
		font = ctx.applyInheritedFont(font, level, inherited)
		if font != nil && isEmptyFont(font) {
			font = nil
		}
		return font, nil
	}

	// 複数ランの場合: すべて同じ書式ならフォント情報のみ
	allSame := true
	firstFont := ctx.rprToFont(runs[0].RPr)
	firstFont = ctx.applyInheritedFont(firstFont, level, inherited)
	for i := 1; i < len(runs); i++ {
		f := ctx.rprToFont(runs[i].RPr)
		f = ctx.applyInheritedFont(f, level, inherited)
		if !fontsEqual(firstFont, f) {
			allSame = false
			break
		}
	}

	if allSame {
		if firstFont != nil && isEmptyFont(firstFont) {
			firstFont = nil
		}
		return firstFont, nil
	}

	// 書式が異なるランがある場合: リッチテキスト
	richText := make([]RichTextRun, 0, len(runs))
	for _, r := range runs {
		if r.T == "" {
			continue
		}
		rt := RichTextRun{Text: r.T}
		font := ctx.rprToFont(r.RPr)
		font = ctx.applyInheritedFont(font, level, inherited)
		if font != nil && !isEmptyFont(font) {
			rt.Font = font
		}
		richText = append(richText, rt)
	}

	return nil, richText
}

// applyInheritedFont は継承チェーンからフォント情報を補完する。
// font が nil の場合は新たに作成する。inherited が nil の場合は何もしない。
// 各フォントフィールドごとに継承チェーンを個別に辿り、空の defRPr による遮断を防ぐ。
func (ctx *parseContext) applyInheritedFont(font *FontStyle, level int, inherited *inheritedStyle) *FontStyle {
	if inherited == nil {
		return font
	}

	if font == nil {
		font = &FontStyle{}
	}

	tc := ctx.f.getTheme()
	modified := false

	// フォント名: 継承チェーン上で最初に見つかったフォント名を使用
	if font.Name == "" {
		for _, ls := range inherited.lstStyles {
			if ppr := ls.GetLevel(level); ppr != nil && ppr.DefRPr != nil {
				if ppr.DefRPr.Latin != nil && ppr.DefRPr.Latin.Typeface != "" {
					font.Name = tc.ResolveThemeFont(ppr.DefRPr.Latin.Typeface)
					modified = true
					break
				}
				if ppr.DefRPr.Ea != nil && ppr.DefRPr.Ea.Typeface != "" {
					font.Name = tc.ResolveThemeFont(ppr.DefRPr.Ea.Typeface)
					modified = true
					break
				}
			}
		}
	}

	// サイズ: 継承チェーン上で最初に見つかったサイズを使用
	if font.Size == 0 {
		for _, ls := range inherited.lstStyles {
			if ppr := ls.GetLevel(level); ppr != nil && ppr.DefRPr != nil && ppr.DefRPr.Sz > 0 {
				font.Size = int64(ppr.DefRPr.Sz) * 127
				modified = true
				break
			}
		}
	}

	// 色: 継承チェーン上で最初に見つかった色を使用
	if font.Color == "" {
		for _, ls := range inherited.lstStyles {
			if ppr := ls.GetLevel(level); ppr != nil && ppr.DefRPr != nil && ppr.DefRPr.SolidFill != nil {
				color := ctx.resolveSolidFillColor(ppr.DefRPr.SolidFill)
				if color != "" {
					font.Color = color
					modified = true
					break
				}
			}
		}
	}

	// 何も補完されなかった場合は元の状態を維持
	if !modified && font.Name == "" && font.Size == 0 && !font.Bold && !font.Italic && font.Underline == "" && !font.Strikethrough && font.Color == "" {
		return nil
	}

	return font
}

// rprToFont は rPr からフォント情報を抽出する
func (ctx *parseContext) rprToFont(rpr *xmlRPr) *FontStyle {
	if rpr == nil {
		return nil
	}

	f := &FontStyle{}

	// フォント名（テーマフォント参照を解決）
	tc := ctx.f.getTheme()
	if rpr.Latin != nil && rpr.Latin.Typeface != "" {
		f.Name = tc.ResolveThemeFont(rpr.Latin.Typeface)
	} else if rpr.Ea != nil && rpr.Ea.Typeface != "" {
		f.Name = tc.ResolveThemeFont(rpr.Ea.Typeface)
	}

	// サイズ（hundredths of point → EMU: ×127）
	if rpr.Sz > 0 {
		f.Size = int64(rpr.Sz) * 127
	}

	// 太字
	if rpr.B == "1" || rpr.B == "true" {
		f.Bold = true
	}

	// 斜体
	if rpr.I == "1" || rpr.I == "true" {
		f.Italic = true
	}

	// 下線
	if rpr.U != "" && rpr.U != "none" {
		f.Underline = rpr.U
	}

	// 取り消し線
	if rpr.Strike != "" && rpr.Strike != "noStrike" {
		f.Strikethrough = true
	}

	// 色
	if rpr.SolidFill != nil {
		f.Color = ctx.resolveSolidFillColor(rpr.SolidFill)
	}

	return f
}

func isEmptyFont(f *FontStyle) bool {
	return f.Name == "" && f.Size == 0 && !f.Bold && !f.Italic &&
		f.Underline == "" && !f.Strikethrough && f.Color == ""
}

func fontsEqual(a, b *FontStyle) bool {
	if a == nil && b == nil {
		return true
	}
	if a == nil || b == nil {
		return false
	}
	return a.Name == b.Name && a.Size == b.Size && a.Bold == b.Bold &&
		a.Italic == b.Italic && a.Underline == b.Underline &&
		a.Strikethrough == b.Strikethrough && a.Color == b.Color
}

// formatAutoNum は自動番号を書式化する
func formatAutoNum(numType string, num int) string {
	switch numType {
	case "arabicPeriod":
		return fmt.Sprintf("%d.", num)
	case "arabicParenR":
		return fmt.Sprintf("%d)", num)
	case "alphaLcPeriod":
		return fmt.Sprintf("%s.", toLowerAlpha(num))
	case "alphaUcPeriod":
		return fmt.Sprintf("%s.", toUpperAlpha(num))
	case "romanLcPeriod":
		return fmt.Sprintf("%s.", toLowerRoman(num))
	case "romanUcPeriod":
		return fmt.Sprintf("%s.", toUpperRoman(num))
	default:
		return fmt.Sprintf("%d.", num)
	}
}

func toLowerAlpha(n int) string {
	if n < 1 {
		return fmt.Sprintf("%d", n)
	}
	var buf [8]byte
	i := len(buf)
	for n > 0 {
		n--
		i--
		buf[i] = byte('a' + n%26)
		n /= 26
	}
	return string(buf[i:])
}

func toUpperAlpha(n int) string {
	return strings.ToUpper(toLowerAlpha(n))
}

func toLowerRoman(n int) string {
	return strings.ToLower(toUpperRoman(n))
}

func toUpperRoman(n int) string {
	vals := []int{1000, 900, 500, 400, 100, 90, 50, 40, 10, 9, 5, 4, 1}
	syms := []string{"M", "CM", "D", "CD", "C", "XC", "L", "XL", "X", "IX", "V", "IV", "I"}
	var sb strings.Builder
	for i, v := range vals {
		for n >= v {
			sb.WriteString(syms[i])
			n -= v
		}
	}
	return sb.String()
}

func mapAlignment(algn string) string {
	switch algn {
	case "l":
		return "left"
	case "r":
		return "right"
	case "ctr":
		return "center"
	case "just":
		return "justify"
	default:
		return algn
	}
}

// extractShapeLevelAlignment はテキストボディレベルの垂直配置を抽出する
func (ctx *parseContext) extractShapeLevelAlignment(txBody *xmlTxBody) *Alignment {
	if txBody.BodyPr.Anchor == "" {
		return nil
	}
	v := mapVerticalAnchor(txBody.BodyPr.Anchor)
	if v == "" {
		return nil
	}
	return &Alignment{Vertical: v}
}

func mapVerticalAnchor(anchor string) string {
	switch anchor {
	case "t":
		return "top"
	case "ctr":
		return "center"
	case "b":
		return "bottom"
	default:
		return ""
	}
}
