package pptx

import (
	"fmt"
	"strings"
)

// parseParagraphs は段落の配列をパースする
func (ctx *parseContext) parseParagraphs(ps []xmlP) []Paragraph {
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

		// 箇条書き
		if p.PPr != nil && p.PPr.BuChar != nil {
			para.Bullet = p.PPr.BuChar.Char
			autoNumCounters = make(map[int]int)
			lastAutoNumLevel = -1
		} else if p.PPr != nil && p.PPr.BuAutoNum != nil {
			an := p.PPr.BuAutoNum
			// レベルが変わるか、前が自動番号でなければカウンタリセット
			if level != lastAutoNumLevel {
				autoNumCounters[level] = 0
			}
			startAt := an.StartAt
			if startAt == 0 {
				startAt = 1
			}
			autoNumCounters[level]++
			num := startAt + autoNumCounters[level] - 1
			para.Bullet = formatAutoNum(an.Type, num)
			lastAutoNumLevel = level
		} else {
			// 箇条書き指定なし（BuNone・PPr なし含む）→ 自動番号をリセット
			autoNumCounters = make(map[int]int)
			lastAutoNumLevel = -1
		}

		// 配置
		if p.PPr != nil && p.PPr.Algn != "" {
			para.Alignment = &Alignment{Horizontal: mapAlignment(p.PPr.Algn)}
		}

		// フォント情報・リッチテキスト
		para.Font, para.RichText = ctx.parseRunStyles(p)

		result = append(result, para)
	}

	return result
}

// parseRunStyles は段落のランからフォント情報とリッチテキストを抽出する
func (ctx *parseContext) parseRunStyles(p xmlP) (*FontStyle, []RichTextRun) {
	runs := p.Rs
	if len(runs) == 0 {
		return nil, nil
	}

	// 単一ランの場合はフォント情報のみ
	if len(runs) == 1 {
		font := ctx.rprToFont(runs[0].RPr)
		if font != nil && isEmptyFont(font) {
			font = nil
		}
		return font, nil
	}

	// 複数ランの場合: すべて同じ書式ならフォント情報のみ
	allSame := true
	firstFont := ctx.rprToFont(runs[0].RPr)
	for i := 1; i < len(runs); i++ {
		f := ctx.rprToFont(runs[i].RPr)
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
		if font != nil && !isEmptyFont(font) {
			rt.Font = font
		}
		richText = append(richText, rt)
	}

	return nil, richText
}

// rprToFont は rPr からフォント情報を抽出する
func (ctx *parseContext) rprToFont(rpr *xmlRPr) *FontStyle {
	if rpr == nil {
		return nil
	}

	f := &FontStyle{}

	// フォント名
	if rpr.Latin != nil && rpr.Latin.Typeface != "" {
		f.Name = rpr.Latin.Typeface
	} else if rpr.Ea != nil && rpr.Ea.Typeface != "" {
		f.Name = rpr.Ea.Typeface
	}

	// サイズ（hundredths of point → point）
	if rpr.Sz > 0 {
		f.Size = float64(rpr.Sz) / 100.0
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
	if n < 1 || n > 26 {
		return fmt.Sprintf("%d", n)
	}
	return string(rune('a' + n - 1))
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
