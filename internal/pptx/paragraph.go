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

		// 段落インデント（marL/indent）
		para.MarginL, para.Indent = resolveParaIndent(p.PPr, level, inherited)

		// 配置（デフォルトの左揃えは省略）
		if p.PPr != nil && p.PPr.Algn != "" && p.PPr.Algn != "l" {
			para.Alignment = &Alignment{Horizontal: mapAlignment(p.PPr.Algn)}
		}

		// フォント情報・リッチテキスト・ハイパーリンク
		para.Font, para.Link, para.RichText = ctx.parseRunStyles(p, level, inherited)

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

// parseRunStyles は段落のランからフォント情報・リッチテキスト・ハイパーリンクを抽出する。
// level は段落のインデントレベル、inherited は継承スタイル（フォント補完用）。
// 戻り値: font（段落統一フォント）, link（段落統一リンク）, richText（書式/リンクが異なる場合）
func (ctx *parseContext) parseRunStyles(p xmlP, level int, inherited *inheritedStyle) (*FontStyle, *HyperlinkData, []RichTextRun) {
	// テキストランのみを抽出（a:br はリッチテキストモードでのみ使用）
	runs := collectTextRuns(p)
	if len(runs) == 0 {
		// ランがなくても継承からフォント情報を取得できる場合がある
		font := ctx.applyInheritedFont(nil, level, inherited)
		if font != nil && isEmptyFont(font) {
			font = nil
		}
		return font, nil, nil
	}

	// 単一ランの場合はフォント情報+リンクのみ
	if len(runs) == 1 && runs[0].R != nil {
		font := ctx.rprToFont(runs[0].R.RPr)
		font = ctx.applyInheritedFont(font, level, inherited)
		if font != nil && isEmptyFont(font) {
			font = nil
		}
		link := ctx.resolveRunHyperlink(runs[0].R.RPr)
		return font, link, nil
	}

	// 複数要素の場合: すべて同じ書式・リンクで改行なしならフォント+リンク情報のみ
	hasBr := false
	allSame := true
	var firstFont *FontStyle
	var firstLink *HyperlinkData
	firstSet := false
	for _, elem := range runs {
		if elem.Br {
			hasBr = true
			continue
		}
		var rpr *xmlRPr
		if elem.R != nil {
			rpr = elem.R.RPr
		} else if elem.Fld != nil {
			rpr = elem.Fld.RPr
		}
		f := ctx.rprToFont(rpr)
		f = ctx.applyInheritedFont(f, level, inherited)
		l := ctx.resolveRunHyperlink(rpr)
		if !firstSet {
			firstFont = f
			firstLink = l
			firstSet = true
		} else if !fontsEqual(firstFont, f) || !linksEqual(firstLink, l) {
			allSame = false
			break
		}
	}

	if allSame && !hasBr {
		if firstFont != nil && isEmptyFont(firstFont) {
			firstFont = nil
		}
		return firstFont, firstLink, nil
	}

	// 書式・リンクが異なるか改行を含む場合: リッチテキスト
	richText := make([]RichTextRun, 0, len(runs))
	for _, elem := range runs {
		if elem.Br {
			richText = append(richText, RichTextRun{Text: "\n"})
			continue
		}
		var text string
		var rpr *xmlRPr
		if elem.R != nil {
			text = elem.R.T
			rpr = elem.R.RPr
		} else if elem.Fld != nil {
			text = elem.Fld.T
			rpr = elem.Fld.RPr
		}
		if text == "" {
			continue
		}
		rt := RichTextRun{Text: text}
		font := ctx.rprToFont(rpr)
		font = ctx.applyInheritedFont(font, level, inherited)
		if font != nil && !isEmptyFont(font) {
			rt.Font = font
		}
		rt.Link = ctx.resolveRunHyperlink(rpr)
		richText = append(richText, rt)
	}

	return nil, nil, richText
}

// collectTextRuns は段落からテキストに関係する要素（r, br, fld）を出現順で返す。
// Elements が空の場合は Rs から構築する（テスト互換）。
func collectTextRuns(p xmlP) []xmlParagraphElement {
	if len(p.Elements) > 0 {
		var result []xmlParagraphElement
		for _, elem := range p.Elements {
			if elem.R != nil || elem.Br || elem.Fld != nil {
				result = append(result, elem)
			}
		}
		return result
	}
	// フォールバック: Rs/Fld から構築
	result := make([]xmlParagraphElement, 0, len(p.Rs)+len(p.Fld))
	for i := range p.Rs {
		result = append(result, xmlParagraphElement{R: &p.Rs[i]})
	}
	for i := range p.Fld {
		result = append(result, xmlParagraphElement{Fld: &p.Fld[i]})
	}
	return result
}

// resolveRunHyperlink はランの rPr からハイパーリンクを解決する
func (ctx *parseContext) resolveRunHyperlink(rpr *xmlRPr) *HyperlinkData {
	if rpr == nil {
		return nil
	}
	return ctx.resolveHyperlink(rpr.HlinkClick)
}

// linksEqual は2つのハイパーリンクが等しいか判定する
func linksEqual(a, b *HyperlinkData) bool {
	if a == nil && b == nil {
		return true
	}
	if a == nil || b == nil {
		return false
	}
	return a.URL == b.URL && a.Slide == b.Slide
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

// resolveParaIndent は段落の marL と indent を解決する。
// スライド上の pPr を優先し、なければ継承チェーンから取得する。
func resolveParaIndent(ppr *xmlPPr, level int, inherited *inheritedStyle) (marL *int64, indent *int64) {
	// スライド上で明示指定されていれば優先
	if ppr != nil {
		marL = ppr.MarL
		indent = ppr.Indent
	}

	// 継承チェーンから補完
	if inherited != nil {
		for _, ls := range inherited.lstStyles {
			if lvlPPr := ls.GetLevel(level); lvlPPr != nil {
				if marL == nil && lvlPPr.MarL != nil {
					marL = lvlPPr.MarL
				}
				if indent == nil && lvlPPr.Indent != nil {
					indent = lvlPPr.Indent
				}
				if marL != nil && indent != nil {
					break
				}
			}
		}
	}

	// 値が0の場合は省略（デフォルト）
	if marL != nil && *marL == 0 {
		marL = nil
	}
	if indent != nil && *indent == 0 {
		indent = nil
	}

	return marL, indent
}

// extractTextMargin は bodyPr からテキストマージンを抽出する。
// OOXML のデフォルト値（91440, 91440, 45720, 45720）と同じ場合は省略する。
func extractTextMargin(bodyPr xmlBodyPr) *TextMargin {
	l := bodyPr.LIns
	r := bodyPr.RIns
	t := bodyPr.TIns
	b := bodyPr.BIns

	// すべて未指定ならデフォルト → 省略
	if l == nil && r == nil && t == nil && b == nil {
		return nil
	}

	// デフォルト値（EMU: left=91440, right=91440, top=45720, bottom=45720）
	isDefault := func(v *int64, def int64) bool {
		return v == nil || *v == def
	}
	if isDefault(l, 91440) && isDefault(r, 91440) && isDefault(t, 45720) && isDefault(b, 45720) {
		return nil
	}

	tm := &TextMargin{}
	if l != nil {
		tm.Left = l
	}
	if r != nil {
		tm.Right = r
	}
	if t != nil {
		tm.Top = t
	}
	if b != nil {
		tm.Bottom = b
	}
	return tm
}

// extractShapeLevelAlignment はテキストボディレベルの垂直配置を抽出する。
// デフォルトの上揃え（"t"）は省略する。
func (ctx *parseContext) extractShapeLevelAlignment(txBody *xmlTxBody) *Alignment {
	if txBody.BodyPr.Anchor == "" || txBody.BodyPr.Anchor == "t" {
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
