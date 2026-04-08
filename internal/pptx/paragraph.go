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
	resetAutoNum := func() {
		autoNumCounters = make(map[int]int)
		lastAutoNumLevel = -1
	}

	for _, p := range ps {
		text := extractParagraphText(p)
		if text == "" {
			resetAutoNum()
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
		buChar, buAutoNum, _ := resolveBullet(p.PPr, level, inherited)
		if buChar != nil {
			para.Bullet = buChar.Char
			resetAutoNum()
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
		} else {
			resetAutoNum()
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
		font := ctx.applyInheritedFont(nil, nil, level, inherited)
		if font != nil && isEmptyFont(font) {
			font = nil
		}
		return font, nil, nil
	}

	// 単一ランの場合はフォント情報+リンクのみ
	if len(runs) == 1 && runs[0].R != nil {
		font := ctx.rprToFont(runs[0].R.RPr)
		font = ctx.applyInheritedFont(font, runs[0].R.RPr, level, inherited)
		if font != nil && isEmptyFont(font) {
			font = nil
		}
		link := ctx.resolveRunHyperlink(runs[0].R.RPr)
		return font, link, nil
	}

	// 1パスで RichTextRun を構築しつつ均一性を判定する
	richText := make([]RichTextRun, 0, len(runs))
	hasBr := false
	allSame := true
	var firstFont *FontStyle
	var firstLink *HyperlinkData
	firstSet := false

	for _, elem := range runs {
		if elem.Br {
			hasBr = true
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
		font := ctx.rprToFont(rpr)
		font = ctx.applyInheritedFont(font, rpr, level, inherited)
		link := ctx.resolveRunHyperlink(rpr)

		if !firstSet {
			firstFont = font
			firstLink = link
			firstSet = true
		} else if allSame && (!fontsEqual(firstFont, font) || !linksEqual(firstLink, link)) {
			allSame = false
		}

		if text != "" {
			rt := RichTextRun{Text: text, Link: link}
			if font != nil && !isEmptyFont(font) {
				rt.Font = font
			}
			richText = append(richText, rt)
		}
	}

	// すべて同じ書式・リンクで改行なしなら統一フォント+リンクのみ
	if allSame && !hasBr {
		if firstFont != nil && isEmptyFont(firstFont) {
			firstFont = nil
		}
		return firstFont, firstLink, nil
	}

	return nil, nil, richText
}

// collectTextRuns は段落からテキストに関係する要素（r, br, fld）を出現順で返す。
func collectTextRuns(p xmlP) []xmlParagraphElement {
	var result []xmlParagraphElement
	for _, elem := range p.Elements {
		if elem.R != nil || elem.Br || elem.Fld != nil {
			result = append(result, elem)
		}
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

// findInheritedDefRPr は継承チェーンから条件に合う最初の defRPr を返す
func findInheritedDefRPr(inherited *inheritedStyle, level int, pred func(*xmlRPr) bool) *xmlRPr {
	for _, ls := range inherited.lstStyles {
		if ppr := ls.GetLevel(level); ppr != nil && ppr.DefRPr != nil && pred(ppr.DefRPr) {
			return ppr.DefRPr
		}
	}
	return nil
}

// applyInheritedFont は継承チェーンからフォント情報を補完する。
// font が nil の場合は新たに作成する。inherited が nil の場合は何もしない。
// rpr は元のランプロパティ（明示指定の有無を判定するため）。nil の場合はすべて未指定として扱う。
// 各フォントフィールドごとに継承チェーンを個別に辿り、空の defRPr による遮断を防ぐ。
func (ctx *parseContext) applyInheritedFont(font *FontStyle, rpr *xmlRPr, level int, inherited *inheritedStyle) *FontStyle {
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
		if drp := findInheritedDefRPr(inherited, level, func(r *xmlRPr) bool {
			return (r.Latin != nil && r.Latin.Typeface != "") || (r.Ea != nil && r.Ea.Typeface != "")
		}); drp != nil {
			if drp.Latin != nil && drp.Latin.Typeface != "" {
				font.Name = tc.ResolveThemeFont(drp.Latin.Typeface)
			} else {
				font.Name = tc.ResolveThemeFont(drp.Ea.Typeface)
			}
			modified = true
		}
	}

	// サイズ: 継承チェーン上で最初に見つかったサイズを使用
	if font.Size == 0 {
		if drp := findInheritedDefRPr(inherited, level, func(r *xmlRPr) bool { return r.Sz > 0 }); drp != nil {
			font.Size = int64(drp.Sz) * 127
			modified = true
		}
	}

	// 色: 継承チェーン上で最初に見つかった色を使用
	if font.Color == "" {
		if drp := findInheritedDefRPr(inherited, level, func(r *xmlRPr) bool { return r.SolidFill != nil }); drp != nil {
			if color := ctx.resolveSolidFillColor(drp.SolidFill); color != "" {
				font.Color = color
				modified = true
			}
		}
	}

	// 太字: rPr で明示指定されていない場合（B == ""）のみ継承
	if !font.Bold && (rpr == nil || rpr.B == "") {
		if drp := findInheritedDefRPr(inherited, level, func(r *xmlRPr) bool { return r.B != "" }); drp != nil {
			if drp.B == "1" || drp.B == "true" {
				font.Bold = true
				modified = true
			}
		}
	}

	// 斜体: rPr で明示指定されていない場合のみ継承
	if !font.Italic && (rpr == nil || rpr.I == "") {
		if drp := findInheritedDefRPr(inherited, level, func(r *xmlRPr) bool { return r.I != "" }); drp != nil {
			if drp.I == "1" || drp.I == "true" {
				font.Italic = true
				modified = true
			}
		}
	}

	// 下線: rPr で明示指定されていない場合のみ継承
	if font.Underline == "" && (rpr == nil || rpr.U == "") {
		if drp := findInheritedDefRPr(inherited, level, func(r *xmlRPr) bool { return r.U != "" && r.U != "none" }); drp != nil {
			font.Underline = drp.U
			modified = true
		}
	}

	// 取り消し線: rPr で明示指定されていない場合のみ継承
	if !font.Strikethrough && (rpr == nil || rpr.Strike == "") {
		if drp := findInheritedDefRPr(inherited, level, func(r *xmlRPr) bool { return r.Strike != "" && r.Strike != "noStrike" }); drp != nil {
			font.Strikethrough = true
			modified = true
		}
	}

	if !modified && isEmptyFont(font) {
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
