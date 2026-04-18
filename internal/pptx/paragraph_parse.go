package pptx

import (
	"strconv"
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

		// 段落インデント（marL/indent、EMU → pt）
		marL, indent := resolveParaIndent(p.PPr, level, inherited)
		para.MarginL = emuToPtPtr(marL)
		para.Indent = emuToPtPtr(indent)

		// 行間・段落前後スペース
		lnSpc, spcBef, spcAft := resolveParaSpacing(p.PPr, level, inherited)
		para.LineSpacing = formatSpacing(lnSpc)
		para.SpaceBefore = formatSpacing(spcBef)
		para.SpaceAfter = formatSpacing(spcAft)

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
		font := nilIfEmpty(ctx.applyInheritedFont(nil, nil, level, inherited))
		return font, nil, nil
	}

	// 単一ランの場合はフォント情報+リンクのみ
	if len(runs) == 1 && runs[0].R != nil {
		font := ctx.rprToFont(runs[0].R.RPr)
		font = nilIfEmpty(ctx.applyInheritedFont(font, runs[0].R.RPr, level, inherited))
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
		return nilIfEmpty(firstFont), firstLink, nil
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

// applyInheritedFont は継承チェーンからフォント情報を補完する。
// font が nil の場合は新たに作成する。inherited が nil の場合は何もしない。
// rpr は元のランプロパティ（明示指定の有無を判定するため）。nil の場合はすべて未指定として扱う。
// 継承チェーンを1回走査し、未解決のプロパティをまとめて収集する。
func (ctx *parseContext) applyInheritedFont(font *FontStyle, rpr *xmlRPr, level int, inherited *inheritedStyle) *FontStyle {
	if inherited == nil {
		return font
	}

	if font == nil {
		font = &FontStyle{}
	}

	tc := ctx.f.getTheme()
	resolvers := ctx.inheritedFontResolvers(font, rpr, tc)
	modified := false

	for _, ls := range inherited.lstStyles {
		if allResolved(resolvers) {
			break
		}
		ppr := ls.GetLevel(level)
		if ppr == nil || ppr.DefRPr == nil {
			continue
		}
		for i := range resolvers {
			if resolvers[i].resolved {
				continue
			}
			applied, done := resolvers[i].apply(ppr.DefRPr)
			if applied {
				modified = true
			}
			if done {
				resolvers[i].resolved = true
			}
		}
	}

	if !modified && isEmptyFont(font) {
		return nil
	}

	return font
}

// fontPropResolver は applyInheritedFont の1プロパティに対する解決ロジック。
// resolved=true になった時点でそのプロパティは探索を打ち切る。
// apply は (値を font に適用したか, 探索を打ち切ってよいか) を返す。
type fontPropResolver struct {
	resolved bool
	apply    func(drp *xmlRPr) (applied bool, done bool)
}

func allResolved(rs []fontPropResolver) bool {
	for _, r := range rs {
		if !r.resolved {
			return false
		}
	}
	return true
}

func (ctx *parseContext) inheritedFontResolvers(font *FontStyle, rpr *xmlRPr, tc *themeColors) []fontPropResolver {
	return []fontPropResolver{
		{
			resolved: font.Name != "",
			apply: func(drp *xmlRPr) (bool, bool) {
				if drp.Latin != nil && drp.Latin.Typeface != "" {
					font.Name = tc.ResolveThemeFont(drp.Latin.Typeface)
					return true, true
				}
				if drp.Ea != nil && drp.Ea.Typeface != "" {
					font.Name = tc.ResolveThemeFont(drp.Ea.Typeface)
					return true, true
				}
				return false, false
			},
		},
		{
			resolved: font.Size != 0,
			apply: func(drp *xmlRPr) (bool, bool) {
				if drp.Sz > 0 {
					font.Size = float64(drp.Sz) / 100
					return true, true
				}
				return false, false
			},
		},
		{
			resolved: font.Color != "",
			apply: func(drp *xmlRPr) (bool, bool) {
				if drp.SolidFill == nil {
					return false, false
				}
				if color := ctx.resolveSolidFillColor(drp.SolidFill); color != "" {
					font.Color = color
					return true, true
				}
				return false, false
			},
		},
		{
			resolved: font.Highlight != "",
			apply: func(drp *xmlRPr) (bool, bool) {
				if drp.Highlight == nil {
					return false, false
				}
				if color := ctx.resolveSolidFillColor(drp.Highlight); color != "" {
					font.Highlight = color
					return true, true
				}
				return false, false
			},
		},
		{
			resolved: font.Baseline != "" || (rpr != nil && rpr.Baseline != nil),
			apply: func(drp *xmlRPr) (bool, bool) {
				if drp.Baseline == nil {
					return false, false
				}
				if label := baselineLabel(drp.Baseline); label != "" {
					font.Baseline = label
					return true, true
				}
				return false, true
			},
		},
		{
			resolved: font.Cap != "" || (rpr != nil && rpr.Cap != ""),
			apply: func(drp *xmlRPr) (bool, bool) {
				if drp.Cap == "" {
					return false, false
				}
				if drp.Cap == "all" || drp.Cap == "small" {
					font.Cap = drp.Cap
					return true, true
				}
				return false, true
			},
		},
		{
			// 太字: B が明示指定（"0"含む）されていれば、値に関わらず探索を停止
			resolved: font.Bold || (rpr != nil && rpr.B != ""),
			apply: func(drp *xmlRPr) (bool, bool) {
				if drp.B == "" {
					return false, false
				}
				if drp.B == "1" || drp.B == "true" {
					font.Bold = true
					return true, true
				}
				return false, true
			},
		},
		{
			// 斜体: 太字と同様
			resolved: font.Italic || (rpr != nil && rpr.I != ""),
			apply: func(drp *xmlRPr) (bool, bool) {
				if drp.I == "" {
					return false, false
				}
				if drp.I == "1" || drp.I == "true" {
					font.Italic = true
					return true, true
				}
				return false, true
			},
		},
		{
			// 下線: "none" は未指定扱いで探索を継続
			resolved: font.Underline != "" || (rpr != nil && rpr.U != ""),
			apply: func(drp *xmlRPr) (bool, bool) {
				if drp.U == "" || drp.U == "none" {
					return false, false
				}
				font.Underline = drp.U
				return true, true
			},
		},
		{
			// 取り消し線: "noStrike" は未指定扱いで探索を継続
			resolved: font.Strikethrough || (rpr != nil && rpr.Strike != ""),
			apply: func(drp *xmlRPr) (bool, bool) {
				if drp.Strike == "" || drp.Strike == "noStrike" {
					return false, false
				}
				font.Strikethrough = true
				return true, true
			},
		},
	}
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

	// サイズ（hundredths of point → pt: ÷100）
	if rpr.Sz > 0 {
		f.Size = float64(rpr.Sz) / 100
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

	// 背景色（ハイライト）
	if rpr.Highlight != nil {
		f.Highlight = ctx.resolveSolidFillColor(rpr.Highlight)
	}

	// 上付き/下付き文字
	f.Baseline = baselineLabel(rpr.Baseline)

	// 英字大文字化（"none" は省略）
	if rpr.Cap == "all" || rpr.Cap == "small" {
		f.Cap = rpr.Cap
	}

	return f
}

// baselineLabel は baseline 属性値（パーセント×1000）を "super"/"sub" に変換する。
// 0 または nil の場合は空文字列。
func baselineLabel(b *int) string {
	if b == nil || *b == 0 {
		return ""
	}
	if *b > 0 {
		return "super"
	}
	return "sub"
}

// nilIfEmpty はフォントが空なら nil を返す
func nilIfEmpty(f *FontStyle) *FontStyle {
	if f == nil || isEmptyFont(f) {
		return nil
	}
	return f
}

func isEmptyFont(f *FontStyle) bool {
	return f.Name == "" && f.Size == 0 && !f.Bold && !f.Italic &&
		f.Underline == "" && !f.Strikethrough && f.Color == "" && f.Highlight == "" &&
		f.Baseline == "" && f.Cap == ""
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
		a.Strikethrough == b.Strikethrough && a.Color == b.Color &&
		a.Highlight == b.Highlight && a.Baseline == b.Baseline && a.Cap == b.Cap
}


// resolveParaSpacing は段落の行間・前スペース・後スペースを解決する。
// スライド上の pPr を優先し、なければ継承チェーンから取得する。
func resolveParaSpacing(ppr *xmlPPr, level int, inherited *inheritedStyle) (lnSpc, spcBef, spcAft *xmlSpacing) {
	if ppr != nil {
		lnSpc = ppr.LnSpc
		spcBef = ppr.SpcBef
		spcAft = ppr.SpcAft
	}

	if inherited != nil {
		for _, ls := range inherited.lstStyles {
			if lnSpc != nil && spcBef != nil && spcAft != nil {
				break
			}
			lvlPPr := ls.GetLevel(level)
			if lvlPPr == nil {
				continue
			}
			if lnSpc == nil && lvlPPr.LnSpc != nil {
				lnSpc = lvlPPr.LnSpc
			}
			if spcBef == nil && lvlPPr.SpcBef != nil {
				spcBef = lvlPPr.SpcBef
			}
			if spcAft == nil && lvlPPr.SpcAft != nil {
				spcAft = lvlPPr.SpcAft
			}
		}
	}

	return lnSpc, spcBef, spcAft
}

// formatSpacing は xmlSpacing を人間可読な文字列に変換する。
// a:spcPct は "150%"（パーセント×1000 → %）、a:spcPts は "6pt"（ポイント×100 → pt）。
// デフォルト値（100% / 0pt）は空文字列を返してノイズを省く。
func formatSpacing(s *xmlSpacing) string {
	if s == nil {
		return ""
	}
	if s.SpcPct != nil {
		if s.SpcPct.Val == 100000 { // 100% はデフォルト
			return ""
		}
		return strconv.FormatFloat(float64(s.SpcPct.Val)/1000.0, 'f', -1, 64) + "%"
	}
	if s.SpcPts != nil {
		if s.SpcPts.Val == 0 { // 0pt はデフォルト
			return ""
		}
		return strconv.FormatFloat(float64(s.SpcPts.Val)/100.0, 'f', -1, 64) + "pt"
	}
	return ""
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


