package pptx

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
	modified := false

	// 各プロパティの解決状態を追跡
	needName := font.Name == ""
	needSize := font.Size == 0
	needColor := font.Color == ""
	needBold := !font.Bold && (rpr == nil || rpr.B == "")
	needItalic := !font.Italic && (rpr == nil || rpr.I == "")
	needUnderline := font.Underline == "" && (rpr == nil || rpr.U == "")
	needStrike := !font.Strikethrough && (rpr == nil || rpr.Strike == "")

	for _, ls := range inherited.lstStyles {
		if !needName && !needSize && !needColor && !needBold && !needItalic && !needUnderline && !needStrike {
			break
		}

		ppr := ls.GetLevel(level)
		if ppr == nil || ppr.DefRPr == nil {
			continue
		}
		drp := ppr.DefRPr

		if needName && ((drp.Latin != nil && drp.Latin.Typeface != "") || (drp.Ea != nil && drp.Ea.Typeface != "")) {
			if drp.Latin != nil && drp.Latin.Typeface != "" {
				font.Name = tc.ResolveThemeFont(drp.Latin.Typeface)
			} else {
				font.Name = tc.ResolveThemeFont(drp.Ea.Typeface)
			}
			needName = false
			modified = true
		}

		if needSize && drp.Sz > 0 {
			font.Size = float64(drp.Sz) / 100
			needSize = false
			modified = true
		}

		if needColor && drp.SolidFill != nil {
			if color := ctx.resolveSolidFillColor(drp.SolidFill); color != "" {
				font.Color = color
				needColor = false
				modified = true
			}
		}

		// 太字: B が明示指定（"0"含む）されていれば、値に関わらず探索を停止
		if needBold && drp.B != "" {
			if drp.B == "1" || drp.B == "true" {
				font.Bold = true
				modified = true
			}
			needBold = false
		}

		// 斜体: 太字と同様
		if needItalic && drp.I != "" {
			if drp.I == "1" || drp.I == "true" {
				font.Italic = true
				modified = true
			}
			needItalic = false
		}

		// 下線: "none" は未指定扱いで探索を継続
		if needUnderline && drp.U != "" && drp.U != "none" {
			font.Underline = drp.U
			needUnderline = false
			modified = true
		}

		// 取り消し線: "noStrike" は未指定扱いで探索を継続
		if needStrike && drp.Strike != "" && drp.Strike != "noStrike" {
			font.Strikethrough = true
			needStrike = false
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

	return f
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
		tm.Left = emuToPtPtr(l)
	}
	if r != nil {
		tm.Right = emuToPtPtr(r)
	}
	if t != nil {
		tm.Top = emuToPtPtr(t)
	}
	if b != nil {
		tm.Bottom = emuToPtPtr(b)
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

