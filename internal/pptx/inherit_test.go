package pptx

import (
	"testing"
)

func TestPhKeyMatching_ExactMatch(t *testing.T) {
	m := map[phKey]*placeholderDef{
		{Type: "body", Idx: "1"}: {xfrm: &xmlXfrm{Off: xmlOff{X: 100}}},
		{Type: "body"}:           {xfrm: &xmlXfrm{Off: xmlOff{X: 200}}},
	}

	// 完全一致
	def := findPlaceholder(m, phKey{Type: "body", Idx: "1"})
	if def == nil || def.xfrm.Off.X != 100 {
		t.Errorf("完全一致で正しいプレースホルダーが取得できない")
	}
}

func TestPhKeyMatching_TypeOnlyFallback(t *testing.T) {
	m := map[phKey]*placeholderDef{
		{Type: "body"}: {xfrm: &xmlXfrm{Off: xmlOff{X: 200}}},
	}

	// idx 不一致時は type のみでフォールバック
	def := findPlaceholder(m, phKey{Type: "body", Idx: "99"})
	if def == nil || def.xfrm.Off.X != 200 {
		t.Errorf("type のみのフォールバックが動作しない")
	}
}

func TestPhKeyMatching_NoMatch(t *testing.T) {
	m := map[phKey]*placeholderDef{
		{Type: "title"}: {xfrm: &xmlXfrm{Off: xmlOff{X: 100}}},
	}

	def := findPlaceholder(m, phKey{Type: "body"})
	if def != nil {
		t.Errorf("一致しないキーで nil が返されるべき")
	}
}

func TestPhKeyMatching_NilMap(t *testing.T) {
	def := findPlaceholder(nil, phKey{Type: "body"})
	if def != nil {
		t.Errorf("nil マップで nil が返されるべき")
	}
}

func TestResolveInheritedXfrm_SlideNil_LayoutHasValue(t *testing.T) {
	layout := &layoutData{
		placeholders: map[phKey]*placeholderDef{
			{Type: "title"}: {xfrm: &xmlXfrm{Off: xmlOff{X: 500, Y: 600}, Ext: xmlExt{Cx: 700, Cy: 800}}},
		},
	}
	master := &masterData{
		placeholders: map[phKey]*placeholderDef{
			{Type: "title"}: {xfrm: &xmlXfrm{Off: xmlOff{X: 100, Y: 200}, Ext: xmlExt{Cx: 300, Cy: 400}}},
		},
	}

	ph := &xmlPh{Type: "title"}
	is := resolveInheritedStyle(ph, nil, layout, master)

	if is == nil {
		t.Fatal("inheritedStyle が nil")
	}
	if is.xfrm == nil {
		t.Fatal("xfrm が継承されていない")
	}
	// レイアウトが優先
	if is.xfrm.Off.X != 500 {
		t.Errorf("xfrm.Off.X: got %d, want 500（レイアウトが優先）", is.xfrm.Off.X)
	}
}

func TestResolveInheritedXfrm_LayoutNil_MasterHasValue(t *testing.T) {
	layout := &layoutData{
		placeholders: make(map[phKey]*placeholderDef),
	}
	master := &masterData{
		placeholders: map[phKey]*placeholderDef{
			{Type: "title"}: {xfrm: &xmlXfrm{Off: xmlOff{X: 100}}},
		},
	}

	ph := &xmlPh{Type: "title"}
	is := resolveInheritedStyle(ph, nil, layout, master)

	if is == nil || is.xfrm == nil {
		t.Fatal("マスターから xfrm が継承されていない")
	}
	if is.xfrm.Off.X != 100 {
		t.Errorf("xfrm.Off.X: got %d, want 100", is.xfrm.Off.X)
	}
}

func TestResolveInheritedFont_FromDefRPr(t *testing.T) {
	ctx := newTestContext()

	inherited := &inheritedStyle{
		lstStyles: []*xmlLstStyle{
			{
				Lvl1pPr: &xmlLvlPPr{
					DefRPr: &xmlRPr{
						Sz:    2400,
						Latin: &xmlFont{Typeface: "メイリオ"},
						SolidFill: &xmlSolidFill{
							SrgbClr: &xmlSrgbClr{Val: "333333"},
						},
					},
				},
			},
		},
	}

	// nil のフォントに継承を適用
	font := ctx.applyInheritedFont(nil, 0, inherited)
	if font == nil {
		t.Fatal("継承フォントが返されない")
	}
	if font.Name != "メイリオ" {
		t.Errorf("Name: got %q, want %q", font.Name, "メイリオ")
	}
	if font.Size != 2400*127 {
		t.Errorf("Size: got %d, want %d", font.Size, 2400*127)
	}
	if font.Color != "#333333" {
		t.Errorf("Color: got %q, want %q", font.Color, "#333333")
	}
}

func TestResolveInheritedFont_NoOverrideExplicit(t *testing.T) {
	ctx := newTestContext()

	inherited := &inheritedStyle{
		lstStyles: []*xmlLstStyle{
			{
				Lvl1pPr: &xmlLvlPPr{
					DefRPr: &xmlRPr{
						Sz:    2400,
						Latin: &xmlFont{Typeface: "ゴシック"},
					},
				},
			},
		},
	}

	// 既にフォント名が設定されている場合は上書きしない
	font := &FontStyle{Name: "メイリオ", Size: 127000}
	result := ctx.applyInheritedFont(font, 0, inherited)
	if result.Name != "メイリオ" {
		t.Errorf("明示的な値が上書きされた: got %q, want %q", result.Name, "メイリオ")
	}
	if result.Size != 127000 {
		t.Errorf("明示的なサイズが上書きされた: got %d, want %d", result.Size, 127000)
	}
}

func TestResolveInheritedFont_Level2(t *testing.T) {
	ctx := newTestContext()

	inherited := &inheritedStyle{
		lstStyles: []*xmlLstStyle{
			{
				Lvl1pPr: &xmlLvlPPr{DefRPr: &xmlRPr{Sz: 2400}},
				Lvl3pPr: &xmlLvlPPr{DefRPr: &xmlRPr{Sz: 1600}},
			},
		},
	}

	// level=2 → lvl3pPr
	font := ctx.applyInheritedFont(nil, 2, inherited)
	if font == nil {
		t.Fatal("継承フォントが返されない")
	}
	if font.Size != 1600*127 {
		t.Errorf("Size: got %d, want %d (level 2 → lvl3pPr)", font.Size, 1600*127)
	}
}

func TestResolveInheritedBullet(t *testing.T) {
	inherited := &inheritedStyle{
		lstStyles: []*xmlLstStyle{
			{
				Lvl1pPr: &xmlLvlPPr{
					BuChar: &xmlBuChar{Char: "•"},
				},
			},
		},
	}

	buChar, buAutoNum, buNone := resolveBullet(nil, 0, inherited)
	if buChar == nil || buChar.Char != "•" {
		t.Errorf("継承から箇条書き文字が取得されるべき")
	}
	if buAutoNum != nil {
		t.Errorf("buAutoNum は nil であるべき")
	}
	if buNone {
		t.Errorf("buNone は false であるべき")
	}
}

func TestResolveInheritedBullet_ExplicitOverride(t *testing.T) {
	inherited := &inheritedStyle{
		lstStyles: []*xmlLstStyle{
			{
				Lvl1pPr: &xmlLvlPPr{
					BuChar: &xmlBuChar{Char: "•"},
				},
			},
		},
	}

	// 段落で BuNone を指定 → 継承を上書き
	ppr := &xmlPPr{BuNone: &struct{}{}}
	buChar, buAutoNum, buNone := resolveBullet(ppr, 0, inherited)
	if buChar != nil {
		t.Errorf("BuNone 指定時は buChar が nil であるべき")
	}
	if buAutoNum != nil {
		t.Errorf("BuNone 指定時は buAutoNum が nil であるべき")
	}
	if !buNone {
		t.Errorf("BuNone 指定時は buNone が true であるべき")
	}
}

func TestResolveInheritedStyle_NilPh(t *testing.T) {
	is := resolveInheritedStyle(nil, nil, nil, nil)
	if is != nil {
		t.Errorf("ph が nil の場合は nil を返すべき")
	}
}

func TestResolveInheritedStyle_NilLayoutMaster(t *testing.T) {
	ph := &xmlPh{Type: "title"}
	is := resolveInheritedStyle(ph, nil, nil, nil)
	if is == nil {
		t.Fatal("nil layout/master でも inheritedStyle は返されるべき")
	}
}

func TestMasterTxStyleForPh(t *testing.T) {
	titleStyle := &xmlLstStyle{Lvl1pPr: &xmlLvlPPr{DefRPr: &xmlRPr{Sz: 4400}}}
	bodyStyle := &xmlLstStyle{Lvl1pPr: &xmlLvlPPr{DefRPr: &xmlRPr{Sz: 2400}}}
	otherStyle := &xmlLstStyle{Lvl1pPr: &xmlLvlPPr{DefRPr: &xmlRPr{Sz: 1800}}}
	txStyles := &xmlTxStyles{
		TitleStyle: titleStyle,
		BodyStyle:  bodyStyle,
		OtherStyle: otherStyle,
	}

	tests := []struct {
		phType string
		want   *xmlLstStyle
	}{
		{"title", titleStyle},
		{"ctrTitle", titleStyle},
		{"body", bodyStyle},
		{"subTitle", bodyStyle},
		{"dt", otherStyle},
		{"ftr", otherStyle},
		{"sldNum", otherStyle},
	}
	for _, tt := range tests {
		got := masterTxStyleForPh(txStyles, tt.phType)
		if got != tt.want {
			t.Errorf("masterTxStyleForPh(%q): 不一致", tt.phType)
		}
	}
}

func TestLstStyleGetLevel(t *testing.T) {
	ls := &xmlLstStyle{
		Lvl1pPr: &xmlLvlPPr{Algn: "l"},
		Lvl5pPr: &xmlLvlPPr{Algn: "ctr"},
	}

	if ls.GetLevel(0) == nil || ls.GetLevel(0).Algn != "l" {
		t.Errorf("level 0 → Lvl1pPr")
	}
	if ls.GetLevel(4) == nil || ls.GetLevel(4).Algn != "ctr" {
		t.Errorf("level 4 → Lvl5pPr")
	}
	if ls.GetLevel(1) != nil {
		t.Errorf("level 1 → nil（未設定）")
	}
	if ls.GetLevel(9) != nil {
		t.Errorf("level 9 → nil（範囲外）")
	}
	if ls.GetLevel(-1) != nil {
		t.Errorf("level -1 → nil（負値）")
	}
}

func TestLstStyleGetLevel_Nil(t *testing.T) {
	var ls *xmlLstStyle
	if ls.GetLevel(0) != nil {
		t.Errorf("nil receiver で nil が返されるべき")
	}
}

func TestInheritedStyleGetLevelPPr_Cascade(t *testing.T) {
	// 最初の lstStyle には level 0 がない → 2番目から取得
	is := &inheritedStyle{
		lstStyles: []*xmlLstStyle{
			{Lvl2pPr: &xmlLvlPPr{Algn: "ctr"}},
			{Lvl1pPr: &xmlLvlPPr{Algn: "l"}},
		},
	}

	ppr := is.getLevelPPr(0)
	if ppr == nil || ppr.Algn != "l" {
		t.Errorf("カスケードで2番目の lstStyle から取得されるべき")
	}
}

func TestResolveThemeFont(t *testing.T) {
	tc := &themeColors{
		colors:    make(map[int]string),
		majorFont: "游ゴシック",
		minorFont: "游明朝",
	}

	if tc.ResolveThemeFont("+mj-lt") != "游ゴシック" {
		t.Errorf("+mj-lt → majorFont")
	}
	if tc.ResolveThemeFont("+mn-ea") != "游明朝" {
		t.Errorf("+mn-ea → minorFont")
	}
	if tc.ResolveThemeFont("Arial") != "Arial" {
		t.Errorf("通常のフォント名はそのまま返す")
	}

	// nil の場合
	var nilTC *themeColors
	if nilTC.ResolveThemeFont("+mj-lt") != "+mj-lt" {
		t.Errorf("nil themeColors ではそのまま返す")
	}
}

func TestCollectPlaceholders(t *testing.T) {
	children := []xmlSpTreeChild{
		{
			Sp: &xmlSp{
				NvSpPr: xmlNvSpPr{
					CNvPr: xmlCNvPr{ID: 1, Name: "Title"},
					NvPr:  xmlNvPr{Ph: &xmlPh{Type: "title"}},
				},
				SpPr: xmlSpPr{
					Xfrm: &xmlXfrm{Off: xmlOff{X: 100, Y: 200}},
				},
			},
		},
		{
			Sp: &xmlSp{
				NvSpPr: xmlNvSpPr{
					CNvPr: xmlCNvPr{ID: 2, Name: "Rect"},
					NvPr:  xmlNvPr{}, // プレースホルダーなし
				},
			},
		},
	}

	out := make(map[phKey]*placeholderDef)
	collectPlaceholders(children, out)

	if len(out) != 1 {
		t.Fatalf("プレースホルダー数: got %d, want 1", len(out))
	}
	def, ok := out[phKey{Type: "title"}]
	if !ok {
		t.Fatal("title プレースホルダーが収集されていない")
	}
	if def.xfrm == nil || def.xfrm.Off.X != 100 {
		t.Errorf("xfrm が正しく収集されていない")
	}
}
