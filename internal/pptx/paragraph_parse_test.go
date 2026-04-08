package pptx

import (
	"testing"
)

func TestFormatAutoNum(t *testing.T) {
	tests := []struct {
		numType string
		num     int
		want    string
	}{
		{"arabicPeriod", 1, "1."},
		{"arabicPeriod", 10, "10."},
		{"arabicParenR", 3, "3)"},
		{"alphaLcPeriod", 1, "a."},
		{"alphaLcPeriod", 26, "z."},
		{"alphaUcPeriod", 1, "A."},
		{"romanLcPeriod", 4, "iv."},
		{"romanUcPeriod", 4, "IV."},
		{"unknown", 5, "5."},
	}
	for _, tt := range tests {
		got := formatAutoNum(tt.numType, tt.num)
		if got != tt.want {
			t.Errorf("formatAutoNum(%q, %d) = %q, want %q", tt.numType, tt.num, got, tt.want)
		}
	}
}

func TestToUpperRoman(t *testing.T) {
	tests := []struct {
		n    int
		want string
	}{
		{1, "I"},
		{4, "IV"},
		{9, "IX"},
		{14, "XIV"},
		{42, "XLII"},
		{99, "XCIX"},
		{2024, "MMXXIV"},
	}
	for _, tt := range tests {
		got := toUpperRoman(tt.n)
		if got != tt.want {
			t.Errorf("toUpperRoman(%d) = %q, want %q", tt.n, got, tt.want)
		}
	}
}

func TestMapAlignment(t *testing.T) {
	tests := []struct {
		input string
		want  string
	}{
		{"l", "left"},
		{"r", "right"},
		{"ctr", "center"},
		{"just", "justify"},
		{"other", "other"},
	}
	for _, tt := range tests {
		got := mapAlignment(tt.input)
		if got != tt.want {
			t.Errorf("mapAlignment(%q) = %q, want %q", tt.input, got, tt.want)
		}
	}
}

func TestParseParagraphs_AutoNumResetOnNonBullet(t *testing.T) {
	ctx := newTestContext()

	ps := []xmlP{
		// autoNum: 1.
		mkPWithPPr(&xmlPPr{BuAutoNum: &xmlBuAutoNum{Type: "arabicPeriod"}}, xmlR{T: "first"}),
		// autoNum: 2.
		mkPWithPPr(&xmlPPr{BuAutoNum: &xmlBuAutoNum{Type: "arabicPeriod"}}, xmlR{T: "second"}),
		// PPr なし（箇条書き指定なし）→ リセット
		mkP(xmlR{T: "plain"}),
		// リセット後なので 1. から再開
		mkPWithPPr(&xmlPPr{BuAutoNum: &xmlBuAutoNum{Type: "arabicPeriod"}}, xmlR{T: "third"}),
	}

	paras := ctx.parseParagraphs(ps, nil)
	if len(paras) != 4 {
		t.Fatalf("got %d paragraphs, want 4", len(paras))
	}
	if paras[0].Bullet != "1." {
		t.Errorf("paras[0].Bullet: got %q, want %q", paras[0].Bullet, "1.")
	}
	if paras[1].Bullet != "2." {
		t.Errorf("paras[1].Bullet: got %q, want %q", paras[1].Bullet, "2.")
	}
	if paras[2].Bullet != "" {
		t.Errorf("paras[2].Bullet: got %q, want empty", paras[2].Bullet)
	}
	if paras[3].Bullet != "1." {
		t.Errorf("paras[3].Bullet: got %q, want %q (リセット後)", paras[3].Bullet, "1.")
	}
}

func TestResolveParaIndent_ExplicitValues(t *testing.T) {
	marL := int64(457200)
	indent := int64(-228600)
	ppr := &xmlPPr{MarL: &marL, Indent: &indent}

	gotMarL, gotIndent := resolveParaIndent(ppr, 0, nil)
	if gotMarL == nil || *gotMarL != 457200 {
		t.Errorf("MarL: got %v, want 457200", gotMarL)
	}
	if gotIndent == nil || *gotIndent != -228600 {
		t.Errorf("Indent: got %v, want -228600", gotIndent)
	}
}

func TestResolveParaIndent_InheritedValues(t *testing.T) {
	marL := int64(914400)
	indent := int64(-342900)
	inherited := &inheritedStyle{
		lstStyles: []*xmlLstStyle{
			{
				Lvl2pPr: &xmlLvlPPr{MarL: &marL, Indent: &indent},
			},
		},
	}

	gotMarL, gotIndent := resolveParaIndent(nil, 1, inherited)
	if gotMarL == nil || *gotMarL != 914400 {
		t.Errorf("MarL: got %v, want 914400", gotMarL)
	}
	if gotIndent == nil || *gotIndent != -342900 {
		t.Errorf("Indent: got %v, want -342900", gotIndent)
	}
}

func TestResolveParaIndent_ZeroOmitted(t *testing.T) {
	zero := int64(0)
	ppr := &xmlPPr{MarL: &zero, Indent: &zero}

	gotMarL, gotIndent := resolveParaIndent(ppr, 0, nil)
	if gotMarL != nil {
		t.Errorf("MarL=0 は省略されるべき: got %v", gotMarL)
	}
	if gotIndent != nil {
		t.Errorf("Indent=0 は省略されるべき: got %v", gotIndent)
	}
}

func TestExtractTextMargin_NonDefault(t *testing.T) {
	l := int64(0)
	r := int64(0)
	ti := int64(0)
	b := int64(0)
	bodyPr := xmlBodyPr{LIns: &l, RIns: &r, TIns: &ti, BIns: &b}

	tm := extractTextMargin(bodyPr)
	if tm == nil {
		t.Fatal("デフォルトと異なるマージンは出力されるべき")
	}
}

func TestExtractTextMargin_Default(t *testing.T) {
	l := int64(91440)
	r := int64(91440)
	ti := int64(45720)
	b := int64(45720)
	bodyPr := xmlBodyPr{LIns: &l, RIns: &r, TIns: &ti, BIns: &b}

	tm := extractTextMargin(bodyPr)
	if tm != nil {
		t.Errorf("デフォルト値のマージンは省略されるべき: got %+v", tm)
	}
}

func TestExtractTextMargin_Nil(t *testing.T) {
	bodyPr := xmlBodyPr{}
	tm := extractTextMargin(bodyPr)
	if tm != nil {
		t.Errorf("未指定のマージンは省略されるべき: got %+v", tm)
	}
}

func TestToLowerAlpha(t *testing.T) {
	tests := []struct {
		n    int
		want string
	}{
		{1, "a"},
		{26, "z"},
		{27, "aa"},
		{28, "ab"},
		{52, "az"},
		{53, "ba"},
		{702, "zz"},
		{703, "aaa"},
	}
	for _, tt := range tests {
		got := toLowerAlpha(tt.n)
		if got != tt.want {
			t.Errorf("toLowerAlpha(%d) = %q, want %q", tt.n, got, tt.want)
		}
	}
}

func TestToLowerAlpha_Invalid(t *testing.T) {
	got := toLowerAlpha(0)
	if got != "0" {
		t.Errorf("toLowerAlpha(0) = %q, want %q", got, "0")
	}
	got = toLowerAlpha(-1)
	if got != "-1" {
		t.Errorf("toLowerAlpha(-1) = %q, want %q", got, "-1")
	}
}

func TestLinksEqual(t *testing.T) {
	if !linksEqual(nil, nil) {
		t.Error("nil == nil should be true")
	}
	if linksEqual(nil, &HyperlinkData{URL: "https://example.com"}) {
		t.Error("nil != non-nil should be false")
	}
	a := &HyperlinkData{URL: "https://example.com"}
	b := &HyperlinkData{URL: "https://example.com"}
	if !linksEqual(a, b) {
		t.Error("identical URLs should be equal")
	}
	c := &HyperlinkData{Slide: 3}
	d := &HyperlinkData{Slide: 3}
	if !linksEqual(c, d) {
		t.Error("identical slides should be equal")
	}
	if linksEqual(a, c) {
		t.Error("URL != Slide should be false")
	}
}

func TestIsEmptyFont(t *testing.T) {
	if !isEmptyFont(&FontStyle{}) {
		t.Error("empty FontStyle should be detected as empty")
	}
	if isEmptyFont(&FontStyle{Bold: true}) {
		t.Error("FontStyle with Bold should not be empty")
	}
}

func TestFontsEqual(t *testing.T) {
	if !fontsEqual(nil, nil) {
		t.Error("nil == nil should be true")
	}
	if fontsEqual(nil, &FontStyle{}) {
		t.Error("nil != &FontStyle{} should be false")
	}
	a := &FontStyle{Name: "Arial", Size: 12, Bold: true}
	b := &FontStyle{Name: "Arial", Size: 12, Bold: true}
	if !fontsEqual(a, b) {
		t.Error("identical fonts should be equal")
	}
	b.Size = 14
	if fontsEqual(a, b) {
		t.Error("different size should not be equal")
	}
}
