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
