package pptx

import (
	"testing"
)

func TestNormalizeHexColor(t *testing.T) {
	tests := []struct {
		input string
		want  string
	}{
		{"FF4472C4", "#4472C4"},
		{"4472C4", "#4472C4"},
		{"#4472C4", "#4472C4"},
		{"#FF4472C4", "#4472C4"},
		{"abc", ""},
		{"", ""},
	}
	for _, tt := range tests {
		got := normalizeHexColor(tt.input)
		if got != tt.want {
			t.Errorf("normalizeHexColor(%q) = %q, want %q", tt.input, got, tt.want)
		}
	}
}

func TestParseHexRGB(t *testing.T) {
	r, g, b, ok := parseHexRGB("FF0000")
	if !ok || r != 1.0 || g != 0.0 || b != 0.0 {
		t.Errorf("parseHexRGB(FF0000) = (%v, %v, %v, %v)", r, g, b, ok)
	}

	_, _, _, ok = parseHexRGB("ZZ")
	if ok {
		t.Error("parseHexRGB should fail for invalid input")
	}
}

func TestRGBHSLRoundTrip(t *testing.T) {
	tests := [][3]float64{
		{1.0, 0.0, 0.0},
		{0.0, 1.0, 0.0},
		{0.0, 0.0, 1.0},
		{0.5, 0.5, 0.5},
	}
	for _, tt := range tests {
		h, s, l := rgbToHSL(tt[0], tt[1], tt[2])
		r, g, b := hslToRGB(h, s, l)
		if diff(r, tt[0]) > 0.01 || diff(g, tt[1]) > 0.01 || diff(b, tt[2]) > 0.01 {
			t.Errorf("RGB→HSL→RGB roundtrip: (%v,%v,%v) → (%v,%v,%v)", tt[0], tt[1], tt[2], r, g, b)
		}
	}
}

func TestApplyTint(t *testing.T) {
	// tint=0 は変化なし
	got := applyTint("#FF0000", 0)
	if got != "#FF0000" {
		t.Errorf("applyTint with 0 = %q, want #FF0000", got)
	}

	// tint>0 で明るくなる
	got = applyTint("#000000", 1.0)
	if got != "#FFFFFF" {
		t.Errorf("applyTint(#000000, 1.0) = %q, want #FFFFFF", got)
	}
}

func TestApplyLuminance(t *testing.T) {
	// lumMod=1, lumOff=0 は変化なし
	got := applyLuminance("#4472C4", 1.0, 0.0)
	if got != "#4472C4" {
		t.Errorf("applyLuminance identity = %q, want #4472C4", got)
	}
}

func diff(a, b float64) float64 {
	d := a - b
	if d < 0 {
		return -d
	}
	return d
}
