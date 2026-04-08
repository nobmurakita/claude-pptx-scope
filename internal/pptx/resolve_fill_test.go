package pptx

import (
	"testing"
)

func TestSchemeClrToThemeIndex(t *testing.T) {
	tests := []struct {
		input string
		want  int
	}{
		{"bg1", 0},
		{"lt1", 0},
		{"tx1", 1},
		{"dk1", 1},
		{"accent1", 4},
		{"accent6", 9},
		{"hlink", 10},
		{"folHlink", 11},
		{"phClr", -1},
		{"unknown", -1},
	}
	for _, tt := range tests {
		got := schemeClrToThemeIndex(tt.input)
		if got != tt.want {
			t.Errorf("schemeClrToThemeIndex(%q) = %d, want %d", tt.input, got, tt.want)
		}
	}
}
