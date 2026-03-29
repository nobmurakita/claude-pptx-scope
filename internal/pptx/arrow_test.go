package pptx

import (
	"testing"
)

func TestResolveArrowType(t *testing.T) {
	arrow := func(typ string) *xmlLineEnd { return &xmlLineEnd{Type: typ} }

	tests := []struct {
		name string
		head *xmlLineEnd
		tail *xmlLineEnd
		want string
	}{
		{"both nil", nil, nil, "none"},
		{"tail only", nil, arrow("triangle"), "end"},
		{"head only", arrow("triangle"), nil, "start"},
		{"both", arrow("triangle"), arrow("triangle"), "both"},
		{"none type", arrow("none"), arrow("none"), "none"},
	}
	for _, tt := range tests {
		got := resolveArrowType(tt.head, tt.tail)
		if got != tt.want {
			t.Errorf("%s: resolveArrowType = %q, want %q", tt.name, got, tt.want)
		}
	}
}
