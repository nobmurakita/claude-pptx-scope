package pptx

import (
	"testing"
)

func TestResolveCalloutPointer_NotCallout(t *testing.T) {
	geom := &xmlPrstGeom{Prst: "rect"}
	pos := &Position{X: 0, Y: 0, W: 1000, H: 1000}

	pt := resolveCalloutPointer(geom, pos)
	if pt != nil {
		t.Error("非吹き出し図形はnilを返すべき")
	}
}

func TestResolveCalloutPointer_NilInputs(t *testing.T) {
	pos := &Position{X: 0, Y: 0, W: 1000, H: 1000}

	if resolveCalloutPointer(nil, pos) != nil {
		t.Error("geom が nil の場合は nil を返すべき")
	}

	geom := &xmlPrstGeom{Prst: "wedgeRectCallout"}
	if resolveCalloutPointer(geom, nil) != nil {
		t.Error("pos が nil の場合は nil を返すべき")
	}
}

func TestResolveCalloutPointer_DefaultValues(t *testing.T) {
	geom := &xmlPrstGeom{Prst: "wedgeRectCallout"}
	pos := &Position{X: 1000, Y: 2000, W: 100000, H: 100000}

	pt := resolveCalloutPointer(geom, pos)
	if pt == nil {
		t.Fatal("吹き出し図形はポインタを返すべき")
	}

	// デフォルト: adj1=-20833, adj2=62500
	// px = 1000 + 100000/2 + (-20833)*100000/100000 = 1000 + 50000 - 20833 = 30167
	// py = 2000 + 100000/2 + 62500*100000/100000 = 2000 + 50000 + 62500 = 114500
	if pt.X != 30167 {
		t.Errorf("X: got %g, want 30167", pt.X)
	}
	if pt.Y != 114500 {
		t.Errorf("Y: got %g, want 114500", pt.Y)
	}
}

func TestResolveCalloutPointer_CustomAdj(t *testing.T) {
	geom := &xmlPrstGeom{
		Prst: "wedgeRectCallout",
		AvLst: &xmlAvLst{
			Gd: []xmlGd{
				{Name: "adj1", Fmla: "val 0"},
				{Name: "adj2", Fmla: "val 0"},
			},
		},
	}
	pos := &Position{X: 0, Y: 0, W: 200000, H: 200000}

	pt := resolveCalloutPointer(geom, pos)
	if pt == nil {
		t.Fatal("ポインタを返すべき")
	}

	// adj1=0, adj2=0 → ポインタは図形の中心
	if pt.X != 100000 {
		t.Errorf("X: got %g, want 100000", pt.X)
	}
	if pt.Y != 100000 {
		t.Errorf("Y: got %g, want 100000", pt.Y)
	}
}

func TestResolveCalloutPointer_AllTypes(t *testing.T) {
	pos := &Position{X: 0, Y: 0, W: 100000, H: 100000}

	calloutTypes := []string{
		"wedgeRectCallout", "wedgeRoundRectCallout", "wedgeEllipseCallout",
		"cloudCallout", "borderCallout1", "borderCallout2", "borderCallout3",
	}

	for _, typ := range calloutTypes {
		geom := &xmlPrstGeom{Prst: typ}
		pt := resolveCalloutPointer(geom, pos)
		if pt == nil {
			t.Errorf("%s: ポインタが nil", typ)
		}
	}
}

func TestParseGdVal(t *testing.T) {
	tests := []struct {
		fmla string
		want *int64
	}{
		{"val 12345", intPtr(12345)},
		{"val -5000", intPtr(-5000)},
		{"val 0", intPtr(0)},
		{"mod x y z", nil},
		{"", nil},
	}

	for _, tt := range tests {
		got := parseGdVal(tt.fmla)
		if tt.want == nil {
			if got != nil {
				t.Errorf("parseGdVal(%q): got %d, want nil", tt.fmla, *got)
			}
		} else {
			if got == nil {
				t.Errorf("parseGdVal(%q): got nil, want %d", tt.fmla, *tt.want)
			} else if *got != *tt.want {
				t.Errorf("parseGdVal(%q): got %d, want %d", tt.fmla, *got, *tt.want)
			}
		}
	}
}

func intPtr(v int64) *int64 {
	return &v
}
