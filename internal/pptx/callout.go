package pptx

import (
	"strconv"
	"strings"
)

// calloutDefaults は吹き出し図形のデフォルトadj値（OOXML仕様準拠）
// adj1=水平方向オフセット, adj2=垂直方向オフセット（100000分率）
var calloutDefaults = map[string][2]int64{
	"wedgeRectCallout":      {-20833, 62500},
	"wedgeRoundRectCallout": {-20833, 62500},
	"wedgeEllipseCallout":   {-20833, 62500},
	"cloudCallout":          {-20833, 62500},
	"borderCallout1":        {18750, -8333},
	"borderCallout2":        {18750, -8333},
	"borderCallout3":        {18750, -8333},
}

// resolveCalloutPointer は吹き出し図形のポインタ位置を計算する
func resolveCalloutPointer(geom *xmlPrstGeom, pos *Position) *Point {
	if geom == nil || pos == nil {
		return nil
	}

	defaults, isCallout := calloutDefaults[geom.Prst]
	if !isCallout {
		return nil
	}

	a1, a2 := defaults[0], defaults[1]

	if geom.AvLst != nil {
		for _, gd := range geom.AvLst.Gd {
			val := parseGdVal(gd.Fmla)
			if val == nil {
				continue
			}
			switch gd.Name {
			case "adj1":
				a1 = *val
			case "adj2":
				a2 = *val
			}
		}
	}

	px := pos.X + pos.W/2 + float64(a1)*pos.W/100000
	py := pos.Y + pos.H/2 + float64(a2)*pos.H/100000

	return &Point{X: px, Y: py}
}

// parseGdVal は "val N" 形式の数式から値を取得する
func parseGdVal(fmla string) *int64 {
	if !strings.HasPrefix(fmla, "val ") {
		return nil
	}
	v, err := strconv.ParseInt(strings.TrimPrefix(fmla, "val "), 10, 64)
	if err != nil {
		return nil
	}
	return &v
}
