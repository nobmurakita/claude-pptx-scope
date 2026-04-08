package pptx

// xfrmToPosition は xfrm 要素を Position に変換する
func xfrmToPosition(xfrm *xmlXfrm) *Position {
	if xfrm == nil {
		return nil
	}
	return &Position{
		X: xfrm.Off.X,
		Y: xfrm.Off.Y,
		W: xfrm.Ext.Cx,
		H: xfrm.Ext.Cy,
	}
}

// xfrmFlip は xfrm の反転情報を文字列に変換する
func xfrmFlip(xfrm *xmlXfrm) string {
	if xfrm == nil {
		return ""
	}
	h := xfrm.FlipH
	v := xfrm.FlipV
	if h && v {
		return "hv"
	}
	if h {
		return "h"
	}
	if v {
		return "v"
	}
	return ""
}

// coordTransform は子座標空間から親座標空間への変換パラメータ
type coordTransform struct {
	chOffX, chOffY int64
	chExtW, chExtH int64
	grpX, grpY     int64
	grpW, grpH     int64
}

// transformGroupChildren はグループ内の子要素の座標を絶対座標に変換する
// グループの子座標空間(ChOff/ChExt)からスライド座標空間(Off/Ext)へマッピングする
func transformGroupChildren(children []Shape, xfrm *xmlGrpXfrm) {
	ct := coordTransform{
		chOffX: xfrm.ChOff.X, chOffY: xfrm.ChOff.Y,
		chExtW: xfrm.ChExt.Cx, chExtH: xfrm.ChExt.Cy,
		grpX: xfrm.Off.X, grpY: xfrm.Off.Y,
		grpW: xfrm.Ext.Cx, grpH: xfrm.Ext.Cy,
	}
	// ChExt が 0 の場合はスケール計算でゼロ除算になるため変換をスキップする。
	// 通常の PPTX では発生しないが、壊れたファイルへの防御として残す。
	if ct.chExtW == 0 || ct.chExtH == 0 {
		return
	}
	ct.applyToChildren(children)
}

// applyToChildren は子要素の座標を再帰的に親座標空間に変換する
func (ct *coordTransform) applyToChildren(children []Shape) {
	for i := range children {
		ct.transformShapePos(children[i].Pos)
		ct.transformPointPos(children[i].CalloutPointer)
		ct.transformPointPos(children[i].Start)
		ct.transformPointPos(children[i].End)
		if children[i].Type == "group" {
			ct.applyToChildren(children[i].Children)
		}
	}
}

// transformShapePos は Position を子座標空間から親座標空間に変換する
func (ct *coordTransform) transformShapePos(pos *Position) {
	if pos == nil {
		return
	}
	pos.X = ct.grpX + (pos.X-ct.chOffX)*ct.grpW/ct.chExtW
	pos.Y = ct.grpY + (pos.Y-ct.chOffY)*ct.grpH/ct.chExtH
	pos.W = pos.W * ct.grpW / ct.chExtW
	pos.H = pos.H * ct.grpH / ct.chExtH
}

// transformPointPos は Point を子座標空間から親座標空間に変換する
func (ct *coordTransform) transformPointPos(pt *Point) {
	if pt == nil {
		return
	}
	pt.X = ct.grpX + (pt.X-ct.chOffX)*ct.grpW/ct.chExtW
	pt.Y = ct.grpY + (pt.Y-ct.chOffY)*ct.grpH/ct.chExtH
}
