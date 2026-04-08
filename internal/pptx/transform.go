package pptx

import "math"

// emuToPt は EMU（English Metric Units）を pt（ポイント）に変換する。
// 1pt = 12700 EMU。小数点以下2桁に丸める。
func emuToPt(emu int64) float64 {
	return math.Round(float64(emu)/127) / 100
}

// emuToPtPtr は *int64（EMU）を *float64（pt）に変換する
func emuToPtPtr(p *int64) *float64 {
	if p == nil {
		return nil
	}
	v := emuToPt(*p)
	return &v
}

// xfrmToPosition は xfrm 要素を Position（pt単位）に変換する
func xfrmToPosition(xfrm *xmlXfrm) *Position {
	if xfrm == nil {
		return nil
	}
	return &Position{
		X: emuToPt(xfrm.Off.X),
		Y: emuToPt(xfrm.Off.Y),
		W: emuToPt(xfrm.Ext.Cx),
		H: emuToPt(xfrm.Ext.Cy),
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

// coordTransform は子座標空間から親座標空間への変換パラメータ（pt単位）
type coordTransform struct {
	chOffX, chOffY float64
	chExtW, chExtH float64
	grpX, grpY     float64
	grpW, grpH     float64
}

// transformGroupChildren はグループ内の子要素の座標を絶対座標に変換する
// グループの子座標空間(ChOff/ChExt)からスライド座標空間(Off/Ext)へマッピングする
func transformGroupChildren(children []Shape, xfrm *xmlGrpXfrm) {
	ct := coordTransform{
		chOffX: emuToPt(xfrm.ChOff.X), chOffY: emuToPt(xfrm.ChOff.Y),
		chExtW: emuToPt(xfrm.ChExt.Cx), chExtH: emuToPt(xfrm.ChExt.Cy),
		grpX: emuToPt(xfrm.Off.X), grpY: emuToPt(xfrm.Off.Y),
		grpW: emuToPt(xfrm.Ext.Cx), grpH: emuToPt(xfrm.Ext.Cy),
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
