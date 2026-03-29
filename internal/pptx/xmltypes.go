package pptx

import (
	"encoding/xml"
	"fmt"
)

// スライドXMLの型定義

// xmlSlide は p:sld 要素
type xmlSlide struct {
	XMLName xml.Name `xml:"sld"`
	Show    string   `xml:"show,attr"`
	CSld    xmlCSld  `xml:"cSld"`
}

// xmlCSld は p:cSld 要素
type xmlCSld struct {
	SpTree xmlSpTree `xml:"spTree"`
}

// xmlSpTree は p:spTree 要素（子要素をXML出現順に保持する）
type xmlSpTree struct {
	Children []xmlSpTreeChild
}

// xmlSpTreeChild は spTree/grpSp の子要素（いずれか1つが非nil）
type xmlSpTreeChild struct {
	Sp           *xmlSp
	CxnSp        *xmlCxnSp
	Pic          *xmlPic
	GrpSp        *xmlGrpSp
	GraphicFrame *xmlGraphicFrame
}

func (t *xmlSpTree) UnmarshalXML(d *xml.Decoder, start xml.StartElement) error {
	for {
		tok, err := d.Token()
		if err != nil {
			return fmt.Errorf("spTree のパースに失敗: %w", err)
		}
		switch el := tok.(type) {
		case xml.StartElement:
			child, ok, err := decodeSpTreeChild(d, el)
			if err != nil {
				return err
			}
			if ok {
				t.Children = append(t.Children, child)
			} else {
				if err := d.Skip(); err != nil {
					return err
				}
			}
		case xml.EndElement:
			return nil
		}
	}
}

// decodeSpTreeChild はタグ名に応じて子要素をデコードする
func decodeSpTreeChild(d *xml.Decoder, el xml.StartElement) (xmlSpTreeChild, bool, error) {
	var child xmlSpTreeChild
	switch el.Name.Local {
	case "sp":
		var v xmlSp
		if err := d.DecodeElement(&v, &el); err != nil {
			return child, false, err
		}
		child.Sp = &v
	case "cxnSp":
		var v xmlCxnSp
		if err := d.DecodeElement(&v, &el); err != nil {
			return child, false, err
		}
		child.CxnSp = &v
	case "pic":
		var v xmlPic
		if err := d.DecodeElement(&v, &el); err != nil {
			return child, false, err
		}
		child.Pic = &v
	case "grpSp":
		var v xmlGrpSp
		if err := d.DecodeElement(&v, &el); err != nil {
			return child, false, err
		}
		child.GrpSp = &v
	case "graphicFrame":
		var v xmlGraphicFrame
		if err := d.DecodeElement(&v, &el); err != nil {
			return child, false, err
		}
		child.GraphicFrame = &v
	default:
		return child, false, nil
	}
	return child, true, nil
}

// xmlSp は p:sp 要素（通常の図形）
type xmlSp struct {
	NvSpPr xmlNvSpPr  `xml:"nvSpPr"`
	SpPr   xmlSpPr    `xml:"spPr"`
	TxBody *xmlTxBody `xml:"txBody"`
}

// xmlNvSpPr は p:nvSpPr 要素
type xmlNvSpPr struct {
	CNvPr xmlCNvPr `xml:"cNvPr"`
	NvPr  xmlNvPr  `xml:"nvPr"`
}

// xmlCNvPr は p:cNvPr 要素
type xmlCNvPr struct {
	ID    int    `xml:"id,attr"`
	Name  string `xml:"name,attr"`
	Descr string `xml:"descr,attr"`
}

// xmlNvPr は p:nvPr 要素
type xmlNvPr struct {
	Ph *xmlPh `xml:"ph"`
}

// xmlPh はプレースホルダー要素
type xmlPh struct {
	Type string `xml:"type,attr"`
	Idx  string `xml:"idx,attr"`
}

// xmlSpPr は p:spPr 要素（図形のプロパティ）
type xmlSpPr struct {
	Xfrm      *xmlXfrm      `xml:"xfrm"`
	PrstGeom  *xmlPrstGeom  `xml:"prstGeom"`
	CustGeom  *struct{}     `xml:"custGeom"`
	SolidFill *xmlSolidFill `xml:"solidFill"`
	NoFill    *struct{}     `xml:"noFill"`
	Ln        *xmlLn        `xml:"ln"`
}

// xmlXfrm は a:xfrm 要素
type xmlXfrm struct {
	Rot   int64  `xml:"rot,attr"`
	FlipH bool   `xml:"flipH,attr"`
	FlipV bool   `xml:"flipV,attr"`
	Off   xmlOff `xml:"off"`
	Ext   xmlExt `xml:"ext"`
}

type xmlOff struct {
	X int64 `xml:"x,attr"`
	Y int64 `xml:"y,attr"`
}

type xmlExt struct {
	Cx int64 `xml:"cx,attr"`
	Cy int64 `xml:"cy,attr"`
}

// xmlPrstGeom は a:prstGeom 要素
type xmlPrstGeom struct {
	Prst  string    `xml:"prst,attr"`
	AvLst *xmlAvLst `xml:"avLst"`
}

// xmlAvLst は a:avLst 要素（調整ハンドル）
type xmlAvLst struct {
	Gd []xmlGd `xml:"gd"`
}

// xmlGd は a:gd 要素
type xmlGd struct {
	Name string `xml:"name,attr"`
	Fmla string `xml:"fmla,attr"`
}

// xmlSolidFill は a:solidFill 要素
type xmlSolidFill struct {
	SrgbClr   *xmlSrgbClr   `xml:"srgbClr"`
	SchemeClr *xmlSchemeClr `xml:"schemeClr"`
}

// xmlSrgbClr は a:srgbClr 要素
type xmlSrgbClr struct {
	Val    string         `xml:"val,attr"`
	LumMod *xmlPercentage `xml:"lumMod"`
	LumOff *xmlPercentage `xml:"lumOff"`
	Tint   *xmlPercentage `xml:"tint"`
	Shade  *xmlPercentage `xml:"shade"`
}

// xmlSchemeClr は a:schemeClr 要素
type xmlSchemeClr struct {
	Val    string         `xml:"val,attr"`
	LumMod *xmlPercentage `xml:"lumMod"`
	LumOff *xmlPercentage `xml:"lumOff"`
	Tint   *xmlPercentage `xml:"tint"`
	Shade  *xmlPercentage `xml:"shade"`
}

type xmlPercentage struct {
	Val int `xml:"val,attr"`
}

// xmlLn は a:ln 要素
type xmlLn struct {
	W         int64         `xml:"w,attr"`
	SolidFill *xmlSolidFill `xml:"solidFill"`
	PrstDash  *xmlPrstDash  `xml:"prstDash"`
	NoFill    *struct{}     `xml:"noFill"`
	HeadEnd   *xmlLineEnd   `xml:"headEnd"`
	TailEnd   *xmlLineEnd   `xml:"tailEnd"`
}

type xmlPrstDash struct {
	Val string `xml:"val,attr"`
}

type xmlLineEnd struct {
	Type string `xml:"type,attr"`
}

// ---------- テキスト ----------

// xmlTxBody は p:txBody 要素
type xmlTxBody struct {
	BodyPr xmlBodyPr `xml:"bodyPr"`
	Ps     []xmlP    `xml:"p"`
}

type xmlBodyPr struct {
	Anchor string `xml:"anchor,attr"`
}

// xmlP は a:p 要素（段落）
type xmlP struct {
	PPr        *xmlPPr  `xml:"pPr"`
	Rs         []xmlR   `xml:"r"`
	Fld        []xmlFld `xml:"fld"`
	EndParaRPr *xmlRPr  `xml:"endParaRPr"`
}

// xmlPPr は a:pPr 要素（段落プロパティ）
type xmlPPr struct {
	Lvl       int           `xml:"lvl,attr"`
	Algn      string        `xml:"algn,attr"`
	BuNone    *struct{}     `xml:"buNone"`
	BuChar    *xmlBuChar    `xml:"buChar"`
	BuAutoNum *xmlBuAutoNum `xml:"buAutoNum"`
	DefRPr    *xmlRPr       `xml:"defRPr"`
}

type xmlBuChar struct {
	Char string `xml:"char,attr"`
}

type xmlBuAutoNum struct {
	Type    string `xml:"type,attr"`
	StartAt int    `xml:"startAt,attr"`
}

// xmlR は a:r 要素（テキストラン）
type xmlR struct {
	RPr *xmlRPr `xml:"rPr"`
	T   string  `xml:"t"`
}

// xmlFld は a:fld 要素（フィールド）
type xmlFld struct {
	RPr *xmlRPr `xml:"rPr"`
	T   string  `xml:"t"`
}

// xmlRPr は a:rPr / a:endParaRPr 要素（ランプロパティ）
type xmlRPr struct {
	Lang      string        `xml:"lang,attr"`
	Sz        int           `xml:"sz,attr"`
	B         string        `xml:"b,attr"`
	I         string        `xml:"i,attr"`
	U         string        `xml:"u,attr"`
	Strike    string        `xml:"strike,attr"`
	SolidFill *xmlSolidFill `xml:"solidFill"`
	Latin     *xmlFont      `xml:"latin"`
	Ea        *xmlFont      `xml:"ea"`
	Cs        *xmlFont      `xml:"cs"`
}

type xmlFont struct {
	Typeface string `xml:"typeface,attr"`
}

