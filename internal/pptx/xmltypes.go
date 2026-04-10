package pptx

import (
	"encoding/xml"
	"fmt"
	"strconv"
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
	ID         int            `xml:"id,attr"`
	Name       string         `xml:"name,attr"`
	Descr      string         `xml:"descr,attr"`
	Hidden     bool           `xml:"hidden,attr"`
	HlinkClick *xmlHlinkClick `xml:"hlinkClick"`
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
	GradFill  *xmlGradFill  `xml:"gradFill"`
	NoFill    *struct{}     `xml:"noFill"`
	Ln        *xmlLn        `xml:"ln"`
}

// xmlGradFill は a:gradFill 要素（グラデーション塗りつぶし）
type xmlGradFill struct {
	GsLst []xmlGradStop `xml:"gsLst>gs"`
}

// xmlGradStop は a:gs 要素（グラデーションストップ）
type xmlGradStop struct {
	Pos       int           `xml:"pos,attr"`
	SolidFill xmlSolidFill  // gs の子に直接 srgbClr/schemeClr が来る
}

func (gs *xmlGradStop) UnmarshalXML(d *xml.Decoder, start xml.StartElement) error {
	for _, attr := range start.Attr {
		if attr.Name.Local == "pos" {
			v, err := strconv.Atoi(attr.Value)
			if err != nil {
				return fmt.Errorf("グラデーションストップの pos 属性が不正です: %w", err)
			}
			gs.Pos = v
		}
	}
	// gs の子要素は solidFill と同じ色要素（srgbClr, schemeClr）
	for {
		tok, err := d.Token()
		if err != nil {
			return err
		}
		switch el := tok.(type) {
		case xml.StartElement:
			switch el.Name.Local {
			case "srgbClr":
				var v xmlSrgbClr
				if err := d.DecodeElement(&v, &el); err != nil {
					return err
				}
				gs.SolidFill.SrgbClr = &v
			case "schemeClr":
				var v xmlSchemeClr
				if err := d.DecodeElement(&v, &el); err != nil {
					return err
				}
				gs.SolidFill.SchemeClr = &v
			default:
				if err := d.Skip(); err != nil {
					return err
				}
			}
		case xml.EndElement:
			return nil
		}
	}
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

// colorTransform は色変換操作（XML出現順を保持する）
type colorTransform struct {
	Op  string // "lumMod", "lumOff", "tint", "shade"
	Val int
}

// xmlSrgbClr は a:srgbClr 要素
type xmlSrgbClr struct {
	Val        string
	Transforms []colorTransform
}

func (c *xmlSrgbClr) UnmarshalXML(d *xml.Decoder, start xml.StartElement) error {
	for _, attr := range start.Attr {
		if attr.Name.Local == "val" {
			c.Val = attr.Value
		}
	}
	return unmarshalColorTransforms(d, &c.Transforms)
}

// xmlSchemeClr は a:schemeClr 要素
type xmlSchemeClr struct {
	Val        string
	Transforms []colorTransform
}

func (c *xmlSchemeClr) UnmarshalXML(d *xml.Decoder, start xml.StartElement) error {
	for _, attr := range start.Attr {
		if attr.Name.Local == "val" {
			c.Val = attr.Value
		}
	}
	return unmarshalColorTransforms(d, &c.Transforms)
}

// unmarshalColorTransforms は色変換子要素をXML出現順にパースする
func unmarshalColorTransforms(d *xml.Decoder, transforms *[]colorTransform) error {
	for {
		tok, err := d.Token()
		if err != nil {
			return err
		}
		switch el := tok.(type) {
		case xml.StartElement:
			switch el.Name.Local {
			case "lumMod", "lumOff", "tint", "shade":
				var pct xmlPercentage
				if err := d.DecodeElement(&pct, &el); err != nil {
					return err
				}
				*transforms = append(*transforms, colorTransform{Op: el.Name.Local, Val: pct.Val})
			default:
				if err := d.Skip(); err != nil {
					return err
				}
			}
		case xml.EndElement:
			return nil
		}
	}
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

// ---------- レイアウト・マスター ----------

// xmlSldLayout は p:sldLayout 要素
type xmlSldLayout struct {
	XMLName xml.Name `xml:"sldLayout"`
	CSld    xmlCSld  `xml:"cSld"`
}

// xmlSldMaster は p:sldMaster 要素
type xmlSldMaster struct {
	XMLName  xml.Name     `xml:"sldMaster"`
	CSld     xmlCSld      `xml:"cSld"`
	TxStyles *xmlTxStyles `xml:"txStyles"`
}

// xmlTxStyles は p:txStyles 要素（マスターレベルのデフォルトテキストスタイル）
type xmlTxStyles struct {
	TitleStyle *xmlLstStyle `xml:"titleStyle"`
	BodyStyle  *xmlLstStyle `xml:"bodyStyle"`
	OtherStyle *xmlLstStyle `xml:"otherStyle"`
}

// xmlLstStyle は a:lstStyle 要素（レベル別段落プロパティのリスト）
type xmlLstStyle struct {
	Lvl1pPr *xmlLvlPPr `xml:"lvl1pPr"`
	Lvl2pPr *xmlLvlPPr `xml:"lvl2pPr"`
	Lvl3pPr *xmlLvlPPr `xml:"lvl3pPr"`
	Lvl4pPr *xmlLvlPPr `xml:"lvl4pPr"`
	Lvl5pPr *xmlLvlPPr `xml:"lvl5pPr"`
	Lvl6pPr *xmlLvlPPr `xml:"lvl6pPr"`
	Lvl7pPr *xmlLvlPPr `xml:"lvl7pPr"`
	Lvl8pPr *xmlLvlPPr `xml:"lvl8pPr"`
	Lvl9pPr *xmlLvlPPr `xml:"lvl9pPr"`
}

// GetLevel はレベル番号（0始まり）に対応する xmlLvlPPr を返す
func (ls *xmlLstStyle) GetLevel(level int) *xmlLvlPPr {
	if ls == nil {
		return nil
	}
	switch level {
	case 0:
		return ls.Lvl1pPr
	case 1:
		return ls.Lvl2pPr
	case 2:
		return ls.Lvl3pPr
	case 3:
		return ls.Lvl4pPr
	case 4:
		return ls.Lvl5pPr
	case 5:
		return ls.Lvl6pPr
	case 6:
		return ls.Lvl7pPr
	case 7:
		return ls.Lvl8pPr
	case 8:
		return ls.Lvl9pPr
	default:
		return nil
	}
}

// xmlLvlPPr は a:lvl1pPr 〜 a:lvl9pPr 要素
type xmlLvlPPr struct {
	Algn      string        `xml:"algn,attr"`
	MarL      *int64        `xml:"marL,attr"`
	Indent    *int64        `xml:"indent,attr"`
	BuNone    *struct{}     `xml:"buNone"`
	BuChar    *xmlBuChar    `xml:"buChar"`
	BuAutoNum *xmlBuAutoNum `xml:"buAutoNum"`
	LnSpc     *xmlSpacing   `xml:"lnSpc"`
	SpcBef    *xmlSpacing   `xml:"spcBef"`
	SpcAft    *xmlSpacing   `xml:"spcAft"`
	DefRPr    *xmlRPr       `xml:"defRPr"`
}

// xmlSpacing は a:lnSpc/a:spcBef/a:spcAft 要素。
// 子として a:spcPct（パーセント×1000）または a:spcPts（ポイント×100）を持つ
type xmlSpacing struct {
	SpcPct *xmlSpacingVal `xml:"spcPct"`
	SpcPts *xmlSpacingVal `xml:"spcPts"`
}

type xmlSpacingVal struct {
	Val int `xml:"val,attr"`
}

// ---------- テキスト ----------

// xmlTxBody は p:txBody 要素
type xmlTxBody struct {
	BodyPr   xmlBodyPr    `xml:"bodyPr"`
	LstStyle *xmlLstStyle `xml:"lstStyle"`
	Ps       []xmlP       `xml:"p"`
}

type xmlBodyPr struct {
	Anchor string `xml:"anchor,attr"`
	LIns   *int64 `xml:"lIns,attr"`
	RIns   *int64 `xml:"rIns,attr"`
	TIns   *int64 `xml:"tIns,attr"`
	BIns   *int64 `xml:"bIns,attr"`
}

// xmlP は a:p 要素（段落）。子要素をXML出現順に保持する。
type xmlP struct {
	PPr        *xmlPPr
	EndParaRPr *xmlRPr
	Elements   []xmlParagraphElement // 出現順の要素リスト（r, br, fld）
}

// xmlParagraphElement は段落の子要素（いずれか1つが非nil）
type xmlParagraphElement struct {
	R   *xmlR
	Br  bool // a:br 要素
	Fld *xmlFld
}

func (p *xmlP) UnmarshalXML(d *xml.Decoder, start xml.StartElement) error {
	for {
		tok, err := d.Token()
		if err != nil {
			return fmt.Errorf("a:p のパースに失敗: %w", err)
		}
		switch el := tok.(type) {
		case xml.StartElement:
			switch el.Name.Local {
			case "pPr":
				var v xmlPPr
				if err := d.DecodeElement(&v, &el); err != nil {
					return err
				}
				p.PPr = &v
			case "r":
				var v xmlR
				if err := d.DecodeElement(&v, &el); err != nil {
					return err
				}
				p.Elements = append(p.Elements, xmlParagraphElement{R: &v})
			case "br":
				p.Elements = append(p.Elements, xmlParagraphElement{Br: true})
				if err := d.Skip(); err != nil {
					return err
				}
			case "fld":
				var v xmlFld
				if err := d.DecodeElement(&v, &el); err != nil {
					return err
				}
				p.Elements = append(p.Elements, xmlParagraphElement{Fld: &v})
			case "endParaRPr":
				var v xmlRPr
				if err := d.DecodeElement(&v, &el); err != nil {
					return err
				}
				p.EndParaRPr = &v
			default:
				if err := d.Skip(); err != nil {
					return err
				}
			}
		case xml.EndElement:
			return nil
		}
	}
}

// xmlPPr は a:pPr 要素（段落プロパティ）
type xmlPPr struct {
	Lvl       int           `xml:"lvl,attr"`
	Algn      string        `xml:"algn,attr"`
	MarL      *int64        `xml:"marL,attr"`
	Indent    *int64        `xml:"indent,attr"`
	BuNone    *struct{}     `xml:"buNone"`
	BuChar    *xmlBuChar    `xml:"buChar"`
	BuAutoNum *xmlBuAutoNum `xml:"buAutoNum"`
	LnSpc     *xmlSpacing   `xml:"lnSpc"`
	SpcBef    *xmlSpacing   `xml:"spcBef"`
	SpcAft    *xmlSpacing   `xml:"spcAft"`
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
	Lang       string         `xml:"lang,attr"`
	Sz         int            `xml:"sz,attr"`
	B          string         `xml:"b,attr"`
	I          string         `xml:"i,attr"`
	U          string         `xml:"u,attr"`
	Strike     string         `xml:"strike,attr"`
	Baseline   *int           `xml:"baseline,attr"` // 上付き/下付き文字（パーセント×1000。正=上付き、負=下付き）
	Cap        string         `xml:"cap,attr"`      // 英字大文字化（"all"/"small"/"none"）
	SolidFill  *xmlSolidFill  `xml:"solidFill"`
	Highlight  *xmlSolidFill  `xml:"highlight"` // 文字の背景色（a:highlight）。中身は solidFill と同じ構造
	Latin      *xmlFont       `xml:"latin"`
	Ea         *xmlFont       `xml:"ea"`
	Cs         *xmlFont       `xml:"cs"`
	HlinkClick *xmlHlinkClick `xml:"hlinkClick"`
}

// xmlHlinkClick は a:hlinkClick 要素
type xmlHlinkClick struct {
	RID    string `xml:"http://schemas.openxmlformats.org/officeDocument/2006/relationships id,attr"`
	Action string `xml:"action,attr"`
}

type xmlFont struct {
	Typeface string `xml:"typeface,attr"`
}

// ---------- プレゼンテーション ----------

// xmlPresentation は presentation.xml の構造
type xmlPresentation struct {
	XMLName          xml.Name     `xml:"presentation"`
	SldSz            xmlSldSz     `xml:"sldSz"`
	SldIdLst         struct {
		SldId []xmlSldId `xml:"sldId"`
	} `xml:"sldIdLst"`
	DefaultTextStyle *xmlLstStyle `xml:"defaultTextStyle"`
}

type xmlSldSz struct {
	Cx int64 `xml:"cx,attr"`
	Cy int64 `xml:"cy,attr"`
}

type xmlSldId struct {
	ID  string `xml:"id,attr"`
	RID string `xml:"http://schemas.openxmlformats.org/officeDocument/2006/relationships id,attr"`
}

