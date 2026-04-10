package pptx

import (
	"encoding/xml"
	"fmt"
)

// ---------- コネクタ ----------

// xmlCxnSp は p:cxnSp 要素
type xmlCxnSp struct {
	NvCxnSpPr xmlNvCxnSpPr `xml:"nvCxnSpPr"`
	SpPr      xmlSpPr      `xml:"spPr"`
	TxBody    *xmlTxBody   `xml:"txBody"`
}

type xmlNvCxnSpPr struct {
	CNvPr      xmlCNvPr      `xml:"cNvPr"`
	CNvCxnSpPr xmlCNvCxnSpPr `xml:"cNvCxnSpPr"`
}

type xmlCNvCxnSpPr struct {
	StCxn  *xmlCxnRef `xml:"stCxn"`
	EndCxn *xmlCxnRef `xml:"endCxn"`
}

type xmlCxnRef struct {
	ID  int `xml:"id,attr"`
	Idx int `xml:"idx,attr"`
}

// ---------- 画像 ----------

// xmlPic は p:pic 要素
type xmlPic struct {
	NvPicPr  xmlNvPicPr  `xml:"nvPicPr"`
	BlipFill xmlBlipFill `xml:"blipFill"`
	SpPr     xmlSpPr     `xml:"spPr"`
}

type xmlNvPicPr struct {
	CNvPr xmlCNvPr `xml:"cNvPr"`
}

type xmlBlipFill struct {
	Blip xmlBlip `xml:"blip"`
}

type xmlBlip struct {
	Embed string `xml:"http://schemas.openxmlformats.org/officeDocument/2006/relationships embed,attr"`
}

// ---------- グループ ----------

// xmlGrpSp は p:grpSp 要素（子要素をXML出現順に保持する）
type xmlGrpSp struct {
	NvGrpSpPr xmlNvGrpSpPr
	GrpSpPr   xmlGrpSpPr
	Children  []xmlSpTreeChild
}

func (g *xmlGrpSp) UnmarshalXML(d *xml.Decoder, start xml.StartElement) error {
	for {
		tok, err := d.Token()
		if err != nil {
			return fmt.Errorf("grpSp のパースに失敗: %w", err)
		}
		switch el := tok.(type) {
		case xml.StartElement:
			switch el.Name.Local {
			case "nvGrpSpPr":
				if err := d.DecodeElement(&g.NvGrpSpPr, &el); err != nil {
					return err
				}
			case "grpSpPr":
				if err := d.DecodeElement(&g.GrpSpPr, &el); err != nil {
					return err
				}
			default:
				child, ok, err := decodeSpTreeChild(d, el)
				if err != nil {
					return err
				}
				if ok {
					g.Children = append(g.Children, child)
				} else {
					if err := d.Skip(); err != nil {
						return err
					}
				}
			}
		case xml.EndElement:
			return nil
		}
	}
}

type xmlNvGrpSpPr struct {
	CNvPr xmlCNvPr `xml:"cNvPr"`
}

type xmlGrpSpPr struct {
	Xfrm *xmlGrpXfrm `xml:"xfrm"`
}

type xmlGrpXfrm struct {
	Off   xmlOff `xml:"off"`
	Ext   xmlExt `xml:"ext"`
	ChOff xmlOff `xml:"chOff"`
	ChExt xmlExt `xml:"chExt"`
}

// ---------- テーブル ----------

// xmlGraphicFrame は p:graphicFrame 要素
type xmlGraphicFrame struct {
	NvGraphicFramePr xmlNvGraphicFramePr `xml:"nvGraphicFramePr"`
	Xfrm             *xmlXfrm            `xml:"xfrm"`
	Graphic          xmlGraphic          `xml:"graphic"`
}

type xmlNvGraphicFramePr struct {
	CNvPr xmlCNvPr `xml:"cNvPr"`
}

type xmlGraphic struct {
	GraphicData xmlGraphicData `xml:"graphicData"`
}

type xmlGraphicData struct {
	URI string  `xml:"uri,attr"`
	Tbl *xmlTbl `xml:"tbl"`
}

// xmlTbl は a:tbl 要素
type xmlTbl struct {
	TblGrid xmlTblGrid `xml:"tblGrid"`
	Trs     []xmlTr    `xml:"tr"`
}

type xmlTblGrid struct {
	GridCols []xmlGridCol `xml:"gridCol"`
}

type xmlGridCol struct {
	W int64 `xml:"w,attr"`
}

type xmlTr struct {
	H   int64   `xml:"h,attr"`
	Tcs []xmlTc `xml:"tc"`
}

type xmlTc struct {
	GridSpan int        `xml:"gridSpan,attr"`
	RowSpan  int        `xml:"rowSpan,attr"`
	VMerge   string     `xml:"vMerge,attr"`
	HMerge   string     `xml:"hMerge,attr"`
	TxBody   *xmlTxBody `xml:"txBody"`
	TcPr     *xmlTcPr   `xml:"tcPr"`
}

// xmlTcPr は a:tcPr 要素（テーブルセルプロパティ）
type xmlTcPr struct {
	LnL       *xmlLn        `xml:"lnL"` // 左罫線
	LnR       *xmlLn        `xml:"lnR"` // 右罫線
	LnT       *xmlLn        `xml:"lnT"` // 上罫線
	LnB       *xmlLn        `xml:"lnB"` // 下罫線
	SolidFill *xmlSolidFill `xml:"solidFill"`
	GradFill  *xmlGradFill  `xml:"gradFill"`
	NoFill    *struct{}     `xml:"noFill"`
}

// ---------- ノート ----------

// xmlNotes は p:notes 要素
type xmlNotes struct {
	XMLName xml.Name `xml:"notes"`
	CSld    xmlCSld  `xml:"cSld"`
}
