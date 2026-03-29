package pptx

import "encoding/xml"

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

// xmlSpTree は p:spTree 要素
type xmlSpTree struct {
	Shapes        []xmlSp           `xml:"sp"`
	GroupShapes   []xmlGrpSp        `xml:"grpSp"`
	Connectors    []xmlCxnSp        `xml:"cxnSp"`
	Pictures      []xmlPic          `xml:"pic"`
	GraphicFrames []xmlGraphicFrame `xml:"graphicFrame"`
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

// xmlGrpSp は p:grpSp 要素
type xmlGrpSp struct {
	NvGrpSpPr     xmlNvGrpSpPr      `xml:"nvGrpSpPr"`
	GrpSpPr       xmlGrpSpPr        `xml:"grpSpPr"`
	Shapes        []xmlSp           `xml:"sp"`
	GroupShapes   []xmlGrpSp        `xml:"grpSp"`
	Connectors    []xmlCxnSp        `xml:"cxnSp"`
	Pictures      []xmlPic          `xml:"pic"`
	GraphicFrames []xmlGraphicFrame `xml:"graphicFrame"`
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
}

// ---------- ノート ----------

// xmlNotes は p:notes 要素
type xmlNotes struct {
	XMLName xml.Name `xml:"notes"`
	CSld    xmlCSld  `xml:"cSld"`
}
