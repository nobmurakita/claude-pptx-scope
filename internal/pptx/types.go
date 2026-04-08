package pptx

// 出力用のデータ型

// SlideInfo はスライドのヘッダ情報（info/slides/search 共通）
type SlideInfo struct {
	Slide     int    `json:"slide"`
	Title     string `json:"title,omitempty"`
	Shapes    *int   `json:"shapes,omitempty"`
	HasNotes  bool   `json:"has_notes,omitempty"`
	HasImages bool   `json:"has_images,omitempty"`
	Hidden    bool   `json:"hidden,omitempty"`
}

// Shape は図形の出力データ
type Shape struct {
	ID             int         `json:"shape"`
	Type           string      `json:"type"`
	Name           string      `json:"name,omitempty"`
	Placeholder    string      `json:"placeholder,omitempty"`
	Pos            *Position   `json:"pos,omitempty"`
	Z              int         `json:"z"` // 描画順（0始まり）。0が有効値のため omitempty なし
	Rotation       float64     `json:"rotation,omitempty"`
	Flip           string      `json:"flip,omitempty"`
	Fill           string      `json:"fill,omitempty"`
	Line           *LineStyle  `json:"line,omitempty"`
	CalloutPointer *Point      `json:"callout_pointer,omitempty"`
	Paragraphs     []Paragraph `json:"paragraphs,omitempty"`
	Table          *TableData  `json:"table,omitempty"`
	Alignment      *Alignment  `json:"alignment,omitempty"`
	// テキストマージン（bodyPr の lIns/rIns/tIns/bIns）
	TextMargin *TextMargin `json:"text_margin,omitempty"`
	// コネクタ
	From          int            `json:"from,omitempty"`
	To            int            `json:"to,omitempty"`
	FromIdx       *int           `json:"from_idx,omitempty"`
	ToIdx         *int           `json:"to_idx,omitempty"`
	ConnectorType string         `json:"connector_type,omitempty"`
	Adj           map[string]int `json:"adj,omitempty"`
	Arrow         string         `json:"arrow,omitempty"`
	Start         *Point         `json:"start,omitempty"`
	End           *Point         `json:"end,omitempty"`
	Label         string         `json:"label,omitempty"`
	// ハイパーリンク（図形全体に設定されたリンク）
	Link *HyperlinkData `json:"link,omitempty"`
	// 画像
	AltText string `json:"alt_text,omitempty"`
	ImageID string `json:"image_id,omitempty"`
	// グループ
	Children []Shape `json:"children,omitempty"`
}

// Position は位置とサイズ（pt単位）
type Position struct {
	X  float64 `json:"x"`
	Y  float64 `json:"y"`
	W  float64 `json:"w"`
	H  float64 `json:"h"`
}

// Point は座標（pt単位）
type Point struct {
	X float64 `json:"x"`
	Y float64 `json:"y"`
}

// LineStyle は枠線情報
type LineStyle struct {
	Color string  `json:"color,omitempty"`
	Style string  `json:"style,omitempty"`
	Width float64 `json:"width,omitempty"`
}

// Paragraph は段落
type Paragraph struct {
	Text      string         `json:"text"`
	Bullet    string         `json:"bullet,omitempty"`
	Level     int            `json:"level,omitempty"`
	MarginL   *float64       `json:"margin_left,omitempty"`
	Indent    *float64       `json:"indent,omitempty"`
	Font      *FontStyle     `json:"font,omitempty"`
	StyleRef  int            `json:"s,omitempty"`
	Alignment *Alignment     `json:"alignment,omitempty"`
	Link      *HyperlinkData `json:"link,omitempty"`
	RichText  []RichTextRun  `json:"rich_text,omitempty"`
}

// RichTextRun はリッチテキストラン
type RichTextRun struct {
	Text     string         `json:"text"`
	Font     *FontStyle     `json:"font,omitempty"`
	StyleRef int            `json:"s,omitempty"`
	Link     *HyperlinkData `json:"link,omitempty"`
}

// HyperlinkData はハイパーリンク情報
type HyperlinkData struct {
	URL   string `json:"url,omitempty"`   // 外部URL（http://, https://, mailto: 等）
	Slide int    `json:"slide,omitempty"` // スライド内リンクのスライド番号
}

// FontStyle はフォント情報
type FontStyle struct {
	Name          string `json:"name,omitempty"`
	Size          float64 `json:"size,omitempty"`
	Bold          bool   `json:"bold,omitempty"`
	Italic        bool   `json:"italic,omitempty"`
	Strikethrough bool   `json:"strikethrough,omitempty"`
	Underline     string `json:"underline,omitempty"`
	Color         string `json:"color,omitempty"`
}

// Alignment は配置情報
type Alignment struct {
	Horizontal string `json:"horizontal,omitempty"`
	Vertical   string `json:"vertical,omitempty"`
}

// TextMargin はテキストボディの内部マージン（pt単位）
type TextMargin struct {
	Left   *float64 `json:"left,omitempty"`
	Right  *float64 `json:"right,omitempty"`
	Top    *float64 `json:"top,omitempty"`
	Bottom *float64 `json:"bottom,omitempty"`
}

// TableData はテーブルデータ
type TableData struct {
	Cols int            `json:"cols"`
	Rows [][]*TableCell `json:"rows"`
}

// TableCell はテーブルのセル
type TableCell struct {
	Text       string      `json:"text"`
	Paragraphs []Paragraph `json:"paragraphs,omitempty"`
}

