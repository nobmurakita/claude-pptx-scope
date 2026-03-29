package pptx

// 出力用のデータ型

// SlideInfo は info コマンドの出力用
type SlideInfo struct {
	Number    int    `json:"number"`
	Title     string `json:"title,omitempty"`
	HasNotes  bool   `json:"has_notes,omitempty"`
	HasImages bool   `json:"has_images,omitempty"`
	Hidden    bool   `json:"hidden,omitempty"`
}

// Shape は図形の出力データ
type Shape struct {
	ID             int         `json:"id"`
	Type           string      `json:"type"`
	Name           string      `json:"name,omitempty"`
	Placeholder    string      `json:"placeholder,omitempty"`
	Pos            *Position   `json:"pos,omitempty"`
	Z              int         `json:"z"`
	Rotation       float64     `json:"rotation,omitempty"`
	Flip           string      `json:"flip,omitempty"`
	Fill           string      `json:"fill,omitempty"`
	Line           *LineStyle  `json:"line,omitempty"`
	CalloutPointer *Point      `json:"callout_pointer,omitempty"`
	Paragraphs     []Paragraph `json:"paragraphs,omitempty"`
	Table          *TableData  `json:"table,omitempty"`
	Font           *FontStyle  `json:"font,omitempty"`
	Alignment      *Alignment  `json:"alignment,omitempty"`
	// コネクタ
	From          int    `json:"from,omitempty"`
	To            int    `json:"to,omitempty"`
	ConnectorType string `json:"connector_type,omitempty"`
	Arrow         string `json:"arrow,omitempty"`
	Label         string `json:"label,omitempty"`
	// 画像
	AltText   string `json:"alt_text,omitempty"`
	ImagePath string `json:"image_path,omitempty"`
	// グループ
	Children []Shape `json:"children,omitempty"`
}

// Position は位置とサイズ
type Position struct {
	X  int64 `json:"x"`
	Y  int64 `json:"y"`
	W  int64 `json:"w"`
	H  int64 `json:"h"`
}

// Point は座標
type Point struct {
	X int64 `json:"x"`
	Y int64 `json:"y"`
}

// LineStyle は枠線情報
type LineStyle struct {
	Color string  `json:"color,omitempty"`
	Style string  `json:"style,omitempty"`
	Width float64 `json:"width,omitempty"`
}

// Paragraph は段落
type Paragraph struct {
	Text      string        `json:"text"`
	Bullet    string        `json:"bullet,omitempty"`
	Level     int           `json:"level,omitempty"`
	Font      *FontStyle    `json:"font,omitempty"`
	Alignment *Alignment    `json:"alignment,omitempty"`
	RichText  []RichTextRun `json:"rich_text,omitempty"`
}

// RichTextRun はリッチテキストラン
type RichTextRun struct {
	Text string     `json:"text"`
	Font *FontStyle `json:"font,omitempty"`
}

// FontStyle はフォント情報
type FontStyle struct {
	Name          string  `json:"name,omitempty"`
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

// TableData はテーブルデータ
type TableData struct {
	Cols int          `json:"cols"`
	Rows [][]*string  `json:"rows"`
}

