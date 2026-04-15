package pptx

// StyleDef はスタイル定義行の出力用。
// style フィールドを先頭に、FontStyle の各フィールドをフラットに並べる。
type StyleDef struct {
	ID            int     `json:"style"`
	Name          string  `json:"name,omitempty"`
	Size          float64 `json:"size,omitempty"`
	Bold          bool    `json:"bold,omitempty"`
	Italic        bool    `json:"italic,omitempty"`
	Strikethrough bool    `json:"strikethrough,omitempty"`
	Underline     string  `json:"underline,omitempty"`
	Color         string  `json:"color,omitempty"`
	Highlight     string  `json:"highlight,omitempty"`
	Baseline      string  `json:"baseline,omitempty"`
	Cap           string  `json:"cap,omitempty"`
}

// StyleDeduplicator はスライド横断でフォントスタイルの重複排除を行う。
// 複数スライドの出力で共有し、初出のスタイル定義のみを返す。
type StyleDeduplicator struct {
	// FontStyle は全フィールドがスカラーなので comparable。値そのものをキーにする。
	styleMap map[FontStyle]int
	nextID   int
}

// NewStyleDeduplicator はスタイル重複排除器を生成する
func NewStyleDeduplicator() *StyleDeduplicator {
	return &StyleDeduplicator{
		styleMap: make(map[FontStyle]int),
	}
}

// Deduplicate はスライドデータ内のフォント情報を重複排除する。
// すべてのフォントをスタイル定義に抽出し、参照IDに置き換える。
// スライド横断で既出のフォントは既存IDを再利用する。
// 戻り値はこのスライドで新規に定義されたスタイル（個別行として出力する）。
// 元の SlideData を直接変更する。
func (sd2 *StyleDeduplicator) Deduplicate(sd *SlideData) []StyleDef {
	var newStyles []StyleDef
	replaceMap := make(map[FontStyle]int)

	assign := func(font **FontStyle, styleRef *int) {
		if *font == nil {
			return
		}
		key := **font
		id, ok := replaceMap[key]
		if !ok {
			if existing, found := sd2.styleMap[key]; found {
				id = existing
			} else {
				sd2.nextID++
				id = sd2.nextID
				sd2.styleMap[key] = id
				newStyles = append(newStyles, StyleDef{
					ID: id, Name: key.Name, Size: key.Size,
					Bold: key.Bold, Italic: key.Italic, Strikethrough: key.Strikethrough,
					Underline: key.Underline, Color: key.Color, Highlight: key.Highlight,
					Baseline: key.Baseline, Cap: key.Cap,
				})
			}
			replaceMap[key] = id
		}
		*styleRef = id
		*font = nil
	}

	WalkSlideParagraphs(sd.Shapes, sd.Notes, func(p *Paragraph) {
		assign(&p.Font, &p.StyleRef)
		for i := range p.RichText {
			assign(&p.RichText[i].Font, &p.RichText[i].StyleRef)
		}
	})

	return newStyles
}
