package pptx

import (
	"fmt"
)

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

// fontKey はフォントスタイルのフィールドを連結したキー文字列を返す
func fontKey(font *FontStyle) string {
	if font == nil {
		return ""
	}
	return fmt.Sprintf("%s\x00%g\x00%t\x00%t\x00%t\x00%s\x00%s\x00%s\x00%s\x00%s",
		font.Name, font.Size, font.Bold, font.Italic, font.Strikethrough, font.Underline, font.Color, font.Highlight, font.Baseline, font.Cap)
}

// StyleDeduplicator はスライド横断でフォントスタイルの重複排除を行う。
// 複数スライドの出力で共有し、初出のスタイル定義のみを返す。
type StyleDeduplicator struct {
	styleMap map[string]int // fontKey → styleID
	nextID   int
}

// NewStyleDeduplicator はスタイル重複排除器を生成する
func NewStyleDeduplicator() *StyleDeduplicator {
	return &StyleDeduplicator{
		styleMap: make(map[string]int),
	}
}

// fontSlot はフォント参照のスロット（読み書き可能）
type fontSlot struct {
	font     **FontStyle
	styleRef *int
}

// Deduplicate はスライドデータ内のフォント情報を重複排除する。
// すべてのフォントをスタイル定義に抽出し、参照IDに置き換える。
// スライド横断で既出のフォントは既存IDを再利用する。
// 戻り値はこのスライドで新規に定義されたスタイル（個別行として出力する）。
// 元の SlideData を直接変更する。
func (sd2 *StyleDeduplicator) Deduplicate(sd *SlideData) []StyleDef {
	// 1パス目: フォントを収集し、新規スタイルを登録
	var newStyles []StyleDef
	replaceMap := make(map[string]int)

	walkFontSlots(sd.Shapes, sd.Notes, func(s fontSlot) {
		key := fontKey(*s.font)
		if key == "" {
			return
		}
		if _, ok := replaceMap[key]; ok {
			return
		}
		if id, ok := sd2.styleMap[key]; ok {
			// スライド横断で既出 → 既存IDを再利用
			replaceMap[key] = id
		} else {
			// 新規スタイル
			sd2.nextID++
			id := sd2.nextID
			sd2.styleMap[key] = id
			replaceMap[key] = id
			f := *s.font
			newStyles = append(newStyles, StyleDef{
				ID:            id,
				Name:          f.Name,
				Size:          f.Size,
				Bold:          f.Bold,
				Italic:        f.Italic,
				Strikethrough: f.Strikethrough,
				Underline:     f.Underline,
				Color:         f.Color,
				Highlight:     f.Highlight,
				Baseline:      f.Baseline,
				Cap:           f.Cap,
			})
		}
	})

	if len(replaceMap) == 0 {
		return nil
	}

	// 2パス目: フォントを参照IDに置き換え
	walkFontSlots(sd.Shapes, sd.Notes, func(s fontSlot) {
		if id, ok := replaceMap[fontKey(*s.font)]; ok {
			*s.styleRef = id
			*s.font = nil
		}
	})

	return newStyles
}

// walkFontSlots はスライドデータ内の全フォントスロットを走査する。
// fn は font が非 nil のスロットに対してのみ呼び出される。
func walkFontSlots(shapes []Shape, notes []Paragraph, fn func(fontSlot)) {
	walkShapeFontSlots(shapes, fn)
	walkParaFontSlots(notes, fn)
}

// walkShapeFontSlots は図形ツリー内のフォントスロットを再帰的に走査する
func walkShapeFontSlots(shapes []Shape, fn func(fontSlot)) {
	for i := range shapes {
		walkParaFontSlots(shapes[i].Paragraphs, fn)
		if shapes[i].Table != nil {
			for _, row := range shapes[i].Table.Rows {
				for _, cell := range row {
					if cell != nil {
						walkParaFontSlots(cell.Paragraphs, fn)
					}
				}
			}
		}
		if shapes[i].Type == "group" {
			walkShapeFontSlots(shapes[i].Children, fn)
		}
	}
}

// walkParaFontSlots は段落内のフォントスロットを走査する
func walkParaFontSlots(paras []Paragraph, fn func(fontSlot)) {
	for i := range paras {
		if paras[i].Font != nil {
			fn(fontSlot{font: &paras[i].Font, styleRef: &paras[i].StyleRef})
		}
		for j := range paras[i].RichText {
			if paras[i].RichText[j].Font != nil {
				fn(fontSlot{font: &paras[i].RichText[j].Font, styleRef: &paras[i].RichText[j].StyleRef})
			}
		}
	}
}
