package pptx

import (
	"encoding/json"
	"fmt"
)

// StyleDef はスタイル定義（スライドレベルの重複排除用）
type StyleDef struct {
	ID int `json:"_s"`
	*FontStyle
}

// MarshalJSON は _s フィールドと FontStyle のフィールドをフラットに結合する
func (sd StyleDef) MarshalJSON() ([]byte, error) {
	m := make(map[string]any)
	m["_s"] = sd.ID
	if sd.FontStyle != nil {
		if sd.Name != "" {
			m["name"] = sd.Name
		}
		if sd.Size != 0 {
			m["size"] = sd.Size
		}
		if sd.Bold {
			m["bold"] = true
		}
		if sd.Italic {
			m["italic"] = true
		}
		if sd.Strikethrough {
			m["strikethrough"] = true
		}
		if sd.Underline != "" {
			m["underline"] = sd.Underline
		}
		if sd.Color != "" {
			m["color"] = sd.Color
		}
	}
	return json.Marshal(m)
}

// fontKey はフォントスタイルのフィールドを連結したキー文字列を返す
func fontKey(font *FontStyle) string {
	if font == nil {
		return ""
	}
	return fmt.Sprintf("%s\x00%d\x00%t\x00%t\x00%t\x00%s\x00%s",
		font.Name, font.Size, font.Bold, font.Italic, font.Strikethrough, font.Underline, font.Color)
}

// StyleDeduplicator はスライド横断でフォントスタイルの重複排除を行う。
// 複数スライドの出力で共有し、初出のスタイル定義のみを返す。
type StyleDeduplicator struct {
	styleMap map[string]int       // fontKey → styleID
	fontMap  map[string]*FontStyle // fontKey → FontStyle
	nextID   int
}

// NewStyleDeduplicator はスタイル重複排除器を生成する
func NewStyleDeduplicator() *StyleDeduplicator {
	return &StyleDeduplicator{
		styleMap: make(map[string]int),
		fontMap:  make(map[string]*FontStyle),
	}
}

// fontSlot はフォント参照のスロット（読み書き可能）
type fontSlot struct {
	font     **FontStyle
	styleRef *int
}

// Deduplicate はスライドデータ内のフォント情報を重複排除する。
// スライド内で2回以上、またはスライド横断で既出のフォントを参照IDに置き換える。
// 戻り値はこのスライドで新規に定義されたスタイルのみ。
// 元の SlideData を直接変更する。
func (sd2 *StyleDeduplicator) Deduplicate(sd *SlideData) []StyleDef {
	// 1パス目: フォント出現回数のカウントとフォントマップの収集
	counts := make(map[string]int)
	localFontMap := make(map[string]*FontStyle)
	walkFontSlots(sd.Shapes, sd.Notes, func(s fontSlot) {
		key := fontKey(*s.font)
		if key == "" {
			return
		}
		counts[key]++
		if _, ok := localFontMap[key]; !ok {
			localFontMap[key] = *s.font
		}
	})

	// スタイル定義に登録するフォントを決定:
	// - スライド内で2回以上使われるフォント
	// - スライド横断で既に登録済みのフォント（1回でも参照IDに置き換え可能）
	var newStyles []StyleDef
	replaceMap := make(map[string]int)

	for key, count := range counts {
		if id, ok := sd2.styleMap[key]; ok {
			// 既にスライド横断で登録済み → 参照IDを使う
			replaceMap[key] = id
		} else if count >= 2 {
			// スライド内で2回以上 → 新規登録
			sd2.nextID++
			id := sd2.nextID
			sd2.styleMap[key] = id
			sd2.fontMap[key] = localFontMap[key]
			replaceMap[key] = id
			newStyles = append(newStyles, StyleDef{ID: id, FontStyle: localFontMap[key]})
		}
	}

	if len(replaceMap) == 0 {
		return nil
	}

	// 2パス目: 対象フォントを参照IDに置き換え
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
