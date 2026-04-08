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

// Deduplicate はスライドデータ内のフォント情報を重複排除する。
// スライド内で2回以上、またはスライド横断で既出のフォントを参照IDに置き換える。
// 戻り値はこのスライドで新規に定義されたスタイルのみ。
// 元の SlideData を直接変更する。
func (sd2 *StyleDeduplicator) Deduplicate(sd *SlideData) []StyleDef {
	// 1パス目: スライド内のフォント出現回数をカウント
	counts := make(map[string]int)
	countShapes(counts, sd.Shapes)
	countParas(counts, sd.Notes)

	// スタイル定義に登録するフォントを決定:
	// - スライド内で2回以上使われるフォント
	// - スライド横断で既に登録済みのフォント（1回でも参照IDに置き換え可能）
	localFontMap := make(map[string]*FontStyle)
	collectFontMap(localFontMap, sd.Shapes)
	collectFontMapParas(localFontMap, sd.Notes)

	var newStyles []StyleDef
	replaceMap := make(map[string]int) // このスライドで置き換え対象のフォント

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
	replaceShapes(replaceMap, sd.Shapes)
	replaceParas(replaceMap, sd.Notes)

	return newStyles
}

// --- 1パス目: カウント ---

func countShapes(counts map[string]int, shapes []Shape) {
	for _, s := range shapes {
		countParas(counts, s.Paragraphs)
		if s.Table != nil {
			countTableCells(counts, s.Table)
		}
		if s.Type == "group" {
			countShapes(counts, s.Children)
		}
	}
}

func countTableCells(counts map[string]int, table *TableData) {
	for _, row := range table.Rows {
		for _, cell := range row {
			if cell != nil {
				countParas(counts, cell.Paragraphs)
			}
		}
	}
}

func countParas(counts map[string]int, paras []Paragraph) {
	for _, p := range paras {
		if key := fontKey(p.Font); key != "" {
			counts[key]++
		}
		for _, rt := range p.RichText {
			if key := fontKey(rt.Font); key != "" {
				counts[key]++
			}
		}
	}
}

// --- FontStyle マッピング収集 ---

func collectFontMap(fontMap map[string]*FontStyle, shapes []Shape) {
	for _, s := range shapes {
		collectFontMapParas(fontMap, s.Paragraphs)
		if s.Table != nil {
			collectFontMapTable(fontMap, s.Table)
		}
		if s.Type == "group" {
			collectFontMap(fontMap, s.Children)
		}
	}
}

func collectFontMapTable(fontMap map[string]*FontStyle, table *TableData) {
	for _, row := range table.Rows {
		for _, cell := range row {
			if cell != nil {
				collectFontMapParas(fontMap, cell.Paragraphs)
			}
		}
	}
}

func collectFontMapParas(fontMap map[string]*FontStyle, paras []Paragraph) {
	for _, p := range paras {
		if p.Font != nil {
			key := fontKey(p.Font)
			if _, ok := fontMap[key]; !ok {
				fontMap[key] = p.Font
			}
		}
		for _, rt := range p.RichText {
			if rt.Font != nil {
				key := fontKey(rt.Font)
				if _, ok := fontMap[key]; !ok {
					fontMap[key] = rt.Font
				}
			}
		}
	}
}

// --- 2パス目: 置き換え ---

func replaceShapes(styleMap map[string]int, shapes []Shape) {
	for i := range shapes {
		replaceParas(styleMap, shapes[i].Paragraphs)
		if shapes[i].Table != nil {
			replaceTableCells(styleMap, shapes[i].Table)
		}
		if shapes[i].Type == "group" {
			replaceShapes(styleMap, shapes[i].Children)
		}
	}
}

func replaceTableCells(styleMap map[string]int, table *TableData) {
	for _, row := range table.Rows {
		for _, cell := range row {
			if cell != nil {
				replaceParas(styleMap, cell.Paragraphs)
			}
		}
	}
}

func replaceParas(styleMap map[string]int, paras []Paragraph) {
	for i := range paras {
		if paras[i].Font != nil {
			if id, ok := styleMap[fontKey(paras[i].Font)]; ok {
				paras[i].StyleRef = id
				paras[i].Font = nil
			}
		}
		for j := range paras[i].RichText {
			if paras[i].RichText[j].Font != nil {
				if id, ok := styleMap[fontKey(paras[i].RichText[j].Font)]; ok {
					paras[i].RichText[j].StyleRef = id
					paras[i].RichText[j].Font = nil
				}
			}
		}
	}
}
