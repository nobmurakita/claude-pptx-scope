package pptx

import "testing"

func TestDeduplicate_AllFontsExtracted(t *testing.T) {
	sd := &SlideData{
		Shapes: []Shape{
			{Paragraphs: []Paragraph{
				{Text: "a", Font: &FontStyle{Name: "Arial", Size: 10}},
				{Text: "b", Font: &FontStyle{Name: "Arial", Size: 10}}, // 同じ
				{Text: "c", Font: &FontStyle{Name: "Meiryo", Size: 20}},
			}},
		},
	}

	dedup := NewStyleDeduplicator()
	styles := dedup.Deduplicate(sd)

	// すべてのフォントがスタイル定義に抽出される
	if len(styles) != 2 {
		t.Fatalf("styles: got %d, want 2", len(styles))
	}
	if styles[0].Name != "Arial" {
		t.Errorf("styles[0].Name: got %q, want %q", styles[0].Name, "Arial")
	}
	if styles[1].Name != "Meiryo" {
		t.Errorf("styles[1].Name: got %q, want %q", styles[1].Name, "Meiryo")
	}

	// すべて s 参照に置き換え
	for i, p := range sd.Shapes[0].Paragraphs {
		if p.Font != nil {
			t.Errorf("para[%d].Font should be nil", i)
		}
		if p.StyleRef == 0 {
			t.Errorf("para[%d].StyleRef should be set", i)
		}
	}

	// Arial の2つは同じID
	if sd.Shapes[0].Paragraphs[0].StyleRef != sd.Shapes[0].Paragraphs[1].StyleRef {
		t.Error("para[0] and para[1] should have the same StyleRef")
	}
	// Meiryo は別のID
	if sd.Shapes[0].Paragraphs[2].StyleRef == sd.Shapes[0].Paragraphs[0].StyleRef {
		t.Error("para[2] should have a different StyleRef from para[0]")
	}
}

func TestDeduplicate_RichText(t *testing.T) {
	sd := &SlideData{
		Shapes: []Shape{
			{Paragraphs: []Paragraph{
				{Text: "ab", RichText: []RichTextRun{
					{Text: "a", Font: &FontStyle{Name: "Arial", Bold: true}},
					{Text: "b"},
				}},
			}},
		},
	}

	dedup := NewStyleDeduplicator()
	styles := dedup.Deduplicate(sd)

	if len(styles) != 1 {
		t.Fatalf("styles: got %d, want 1", len(styles))
	}

	rt := sd.Shapes[0].Paragraphs[0].RichText[0]
	if rt.StyleRef != 1 {
		t.Errorf("richtext[0].StyleRef: got %d, want 1", rt.StyleRef)
	}
	if rt.Font != nil {
		t.Error("richtext[0].Font should be nil")
	}
}

func TestDeduplicate_NoFonts(t *testing.T) {
	sd := &SlideData{
		Shapes: []Shape{
			{Paragraphs: []Paragraph{{Text: "plain"}}},
		},
	}

	dedup := NewStyleDeduplicator()
	styles := dedup.Deduplicate(sd)
	if styles != nil {
		t.Errorf("styles should be nil, got %d", len(styles))
	}
}

func TestDeduplicate_CrossSlide(t *testing.T) {
	sd1 := &SlideData{
		Shapes: []Shape{
			{Paragraphs: []Paragraph{
				{Text: "a", Font: &FontStyle{Name: "Arial"}},
			}},
		},
	}
	sd2 := &SlideData{
		Shapes: []Shape{
			{Paragraphs: []Paragraph{
				{Text: "b", Font: &FontStyle{Name: "Arial"}}, // スライド横断で既出
				{Text: "c", Font: &FontStyle{Name: "Meiryo"}},
			}},
		},
	}

	dedup := NewStyleDeduplicator()

	// スライド1: Arial → 新規スタイル定義
	styles1 := dedup.Deduplicate(sd1)
	if len(styles1) != 1 {
		t.Fatalf("slide1 styles: got %d, want 1", len(styles1))
	}
	arialID := styles1[0].ID

	// スライド2: Arial は既出なので新規定義なし、Meiryo は新規
	styles2 := dedup.Deduplicate(sd2)
	if len(styles2) != 1 {
		t.Fatalf("slide2 styles: got %d, want 1", len(styles2))
	}
	if styles2[0].Name != "Meiryo" {
		t.Errorf("slide2 new style should be Meiryo, got %q", styles2[0].Name)
	}

	// Arial は既存IDを再利用
	if sd2.Shapes[0].Paragraphs[0].StyleRef != arialID {
		t.Errorf("slide2 Arial StyleRef: got %d, want %d", sd2.Shapes[0].Paragraphs[0].StyleRef, arialID)
	}
	if sd2.Shapes[0].Paragraphs[0].Font != nil {
		t.Error("slide2 Arial Font should be nil")
	}
}

func TestDeduplicate_Group(t *testing.T) {
	sd := &SlideData{
		Shapes: []Shape{
			{
				Type: "group",
				Children: []Shape{
					{Paragraphs: []Paragraph{{Text: "a", Font: &FontStyle{Name: "Arial"}}}},
				},
			},
			{Paragraphs: []Paragraph{{Text: "b", Font: &FontStyle{Name: "Arial"}}}},
		},
	}

	dedup := NewStyleDeduplicator()
	styles := dedup.Deduplicate(sd)

	// Arial は1種類なのでスタイル定義は1つ
	if len(styles) != 1 {
		t.Fatalf("styles: got %d, want 1", len(styles))
	}

	child := sd.Shapes[0].Children[0].Paragraphs[0]
	if child.StyleRef != 1 {
		t.Errorf("group child StyleRef: got %d, want 1", child.StyleRef)
	}
	// 両方とも同じID
	if sd.Shapes[1].Paragraphs[0].StyleRef != 1 {
		t.Errorf("second shape StyleRef: got %d, want 1", sd.Shapes[1].Paragraphs[0].StyleRef)
	}
}
