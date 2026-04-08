package pptx

import "testing"

func TestDeduplicate_MultipleUsage(t *testing.T) {
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

	// Arial は2回使われるのでスタイル定義に抽出
	if len(styles) != 1 {
		t.Fatalf("styles: got %d, want 1", len(styles))
	}
	if styles[0].Name != "Arial" {
		t.Errorf("styles[0].Name: got %q, want %q", styles[0].Name, "Arial")
	}

	// s 参照に置き換え
	if sd.Shapes[0].Paragraphs[0].StyleRef != 1 {
		t.Errorf("para[0].StyleRef: got %d, want 1", sd.Shapes[0].Paragraphs[0].StyleRef)
	}
	if sd.Shapes[0].Paragraphs[0].Font != nil {
		t.Error("para[0].Font should be nil")
	}

	// Meiryo は1回のみなのでインラインのまま
	if sd.Shapes[0].Paragraphs[2].StyleRef != 0 {
		t.Errorf("para[2].StyleRef: got %d, want 0", sd.Shapes[0].Paragraphs[2].StyleRef)
	}
	if sd.Shapes[0].Paragraphs[2].Font == nil {
		t.Error("para[2].Font should remain inline")
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
				{Text: "c", Font: &FontStyle{Name: "Arial", Bold: true}},
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

func TestDeduplicate_SingleUsageStaysInline(t *testing.T) {
	sd := &SlideData{
		Shapes: []Shape{
			{Paragraphs: []Paragraph{
				{Text: "a", Font: &FontStyle{Name: "A"}},
				{Text: "b", Font: &FontStyle{Name: "B"}},
			}},
		},
	}

	dedup := NewStyleDeduplicator()
	styles := dedup.Deduplicate(sd)
	if styles != nil {
		t.Errorf("all single-use should stay inline, got %d styles", len(styles))
	}
	for i, p := range sd.Shapes[0].Paragraphs {
		if p.Font == nil {
			t.Errorf("para[%d].Font should remain inline", i)
		}
	}
}

func TestDeduplicate_CrossSlide(t *testing.T) {
	sd1 := &SlideData{
		Shapes: []Shape{
			{Paragraphs: []Paragraph{
				{Text: "a", Font: &FontStyle{Name: "Arial"}},
				{Text: "b", Font: &FontStyle{Name: "Arial"}},
			}},
		},
	}
	sd2 := &SlideData{
		Shapes: []Shape{
			{Paragraphs: []Paragraph{
				{Text: "c", Font: &FontStyle{Name: "Arial"}}, // スライド横断で既出
			}},
		},
	}

	dedup := NewStyleDeduplicator()

	// スライド1: Arial が2回 → 新規スタイル定義
	styles1 := dedup.Deduplicate(sd1)
	if len(styles1) != 1 {
		t.Fatalf("slide1 styles: got %d, want 1", len(styles1))
	}
	arialID := styles1[0].ID

	// スライド2: Arial が1回だがスライド横断で既出 → 新規定義なし、参照IDを使う
	styles2 := dedup.Deduplicate(sd2)
	if styles2 != nil {
		t.Errorf("slide2 should not produce new styles, got %d", len(styles2))
	}
	if sd2.Shapes[0].Paragraphs[0].StyleRef != arialID {
		t.Errorf("slide2 para StyleRef: got %d, want %d", sd2.Shapes[0].Paragraphs[0].StyleRef, arialID)
	}
	if sd2.Shapes[0].Paragraphs[0].Font != nil {
		t.Error("slide2 para Font should be nil")
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
	if len(styles) != 1 {
		t.Fatalf("styles: got %d, want 1", len(styles))
	}

	child := sd.Shapes[0].Children[0].Paragraphs[0]
	if child.StyleRef != 1 {
		t.Errorf("group child StyleRef: got %d, want 1", child.StyleRef)
	}
}
