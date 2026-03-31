package pptx

import (
	"testing"
)

func TestFilterShapesByText(t *testing.T) {
	shapes := []Shape{
		{ID: 1, Type: "rect", Paragraphs: []Paragraph{
			{Text: "Hello World"},
			{Text: "Goodbye"},
		}},
		{ID: 2, Type: "rect", Paragraphs: []Paragraph{
			{Text: "Nothing here"},
		}},
	}

	matched := filterShapesByText(shapes, "hello")
	if len(matched) != 1 {
		t.Fatalf("got %d matched shapes, want 1", len(matched))
	}
	if matched[0].ID != 1 {
		t.Errorf("matched shape ID: got %d, want 1", matched[0].ID)
	}
	// マッチした段落のみ含まれる
	if len(matched[0].Paragraphs) != 1 {
		t.Errorf("matched paragraphs: got %d, want 1", len(matched[0].Paragraphs))
	}
}

func TestFilterShapesByText_Table(t *testing.T) {
	cell := func(s string) *TableCell { return &TableCell{Text: s} }
	shapes := []Shape{
		{
			ID:   1,
			Type: "table",
			Table: &TableData{
				Cols: 2,
				Rows: [][]*TableCell{
					{cell("Alpha"), cell("Beta")},
					{cell("Gamma"), nil},
				},
			},
		},
	}

	matched := filterShapesByText(shapes, "beta")
	if len(matched) != 1 {
		t.Fatalf("got %d matched shapes, want 1", len(matched))
	}

	// マッチしないテーブル
	matched = filterShapesByText(shapes, "delta")
	if len(matched) != 0 {
		t.Errorf("got %d matched shapes, want 0", len(matched))
	}
}

func TestFilterShapesByText_Connector(t *testing.T) {
	shapes := []Shape{
		{ID: 1, Type: "connector", Label: "接続ラベル"},
		{ID: 2, Type: "connector", Label: "別のラベル"},
	}

	matched := filterShapesByText(shapes, "接続")
	if len(matched) != 1 {
		t.Fatalf("got %d matched shapes, want 1", len(matched))
	}
	if matched[0].ID != 1 {
		t.Errorf("matched shape ID: got %d, want 1", matched[0].ID)
	}
}

func TestFilterShapesByText_Group(t *testing.T) {
	shapes := []Shape{
		{
			ID:   1,
			Type: "group",
			Children: []Shape{
				{ID: 2, Type: "rect", Paragraphs: []Paragraph{{Text: "内部テキスト"}}},
				{ID: 3, Type: "rect", Paragraphs: []Paragraph{{Text: "別のテキスト"}}},
			},
		},
	}

	matched := filterShapesByText(shapes, "内部")
	if len(matched) != 1 {
		t.Fatalf("got %d matched shapes, want 1", len(matched))
	}
	if len(matched[0].Children) != 1 {
		t.Errorf("matched children: got %d, want 1", len(matched[0].Children))
	}
}

func TestFilterShapesByText_NoMatch(t *testing.T) {
	shapes := []Shape{
		{ID: 1, Type: "rect", Paragraphs: []Paragraph{{Text: "テスト"}}},
	}

	matched := filterShapesByText(shapes, "存在しない")
	if len(matched) != 0 {
		t.Errorf("got %d matched shapes, want 0", len(matched))
	}
}

func TestFilterParagraphsByText(t *testing.T) {
	paras := []Paragraph{
		{Text: "最初の段落"},
		{Text: "二番目の段落"},
		{Text: "最初に戻る"},
	}

	matched := filterParagraphsByText(paras, "最初")
	if len(matched) != 2 {
		t.Fatalf("got %d matched paragraphs, want 2", len(matched))
	}
}

func TestFilterParagraphsByText_Empty(t *testing.T) {
	matched := filterParagraphsByText(nil, "test")
	if len(matched) != 0 {
		t.Errorf("got %d matched paragraphs, want 0", len(matched))
	}
}
