package pptx

import (
	"testing"
)

func TestMatchShapesText(t *testing.T) {
	shapes := []Shape{
		{ID: 1, Type: "rect", Paragraphs: []Paragraph{
			{Text: "Hello World"},
			{Text: "Goodbye"},
		}},
		{ID: 2, Type: "rect", Paragraphs: []Paragraph{
			{Text: "Nothing here"},
		}},
	}

	if !matchShapesText(shapes, "hello") {
		t.Error("expected match for 'hello'")
	}
	if matchShapesText(shapes, "missing") {
		t.Error("expected no match for 'missing'")
	}
}

func TestMatchShapesText_Table(t *testing.T) {
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

	if !matchShapesText(shapes, "beta") {
		t.Error("expected match for 'beta'")
	}
	if matchShapesText(shapes, "delta") {
		t.Error("expected no match for 'delta'")
	}
}

func TestMatchShapesText_Connector(t *testing.T) {
	shapes := []Shape{
		{ID: 1, Type: "connector", Label: "接続ラベル"},
		{ID: 2, Type: "connector", Label: "別のラベル"},
	}

	if !matchShapesText(shapes, "接続") {
		t.Error("expected match for '接続'")
	}
	if matchShapesText(shapes, "存在しない") {
		t.Error("expected no match for '存在しない'")
	}
}

func TestMatchShapesText_Group(t *testing.T) {
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

	if !matchShapesText(shapes, "内部") {
		t.Error("expected match for '内部'")
	}
	if matchShapesText(shapes, "外部") {
		t.Error("expected no match for '外部'")
	}
}

func TestMatchSlideText_Notes(t *testing.T) {
	sd := &SlideData{
		Shapes: []Shape{
			{ID: 1, Type: "rect", Paragraphs: []Paragraph{{Text: "本文"}}},
		},
		Notes: []Paragraph{
			{Text: "ノートのテキスト"},
		},
	}

	if !matchSlideText(sd, "ノート") {
		t.Error("expected match for 'ノート' in notes")
	}
	if !matchSlideText(sd, "本文") {
		t.Error("expected match for '本文' in shapes")
	}
	if matchSlideText(sd, "存在しない") {
		t.Error("expected no match for '存在しない'")
	}
}
