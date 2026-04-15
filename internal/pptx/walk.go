package pptx

// WalkSlideParagraphs はスライド内の全段落（図形・テーブルセル・ノート含む）を
// 出現順に走査し、各段落のポインタを fn に渡す。fn は段落を変更してよい。
func WalkSlideParagraphs(shapes []Shape, notes []Paragraph, fn func(*Paragraph)) {
	walkShapeParagraphs(shapes, fn)
	for i := range notes {
		fn(&notes[i])
	}
}

func walkShapeParagraphs(shapes []Shape, fn func(*Paragraph)) {
	for i := range shapes {
		for j := range shapes[i].Paragraphs {
			fn(&shapes[i].Paragraphs[j])
		}
		if shapes[i].Table != nil {
			for _, row := range shapes[i].Table.Rows {
				for _, cell := range row {
					if cell != nil {
						for k := range cell.Paragraphs {
							fn(&cell.Paragraphs[k])
						}
					}
				}
			}
		}
		if len(shapes[i].Children) > 0 {
			walkShapeParagraphs(shapes[i].Children, fn)
		}
	}
}
