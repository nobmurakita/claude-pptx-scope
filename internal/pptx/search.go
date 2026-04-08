package pptx

import (
	"strings"
)

// Search はプレゼンテーション内のテキストを検索し、マッチしたスライドの情報を返す
func (f *File) Search(query string, slideNums []int, includeNotes bool) ([]SlideInfo, error) {
	queryLower := strings.ToLower(query)

	targets, err := f.ResolveSlideNums(slideNums)
	if err != nil {
		return nil, err
	}

	var results []SlideInfo

	for _, num := range targets {
		sd, err := f.LoadSlide(num, includeNotes)
		if err != nil {
			return nil, err
		}

		if matchSlideText(sd, queryLower) {
			results = append(results, sd.Info())
		}
	}

	return results, nil
}

// matchSlideText はスライド内のテキストがクエリにマッチするか判定する
func matchSlideText(sd *SlideData, queryLower string) bool {
	if matchShapesText(sd.Shapes, queryLower) {
		return true
	}
	for _, p := range sd.Notes {
		if strings.Contains(strings.ToLower(p.Text), queryLower) {
			return true
		}
	}
	return false
}

// matchShapesText は図形群のテキストがクエリにマッチするか判定する
func matchShapesText(shapes []Shape, queryLower string) bool {
	for _, s := range shapes {
		switch {
		case s.Type == "table" && s.Table != nil:
			for _, row := range s.Table.Rows {
				for _, cell := range row {
					if cell != nil && strings.Contains(strings.ToLower(cell.Text), queryLower) {
						return true
					}
				}
			}
		case s.Type == "connector":
			if s.Label != "" && strings.Contains(strings.ToLower(s.Label), queryLower) {
				return true
			}
		case s.Type == "group":
			if matchShapesText(s.Children, queryLower) {
				return true
			}
		default:
			for _, p := range s.Paragraphs {
				if strings.Contains(strings.ToLower(p.Text), queryLower) {
					return true
				}
			}
		}
	}
	return false
}
