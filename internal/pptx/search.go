package pptx

import (
	"strings"
)

// Search はプレゼンテーション内のテキストを検索する
func (f *File) Search(query string, slideNums []int, includeNotes bool) ([]SlideData, error) {
	queryLower := strings.ToLower(query)

	targets, err := f.ResolveSlideNums(slideNums)
	if err != nil {
		return nil, err
	}

	var results []SlideData

	for _, num := range targets {
		sd, err := f.LoadSlide(num, includeNotes)
		if err != nil {
			return nil, err
		}

		matchedShapes := filterShapesByText(sd.Shapes, queryLower)
		matchedNotes := filterParagraphsByText(sd.Notes, queryLower)

		if len(matchedShapes) == 0 && len(matchedNotes) == 0 {
			continue
		}

		result := SlideData{
			Number: sd.Number,
			Title:  sd.Title,
			Shapes: matchedShapes,
			Notes:  matchedNotes,
		}
		results = append(results, result)
	}

	return results, nil
}

// filterShapesByText はテキストにマッチする図形のみを抽出する
func filterShapesByText(shapes []Shape, queryLower string) []Shape {
	var matched []Shape
	for _, s := range shapes {
		ms := matchShape(s, queryLower)
		if ms != nil {
			matched = append(matched, *ms)
		}
	}
	return matched
}

// matchShape は図形がクエリにマッチするか判定し、マッチした部分のみを含む図形を返す
func matchShape(s Shape, queryLower string) *Shape {
	switch {
	case s.Type == "table" && s.Table != nil:
		// テーブル: いずれかのセルにマッチすればテーブル全体を返す
		for _, row := range s.Table.Rows {
			for _, cell := range row {
				if cell != nil && strings.Contains(strings.ToLower(*cell), queryLower) {
					return &s
				}
			}
		}
		return nil

	case s.Type == "connector":
		// コネクタ: ラベルにマッチ
		if s.Label != "" && strings.Contains(strings.ToLower(s.Label), queryLower) {
			return &s
		}
		return nil

	case s.Type == "group":
		// グループ: 子要素で再帰検索
		matchedChildren := filterShapesByText(s.Children, queryLower)
		if len(matchedChildren) > 0 {
			result := s
			result.Children = matchedChildren
			return &result
		}
		return nil

	default:
		// 通常の図形: 段落のテキストにマッチ
		var matchedParas []Paragraph
		for _, p := range s.Paragraphs {
			if strings.Contains(strings.ToLower(p.Text), queryLower) {
				matchedParas = append(matchedParas, p)
			}
		}
		if len(matchedParas) == 0 {
			return nil
		}
		result := s
		result.Paragraphs = matchedParas
		return &result
	}
}

// filterParagraphsByText はテキストにマッチする段落のみを抽出する
func filterParagraphsByText(paras []Paragraph, queryLower string) []Paragraph {
	var matched []Paragraph
	for _, p := range paras {
		if strings.Contains(strings.ToLower(p.Text), queryLower) {
			matched = append(matched, p)
		}
	}
	return matched
}
