package pptx

import (
	"strings"
)

// Search はプレゼンテーション内のテキストを検索し、マッチしたスライドの情報を返す。
// 図形のフルパースを行わず、XMLから直接テキストを抽出して軽量に検索する。
func (f *File) Search(query string, slideNums []int, includeNotes bool) ([]SlideInfo, error) {
	queryLower := strings.ToLower(query)

	targets, err := f.ResolveSlideNums(slideNums)
	if err != nil {
		return nil, err
	}

	var results []SlideInfo

	for _, num := range targets {
		idx := num - 1
		sld, err := f.loadSlideXML(idx)
		if err != nil {
			return nil, err
		}
		if sld == nil {
			continue
		}

		matched := matchSpTreeText(sld.CSld.SpTree.Children, queryLower)

		if !matched && includeNotes {
			txBody := f.findNotesBody(idx)
			if txBody != nil {
				text := extractTextFromTxBody(txBody)
				if strings.Contains(strings.ToLower(text), queryLower) {
					matched = true
				}
			}
		}

		if matched {
			results = append(results, SlideInfo{
				Slide:    num,
				Title:    extractTitle(sld.CSld.SpTree.Children),
				HasNotes: f.hasNotes(idx),
				Hidden:   sld.Show == "0",
			})
		}
	}

	return results, nil
}

// matchSpTreeText は spTree の子要素のテキストがクエリにマッチするか判定する
func matchSpTreeText(children []xmlSpTreeChild, queryLower string) bool {
	for _, child := range children {
		switch {
		case child.Sp != nil:
			if matchTxBodyText(child.Sp.TxBody, queryLower) {
				return true
			}
		case child.CxnSp != nil:
			if matchTxBodyText(child.CxnSp.TxBody, queryLower) {
				return true
			}
		case child.GrpSp != nil:
			if matchSpTreeText(child.GrpSp.Children, queryLower) {
				return true
			}
		case child.GraphicFrame != nil:
			tbl := child.GraphicFrame.Graphic.GraphicData.Tbl
			if tbl != nil {
				for _, tr := range tbl.Trs {
					for _, tc := range tr.Tcs {
						if matchTxBodyText(tc.TxBody, queryLower) {
							return true
						}
					}
				}
			}
		}
	}
	return false
}

// matchTxBodyText は txBody 内のテキストがクエリにマッチするか判定する
func matchTxBodyText(txBody *xmlTxBody, queryLower string) bool {
	if txBody == nil {
		return false
	}
	text := extractTextFromTxBody(txBody)
	return strings.Contains(strings.ToLower(text), queryLower)
}
