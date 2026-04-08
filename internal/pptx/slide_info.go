package pptx

import (
	"encoding/xml"
	"fmt"
)

// LoadSlideInfos は全スライドの概要情報を取得する（info コマンド用）
func (f *File) LoadSlideInfos() ([]SlideInfo, error) {
	infos := make([]SlideInfo, 0, len(f.slideEntries))

	for i, entry := range f.slideEntries {
		info := SlideInfo{Slide: i + 1}

		// スライドXMLを読み込み
		data, err := readZipFile(f.zi, entry.Path)
		if err != nil {
			return nil, fmt.Errorf("スライド %d の読み込みに失敗: %w", i+1, err)
		}
		if data == nil {
			continue
		}

		var sld xmlSlide
		if err := xml.Unmarshal(data, &sld); err != nil {
			return nil, fmt.Errorf("スライド %d のパースに失敗: %w", i+1, err)
		}

		// 非表示判定
		if sld.Show == "0" {
			info.Hidden = true
		}

		// タイトル取得
		info.Title = extractTitle(sld.CSld.SpTree.Children)

		// ノート有無チェック
		info.HasNotes = f.hasNotes(i)

		// 画像有無チェック
		info.HasImages = hasImages(sld.CSld.SpTree.Children)

		infos = append(infos, info)
	}

	return infos, nil
}

// hasImages は子要素内に画像が存在するかチェックする（グループ内も再帰的に確認）
func hasImages(children []xmlSpTreeChild) bool {
	for _, child := range children {
		if child.Pic != nil {
			return true
		}
		if child.GrpSp != nil && hasImages(child.GrpSp.Children) {
			return true
		}
	}
	return false
}
