package pptx

import (
	"encoding/xml"
	"fmt"
	"strings"
)

// LoadSlideInfos は全スライドの概要情報を取得する（info コマンド用）
func (f *File) LoadSlideInfos() ([]SlideInfo, error) {
	infos := make([]SlideInfo, 0, len(f.slideEntries))

	for i, entry := range f.slideEntries {
		info := SlideInfo{Number: i + 1}

		// スライドXMLを読み込み
		data, err := readZipFile(f.zr, entry.Path)
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
		info.Title = extractTitle(sld.CSld.SpTree)

		// ノート有無チェック
		info.HasNotes = f.hasNotes(i)

		infos = append(infos, info)
	}

	return infos, nil
}

// extractTitle は spTree からタイトルテキストを取得する
func extractTitle(spTree xmlSpTree) string {
	for _, sp := range spTree.Shapes {
		ph := sp.NvSpPr.NvPr.Ph
		if ph == nil {
			continue
		}
		if ph.Type == "title" || ph.Type == "ctrTitle" {
			return extractTextFromTxBody(sp.TxBody)
		}
	}
	return ""
}

// extractTextFromTxBody は txBody から全テキストを結合して返す
func extractTextFromTxBody(txBody *xmlTxBody) string {
	if txBody == nil {
		return ""
	}
	var parts []string
	for _, p := range txBody.Ps {
		text := extractParagraphText(p)
		if text != "" {
			parts = append(parts, text)
		}
	}
	return strings.Join(parts, " ")
}

// extractParagraphText は段落からプレーンテキストを結合する
func extractParagraphText(p xmlP) string {
	var sb strings.Builder
	for _, r := range p.Rs {
		sb.WriteString(r.T)
	}
	for _, fld := range p.Fld {
		sb.WriteString(fld.T)
	}
	return sb.String()
}

// hasNotes はスライドにノートが存在するか確認する
func (f *File) hasNotes(slideIdx int) bool {
	notesPath := f.notesPath(slideIdx)
	if notesPath == "" {
		return false
	}

	data, err := readZipFile(f.zr, notesPath)
	if err != nil || data == nil {
		return false
	}

	var notes xmlNotes
	if err := xml.Unmarshal(data, &notes); err != nil {
		return false
	}

	// body プレースホルダーからテキストを抽出
	for _, sp := range notes.CSld.SpTree.Shapes {
		ph := sp.NvSpPr.NvPr.Ph
		if ph == nil || ph.Type != "body" {
			continue
		}
		text := extractTextFromTxBody(sp.TxBody)
		if strings.TrimSpace(text) != "" {
			return true
		}
	}
	return false
}

// notesPath はスライドに対応するノートのZIPパスを返す
func (f *File) notesPath(slideIdx int) string {
	if slideIdx < 0 || slideIdx >= len(f.slideEntries) {
		return ""
	}
	entry := f.slideEntries[slideIdx]
	// スライドの .rels からノートのリレーションを探す
	relsPath := slideRelsPath(entry.Path)
	rels := loadRelsTyped(f, relsPath)
	for _, r := range rels {
		if strings.HasSuffix(r.Type, "/notesSlide") {
			return resolveRelTarget(pathDir(entry.Path), r.Target)
		}
	}
	return ""
}

// slideRelsPath はスライドXMLパスから .rels パスを生成する
func slideRelsPath(slidePath string) string {
	dir := pathDir(slidePath)
	base := pathBase(slidePath)
	return dir + "/_rels/" + base + ".rels"
}

// pathDir はパスのディレクトリ部分を返す
func pathDir(p string) string {
	idx := strings.LastIndex(p, "/")
	if idx < 0 {
		return ""
	}
	return p[:idx]
}

// pathBase はパスのファイル名部分を返す
func pathBase(p string) string {
	idx := strings.LastIndex(p, "/")
	if idx < 0 {
		return p
	}
	return p[idx+1:]
}
