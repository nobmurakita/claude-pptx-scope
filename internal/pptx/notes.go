package pptx

import "strings"

// hasNotes はスライドにノートが存在するか確認する
func (f *File) hasNotes(slideIdx int) bool {
	txBody := f.findNotesBody(slideIdx)
	text := extractTextFromTxBody(txBody)
	return strings.TrimSpace(text) != ""
}

// findNotesBody はスライドのノートから body プレースホルダーの txBody を取得する。
// hasNotes と loadNotesParagraphs の共通処理。
func (f *File) findNotesBody(slideIdx int) *xmlTxBody {
	notesPath := f.notesPath(slideIdx)
	if notesPath == "" {
		return nil
	}

	// ノートの読み込み・パース失敗はスライド処理に影響させない
	var notes xmlNotes
	if err := decodeZipXML(f.zi, notesPath, &notes); err != nil {
		return nil
	}

	for _, child := range notes.CSld.SpTree.Children {
		if child.Sp == nil {
			continue
		}
		ph := child.Sp.NvSpPr.NvPr.Ph
		if ph == nil || ph.Type != "body" {
			continue
		}
		return child.Sp.TxBody
	}
	return nil
}

// notesPath はスライドに対応するノートのZIPパスを返す
func (f *File) notesPath(slideIdx int) string {
	if slideIdx < 0 || slideIdx >= len(f.slideEntries) {
		return ""
	}
	entry := f.slideEntries[slideIdx]
	// スライドの .rels からノートのリレーションを探す
	relsPath := relsPathFor(entry.Path)
	rels, _ := loadRelsTyped(f, relsPath) // エラー時はノートなしとして扱う
	for _, r := range rels {
		if strings.HasSuffix(r.Type, "/notesSlide") {
			return resolveRelTarget(pathDir(entry.Path), r.Target)
		}
	}
	return ""
}

// loadNotesParagraphs はスライドのノートの段落を取得する。
// ノートの読み込み・パース失敗時はnilを返す（スライド処理は継続する）。
func (f *File) loadNotesParagraphs(slideIdx int) []Paragraph {
	txBody := f.findNotesBody(slideIdx)
	if txBody == nil {
		return nil
	}
	ctx := newTextOnlyContext(f)
	paras := ctx.parseParagraphs(txBody.Ps, nil)
	if len(paras) > 0 {
		return paras
	}
	return nil
}
