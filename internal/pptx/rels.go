package pptx

import (
	"encoding/xml"
	"fmt"
)

// xmlRelationships は .rels ファイルの構造
type xmlRelationships struct {
	XMLName xml.Name `xml:"Relationships"`
	Rels    []xmlRel `xml:"Relationship"`
}

type xmlRel struct {
	ID     string `xml:"Id,attr"`
	Type   string `xml:"Type,attr"`
	Target string `xml:"Target,attr"`
}

// loadRels は指定パスの .rels をパースしてIDからTargetへのマップを返す。
// ファイルが存在しない場合は nil を返す。パースエラー時もエラーを伝播する。
func loadRels(f *File, relsPath string) (map[string]string, error) {
	var rels xmlRelationships
	if err := decodeZipXML(f.zi, relsPath, &rels); err != nil {
		return nil, fmt.Errorf("%s のパースに失敗: %w", relsPath, err)
	}
	if len(rels.Rels) == 0 {
		return nil, nil
	}
	m := make(map[string]string, len(rels.Rels))
	for _, r := range rels.Rels {
		m[r.ID] = r.Target
	}
	return m, nil
}

// loadRelsTyped は指定パスの .rels をパースしてリレーション一覧を返す。
func loadRelsTyped(f *File, relsPath string) ([]xmlRel, error) {
	var rels xmlRelationships
	if err := decodeZipXML(f.zi, relsPath, &rels); err != nil {
		return nil, fmt.Errorf("%s のパースに失敗: %w", relsPath, err)
	}
	return rels.Rels, nil
}
