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

// loadRelsTyped は指定パスの .rels をパースしてリレーション一覧を返す。
// ファイルが存在しない場合は nil, nil を返す。
// ファイルが存在するがリレーションが空の場合は空スライスを返す。
func loadRelsTyped(f *File, relsPath string) ([]xmlRel, error) {
	data, err := readZipFile(f.zi, relsPath)
	if err != nil {
		return nil, fmt.Errorf("%s の読み込みに失敗: %w", relsPath, err)
	}
	if data == nil {
		return nil, nil
	}
	var rels xmlRelationships
	if err := xml.Unmarshal(data, &rels); err != nil {
		return nil, fmt.Errorf("%s のパースに失敗: %w", relsPath, err)
	}
	if rels.Rels == nil {
		return []xmlRel{}, nil
	}
	return rels.Rels, nil
}

// relsToMap は []xmlRel を ID→Target マップに変換する
func relsToMap(rels []xmlRel) map[string]string {
	m := make(map[string]string, len(rels))
	for _, r := range rels {
		m[r.ID] = r.Target
	}
	return m
}
