package pptx

import "encoding/xml"

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

// loadRels は指定パスの .rels をパースしてIDからTargetへのマップを返す
func loadRels(f *File, relsPath string) map[string]string {
	var rels xmlRelationships
	if err := decodeZipXML(f.zi, relsPath, &rels); err != nil {
		return nil
	}
	m := make(map[string]string, len(rels.Rels))
	for _, r := range rels.Rels {
		m[r.ID] = r.Target
	}
	return m
}

// loadRelsTyped は指定パスの .rels をパースしてリレーション一覧を返す
func loadRelsTyped(f *File, relsPath string) []xmlRel {
	var rels xmlRelationships
	if err := decodeZipXML(f.zi, relsPath, &rels); err != nil {
		return nil
	}
	return rels.Rels
}
