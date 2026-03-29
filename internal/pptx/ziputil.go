package pptx

import (
	"encoding/xml"
	"io"
)

// readZipFile は ZIP 内の指定パスのファイルを読み込む
func readZipFile(zi *zipIndex, path string) ([]byte, error) {
	f := zi.Lookup(path)
	if f == nil {
		return nil, nil // ファイルが存在しない場合は nil を返す
	}
	rc, err := f.Open()
	if err != nil {
		return nil, err
	}
	defer rc.Close()
	return io.ReadAll(rc)
}

// decodeZipXML は ZIP 内の指定パスのXMLをデコードする
func decodeZipXML(zi *zipIndex, path string, v any) error {
	data, err := readZipFile(zi, path)
	if err != nil {
		return err
	}
	if data == nil {
		return nil
	}
	return xml.Unmarshal(data, v)
}

// openZipFile は ZIP 内のファイルを開く
func openZipFile(zi *zipIndex, path string) (io.ReadCloser, int64, error) {
	f := zi.Lookup(path)
	if f == nil {
		return nil, 0, nil
	}
	rc, err := f.Open()
	if err != nil {
		return nil, 0, err
	}
	return rc, int64(f.UncompressedSize64), nil
}
