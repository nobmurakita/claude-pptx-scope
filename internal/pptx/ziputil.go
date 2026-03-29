package pptx

import (
	"archive/zip"
	"encoding/xml"
	"io"
)

// readZipFile は ZIP 内の指定パスのファイルを読み込む
func readZipFile(zr *zip.ReadCloser, path string) ([]byte, error) {
	for _, f := range zr.File {
		if f.Name == path {
			rc, err := f.Open()
			if err != nil {
				return nil, err
			}
			defer rc.Close()
			return io.ReadAll(rc)
		}
	}
	return nil, nil // ファイルが存在しない場合は nil を返す
}

// decodeZipXML は ZIP 内の指定パスのXMLをデコードする
func decodeZipXML(zr *zip.ReadCloser, path string, v any) error {
	data, err := readZipFile(zr, path)
	if err != nil {
		return err
	}
	if data == nil {
		return nil
	}
	return xml.Unmarshal(data, v)
}

// openZipFile は ZIP 内のファイルを開く
func openZipFile(zr *zip.ReadCloser, path string) (io.ReadCloser, int64, error) {
	for _, f := range zr.File {
		if f.Name == path {
			rc, err := f.Open()
			if err != nil {
				return nil, 0, err
			}
			return rc, int64(f.UncompressedSize64), nil
		}
	}
	return nil, 0, nil
}
