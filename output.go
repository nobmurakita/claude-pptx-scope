package main

import (
	"encoding/json"
	"io"
)

// newJSONEncoder は JSON 出力用のエンコーダを生成する
func newJSONEncoder(w io.Writer) *json.Encoder {
	enc := json.NewEncoder(w)
	enc.SetEscapeHTML(false)
	return enc
}
