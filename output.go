package main

import (
	"encoding/json"
	"io"
)

// newJSONLWriter は JSONL 出力用のエンコーダを生成する
func newJSONLWriter(w io.Writer) *json.Encoder {
	enc := json.NewEncoder(w)
	enc.SetEscapeHTML(false)
	return enc
}
