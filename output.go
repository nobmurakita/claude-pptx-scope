package main

import (
	"encoding/json"
	"io"
)

type jsonEncoder = json.Encoder

// newJSONLWriter は JSONL 出力用のエンコーダを生成する
func newJSONLWriter(w io.Writer) *jsonEncoder {
	enc := json.NewEncoder(w)
	enc.SetEscapeHTML(false)
	return enc
}
