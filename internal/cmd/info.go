package cmd

import (
	"fmt"

	"github.com/nobmurakita/claude-pptx-scope/internal/pptx"
	"github.com/spf13/cobra"
)

// NewInfoCmd は info サブコマンドを生成する
func NewInfoCmd() *cobra.Command {
	return &cobra.Command{
		Use:   "info <file>",
		Short: "ファイルの概要（スライド一覧、スライドサイズ）を表示する",
		Args:  cobra.ExactArgs(1),
		RunE:  runInfo,
	}
}

// infoMeta はファイルレベルのメタ情報
type infoMeta struct {
	File      string         `json:"file"`
	SlideSize pptx.SlideSize `json:"slide_size"`
}

func runInfo(cmd *cobra.Command, args []string) error {
	f, err := pptx.OpenFile(args[0])
	if err != nil {
		return err
	}
	defer f.Close()

	infos, err := f.LoadSlideInfos()
	if err != nil {
		return err
	}

	ow, err := newOutputWriter(cmd)
	if err != nil {
		return err
	}
	defer ow.cleanup()

	enc := newJSONEncoder(ow)

	// メタ情報行
	if err := enc.Encode(infoMeta{File: f.Name, SlideSize: f.GetSlideSize()}); err != nil {
		return fmt.Errorf("JSON出力エラー: %w", err)
	}

	// スライド情報行
	for _, info := range infos {
		if err := enc.Encode(info); err != nil {
			return fmt.Errorf("JSON出力エラー: %w", err)
		}
	}

	return ow.finalize()
}
