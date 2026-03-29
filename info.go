package main

import (
	"fmt"
	"os"

	"github.com/nobmurakita/cc-read-pptx/internal/pptx"
	"github.com/spf13/cobra"
)

func init() {
	rootCmd.AddCommand(infoCmd)
}

var infoCmd = &cobra.Command{
	Use:   "info <file>",
	Short: "ファイルの概要（スライド一覧、スライドサイズ）を表示する",
	Args:  cobra.ExactArgs(1),
	RunE:  runInfo,
}

type infoOutput struct {
	File      string           `json:"file"`
	SlideSize pptx.SlideSize   `json:"slide_size"`
	Slides    []pptx.SlideInfo `json:"slides"`
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

	out := infoOutput{
		File:      f.Name,
		SlideSize: f.GetSlideSize(),
		Slides:    infos,
	}

	enc := newJSONLWriter(os.Stdout)
	if err := enc.Encode(out); err != nil {
		return fmt.Errorf("JSON出力エラー: %w", err)
	}
	return nil
}
