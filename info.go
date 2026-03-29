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
	File      string            `json:"file"`
	SlideSize pptx.SlideSize    `json:"slide_size"`
	Slides    []slideInfoOutput `json:"slides"`
}

type slideInfoOutput struct {
	Number    int    `json:"number"`
	Title     string `json:"title,omitempty"`
	HasNotes  bool   `json:"has_notes,omitempty"`
	HasImages bool   `json:"has_images,omitempty"`
	Hidden    bool   `json:"hidden,omitempty"`
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

	slides := make([]slideInfoOutput, len(infos))
	for i, info := range infos {
		slides[i] = slideInfoOutput{
			Number:    info.Number,
			Title:     info.Title,
			HasNotes:  info.HasNotes,
			HasImages: info.HasImages,
			Hidden:    info.Hidden,
		}
	}

	out := infoOutput{
		File:      f.Name,
		SlideSize: f.GetSlideSize(),
		Slides:    slides,
	}

	enc := newJSONLWriter(os.Stdout)
	if err := enc.Encode(out); err != nil {
		return fmt.Errorf("JSON出力エラー: %w", err)
	}
	return nil
}
