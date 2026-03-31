package cmd

import (
	"encoding/json"
	"fmt"
	"os"

	"github.com/nobmurakita/cc-read-pptx/internal/pptx"
	"github.com/spf13/cobra"
)

// NewSlidesCmd は slides サブコマンドを生成する
func NewSlidesCmd() *cobra.Command {
	cmd := &cobra.Command{
		Use:   "slides <file>",
		Short: "スライドの内容をJSONL形式で出力する",
		Args:  cobra.ExactArgs(1),
		RunE:  runSlides,
	}
	cmd.Flags().IntSlice("slide", nil, "対象スライド番号（1始まり、複数指定可: --slide 1,3）")
	cmd.Flags().Bool("notes", false, "ノートも出力する")
	return cmd
}

type slideOutput struct {
	Slide  int              `json:"slide"`
	Title  string           `json:"title,omitempty"`
	Shapes []pptx.Shape     `json:"shapes"`
	Notes  []pptx.Paragraph `json:"notes,omitempty"`
}

// styleDefsOutput はスタイル定義の出力行
type styleDefsOutput struct {
	Styles []pptx.StyleDef `json:"_styles"`
}

func runSlides(cmd *cobra.Command, args []string) error {
	slideNums, err := cmd.Flags().GetIntSlice("slide")
	if err != nil {
		return fmt.Errorf("--slide フラグの解析エラー: %w", err)
	}
	includeNotes, err := cmd.Flags().GetBool("notes")
	if err != nil {
		return fmt.Errorf("--notes フラグの解析エラー: %w", err)
	}

	f, err := pptx.OpenFile(args[0])
	if err != nil {
		return err
	}
	defer f.Close()

	targets, err := f.ResolveSlideNums(slideNums)
	if err != nil {
		return err
	}

	enc := newJSONLWriter(os.Stdout)
	dedup := pptx.NewStyleDeduplicator()

	for _, n := range targets {
		if err := emitSlide(f, enc, dedup, n, includeNotes); err != nil {
			return err
		}
	}
	return nil
}

func emitSlide(f *pptx.File, enc *json.Encoder, dedup *pptx.StyleDeduplicator, slideNum int, includeNotes bool) error {
	sd, err := f.LoadSlide(slideNum, includeNotes)
	if err != nil {
		return err
	}
	return emitSlideData(enc, dedup, sd)
}

func emitSlideData(enc *json.Encoder, dedup *pptx.StyleDeduplicator, sd *pptx.SlideData) error {
	// フォントスタイルの重複排除（スライド横断で共有）
	newStyles := dedup.Deduplicate(sd)

	// 新規スタイル定義があれば独立行として先に出力
	if len(newStyles) > 0 {
		if err := enc.Encode(styleDefsOutput{Styles: newStyles}); err != nil {
			return fmt.Errorf("スタイル定義の出力エラー: %w", err)
		}
	}

	out := slideOutput{
		Slide:  sd.Number,
		Title:  sd.Title,
		Shapes: sd.Shapes,
		Notes:  sd.Notes,
	}
	if err := enc.Encode(out); err != nil {
		return fmt.Errorf("JSON出力エラー: %w", err)
	}
	return nil
}
