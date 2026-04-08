package cmd

import (
	"encoding/json"
	"fmt"

	"github.com/nobmurakita/claude-pptx-scope/internal/pptx"
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

// notesOutput はノートの出力行
type notesOutput struct {
	Notes []pptx.Paragraph `json:"notes"`
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

	ow, err := newOutputWriter(cmd)
	if err != nil {
		return err
	}
	defer ow.cleanup()

	enc := newJSONLWriter(ow)
	dedup := pptx.NewStyleDeduplicator()

	for _, n := range targets {
		if err := emitSlide(f, enc, dedup, n, includeNotes); err != nil {
			return err
		}
	}
	return ow.finalize()
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

	// スライドヘッダ行
	header := sd.Info()
	header.Shapes = intPtr(len(sd.Shapes))
	if err := enc.Encode(header); err != nil {
		return fmt.Errorf("JSON出力エラー: %w", err)
	}

	// 図形を1つずつ個別の行として出力
	for i := range sd.Shapes {
		if err := enc.Encode(sd.Shapes[i]); err != nil {
			return fmt.Errorf("JSON出力エラー: %w", err)
		}
	}

	// ノートを独立行として出力
	if len(sd.Notes) > 0 {
		if err := enc.Encode(notesOutput{Notes: sd.Notes}); err != nil {
			return fmt.Errorf("JSON出力エラー: %w", err)
		}
	}

	return nil
}
