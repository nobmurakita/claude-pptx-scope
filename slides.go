package main

import (
	"encoding/json"
	"fmt"
	"os"

	"github.com/nobmurakita/cc-read-pptx/internal/pptx"
	"github.com/spf13/cobra"
)

func init() {
	slidesCmd.Flags().Int("slide", 0, "対象スライド番号（1始まり）")
	slidesCmd.Flags().Bool("notes", false, "ノートも出力する")
	slidesCmd.Flags().String("extract-images", "", "画像を抽出するディレクトリ")
	rootCmd.AddCommand(slidesCmd)
}

var slidesCmd = &cobra.Command{
	Use:   "slides <file>",
	Short: "スライドの内容をJSONL形式で出力する",
	Args:  cobra.ExactArgs(1),
	RunE:  runSlides,
}

type slideOutput struct {
	Slide  int              `json:"slide"`
	Title  string           `json:"title,omitempty"`
	Shapes []pptx.Shape     `json:"shapes"`
	Notes  []pptx.Paragraph `json:"notes,omitempty"`
}

func runSlides(cmd *cobra.Command, args []string) error {
	slideNum, err := cmd.Flags().GetInt("slide")
	if err != nil {
		return fmt.Errorf("--slide フラグの解析エラー: %w", err)
	}
	includeNotes, err := cmd.Flags().GetBool("notes")
	if err != nil {
		return fmt.Errorf("--notes フラグの解析エラー: %w", err)
	}
	extractDir, err := cmd.Flags().GetString("extract-images")
	if err != nil {
		return fmt.Errorf("--extract-images フラグの解析エラー: %w", err)
	}

	if slideNum < 0 {
		return fmt.Errorf("--slide には1以上の値を指定してください")
	}

	f, err := pptx.OpenFile(args[0])
	if err != nil {
		return err
	}
	defer f.Close()

	// 画像抽出ディレクトリの作成
	if extractDir != "" {
		if err := os.MkdirAll(extractDir, 0755); err != nil {
			return fmt.Errorf("ディレクトリの作成エラー: %w", err)
		}
	}

	enc := newJSONEncoder(os.Stdout)

	if slideNum > 0 {
		// 特定のスライド
		return emitSlide(f, enc, slideNum, includeNotes, extractDir)
	}

	// 全スライド
	for i := 1; i <= f.SlideCount(); i++ {
		if err := emitSlide(f, enc, i, includeNotes, extractDir); err != nil {
			return err
		}
	}
	return nil
}

func emitSlide(f *pptx.File, enc *json.Encoder, slideNum int, includeNotes bool, extractDir string) error {
	sd, err := f.LoadSlide(slideNum, includeNotes, extractDir)
	if err != nil {
		return err
	}
	return emitSlideData(enc, sd)
}

func emitSlideData(enc *json.Encoder, sd *pptx.SlideData) error {
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
