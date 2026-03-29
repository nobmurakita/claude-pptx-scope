package main

import (
	"fmt"
	"os"

	"github.com/nobmurakita/cc-read-ppt/internal/pptx"
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
	slideNum, _ := cmd.Flags().GetInt("slide")
	includeNotes, _ := cmd.Flags().GetBool("notes")
	extractDir, _ := cmd.Flags().GetString("extract-images")

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

	enc := newJSONLWriter(os.Stdout)

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

func emitSlide(f *pptx.File, enc *jsonEncoder, slideNum int, includeNotes bool, extractDir string) error {
	sd, err := f.LoadSlide(slideNum, includeNotes, extractDir)
	if err != nil {
		return err
	}

	out := slideOutput{
		Slide:  sd.Number,
		Title:  sd.Title,
		Shapes: sd.Shapes,
		Notes:  sd.Notes,
	}
	if out.Shapes == nil {
		out.Shapes = []pptx.Shape{}
	}

	if err := enc.Encode(out); err != nil {
		return fmt.Errorf("JSON出力エラー: %w", err)
	}
	return nil
}
