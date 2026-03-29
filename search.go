package main

import (
	"fmt"
	"os"

	"github.com/nobmurakita/cc-read-ppt/internal/pptx"
	"github.com/spf13/cobra"
)

func init() {
	searchCmd.Flags().String("text", "", "検索文字列（部分一致、大文字小文字無視）")
	searchCmd.Flags().Int("slide", 0, "対象スライド番号（1始まり）")
	searchCmd.Flags().Bool("notes", false, "ノートも検索対象にする")
	rootCmd.AddCommand(searchCmd)
}

var searchCmd = &cobra.Command{
	Use:   "search <file>",
	Short: "プレゼンテーション内のテキストを検索する",
	Args:  cobra.ExactArgs(1),
	RunE:  runSearch,
}

func runSearch(cmd *cobra.Command, args []string) error {
	textFlag, _ := cmd.Flags().GetString("text")
	slideNum, _ := cmd.Flags().GetInt("slide")
	includeNotes, _ := cmd.Flags().GetBool("notes")

	if textFlag == "" {
		return fmt.Errorf("--text を指定してください")
	}

	f, err := pptx.OpenFile(args[0])
	if err != nil {
		return err
	}
	defer f.Close()

	results, err := f.Search(textFlag, slideNum, includeNotes)
	if err != nil {
		return err
	}

	enc := newJSONLWriter(os.Stdout)
	for _, r := range results {
		out := slideOutput{
			Slide:  r.Number,
			Title:  r.Title,
			Shapes: r.Shapes,
			Notes:  r.Notes,
		}
		if out.Shapes == nil {
			out.Shapes = []pptx.Shape{}
		}
		if err := enc.Encode(out); err != nil {
			return fmt.Errorf("JSON出力エラー: %w", err)
		}
	}

	return nil
}
