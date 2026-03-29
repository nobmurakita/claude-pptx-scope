package main

import (
	"fmt"
	"os"

	"github.com/nobmurakita/cc-read-pptx/internal/pptx"
	"github.com/spf13/cobra"
)

func init() {
	searchCmd.Flags().String("text", "", "検索文字列（部分一致、大文字小文字無視）")
	_ = searchCmd.MarkFlagRequired("text")
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
	textFlag, err := cmd.Flags().GetString("text")
	if err != nil {
		return fmt.Errorf("--text フラグの解析エラー: %w", err)
	}
	slideNum, err := cmd.Flags().GetInt("slide")
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

	results, err := f.Search(textFlag, slideNum, includeNotes)
	if err != nil {
		return err
	}

	enc := newJSONLWriter(os.Stdout)
	for i := range results {
		if err := emitSlideData(enc, &results[i]); err != nil {
			return err
		}
	}

	return nil
}
