package cmd

import (
	"fmt"

	"github.com/nobmurakita/claude-pptx-scope/internal/pptx"
	"github.com/spf13/cobra"
)

// NewSearchCmd は search サブコマンドを生成する
func NewSearchCmd() *cobra.Command {
	cmd := &cobra.Command{
		Use:   "search <file>",
		Short: "プレゼンテーション内のテキストを検索する",
		Args:  cobra.ExactArgs(1),
		RunE:  runSearch,
	}
	cmd.Flags().String("text", "", "検索文字列（部分一致、大文字小文字無視）")
	_ = cmd.MarkFlagRequired("text")
	cmd.Flags().IntSlice("slide", nil, "対象スライド番号（1始まり、複数指定可: --slide 1,3）")
	cmd.Flags().Bool("notes", false, "ノートも検索対象にする")
	return cmd
}

func runSearch(cmd *cobra.Command, args []string) error {
	textFlag, err := cmd.Flags().GetString("text")
	if err != nil {
		return fmt.Errorf("--text フラグの解析エラー: %w", err)
	}
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

	results, err := f.Search(textFlag, slideNums, includeNotes)
	if err != nil {
		return err
	}

	ow, err := newOutputWriter(cmd)
	if err != nil {
		return err
	}
	defer ow.cleanup()

	enc := newJSONEncoder(ow)
	for _, info := range results {
		if err := enc.Encode(info); err != nil {
			return fmt.Errorf("JSON出力エラー: %w", err)
		}
	}

	return ow.finalize()
}
