package cmd

import (
	"fmt"
	"os"
	"path/filepath"

	"github.com/nobmurakita/claude-pptx-scope/internal/pptx"
	"github.com/spf13/cobra"
)

// NewImageCmd は image サブコマンドを生成する
func NewImageCmd() *cobra.Command {
	return &cobra.Command{
		Use:   "image <file> <image_id>",
		Short: "画像をファイルに保存する",
		Args:  cobra.ExactArgs(2),
		RunE:  runImage,
	}
}

func runImage(cmd *cobra.Command, args []string) error {
	f, err := pptx.OpenFile(args[0])
	if err != nil {
		return err
	}
	defer f.Close()

	imageID := args[1]

	// 一時ファイルを自動生成（拡張子は image_id から取得）
	ext := filepath.Ext(imageID)
	out, err := os.CreateTemp("", "pptx-scope-*"+ext)
	if err != nil {
		return fmt.Errorf("一時ファイルの作成エラー: %w", err)
	}
	outputPath := out.Name()
	succeeded := false
	defer func() {
		if !succeeded {
			out.Close()
			os.Remove(outputPath)
		}
	}()

	if err := f.ExtractImage(imageID, out); err != nil {
		return err
	}

	// 書き込み完了後、パスを報告する前にファイルを閉じる
	if err := out.Close(); err != nil {
		return fmt.Errorf("一時ファイルの書き込みエラー: %w", err)
	}
	succeeded = true

	useStdout, _ := cmd.Root().PersistentFlags().GetBool("stdout")
	if useStdout {
		fmt.Println(outputPath)
		return nil
	}
	enc := newJSONEncoder(os.Stdout)
	return enc.Encode(outputResult{File: outputPath})
}
