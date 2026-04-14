package main

import (
	"fmt"
	"os"

	"github.com/nobmurakita/claude-pptx-scope/internal/cmd"
	"github.com/spf13/cobra"
)

func main() {
	rootCmd := &cobra.Command{
		Use:           "pptx-scope",
		Short:         "PowerPoint ファイル（.pptx）の内容をAIエージェント向けに読み取るツール",
		SilenceUsage:  true,
		SilenceErrors: true,
	}
	rootCmd.PersistentFlags().Bool("stdout", false, "出力を標準出力に直接書き出す（デバッグ用）")
	rootCmd.AddCommand(
		cmd.NewInfoCmd(),
		cmd.NewSlidesCmd(),
		cmd.NewSearchCmd(),
		cmd.NewImageCmd(),
		cmd.NewCleanupCmd(),
		cmd.NewVersionCmd(),
	)

	if err := rootCmd.Execute(); err != nil {
		fmt.Fprintf(os.Stderr, "pptx-scope: %s\n", err)
		os.Exit(1)
	}
}
