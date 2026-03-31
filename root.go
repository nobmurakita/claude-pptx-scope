package main

import (
	"fmt"
	"os"

	"github.com/nobmurakita/cc-read-pptx/internal/cmd"
	"github.com/spf13/cobra"
)

const (
	exitOK    = 0
	exitError = 1
)

var rootCmd = &cobra.Command{
	Use:           "cc-read-pptx",
	Short:         "PowerPoint ファイル（.pptx）の内容をAIエージェント向けに読み取るツール",
	SilenceUsage:  true,
	SilenceErrors: true,
}

func init() {
	rootCmd.AddCommand(
		cmd.NewInfoCmd(),
		cmd.NewSlidesCmd(),
		cmd.NewSearchCmd(),
		cmd.NewImageCmd(),
	)
}

func execute() int {
	err := rootCmd.Execute()
	if err == nil {
		return exitOK
	}
	fmt.Fprintf(os.Stderr, "cc-read-pptx: %s\n", err)
	return exitError
}
