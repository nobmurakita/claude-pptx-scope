package main

import (
	"fmt"
	"os"

	"github.com/spf13/cobra"
)

const (
	exitOK    = 0
	exitError = 1
)

var rootCmd = &cobra.Command{
	Use:           "cc-read-ppt",
	Short:         "PowerPoint ファイル（.pptx）の内容をAIエージェント向けに読み取るツール",
	SilenceUsage:  true,
	SilenceErrors: true,
}

func execute() int {
	err := rootCmd.Execute()
	if err == nil {
		return exitOK
	}
	fmt.Fprintf(os.Stderr, "cc-read-ppt: %s\n", err)
	return exitError
}
