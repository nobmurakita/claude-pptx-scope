package cmd

import (
	"encoding/json"
	"fmt"
	"os"
	"path/filepath"
	"strings"

	"github.com/spf13/cobra"
)

// tmpFilePrefix は cleanup の安全確認に使う一時ファイルのプレフィックス
const tmpFilePrefix = "pptx-scope-tmp-"

// NewCleanupCmd は cleanup サブコマンドを生成する
func NewCleanupCmd() *cobra.Command {
	return &cobra.Command{
		Use:   "cleanup <file> [file...]",
		Short: "pptx-scope が生成した一時ファイルを削除する",
		Args:  cobra.MinimumNArgs(1),
		RunE:  runCleanup,
	}
}

type cleanupOutput struct {
	Deleted int `json:"deleted"`
}

func runCleanup(cmd *cobra.Command, args []string) error {
	// os.CreateTemp("", ...) で生成された一時ファイルのみを削除対象とする。
	// 他コマンドは os.TempDir() 直下に pptx-scope-tmp-* を作成するため、
	// 親ディレクトリがそれと一致するかを比較する。
	// macOS では os.TempDir() が /var/... を返す一方、実体は /private/var/... で
	// あるため、両辺で EvalSymlinks による正規化を行ってから比較する。
	tmpDir, err := canonicalDir(os.TempDir())
	if err != nil {
		return fmt.Errorf("一時ディレクトリの解決エラー: %w", err)
	}

	deleted := 0
	for _, path := range args {
		abs, err := filepath.Abs(path)
		if err != nil {
			return fmt.Errorf("パスの解決エラー: %w", err)
		}
		if !strings.HasPrefix(filepath.Base(abs), tmpFilePrefix) {
			return fmt.Errorf("pptx-scope が生成した一時ファイルではありません: %s", path)
		}
		parent, err := canonicalDir(filepath.Dir(abs))
		if err != nil {
			return fmt.Errorf("親ディレクトリの解決エラー: %w", err)
		}
		if parent != tmpDir {
			return fmt.Errorf("一時ディレクトリ配下ではありません: %s", path)
		}
		if err := os.Remove(abs); err != nil {
			if os.IsNotExist(err) {
				continue
			}
			return fmt.Errorf("ファイルの削除エラー: %w", err)
		}
		deleted++
	}

	enc := json.NewEncoder(os.Stdout)
	enc.SetEscapeHTML(false)
	return enc.Encode(cleanupOutput{Deleted: deleted})
}

// canonicalDir はディレクトリパスを EvalSymlinks + Clean で正規化する。
// 対象ディレクトリが存在しない場合は Clean した結果をそのまま返す。
func canonicalDir(dir string) (string, error) {
	resolved, err := filepath.EvalSymlinks(dir)
	if err != nil {
		if os.IsNotExist(err) {
			return filepath.Clean(dir), nil
		}
		return "", err
	}
	return filepath.Clean(resolved), nil
}
