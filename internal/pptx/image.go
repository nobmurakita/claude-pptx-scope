package pptx

import (
	"io"
	"os"
	"path/filepath"
	"strings"
)

// extractImage は画像をファイルに抽出しパスを返す
func (ctx *parseContext) extractImage(embedID string) string {
	// リレーション未解決時は画像なしで継続する（スライド処理を止めない）
	if ctx.slideRels == nil {
		return ""
	}
	target, ok := ctx.slideRels[embedID]
	if !ok {
		return ""
	}

	// ZIP内のパスを解決
	mediaPath := resolveRelTarget(pathDir(ctx.slidePath), target)

	// ZIP内のファイルを開く（メディアファイルが欠損していてもスライド処理は継続する）
	rc, _, err := openZipFile(ctx.f.zi, mediaPath)
	if err != nil || rc == nil {
		return ""
	}
	defer rc.Close()

	// 抽出先ファイルを作成（一意なファイル名を自動生成）
	ext := strings.ToLower(filepath.Ext(mediaPath))
	outFile, err := os.CreateTemp(ctx.extractDir, "image_*"+ext)
	if err != nil {
		return ""
	}
	outPath := outFile.Name()
	writeOK := false
	defer func() {
		if !writeOK {
			outFile.Close()
			os.Remove(outPath)
		}
	}()

	// コピー・フラッシュ失敗時も画像なしで継続する（defer で一時ファイルを削除）
	if _, err := io.Copy(outFile, rc); err != nil {
		return ""
	}
	if err := outFile.Close(); err != nil {
		return ""
	}
	writeOK = true

	return outPath
}
