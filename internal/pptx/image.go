package pptx

import (
	"fmt"
	"io"
	"os"
	"path/filepath"
	"strings"
)

// extractImage は画像をファイルに抽出しメタデータを返す
func (ctx *parseContext) extractImage(embedID string, pos *Position) *ImageData {
	if ctx.slideRels == nil {
		return nil
	}
	target, ok := ctx.slideRels[embedID]
	if !ok {
		return nil
	}

	// ZIP内のパスを解決
	mediaPath := resolveRelTarget(pathDir(ctx.slidePath), target)

	// ZIP内のファイルを開く（メディアファイルが欠損していてもスライド処理は継続する）
	rc, size, err := openZipFile(ctx.f.zi, mediaPath)
	if err != nil || rc == nil {
		return nil
	}
	defer rc.Close()

	// ファイル拡張子から形式を判定
	ext := strings.ToLower(filepath.Ext(mediaPath))
	format := strings.TrimPrefix(ext, ".")
	if format == "jpg" {
		format = "jpeg"
	}

	// 抽出先パス
	ctx.imageCount++
	outName := fmt.Sprintf("image_%d%s", ctx.imageCount, ext)
	outPath := filepath.Join(ctx.extractDir, outName)

	// ファイルに書き出し（書き出し失敗時は画像メタデータなしで継続する）
	outFile, err := os.Create(outPath)
	if err != nil {
		return nil
	}
	defer outFile.Close()

	if _, err := io.Copy(outFile, rc); err != nil {
		outFile.Close()
		os.Remove(outPath)
		return nil
	}

	// サイズ（EMU → ピクセル）
	width, height := 0, 0
	if pos != nil {
		width = int(pos.Cx / 9525)
		height = int(pos.Cy / 9525)
	}

	return &ImageData{
		Format: format,
		Width:  width,
		Height: height,
		Size:   size,
		Path:   outPath,
	}
}
