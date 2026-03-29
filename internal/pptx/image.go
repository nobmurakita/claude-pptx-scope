package pptx

import (
	"io"
	"os"
	"path/filepath"
	"strings"
)

// extractImage は画像をファイルに抽出しメタデータを返す
func (ctx *parseContext) extractImage(embedID string, pos *Position) *ImageData {
	// リレーション未解決時は画像なしで継続する（スライド処理を止めない）
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

	// 抽出先ファイルを作成（一意なファイル名を自動生成）
	outFile, err := os.CreateTemp(ctx.extractDir, "image_*"+ext)
	if err != nil {
		return nil
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
		return nil
	}
	if err := outFile.Close(); err != nil {
		return nil
	}
	writeOK = true

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
