package pptx

import (
	"fmt"
	"io"
)

// resolveImagePath はリレーションIDからZIP内の画像パスを解決する。
// 解決できない場合は空文字列を返す（スライド処理を止めない）。
func (ctx *parseContext) resolveImagePath(embedID string) string {
	if ctx.slideRels == nil {
		return ""
	}
	target, ok := ctx.slideRels[embedID]
	if !ok {
		return ""
	}
	return resolveRelTarget(pathDir(ctx.slidePath), target)
}

// ExtractImage はZIP内の画像をwに書き出す。
func (f *File) ExtractImage(mediaPath string, w io.Writer) error {
	rc, _, err := openZipFile(f.zi, mediaPath)
	if err != nil {
		return err
	}
	if rc == nil {
		return fmt.Errorf("画像が見つかりません: %s", mediaPath)
	}
	defer rc.Close()

	_, err = io.Copy(w, rc)
	return err
}
