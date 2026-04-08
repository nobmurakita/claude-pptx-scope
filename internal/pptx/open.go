package pptx

import (
	"archive/zip"
	"encoding/xml"
	"fmt"
	"path"
	"path/filepath"
	"strings"
	"sync"
)

// zipIndex は ZIP 内のファイルを名前で高速検索するためのインデックス
type zipIndex struct {
	files map[string]*zip.File
}

// newZipIndex は zip.ReadCloser からインデックスを構築する
func newZipIndex(zr *zip.ReadCloser) *zipIndex {
	m := make(map[string]*zip.File, len(zr.File))
	for _, f := range zr.File {
		m[f.Name] = f
	}
	return &zipIndex{files: m}
}

// Lookup は指定パスのファイルを返す。見つからなければ nil
func (zi *zipIndex) Lookup(path string) *zip.File {
	return zi.files[path]
}

// File はオープンしたPowerPointファイルを表す
type File struct {
	Name string // ファイル名（パス除去済み）
	zr   *zip.ReadCloser
	zi   *zipIndex

	// presentation.xml からパースした情報
	slideEntries []slideEntry // sldIdLst の順序
	slideSize    SlideSize

	// presentation.xml の defaultTextStyle
	defaultTextStyle *xmlLstStyle

	// テーマ
	themePath string // presentation.xml.rels から解決したテーマファイルのZIPパス
	theme     *themeColors
	themeOnce sync.Once

	// レイアウト/マスター継承キャッシュ
	inherit     *inheritCache
	inheritOnce sync.Once
}

// slideEntry はスライドのエントリ
type slideEntry struct {
	RID  string // リレーションID
	Path string // ZIP内のXMLパス（例: "ppt/slides/slide1.xml"）
}

// SlideSize はスライドのサイズ
type SlideSize struct {
	Width  int64 `json:"width"`
	Height int64 `json:"height"`
}

// OpenFile は PowerPoint ファイルを開く
func OpenFile(path string) (*File, error) {
	ext := strings.ToLower(filepath.Ext(path))
	if ext != ".pptx" {
		return nil, fmt.Errorf(".pptx 形式のみ対応しています")
	}

	zr, err := zip.OpenReader(path)
	if err != nil {
		return nil, err
	}

	f := &File{
		Name: filepath.Base(path),
		zr:   zr,
		zi:   newZipIndex(zr),
	}

	if err := f.loadPresentation(); err != nil {
		zr.Close()
		return nil, err
	}

	return f, nil
}

// Close はファイルを閉じる
func (f *File) Close() error {
	if f.zr != nil {
		return f.zr.Close()
	}
	return nil
}

// SlideCount はスライド数を返す
func (f *File) SlideCount() int {
	return len(f.slideEntries)
}

// ResolveSlideNums はスライド番号のスライスを検証し、対象番号リストを返す。
// 空スライスの場合は全スライドを返す。
func (f *File) ResolveSlideNums(slideNums []int) ([]int, error) {
	if len(slideNums) == 0 {
		all := make([]int, len(f.slideEntries))
		for i := range all {
			all[i] = i + 1
		}
		return all, nil
	}
	for _, n := range slideNums {
		if n < 1 || n > len(f.slideEntries) {
			return nil, fmt.Errorf("スライド番号 %d は範囲外です（1〜%d）", n, len(f.slideEntries))
		}
	}
	return slideNums, nil
}

// GetSlideSize はスライドサイズを返す
func (f *File) GetSlideSize() SlideSize {
	return f.slideSize
}

// getTheme はテーマカラーを遅延ロードして返す。
// テーマファイルが存在しない場合や読み込み失敗時は nil を返す。
func (f *File) getTheme() *themeColors {
	f.themeOnce.Do(func() {
		if f.themePath == "" {
			return
		}
		data, err := readZipFile(f.zi, f.themePath)
		if err != nil || data == nil {
			return
		}
		tc, err := parseThemeColors(data)
		if err != nil {
			return
		}
		f.theme = tc
	})
	return f.theme
}

// slidePathToNum はZIP内のスライドパスからスライド番号（1始まり）を返す。
// 見つからない場合は 0 を返す。
func (f *File) slidePathToNum(slidePath string) int {
	for i, entry := range f.slideEntries {
		if entry.Path == slidePath {
			return i + 1
		}
	}
	return 0
}

// getInheritCache はレイアウト/マスター継承キャッシュを遅延初期化して返す
func (f *File) getInheritCache() *inheritCache {
	f.inheritOnce.Do(func() {
		f.inherit = newInheritCache()
	})
	return f.inherit
}

// loadPresentation は presentation.xml をパースする
func (f *File) loadPresentation() error {
	data, err := readZipFile(f.zi, "ppt/presentation.xml")
	if err != nil {
		return fmt.Errorf("presentation.xml の読み込みに失敗: %w", err)
	}
	if data == nil {
		return fmt.Errorf("presentation.xml が見つかりません")
	}

	var pres xmlPresentation
	if err := xml.Unmarshal(data, &pres); err != nil {
		return fmt.Errorf("presentation.xml のパースに失敗: %w", err)
	}

	// スライドサイズ
	f.slideSize = SlideSize{
		Width:  pres.SldSz.Cx,
		Height: pres.SldSz.Cy,
	}

	// defaultTextStyle
	f.defaultTextStyle = pres.DefaultTextStyle

	// リレーション読み込み
	typedRels, err := loadRelsTyped(f, "ppt/_rels/presentation.xml.rels")
	if err != nil {
		return fmt.Errorf("presentation.xml.rels の読み込みに失敗: %w", err)
	}
	if typedRels == nil {
		return fmt.Errorf("presentation.xml.rels が見つかりません")
	}

	// ID→Target マップを構築 + テーマパスを探索
	relsMap := make(map[string]string, len(typedRels))
	for _, r := range typedRels {
		relsMap[r.ID] = r.Target
		if strings.HasSuffix(r.Type, "/theme") {
			f.themePath = resolveRelTarget("ppt", r.Target)
		}
	}

	// sldIdLst からスライドエントリを構築
	f.slideEntries = make([]slideEntry, 0, len(pres.SldIdLst.SldId))
	for _, sid := range pres.SldIdLst.SldId {
		target, ok := relsMap[sid.RID]
		if !ok {
			continue
		}
		path := resolveRelTarget("ppt", target)
		f.slideEntries = append(f.slideEntries, slideEntry{
			RID:  sid.RID,
			Path: path,
		})
	}

	return nil
}

// resolveRelTarget はリレーションのTargetをZIP内のパスに解決する
func resolveRelTarget(basePath, target string) string {
	if strings.HasPrefix(target, "/") {
		return target[1:]
	}
	combined := basePath + "/" + target
	return cleanPath(combined)
}

// cleanPath はパス内の ".." や "." を解決する
func cleanPath(p string) string {
	if p == "" {
		return ""
	}
	return path.Clean(p)
}
