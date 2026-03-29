package pptx

import (
	"archive/zip"
	"encoding/xml"
	"fmt"
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

	// 遅延ロード
	theme     *themeColors
	themeOnce sync.Once
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

// GetSlideSize はスライドサイズを返す
func (f *File) GetSlideSize() SlideSize {
	return f.slideSize
}

// getTheme はテーマカラーを遅延ロードして返す
func (f *File) getTheme() *themeColors {
	f.themeOnce.Do(func() {
		data, err := readZipFile(f.zi, "ppt/theme/theme1.xml")
		if err == nil && data != nil {
			f.theme = parseThemeColors(data)
		}
	})
	return f.theme
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

	// リレーション読み込み
	rels := loadRels(f, "ppt/_rels/presentation.xml.rels")
	if rels == nil {
		return fmt.Errorf("presentation.xml.rels の読み込みに失敗")
	}

	// sldIdLst からスライドエントリを構築
	f.slideEntries = make([]slideEntry, 0, len(pres.SldIdLst.SldId))
	for _, sid := range pres.SldIdLst.SldId {
		target, ok := rels[sid.RID]
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

// cleanPath はパス内の ".." を解決する
func cleanPath(p string) string {
	parts := strings.Split(p, "/")
	var result []string
	for _, part := range parts {
		if part == ".." {
			if len(result) > 0 {
				result = result[:len(result)-1]
			}
		} else if part != "." && part != "" {
			result = append(result, part)
		}
	}
	return strings.Join(result, "/")
}

// xmlPresentation は presentation.xml の構造
type xmlPresentation struct {
	XMLName  xml.Name `xml:"presentation"`
	SldSz    xmlSldSz `xml:"sldSz"`
	SldIdLst struct {
		SldId []xmlSldId `xml:"sldId"`
	} `xml:"sldIdLst"`
}

type xmlSldSz struct {
	Cx int64 `xml:"cx,attr"`
	Cy int64 `xml:"cy,attr"`
}

type xmlSldId struct {
	ID  string `xml:"id,attr"`
	RID string `xml:"http://schemas.openxmlformats.org/officeDocument/2006/relationships id,attr"`
}
