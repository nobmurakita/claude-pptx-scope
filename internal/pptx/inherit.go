package pptx

import (
	"strings"
	"sync"
)

// phKey はプレースホルダーのマッチングキー
type phKey struct {
	Type string
	Idx  string
}

// placeholderDef はレイアウト/マスター上のプレースホルダー定義
type placeholderDef struct {
	xfrm     *xmlXfrm
	lstStyle *xmlLstStyle // txBody 内の lstStyle
}

// layoutData はスライドレイアウトのパース済みデータ
type layoutData struct {
	placeholders map[phKey]*placeholderDef
	masterPath   string // このレイアウトが参照するマスターのZIPパス
}

// masterData はスライドマスターのパース済みデータ
type masterData struct {
	placeholders map[phKey]*placeholderDef
	txStyles     *xmlTxStyles // マスターレベルのデフォルトテキストスタイル
}

// inheritedStyle はプレースホルダーに対する継承済みスタイル
type inheritedStyle struct {
	xfrm *xmlXfrm
	// 継承チェーン上の lstStyle（優先度順: スライド自身 → レイアウト → マスター → マスターtxStyles）
	lstStyles []*xmlLstStyle
}

// getLevelPPr はレベルに対応する xmlLvlPPr を継承チェーンから探す
func (is *inheritedStyle) getLevelPPr(level int) *xmlLvlPPr {
	if is == nil {
		return nil
	}
	for _, ls := range is.lstStyles {
		if ppr := ls.GetLevel(level); ppr != nil {
			return ppr
		}
	}
	return nil
}

// getDefRPr はレベルに対応する defRPr を継承チェーンから探す
func (is *inheritedStyle) getDefRPr(level int) *xmlRPr {
	if is == nil {
		return nil
	}
	for _, ls := range is.lstStyles {
		if ppr := ls.GetLevel(level); ppr != nil && ppr.DefRPr != nil {
			return ppr.DefRPr
		}
	}
	return nil
}

// File にレイアウト/マスターキャッシュを追加するための拡張

// inheritCache はレイアウト/マスターのキャッシュ
type inheritCache struct {
	mu      sync.Mutex
	layouts map[string]*layoutData // layoutPath → parsed data
	masters map[string]*masterData // masterPath → parsed data
}

func newInheritCache() *inheritCache {
	return &inheritCache{
		layouts: make(map[string]*layoutData),
		masters: make(map[string]*masterData),
	}
}

// getLayout はレイアウトデータを取得する（キャッシュ付き）
func (ic *inheritCache) getLayout(f *File, layoutPath string) *layoutData {
	ic.mu.Lock()
	defer ic.mu.Unlock()

	if ld, ok := ic.layouts[layoutPath]; ok {
		return ld
	}

	ld := loadLayoutData(f, layoutPath)
	ic.layouts[layoutPath] = ld
	return ld
}

// getMaster はマスターデータを取得する（キャッシュ付き）
func (ic *inheritCache) getMaster(f *File, masterPath string) *masterData {
	ic.mu.Lock()
	defer ic.mu.Unlock()

	if md, ok := ic.masters[masterPath]; ok {
		return md
	}

	md := loadMasterData(f, masterPath)
	ic.masters[masterPath] = md
	return md
}

// loadLayoutData はスライドレイアウトXMLをパースする
func loadLayoutData(f *File, layoutPath string) *layoutData {
	ld := &layoutData{
		placeholders: make(map[phKey]*placeholderDef),
	}

	var layout xmlSldLayout
	if err := decodeZipXML(f.zi, layoutPath, &layout); err != nil {
		return ld
	}

	// プレースホルダーを収集
	collectPlaceholders(layout.CSld.SpTree.Children, ld.placeholders)

	// レイアウトの .rels からマスターパスを解決
	relsPath := relsPathFor(layoutPath)
	rels, err := loadRelsTyped(f, relsPath)
	if err == nil {
		for _, r := range rels {
			if strings.HasSuffix(r.Type, "/slideMaster") {
				ld.masterPath = resolveRelTarget(pathDir(layoutPath), r.Target)
				break
			}
		}
	}

	return ld
}

// loadMasterData はスライドマスターXMLをパースする
func loadMasterData(f *File, masterPath string) *masterData {
	md := &masterData{
		placeholders: make(map[phKey]*placeholderDef),
	}

	var master xmlSldMaster
	if err := decodeZipXML(f.zi, masterPath, &master); err != nil {
		return md
	}

	// プレースホルダーを収集
	collectPlaceholders(master.CSld.SpTree.Children, md.placeholders)
	md.txStyles = master.TxStyles

	return md
}

// collectPlaceholders は spTree の子要素からプレースホルダー定義を収集する
func collectPlaceholders(children []xmlSpTreeChild, out map[phKey]*placeholderDef) {
	for _, child := range children {
		if child.Sp == nil {
			continue
		}
		ph := child.Sp.NvSpPr.NvPr.Ph
		if ph == nil {
			continue
		}
		key := phKey{Type: ph.Type, Idx: ph.Idx}
		def := &placeholderDef{
			xfrm: child.Sp.SpPr.Xfrm,
		}
		if child.Sp.TxBody != nil {
			def.lstStyle = child.Sp.TxBody.LstStyle
		}
		out[key] = def
	}
}

// relsPathFor は任意のXMLパスから .rels パスを生成する
func relsPathFor(xmlPath string) string {
	dir := pathDir(xmlPath)
	base := pathBase(xmlPath)
	return dir + "/_rels/" + base + ".rels"
}

// resolveLayoutPath はスライドの .rels からレイアウトのZIPパスを解決する
func resolveLayoutPath(f *File, slidePath string) string {
	relsPath := slideRelsPath(slidePath)
	rels, err := loadRelsTyped(f, relsPath)
	if err != nil {
		return ""
	}
	for _, r := range rels {
		if strings.HasSuffix(r.Type, "/slideLayout") {
			return resolveRelTarget(pathDir(slidePath), r.Target)
		}
	}
	return ""
}

// resolveInheritedStyle はプレースホルダーの継承スタイルを解決する
func resolveInheritedStyle(ph *xmlPh, slideTxBodyLstStyle *xmlLstStyle, layout *layoutData, master *masterData, defaultTextStyle *xmlLstStyle) *inheritedStyle {
	if ph == nil {
		return nil
	}

	is := &inheritedStyle{}

	// スライド自身の txBody.lstStyle
	if slideTxBodyLstStyle != nil {
		is.lstStyles = append(is.lstStyles, slideTxBodyLstStyle)
	}

	phType := ph.Type
	if phType == "" {
		phType = "body"
	}
	key := phKey{Type: phType, Idx: ph.Idx}

	// レイアウトからプレースホルダーを検索
	if layout != nil {
		layoutDef := findPlaceholder(layout.placeholders, key)
		if layoutDef != nil {
			if is.xfrm == nil && layoutDef.xfrm != nil {
				is.xfrm = layoutDef.xfrm
			}
			if layoutDef.lstStyle != nil {
				is.lstStyles = append(is.lstStyles, layoutDef.lstStyle)
			}
		}
	}

	// マスターからプレースホルダーを検索
	if master != nil {
		masterDef := findPlaceholder(master.placeholders, key)
		if masterDef != nil {
			if is.xfrm == nil && masterDef.xfrm != nil {
				is.xfrm = masterDef.xfrm
			}
			if masterDef.lstStyle != nil {
				is.lstStyles = append(is.lstStyles, masterDef.lstStyle)
			}
		}

		// マスターの txStyles からフォールバック
		txStyleLst := masterTxStyleForPh(master.txStyles, phType)
		if txStyleLst != nil {
			is.lstStyles = append(is.lstStyles, txStyleLst)
		}
	}

	// presentation.xml の defaultTextStyle（最終フォールバック）
	if defaultTextStyle != nil {
		is.lstStyles = append(is.lstStyles, defaultTextStyle)
	}

	return is
}

// findPlaceholder はプレースホルダーマップから一致するものを探す。
// 完全一致を試み、見つからなければ type のみで再検索する。
func findPlaceholder(m map[phKey]*placeholderDef, key phKey) *placeholderDef {
	if m == nil {
		return nil
	}
	// 完全一致
	if def, ok := m[key]; ok {
		return def
	}
	// idx を無視して type のみでマッチ
	if key.Idx != "" {
		if def, ok := m[phKey{Type: key.Type}]; ok {
			return def
		}
	}
	return nil
}

// masterTxStyleForPh はプレースホルダー種別に対応する txStyles の lstStyle を返す
func masterTxStyleForPh(txStyles *xmlTxStyles, phType string) *xmlLstStyle {
	if txStyles == nil {
		return nil
	}
	switch phType {
	case "title", "ctrTitle":
		return txStyles.TitleStyle
	case "subTitle", "body":
		return txStyles.BodyStyle
	default:
		return txStyles.OtherStyle
	}
}
