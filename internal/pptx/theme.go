package pptx

import (
	"encoding/xml"
	"fmt"
)

// themeColors はテーマカラーとフォントのマッピング
type themeColors struct {
	colors    map[int]string // テーマインデックス → "#RRGGBB"
	majorFont string         // 見出しフォント（+mj-lt 解決用）
	minorFont string         // 本文フォント（+mn-lt 解決用）
}

// Get はテーマインデックスから色を取得する
func (tc *themeColors) Get(idx int) string {
	if tc == nil {
		return ""
	}
	return tc.colors[idx]
}

// parseThemeColors は theme1.xml からテーマカラーとフォントをパースする。
// パース失敗時はエラーを返す。
func parseThemeColors(data []byte) (*themeColors, error) {
	var theme struct {
		ThemeElements struct {
			ClrScheme struct {
				Dk1      xmlThemeColor `xml:"dk1"`
				Lt1      xmlThemeColor `xml:"lt1"`
				Dk2      xmlThemeColor `xml:"dk2"`
				Lt2      xmlThemeColor `xml:"lt2"`
				Accent1  xmlThemeColor `xml:"accent1"`
				Accent2  xmlThemeColor `xml:"accent2"`
				Accent3  xmlThemeColor `xml:"accent3"`
				Accent4  xmlThemeColor `xml:"accent4"`
				Accent5  xmlThemeColor `xml:"accent5"`
				Accent6  xmlThemeColor `xml:"accent6"`
				Hlink    xmlThemeColor `xml:"hlink"`
				FolHlink xmlThemeColor `xml:"folHlink"`
			} `xml:"clrScheme"`
			FontScheme struct {
				MajorFont xmlThemeFontSet `xml:"majorFont"`
				MinorFont xmlThemeFontSet `xml:"minorFont"`
			} `xml:"fontScheme"`
		} `xml:"themeElements"`
	}
	if err := xml.Unmarshal(data, &theme); err != nil {
		return nil, fmt.Errorf("テーマXMLのパースに失敗: %w", err)
	}

	cs := theme.ThemeElements.ClrScheme
	fs := theme.ThemeElements.FontScheme
	tc := &themeColors{colors: make(map[int]string, 12)}

	// PowerPoint テーマインデックスのマッピング:
	// 0=lt1(bg1), 1=dk1(tx1), 2=lt2(bg2), 3=dk2(tx2), 4-9=accent1-6, 10=hlink, 11=folHlink
	entries := []xmlThemeColor{
		cs.Lt1, cs.Dk1, cs.Lt2, cs.Dk2,
		cs.Accent1, cs.Accent2, cs.Accent3, cs.Accent4, cs.Accent5, cs.Accent6,
		cs.Hlink, cs.FolHlink,
	}
	for i, e := range entries {
		c := extractThemeColorValue(e)
		if c != "" {
			tc.colors[i] = c
		}
	}

	// テーマフォント（ea → latin の優先順で日本語環境対応）
	tc.majorFont = resolveThemeFontTypeface(fs.MajorFont)
	tc.minorFont = resolveThemeFontTypeface(fs.MinorFont)

	return tc, nil
}

// resolveThemeFontTypeface はテーマフォントセットからフォント名を解決する。
// ea（東アジア）を優先し、なければ latin を返す。
func resolveThemeFontTypeface(fs xmlThemeFontSet) string {
	if fs.Ea.Typeface != "" {
		return fs.Ea.Typeface
	}
	return fs.Latin.Typeface
}

// ResolveThemeFont はテーマフォント参照（+mj-lt, +mn-lt 等）を実際のフォント名に解決する。
// テーマフォント参照でない場合はそのまま返す。
func (tc *themeColors) ResolveThemeFont(typeface string) string {
	if tc == nil {
		return typeface
	}
	switch typeface {
	case "+mj-lt", "+mj-ea", "+mj-cs":
		if tc.majorFont != "" {
			return tc.majorFont
		}
	case "+mn-lt", "+mn-ea", "+mn-cs":
		if tc.minorFont != "" {
			return tc.minorFont
		}
	}
	return typeface
}

type xmlThemeColor struct {
	SrgbClr struct {
		Val string `xml:"val,attr"`
	} `xml:"srgbClr"`
	SysClr struct {
		LastClr string `xml:"lastClr,attr"`
		Val     string `xml:"val,attr"`
	} `xml:"sysClr"`
}

// xmlThemeFontSet は a:majorFont / a:minorFont 要素
type xmlThemeFontSet struct {
	Latin xmlFont `xml:"latin"`
	Ea    xmlFont `xml:"ea"`
}

func extractThemeColorValue(tc xmlThemeColor) string {
	if tc.SrgbClr.Val != "" {
		return normalizeHexColor(tc.SrgbClr.Val)
	}
	if tc.SysClr.LastClr != "" {
		return normalizeHexColor(tc.SysClr.LastClr)
	}
	return ""
}
