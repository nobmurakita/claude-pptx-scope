package pptx

import "encoding/xml"

// themeColors はテーマカラーのマッピング
type themeColors struct {
	colors map[int]string // テーマインデックス → "#RRGGBB"
}

// Get はテーマインデックスから色を取得する
func (tc *themeColors) Get(idx int) string {
	if tc == nil {
		return ""
	}
	return tc.colors[idx]
}

// parseThemeColors は theme1.xml からテーマカラーをパースする
func parseThemeColors(data []byte) *themeColors {
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
		} `xml:"themeElements"`
	}
	if err := xml.Unmarshal(data, &theme); err != nil {
		return nil
	}

	cs := theme.ThemeElements.ClrScheme
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
	return tc
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

func extractThemeColorValue(tc xmlThemeColor) string {
	if tc.SrgbClr.Val != "" {
		return normalizeHexColor(tc.SrgbClr.Val)
	}
	if tc.SysClr.LastClr != "" {
		return normalizeHexColor(tc.SysClr.LastClr)
	}
	return ""
}
