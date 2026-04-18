package pptx

// LoadSlideInfos は全スライドの概要情報を取得する（info コマンド用）
func (f *File) LoadSlideInfos() ([]SlideInfo, error) {
	infos := make([]SlideInfo, 0, len(f.slideEntries))

	for i := range f.slideEntries {
		sld, err := f.loadSlideXML(i)
		if err != nil {
			return nil, err
		}
		if sld == nil {
			continue
		}
		infos = append(infos, SlideInfo{
			Slide:    i + 1,
			Title:    extractTitle(sld.CSld.SpTree.Children),
			Hidden:   sld.Show == "0",
			HasNotes: f.hasNotes(i),
		})
	}

	return infos, nil
}
