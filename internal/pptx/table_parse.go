package pptx

// parseGraphicFrame はテーブル等のgraphicFrameをパースする
func (ctx *parseContext) parseGraphicFrame(gf xmlGraphicFrame) *Shape {
	if gf.NvGraphicFramePr.CNvPr.Hidden {
		return nil
	}

	tbl := gf.Graphic.GraphicData.Tbl
	if tbl == nil {
		return nil // テーブル以外のgraphicFrameはスキップ
	}

	s := &Shape{
		ID:   ctx.allocID(gf.NvGraphicFramePr.CNvPr.ID),
		Type: "table",
		Name: gf.NvGraphicFramePr.CNvPr.Name,
	}

	// 位置
	s.Pos = xfrmToPosition(gf.Xfrm)

	// テーブルデータ（被結合セルは null）
	s.Table = ctx.parseTableData(tbl)

	return s
}

// parseTableData はXMLテーブルからTableDataを構築する
func (ctx *parseContext) parseTableData(tbl *xmlTbl) *TableData {
	cols := len(tbl.TblGrid.GridCols)
	var rows [][]*TableCell

	// rowSpan による被結合セルを後のパスで null にするための記録
	type rowSpanArea struct {
		row, col, rowSpan, colSpan int
	}
	var rowSpans []rowSpanArea

	for _, tr := range tbl.Trs {
		row := make([]*TableCell, cols)
		colIdx := 0
		for _, tc := range tr.Tcs {
			if colIdx >= cols {
				break
			}
			if tc.VMerge != "1" && tc.HMerge != "1" {
				cell := ctx.parseTableCell(tc.TxBody)
				row[colIdx] = cell
			}
			span := tc.GridSpan
			if span < 1 {
				span = 1
			}
			if tc.RowSpan > 1 {
				rowSpans = append(rowSpans, rowSpanArea{
					row: len(rows), col: colIdx, rowSpan: tc.RowSpan, colSpan: span,
				})
			}
			colIdx += span
		}
		rows = append(rows, row)
	}

	// rowSpan による被結合セルを null にする
	// （標準XMLでは vMerge で既に null だが、vMerge 省略時のフォールバック）
	for _, rs := range rowSpans {
		for r := rs.row + 1; r < rs.row+rs.rowSpan && r < len(rows); r++ {
			for c := rs.col; c < rs.col+rs.colSpan && c < cols && c < len(rows[r]); c++ {
				rows[r][c] = nil
			}
		}
	}

	return &TableData{
		Cols: cols,
		Rows: rows,
	}
}

// parseTableCell はテーブルセルのテキストと段落情報をパースする。
// リッチテキスト情報がある場合のみ paragraphs を設定する。
func (ctx *parseContext) parseTableCell(txBody *xmlTxBody) *TableCell {
	if txBody == nil {
		return &TableCell{}
	}

	text := extractTextFromTxBody(txBody)
	paras := ctx.parseParagraphs(txBody.Ps, nil)

	cell := &TableCell{Text: text}

	// フォント・リッチテキスト・リンク・箇条書き等の情報がある場合のみ paragraphs を設定
	for _, p := range paras {
		if p.Font != nil || len(p.RichText) > 0 || p.Link != nil || p.Bullet != "" || p.Level > 0 {
			cell.Paragraphs = paras
			return cell
		}
	}

	return cell
}
