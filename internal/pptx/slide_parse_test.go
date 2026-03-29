package pptx

import (
	"testing"
)

// --- parseContext ヘルパーのテスト ---

func newTestContext() *parseContext {
	return &parseContext{
		f:         &File{},
		pptxIDMap: make(map[int]int),
	}
}

func TestAllocID(t *testing.T) {
	ctx := newTestContext()

	// 1つ目のID割り当て
	id1 := ctx.allocID(100)
	if id1 != 1 {
		t.Errorf("1つ目のID: got %d, want 1", id1)
	}

	// pptxIDMap に登録されること
	if mapped, ok := ctx.pptxIDMap[100]; !ok || mapped != 1 {
		t.Errorf("pptxIDMap[100]: got %d, %v, want 1, true", mapped, ok)
	}

	// 2つ目のID割り当て
	id2 := ctx.allocID(200)
	if id2 != 2 {
		t.Errorf("2つ目のID: got %d, want 2", id2)
	}

	// pptxID=0 の場合はマップに登録されない
	id3 := ctx.allocID(0)
	if id3 != 3 {
		t.Errorf("3つ目のID: got %d, want 3", id3)
	}
	if _, ok := ctx.pptxIDMap[0]; ok {
		t.Error("pptxID=0 はマップに登録されるべきではない")
	}
}

func TestAllocZ(t *testing.T) {
	ctx := newTestContext()

	z0 := ctx.allocZ()
	z1 := ctx.allocZ()
	z2 := ctx.allocZ()

	if z0 != 0 || z1 != 1 || z2 != 2 {
		t.Errorf("z-order: got %d, %d, %d, want 0, 1, 2", z0, z1, z2)
	}
}

func TestNewChildContext(t *testing.T) {
	ctx := newTestContext()
	ctx.nextID = 5
	ctx.nextZ = 3
	ctx.allocID(100) // nextID=6, pptxIDMap[100]=6

	child := ctx.newChildContext()

	// カウンタはコピーされる
	if child.nextID != ctx.nextID {
		t.Errorf("child.nextID: got %d, want %d", child.nextID, ctx.nextID)
	}
	if child.nextZ != ctx.nextZ {
		t.Errorf("child.nextZ: got %d, want %d", child.nextZ, ctx.nextZ)
	}

	// pptxIDMap は参照共有
	child.allocID(200) // child で登録
	if _, ok := ctx.pptxIDMap[200]; !ok {
		t.Error("子で登録したIDが親のpptxIDMapに反映されるべき")
	}
}

func TestSyncFromChild(t *testing.T) {
	ctx := newTestContext()
	ctx.nextID = 5
	ctx.nextZ = 3

	child := ctx.newChildContext()
	child.nextID = 10
	child.nextZ = 7

	ctx.syncFromChild(child)

	if ctx.nextID != 10 {
		t.Errorf("syncFromChild nextID: got %d, want 10", ctx.nextID)
	}
	if ctx.nextZ != 7 {
		t.Errorf("syncFromChild nextZ: got %d, want 7", ctx.nextZ)
	}
}

// --- ソート関連のテスト ---

func TestPhPriority(t *testing.T) {
	tests := []struct {
		ph   *xmlPh
		want int
	}{
		{nil, 99},
		{&xmlPh{Type: "title"}, 0},
		{&xmlPh{Type: "ctrTitle"}, 0},
		{&xmlPh{Type: "subTitle"}, 1},
		{&xmlPh{Type: "body"}, 2},
		{&xmlPh{Type: "ftr"}, 3},
		{&xmlPh{Type: ""}, 3},
	}

	for _, tt := range tests {
		name := "nil"
		if tt.ph != nil {
			name = tt.ph.Type
			if name == "" {
				name = "(empty)"
			}
		}
		t.Run(name, func(t *testing.T) {
			got := phPriority(tt.ph)
			if got != tt.want {
				t.Errorf("phPriority: got %d, want %d", got, tt.want)
			}
		})
	}
}

func TestSortShapeItems(t *testing.T) {
	items := []shapeItem{
		{order: 0, shape: Shape{ID: 1}, isPH: false},
		{order: 1, shape: Shape{ID: 2}, isPH: true, phPriority: 2},  // body
		{order: 2, shape: Shape{ID: 3}, isPH: true, phPriority: 0},  // title
		{order: 3, shape: Shape{ID: 4}, isPH: false},
		{order: 4, shape: Shape{ID: 5}, isPH: true, phPriority: 1},  // subTitle
	}

	sortShapeItems(items)

	// プレースホルダー（優先度順）→ 非プレースホルダー（出現順）
	wantIDs := []int{3, 5, 2, 1, 4}
	for i, want := range wantIDs {
		if items[i].shape.ID != want {
			t.Errorf("items[%d].ID: got %d, want %d", i, items[i].shape.ID, want)
		}
	}
}

func TestLessShapeItem(t *testing.T) {
	// プレースホルダーは非プレースホルダーより先
	ph := shapeItem{order: 1, isPH: true, phPriority: 2}
	nonPH := shapeItem{order: 0, isPH: false}
	if !lessShapeItem(ph, nonPH) {
		t.Error("プレースホルダーは非プレースホルダーより先であるべき")
	}
	if lessShapeItem(nonPH, ph) {
		t.Error("非プレースホルダーはプレースホルダーより後であるべき")
	}

	// 同じプレースホルダー間は優先度順
	phTitle := shapeItem{order: 1, isPH: true, phPriority: 0}
	phBody := shapeItem{order: 0, isPH: true, phPriority: 2}
	if !lessShapeItem(phTitle, phBody) {
		t.Error("title(優先度0) は body(優先度2) より先であるべき")
	}

	// 非プレースホルダー間は出現順
	a := shapeItem{order: 0, isPH: false}
	b := shapeItem{order: 1, isPH: false}
	if !lessShapeItem(a, b) {
		t.Error("出現順で先のものが先であるべき")
	}
}

// --- parseSp のテスト ---

func TestParseSp_TextOnly(t *testing.T) {
	ctx := newTestContext()

	sp := xmlSp{
		NvSpPr: xmlNvSpPr{
			CNvPr: xmlCNvPr{ID: 10, Name: "テキスト1"},
		},
		SpPr: xmlSpPr{
			PrstGeom: &xmlPrstGeom{Prst: "roundRect"},
			Xfrm: &xmlXfrm{
				Off: xmlOff{X: 100, Y: 200},
				Ext: xmlExt{Cx: 300, Cy: 400},
			},
		},
		TxBody: &xmlTxBody{
			Ps: []xmlP{
				{Rs: []xmlR{{T: "Hello"}}},
			},
		},
	}

	s := ctx.parseSp(sp)
	if s == nil {
		t.Fatal("parseSp がnilを返した")
	}

	if s.Type != "roundRect" {
		t.Errorf("Type: got %q, want %q", s.Type, "roundRect")
	}
	if s.Name != "テキスト1" {
		t.Errorf("Name: got %q, want %q", s.Name, "テキスト1")
	}
	if s.Pos == nil {
		t.Fatal("Pos がnil")
	}
	if s.Pos.X != 100 || s.Pos.Y != 200 {
		t.Errorf("Pos: got (%d,%d), want (100,200)", s.Pos.X, s.Pos.Y)
	}
	if len(s.Paragraphs) != 1 || s.Paragraphs[0].Text != "Hello" {
		t.Errorf("Paragraphs: got %v", s.Paragraphs)
	}
}

func TestParseSp_Empty_ReturnsNil(t *testing.T) {
	ctx := newTestContext()

	// テキストなし・塗りなし・枠線なし → nil
	sp := xmlSp{
		NvSpPr: xmlNvSpPr{
			CNvPr: xmlCNvPr{ID: 10},
		},
		SpPr: xmlSpPr{},
	}

	s := ctx.parseSp(sp)
	if s != nil {
		t.Error("テキスト・塗り・枠線のない図形はnilであるべき")
	}
}

func TestParseSp_FillOnly(t *testing.T) {
	ctx := newTestContext()

	sp := xmlSp{
		NvSpPr: xmlNvSpPr{
			CNvPr: xmlCNvPr{ID: 10},
		},
		SpPr: xmlSpPr{
			SolidFill: &xmlSolidFill{
				SrgbClr: &xmlSrgbClr{Val: "FF0000"},
			},
		},
	}

	s := ctx.parseSp(sp)
	if s == nil {
		t.Fatal("塗りのみの図形はnilであるべきではない")
	}
	if s.Fill != "#FF0000" {
		t.Errorf("Fill: got %q, want %q", s.Fill, "#FF0000")
	}
}

func TestParseSp_Placeholder(t *testing.T) {
	ctx := newTestContext()

	sp := xmlSp{
		NvSpPr: xmlNvSpPr{
			CNvPr: xmlCNvPr{ID: 10, Name: "Title"},
			NvPr:  xmlNvPr{Ph: &xmlPh{Type: "title"}},
		},
		SpPr: xmlSpPr{},
		TxBody: &xmlTxBody{
			Ps: []xmlP{{Rs: []xmlR{{T: "タイトル"}}}},
		},
	}

	s := ctx.parseSp(sp)
	if s == nil {
		t.Fatal("parseSp がnilを返した")
	}
	if s.Placeholder != "title" {
		t.Errorf("Placeholder: got %q, want %q", s.Placeholder, "title")
	}
	// プレースホルダーの場合、Name は設定されない
	if s.Name != "" {
		t.Errorf("プレースホルダーの場合 Name は空であるべき: got %q", s.Name)
	}
}

func TestParseSp_PlaceholderEmptyType(t *testing.T) {
	ctx := newTestContext()

	sp := xmlSp{
		NvSpPr: xmlNvSpPr{
			CNvPr: xmlCNvPr{ID: 10},
			NvPr:  xmlNvPr{Ph: &xmlPh{Type: ""}},
		},
		SpPr: xmlSpPr{},
		TxBody: &xmlTxBody{
			Ps: []xmlP{{Rs: []xmlR{{T: "本文"}}}},
		},
	}

	s := ctx.parseSp(sp)
	if s == nil {
		t.Fatal("parseSp がnilを返した")
	}
	// type未指定のプレースホルダーはbody扱い
	if s.Placeholder != "body" {
		t.Errorf("Placeholder: got %q, want %q", s.Placeholder, "body")
	}
}

func TestParseSp_Rotation(t *testing.T) {
	ctx := newTestContext()

	sp := xmlSp{
		NvSpPr: xmlNvSpPr{CNvPr: xmlCNvPr{ID: 10}},
		SpPr: xmlSpPr{
			Xfrm: &xmlXfrm{
				Rot:   5400000, // 90度
				FlipH: true,
				Off:   xmlOff{X: 0, Y: 0},
				Ext:   xmlExt{Cx: 100, Cy: 100},
			},
		},
		TxBody: &xmlTxBody{
			Ps: []xmlP{{Rs: []xmlR{{T: "回転"}}}},
		},
	}

	s := ctx.parseSp(sp)
	if s == nil {
		t.Fatal("parseSp がnilを返した")
	}
	if s.Rotation != 90.0 {
		t.Errorf("Rotation: got %f, want 90.0", s.Rotation)
	}
	if s.Flip != "h" {
		t.Errorf("Flip: got %q, want %q", s.Flip, "h")
	}
}

func TestParseSp_CustomGeom(t *testing.T) {
	ctx := newTestContext()

	sp := xmlSp{
		NvSpPr: xmlNvSpPr{CNvPr: xmlCNvPr{ID: 10}},
		SpPr: xmlSpPr{
			CustGeom: &struct{}{},
		},
		TxBody: &xmlTxBody{
			Ps: []xmlP{{Rs: []xmlR{{T: "カスタム"}}}},
		},
	}

	s := ctx.parseSp(sp)
	if s == nil {
		t.Fatal("parseSp がnilを返した")
	}
	if s.Type != "customShape" {
		t.Errorf("Type: got %q, want %q", s.Type, "customShape")
	}
}

func TestParseSp_DefaultRect(t *testing.T) {
	ctx := newTestContext()

	sp := xmlSp{
		NvSpPr: xmlNvSpPr{CNvPr: xmlCNvPr{ID: 10}},
		SpPr:   xmlSpPr{},
		TxBody: &xmlTxBody{
			Ps: []xmlP{{Rs: []xmlR{{T: "テスト"}}}},
		},
	}

	s := ctx.parseSp(sp)
	if s == nil {
		t.Fatal("parseSp がnilを返した")
	}
	if s.Type != "rect" {
		t.Errorf("Type: got %q, want %q", s.Type, "rect")
	}
}

// --- parseCxnSp のテスト ---

func TestParseCxnSp(t *testing.T) {
	ctx := newTestContext()

	cxn := xmlCxnSp{
		NvCxnSpPr: xmlNvCxnSpPr{
			CNvPr: xmlCNvPr{ID: 20, Name: "コネクタ1"},
			CNvCxnSpPr: xmlCNvCxnSpPr{
				StCxn:  &xmlCxnRef{ID: 100},
				EndCxn: &xmlCxnRef{ID: 200},
			},
		},
		SpPr: xmlSpPr{
			PrstGeom: &xmlPrstGeom{Prst: "straightConnector1"},
			Xfrm: &xmlXfrm{
				Off: xmlOff{X: 10, Y: 20},
				Ext: xmlExt{Cx: 300, Cy: 0},
			},
		},
	}

	s := ctx.parseCxnSp(cxn)
	if s == nil {
		t.Fatal("parseCxnSp がnilを返した")
	}

	if s.Type != "connector" {
		t.Errorf("Type: got %q, want %q", s.Type, "connector")
	}
	if s.ConnectorType != "straightConnector1" {
		t.Errorf("ConnectorType: got %q, want %q", s.ConnectorType, "straightConnector1")
	}
	if s.Name != "コネクタ1" {
		t.Errorf("Name: got %q, want %q", s.Name, "コネクタ1")
	}
	// 未解決のfrom/toは負値
	if s.From != -100 {
		t.Errorf("From: got %d, want -100", s.From)
	}
	if s.To != -200 {
		t.Errorf("To: got %d, want -200", s.To)
	}
}

func TestParseCxnSp_WithLabel(t *testing.T) {
	ctx := newTestContext()

	cxn := xmlCxnSp{
		NvCxnSpPr: xmlNvCxnSpPr{
			CNvPr: xmlCNvPr{ID: 20, Name: "コネクタ"},
		},
		SpPr: xmlSpPr{},
		TxBody: &xmlTxBody{
			Ps: []xmlP{
				{Rs: []xmlR{{T: "ラベル1"}}},
				{Rs: []xmlR{{T: "ラベル2"}}},
			},
		},
	}

	s := ctx.parseCxnSp(cxn)
	if s.Label != "ラベル1\nラベル2" {
		t.Errorf("Label: got %q, want %q", s.Label, "ラベル1\nラベル2")
	}
}

// --- parsePic のテスト ---

func TestParsePic(t *testing.T) {
	ctx := newTestContext()

	pic := xmlPic{
		NvPicPr: xmlNvPicPr{
			CNvPr: xmlCNvPr{ID: 30, Name: "画像1", Descr: "テスト画像"},
		},
		BlipFill: xmlBlipFill{
			Blip: xmlBlip{Embed: "rId2"},
		},
		SpPr: xmlSpPr{
			Xfrm: &xmlXfrm{
				Off: xmlOff{X: 500, Y: 600},
				Ext: xmlExt{Cx: 700, Cy: 800},
			},
		},
	}

	s := ctx.parsePic(pic)
	if s == nil {
		t.Fatal("parsePic がnilを返した")
	}

	if s.Type != "picture" {
		t.Errorf("Type: got %q, want %q", s.Type, "picture")
	}
	if s.AltText != "テスト画像" {
		t.Errorf("AltText: got %q, want %q", s.AltText, "テスト画像")
	}
	if s.Pos.X != 500 || s.Pos.W != 700 {
		t.Errorf("Pos: got (%d,%d), want (500,700)", s.Pos.X, s.Pos.W)
	}
	// extractDir が空なので ImagePath は空
	if s.ImagePath != "" {
		t.Error("extractDir が空の場合 ImagePath は空であるべき")
	}
}

// --- parseGrpSp のテスト ---

func TestParseGrpSp(t *testing.T) {
	ctx := newTestContext()

	grp := xmlGrpSp{
		NvGrpSpPr: xmlNvGrpSpPr{
			CNvPr: xmlCNvPr{ID: 40, Name: "グループ1"},
		},
		GrpSpPr: xmlGrpSpPr{
			Xfrm: &xmlGrpXfrm{
				Off: xmlOff{X: 10, Y: 20},
				Ext: xmlExt{Cx: 300, Cy: 400},
			},
		},
		Children: []xmlSpTreeChild{
			{
				Sp: &xmlSp{
					NvSpPr: xmlNvSpPr{CNvPr: xmlCNvPr{ID: 41}},
					SpPr:   xmlSpPr{},
					TxBody: &xmlTxBody{
						Ps: []xmlP{{Rs: []xmlR{{T: "子1"}}}},
					},
				},
			},
		},
	}

	s := ctx.parseGrpSp(grp)
	if s == nil {
		t.Fatal("parseGrpSp がnilを返した")
	}

	if s.Type != "group" {
		t.Errorf("Type: got %q, want %q", s.Type, "group")
	}
	if s.Pos.X != 10 || s.Pos.H != 400 {
		t.Errorf("Pos: got (%d,%d), want (10,400)", s.Pos.X, s.Pos.H)
	}
	if len(s.Children) != 1 {
		t.Fatalf("Children: got %d, want 1", len(s.Children))
	}
	if s.Children[0].Paragraphs[0].Text != "子1" {
		t.Errorf("子要素のテキスト: got %q, want %q", s.Children[0].Paragraphs[0].Text, "子1")
	}
}

func TestParseGrpSp_EmptyChildren_ReturnsNil(t *testing.T) {
	ctx := newTestContext()

	grp := xmlGrpSp{
		NvGrpSpPr: xmlNvGrpSpPr{
			CNvPr: xmlCNvPr{ID: 40},
		},
		GrpSpPr: xmlGrpSpPr{},
	}

	s := ctx.parseGrpSp(grp)
	if s != nil {
		t.Error("子要素がないグループはnilであるべき")
	}
}

// --- parseGraphicFrame (テーブル) のテスト ---

func TestParseGraphicFrame_SimpleTable(t *testing.T) {
	ctx := newTestContext()

	gf := xmlGraphicFrame{
		NvGraphicFramePr: xmlNvGraphicFramePr{
			CNvPr: xmlCNvPr{ID: 50, Name: "テーブル1"},
		},
		Xfrm: &xmlXfrm{
			Off: xmlOff{X: 100, Y: 200},
			Ext: xmlExt{Cx: 500, Cy: 300},
		},
		Graphic: xmlGraphic{
			GraphicData: xmlGraphicData{
				Tbl: &xmlTbl{
					TblGrid: xmlTblGrid{
						GridCols: []xmlGridCol{{W: 100}, {W: 200}},
					},
					Trs: []xmlTr{
						{
							Tcs: []xmlTc{
								{TxBody: &xmlTxBody{Ps: []xmlP{{Rs: []xmlR{{T: "A1"}}}}}},
								{TxBody: &xmlTxBody{Ps: []xmlP{{Rs: []xmlR{{T: "B1"}}}}}},
							},
						},
						{
							Tcs: []xmlTc{
								{TxBody: &xmlTxBody{Ps: []xmlP{{Rs: []xmlR{{T: "A2"}}}}}},
								{TxBody: &xmlTxBody{Ps: []xmlP{{Rs: []xmlR{{T: "B2"}}}}}},
							},
						},
					},
				},
			},
		},
	}

	s := ctx.parseGraphicFrame(gf)
	if s == nil {
		t.Fatal("parseGraphicFrame がnilを返した")
	}

	if s.Type != "table" {
		t.Errorf("Type: got %q, want %q", s.Type, "table")
	}
	if s.Table.Cols != 2 {
		t.Errorf("Cols: got %d, want 2", s.Table.Cols)
	}
	if len(s.Table.Rows) != 2 {
		t.Fatalf("Rows: got %d, want 2", len(s.Table.Rows))
	}
	if *s.Table.Rows[0][0] != "A1" {
		t.Errorf("Rows[0][0]: got %q, want %q", *s.Table.Rows[0][0], "A1")
	}
	if *s.Table.Rows[1][1] != "B2" {
		t.Errorf("Rows[1][1]: got %q, want %q", *s.Table.Rows[1][1], "B2")
	}
}

func TestParseGraphicFrame_MergedCells(t *testing.T) {
	ctx := newTestContext()

	gf := xmlGraphicFrame{
		NvGraphicFramePr: xmlNvGraphicFramePr{
			CNvPr: xmlCNvPr{ID: 50},
		},
		Graphic: xmlGraphic{
			GraphicData: xmlGraphicData{
				Tbl: &xmlTbl{
					TblGrid: xmlTblGrid{
						GridCols: []xmlGridCol{{W: 100}, {W: 100}, {W: 100}},
					},
					Trs: []xmlTr{
						{
							Tcs: []xmlTc{
								// A1: colSpan=2
								{GridSpan: 2, TxBody: &xmlTxBody{Ps: []xmlP{{Rs: []xmlR{{T: "A1-B1"}}}}}},
								{TxBody: &xmlTxBody{Ps: []xmlP{{Rs: []xmlR{{T: "C1"}}}}}},
							},
						},
						{
							Tcs: []xmlTc{
								{TxBody: &xmlTxBody{Ps: []xmlP{{Rs: []xmlR{{T: "A2"}}}}}},
								{TxBody: &xmlTxBody{Ps: []xmlP{{Rs: []xmlR{{T: "B2"}}}}}},
								{TxBody: &xmlTxBody{Ps: []xmlP{{Rs: []xmlR{{T: "C2"}}}}}},
							},
						},
					},
				},
			},
		},
	}

	s := ctx.parseGraphicFrame(gf)
	if s == nil {
		t.Fatal("parseGraphicFrame がnilを返した")
	}

	// A1-B1 は col 0, C1 は col 2 (colSpan=2 なので)
	if *s.Table.Rows[0][0] != "A1-B1" {
		t.Errorf("Rows[0][0]: got %q, want %q", *s.Table.Rows[0][0], "A1-B1")
	}
	// col 1 は colSpan で飛ばされて nil
	if s.Table.Rows[0][1] != nil {
		t.Errorf("Rows[0][1]: got %v, want nil (colSpan による被結合セル)", s.Table.Rows[0][1])
	}
	if *s.Table.Rows[0][2] != "C1" {
		t.Errorf("Rows[0][2]: got %q, want %q", *s.Table.Rows[0][2], "C1")
	}
}

func TestParseGraphicFrame_RowSpan(t *testing.T) {
	ctx := newTestContext()

	gf := xmlGraphicFrame{
		NvGraphicFramePr: xmlNvGraphicFramePr{
			CNvPr: xmlCNvPr{ID: 50},
		},
		Graphic: xmlGraphic{
			GraphicData: xmlGraphicData{
				Tbl: &xmlTbl{
					TblGrid: xmlTblGrid{
						GridCols: []xmlGridCol{{W: 100}, {W: 100}},
					},
					Trs: []xmlTr{
						{
							Tcs: []xmlTc{
								// A1: rowSpan=2
								{RowSpan: 2, TxBody: &xmlTxBody{Ps: []xmlP{{Rs: []xmlR{{T: "A1"}}}}}},
								{TxBody: &xmlTxBody{Ps: []xmlP{{Rs: []xmlR{{T: "B1"}}}}}},
							},
						},
						{
							Tcs: []xmlTc{
								// A2 は vMerge で被結合
								{VMerge: "1", TxBody: &xmlTxBody{Ps: []xmlP{{Rs: []xmlR{{T: ""}}}}}},
								{TxBody: &xmlTxBody{Ps: []xmlP{{Rs: []xmlR{{T: "B2"}}}}}},
							},
						},
					},
				},
			},
		},
	}

	s := ctx.parseGraphicFrame(gf)
	if s == nil {
		t.Fatal("parseGraphicFrame がnilを返した")
	}

	if *s.Table.Rows[0][0] != "A1" {
		t.Errorf("Rows[0][0]: got %q, want %q", *s.Table.Rows[0][0], "A1")
	}
	// A2 は vMerge="1" なので nil
	if s.Table.Rows[1][0] != nil {
		t.Errorf("Rows[1][0]: got %v, want nil (vMerge による被結合セル)", s.Table.Rows[1][0])
	}
	if *s.Table.Rows[1][1] != "B2" {
		t.Errorf("Rows[1][1]: got %q, want %q", *s.Table.Rows[1][1], "B2")
	}
}

func TestParseGraphicFrame_NoTable_ReturnsNil(t *testing.T) {
	ctx := newTestContext()

	gf := xmlGraphicFrame{
		NvGraphicFramePr: xmlNvGraphicFramePr{
			CNvPr: xmlCNvPr{ID: 50},
		},
		Graphic: xmlGraphic{
			GraphicData: xmlGraphicData{
				Tbl: nil, // テーブルなし
			},
		},
	}

	s := ctx.parseGraphicFrame(gf)
	if s != nil {
		t.Error("テーブルのないgraphicFrameはnilであるべき")
	}
}

// --- resolveConnectors のテスト ---

func TestResolveConnectors(t *testing.T) {
	ctx := newTestContext()
	ctx.pptxIDMap[100] = 1
	ctx.pptxIDMap[200] = 2

	shapes := []Shape{
		{ID: 1, Type: "rect"},
		{ID: 2, Type: "rect"},
		{ID: 3, Type: "connector", From: -100, To: -200},
	}

	ctx.resolveConnectors(shapes)

	if shapes[2].From != 1 {
		t.Errorf("From: got %d, want 1", shapes[2].From)
	}
	if shapes[2].To != 2 {
		t.Errorf("To: got %d, want 2", shapes[2].To)
	}
}

func TestResolveConnectors_UnknownID(t *testing.T) {
	ctx := newTestContext()

	shapes := []Shape{
		{ID: 1, Type: "connector", From: -999, To: -888},
	}

	ctx.resolveConnectors(shapes)

	// 未知のIDは0にリセット
	if shapes[0].From != 0 {
		t.Errorf("From: got %d, want 0", shapes[0].From)
	}
	if shapes[0].To != 0 {
		t.Errorf("To: got %d, want 0", shapes[0].To)
	}
}

func TestResolveConnectors_InGroup(t *testing.T) {
	ctx := newTestContext()
	ctx.pptxIDMap[100] = 1

	shapes := []Shape{
		{
			ID:   2,
			Type: "group",
			Children: []Shape{
				{ID: 1, Type: "rect"},
				{ID: 3, Type: "connector", From: -100, To: 0},
			},
		},
	}

	ctx.resolveConnectors(shapes)

	if shapes[0].Children[1].From != 1 {
		t.Errorf("グループ内コネクタの From: got %d, want 1", shapes[0].Children[1].From)
	}
}

// --- parseSpTree のテスト ---

func TestParseSpTree(t *testing.T) {
	ctx := newTestContext()

	children := []xmlSpTreeChild{
		{
			Sp: &xmlSp{
				NvSpPr: xmlNvSpPr{CNvPr: xmlCNvPr{ID: 1}},
				SpPr:   xmlSpPr{},
				TxBody: &xmlTxBody{Ps: []xmlP{{Rs: []xmlR{{T: "テキスト"}}}}},
			},
		},
		{
			Pic: &xmlPic{
				NvPicPr: xmlNvPicPr{CNvPr: xmlCNvPr{ID: 2, Name: "画像"}},
				SpPr:    xmlSpPr{Xfrm: &xmlXfrm{Off: xmlOff{X: 0, Y: 0}, Ext: xmlExt{Cx: 100, Cy: 100}}},
			},
		},
	}

	shapes := ctx.parseSpTree(children)
	if len(shapes) != 2 {
		t.Fatalf("shapes: got %d, want 2", len(shapes))
	}

	// z-order はXML出現順
	if shapes[0].Z != 0 {
		t.Errorf("shapes[0].Z: got %d, want 0", shapes[0].Z)
	}
	if shapes[1].Z != 1 {
		t.Errorf("shapes[1].Z: got %d, want 1", shapes[1].Z)
	}
}

func TestParseSpTree_PlaceholdersSortedFirst(t *testing.T) {
	ctx := newTestContext()

	children := []xmlSpTreeChild{
		// 非プレースホルダー（出現順0）
		{
			Sp: &xmlSp{
				NvSpPr: xmlNvSpPr{CNvPr: xmlCNvPr{ID: 1}},
				SpPr:   xmlSpPr{},
				TxBody: &xmlTxBody{Ps: []xmlP{{Rs: []xmlR{{T: "通常"}}}}},
			},
		},
		// タイトルプレースホルダー（出現順1）
		{
			Sp: &xmlSp{
				NvSpPr: xmlNvSpPr{
					CNvPr: xmlCNvPr{ID: 2},
					NvPr:  xmlNvPr{Ph: &xmlPh{Type: "title"}},
				},
				SpPr:   xmlSpPr{},
				TxBody: &xmlTxBody{Ps: []xmlP{{Rs: []xmlR{{T: "タイトル"}}}}},
			},
		},
	}

	shapes := ctx.parseSpTree(children)
	if len(shapes) != 2 {
		t.Fatalf("shapes: got %d, want 2", len(shapes))
	}

	// タイトルが先に来る（出力順）
	if shapes[0].Placeholder != "title" {
		t.Errorf("shapes[0] はタイトルプレースホルダーであるべき: got placeholder=%q", shapes[0].Placeholder)
	}

	// z-order は元のXML出現順を保持
	// 通常図形(出現順0) → z=0, タイトル(出現順1) → z=1
	if shapes[0].Z != 1 { // タイトルは出現順1なのでz=1
		t.Errorf("shapes[0].Z: got %d, want 1", shapes[0].Z)
	}
	if shapes[1].Z != 0 { // 通常は出現順0なのでz=0
		t.Errorf("shapes[1].Z: got %d, want 0", shapes[1].Z)
	}
}
