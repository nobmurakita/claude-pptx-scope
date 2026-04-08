package pptx

import (
	"testing"
)

// mkSpWithText はテスト用にテキストを持つ図形の xmlSpTreeChild を生成する
func mkSpWithText(text string) xmlSpTreeChild {
	return xmlSpTreeChild{
		Sp: &xmlSp{
			TxBody: &xmlTxBody{
				Ps: []xmlP{{Elements: []xmlParagraphElement{{R: &xmlR{T: text}}}}},
			},
		},
	}
}

func TestMatchSpTreeText(t *testing.T) {
	children := []xmlSpTreeChild{
		mkSpWithText("Hello World"),
		mkSpWithText("Goodbye"),
	}

	if !matchSpTreeText(children, "hello") {
		t.Error("expected match for 'hello'")
	}
	if matchSpTreeText(children, "missing") {
		t.Error("expected no match for 'missing'")
	}
}

func TestMatchSpTreeText_Table(t *testing.T) {
	children := []xmlSpTreeChild{
		{
			GraphicFrame: &xmlGraphicFrame{
				Graphic: xmlGraphic{
					GraphicData: xmlGraphicData{
						Tbl: &xmlTbl{
							TblGrid: xmlTblGrid{GridCols: []xmlGridCol{{W: 100}, {W: 100}}},
							Trs: []xmlTr{
								{Tcs: []xmlTc{
									{TxBody: &xmlTxBody{Ps: []xmlP{{Elements: []xmlParagraphElement{{R: &xmlR{T: "Alpha"}}}}}}},
									{TxBody: &xmlTxBody{Ps: []xmlP{{Elements: []xmlParagraphElement{{R: &xmlR{T: "Beta"}}}}}}},
								}},
							},
						},
					},
				},
			},
		},
	}

	if !matchSpTreeText(children, "beta") {
		t.Error("expected match for 'beta'")
	}
	if matchSpTreeText(children, "delta") {
		t.Error("expected no match for 'delta'")
	}
}

func TestMatchSpTreeText_Connector(t *testing.T) {
	children := []xmlSpTreeChild{
		{
			CxnSp: &xmlCxnSp{
				TxBody: &xmlTxBody{
					Ps: []xmlP{{Elements: []xmlParagraphElement{{R: &xmlR{T: "接続ラベル"}}}}},
				},
			},
		},
	}

	if !matchSpTreeText(children, "接続") {
		t.Error("expected match for '接続'")
	}
	if matchSpTreeText(children, "存在しない") {
		t.Error("expected no match for '存在しない'")
	}
}

func TestMatchSpTreeText_Group(t *testing.T) {
	children := []xmlSpTreeChild{
		{
			GrpSp: &xmlGrpSp{
				Children: []xmlSpTreeChild{
					mkSpWithText("内部テキスト"),
					mkSpWithText("別のテキスト"),
				},
			},
		},
	}

	if !matchSpTreeText(children, "内部") {
		t.Error("expected match for '内部'")
	}
	if matchSpTreeText(children, "外部") {
		t.Error("expected no match for '外部'")
	}
}

func TestMatchTxBodyText(t *testing.T) {
	txBody := &xmlTxBody{
		Ps: []xmlP{
			{Elements: []xmlParagraphElement{{R: &xmlR{T: "本文テキスト"}}}},
			{Elements: []xmlParagraphElement{{R: &xmlR{T: "ノートのテキスト"}}}},
		},
	}

	if !matchTxBodyText(txBody, "ノート") {
		t.Error("expected match for 'ノート'")
	}
	if !matchTxBodyText(txBody, "本文") {
		t.Error("expected match for '本文'")
	}
	if matchTxBodyText(txBody, "存在しない") {
		t.Error("expected no match for '存在しない'")
	}
	if matchTxBodyText(nil, "test") {
		t.Error("expected no match for nil txBody")
	}
}
