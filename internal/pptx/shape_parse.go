package pptx

// parseSp は通常の図形をパースする
func (ctx *parseContext) parseSp(sp xmlSp) *Shape {
	if sp.NvSpPr.CNvPr.Hidden {
		return nil
	}

	ph := sp.NvSpPr.NvPr.Ph

	// プレースホルダーの継承スタイルを解決
	var inherited *inheritedStyle
	if ph != nil {
		var slideLstStyle *xmlLstStyle
		if sp.TxBody != nil {
			slideLstStyle = sp.TxBody.LstStyle
		}
		inherited = resolveInheritedStyle(ph, slideLstStyle, ctx.layout, ctx.master, ctx.f.defaultTextStyle)
	}

	// テキストを先にパースし、結果で有無を判定（extractParagraphText の重複呼び出しを回避）
	var paras []Paragraph
	if sp.TxBody != nil {
		paras = ctx.parseParagraphs(sp.TxBody.Ps, inherited)
	}

	// テキスト・塗りつぶし・枠線のいずれもない図形はスキップ（プレースホルダー含む）
	hasText := len(paras) > 0
	hasFill := sp.SpPr.SolidFill != nil || sp.SpPr.GradFill != nil
	hasLine := sp.SpPr.Ln != nil && sp.SpPr.Ln.NoFill == nil
	if !hasText && !hasFill && !hasLine {
		return nil
	}

	s := &Shape{
		ID: ctx.allocID(sp.NvSpPr.CNvPr.ID),
	}

	// 図形全体のハイパーリンク
	s.Link = ctx.resolveHyperlink(sp.NvSpPr.CNvPr.HlinkClick)

	// 図形種別
	if sp.SpPr.PrstGeom != nil {
		s.Type = sp.SpPr.PrstGeom.Prst
	} else if sp.SpPr.CustGeom != nil {
		s.Type = "customShape"
	} else {
		s.Type = "rect" // デフォルト
	}

	// 名前とプレースホルダー
	if ph != nil {
		s.Placeholder = ph.Type
		if s.Placeholder == "" {
			s.Placeholder = "body" // type未指定のプレースホルダーはbody扱い
		}
	} else {
		s.Name = sp.NvSpPr.CNvPr.Name
	}

	// 位置（スライド上で未指定の場合、レイアウト/マスターから継承）
	xfrm := sp.SpPr.Xfrm
	if xfrm == nil && inherited != nil {
		xfrm = inherited.xfrm
	}
	s.Pos = xfrmToPosition(xfrm)

	// 回転・反転
	if xfrm != nil {
		s.Rotation = float64(xfrm.Rot) / 60000.0
		s.Flip = xfrmFlip(xfrm)
	}

	// 塗りつぶし（solidFill 優先、なければ gradFill の代表色）
	s.Fill = ctx.resolveFillColor(&sp.SpPr)

	// 枠線
	s.Line = ctx.resolveLine(sp.SpPr.Ln)

	// 吹き出しポインタ
	s.CalloutPointer = resolveCalloutPointer(sp.SpPr.PrstGeom, s.Pos)

	// テキスト（パース済みの結果を使用）
	if sp.TxBody != nil {
		s.Paragraphs = paras
		s.Alignment = ctx.extractShapeLevelAlignment(sp.TxBody)
		s.TextMargin = extractTextMargin(sp.TxBody.BodyPr)
	}

	return s
}

// parsePic は画像をパースする
func (ctx *parseContext) parsePic(pic xmlPic) *Shape {
	if pic.NvPicPr.CNvPr.Hidden {
		return nil
	}

	s := &Shape{
		ID:   ctx.allocID(pic.NvPicPr.CNvPr.ID),
		Type: "picture",
		Name: pic.NvPicPr.CNvPr.Name,
	}

	// 図形全体のハイパーリンク
	s.Link = ctx.resolveHyperlink(pic.NvPicPr.CNvPr.HlinkClick)

	// 代替テキスト
	s.AltText = pic.NvPicPr.CNvPr.Descr

	// 位置
	s.Pos = xfrmToPosition(pic.SpPr.Xfrm)

	// 画像IDの解決
	if pic.BlipFill.Blip.Embed != "" {
		s.ImageID = ctx.resolveImagePath(pic.BlipFill.Blip.Embed)
	}

	return s
}

// parseGrpSp はグループをパースする
func (ctx *parseContext) parseGrpSp(grp xmlGrpSp) *Shape {
	if grp.NvGrpSpPr.CNvPr.Hidden {
		return nil
	}

	s := &Shape{
		ID:   ctx.allocID(grp.NvGrpSpPr.CNvPr.ID),
		Type: "group",
		Name: grp.NvGrpSpPr.CNvPr.Name,
	}

	// グループの位置
	if grp.GrpSpPr.Xfrm != nil {
		s.Pos = &Position{
			X: grp.GrpSpPr.Xfrm.Off.X,
			Y: grp.GrpSpPr.Xfrm.Off.Y,
			W: grp.GrpSpPr.Xfrm.Ext.Cx,
			H: grp.GrpSpPr.Xfrm.Ext.Cy,
		}
	}

	// 子要素のパース
	childCtx := ctx.newChildContext()
	s.Children = childCtx.parseSpTree(grp.Children)
	ctx.syncFromChild(childCtx)

	if len(s.Children) == 0 {
		return nil
	}

	// 子要素の座標をグループローカル座標から絶対座標に変換
	if grp.GrpSpPr.Xfrm != nil {
		xfrm := grp.GrpSpPr.Xfrm
		transformGroupChildren(s.Children, xfrm)
	}

	return s
}

// resolveHyperlink は xmlHlinkClick からハイパーリンク情報を解決する
func (ctx *parseContext) resolveHyperlink(hlink *xmlHlinkClick) *HyperlinkData {
	if hlink == nil || hlink.RID == "" {
		return nil
	}
	if ctx.slideRels == nil {
		return nil
	}
	target, ok := ctx.slideRels[hlink.RID]
	if !ok || target == "" {
		return nil
	}

	// スライド内リンク（action が ppaction://hlinksldjump の場合）
	if hlink.Action == "ppaction://hlinksldjump" {
		slidePath := resolveRelTarget(pathDir(ctx.slidePath), target)
		slideNum := ctx.f.slidePathToNum(slidePath)
		if slideNum > 0 {
			return &HyperlinkData{Slide: slideNum}
		}
		return nil
	}

	// 外部リンク
	return &HyperlinkData{URL: target}
}
