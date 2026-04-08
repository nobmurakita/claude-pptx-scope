package cmd

import (
	"encoding/json"
	"fmt"

	"github.com/nobmurakita/claude-pptx-scope/internal/pptx"
	"github.com/spf13/cobra"
)

// NewSlidesCmd は slides サブコマンドを生成する
func NewSlidesCmd() *cobra.Command {
	cmd := &cobra.Command{
		Use:   "slides <file>",
		Short: "スライドの内容をJSONL形式で出力する",
		Args:  cobra.ExactArgs(1),
		RunE:  runSlides,
	}
	cmd.Flags().IntSlice("slide", nil, "対象スライド番号（1始まり、複数指定可: --slide 1,3）")
	cmd.Flags().Bool("notes", false, "ノートも出力する")
	return cmd
}

// notesOutput はノートの出力行
type notesOutput struct {
	Notes []pptx.Paragraph `json:"notes"`
}

func runSlides(cmd *cobra.Command, args []string) error {
	slideNums, err := cmd.Flags().GetIntSlice("slide")
	if err != nil {
		return fmt.Errorf("--slide フラグの解析エラー: %w", err)
	}
	includeNotes, err := cmd.Flags().GetBool("notes")
	if err != nil {
		return fmt.Errorf("--notes フラグの解析エラー: %w", err)
	}

	f, err := pptx.OpenFile(args[0])
	if err != nil {
		return err
	}
	defer f.Close()

	targets, err := f.ResolveSlideNums(slideNums)
	if err != nil {
		return err
	}

	ow, err := newOutputWriter(cmd)
	if err != nil {
		return err
	}
	defer ow.cleanup()

	enc := newJSONEncoder(ow)
	dedup := pptx.NewStyleDeduplicator()

	for _, n := range targets {
		if err := emitSlide(f, enc, dedup, n, includeNotes); err != nil {
			return err
		}
	}
	return ow.finalize()
}

func emitSlide(f *pptx.File, enc *json.Encoder, dedup *pptx.StyleDeduplicator, slideNum int, includeNotes bool) error {
	sd, err := f.LoadSlide(slideNum, includeNotes)
	if err != nil {
		return err
	}
	return emitSlideData(enc, dedup, sd)
}

func emitSlideData(enc *json.Encoder, dedup *pptx.StyleDeduplicator, sd *pptx.SlideData) error {
	// フォントスタイルの重複排除（スライド横断で共有）
	newStyles := dedup.Deduplicate(sd)

	// 新規スタイル定義を初回使用の直前に出力するための準備
	pending := make(map[int]pptx.StyleDef, len(newStyles))
	for _, s := range newStyles {
		pending[s.ID] = s
	}

	// スライドヘッダ行
	header := sd.Info()
	header.Shapes = intPtr(len(sd.Shapes))
	if err := enc.Encode(header); err != nil {
		return fmt.Errorf("JSON出力エラー: %w", err)
	}

	// 図形を1つずつ個別の行として出力（使用するスタイル定義を直前に挿入）
	for i := range sd.Shapes {
		if err := emitPendingStyles(enc, pending, &sd.Shapes[i]); err != nil {
			return err
		}
		if err := enc.Encode(sd.Shapes[i]); err != nil {
			return fmt.Errorf("JSON出力エラー: %w", err)
		}
	}

	// ノートで使用するスタイル定義を出力
	if len(sd.Notes) > 0 {
		if err := emitPendingParaStyles(enc, pending, sd.Notes); err != nil {
			return err
		}
		if err := enc.Encode(notesOutput{Notes: sd.Notes}); err != nil {
			return fmt.Errorf("JSON出力エラー: %w", err)
		}
	}

	return nil
}

// emitPendingStyles は図形が参照するスタイル定義のうち未出力のものを出力する
func emitPendingStyles(enc *json.Encoder, pending map[int]pptx.StyleDef, shape *pptx.Shape) error {
	if len(pending) == 0 {
		return nil
	}
	for _, id := range collectShapeStyleIDs(shape) {
		if s, ok := pending[id]; ok {
			if err := enc.Encode(s); err != nil {
				return fmt.Errorf("スタイル定義の出力エラー: %w", err)
			}
			delete(pending, id)
		}
	}
	return nil
}

// emitPendingParaStyles は段落群が参照するスタイル定義のうち未出力のものを出力する
func emitPendingParaStyles(enc *json.Encoder, pending map[int]pptx.StyleDef, paras []pptx.Paragraph) error {
	if len(pending) == 0 {
		return nil
	}
	for _, id := range collectParaStyleIDs(paras) {
		if s, ok := pending[id]; ok {
			if err := enc.Encode(s); err != nil {
				return fmt.Errorf("スタイル定義の出力エラー: %w", err)
			}
			delete(pending, id)
		}
	}
	return nil
}

// collectShapeStyleIDs は図形が参照するスタイルIDを収集する（重複なし、出現順）
func collectShapeStyleIDs(shape *pptx.Shape) []int {
	seen := make(map[int]bool)
	var ids []int
	collectShapeStyleRefs(shape, func(id int) {
		if !seen[id] {
			seen[id] = true
			ids = append(ids, id)
		}
	})
	return ids
}

// collectShapeStyleRefs は図形内の全スタイル参照を走査する
func collectShapeStyleRefs(shape *pptx.Shape, fn func(int)) {
	for _, p := range shape.Paragraphs {
		collectParaStyleRefs(p, fn)
	}
	if shape.Table != nil {
		for _, row := range shape.Table.Rows {
			for _, cell := range row {
				if cell != nil {
					for _, p := range cell.Paragraphs {
						collectParaStyleRefs(p, fn)
					}
				}
			}
		}
	}
	for i := range shape.Children {
		collectShapeStyleRefs(&shape.Children[i], fn)
	}
}

// collectParaStyleIDs は段落群が参照するスタイルIDを収集する（重複なし、出現順）
func collectParaStyleIDs(paras []pptx.Paragraph) []int {
	seen := make(map[int]bool)
	var ids []int
	for _, p := range paras {
		collectParaStyleRefs(p, func(id int) {
			if !seen[id] {
				seen[id] = true
				ids = append(ids, id)
			}
		})
	}
	return ids
}

// collectParaStyleRefs は段落内のスタイル参照を走査する
func collectParaStyleRefs(p pptx.Paragraph, fn func(int)) {
	if p.StyleRef != 0 {
		fn(p.StyleRef)
	}
	for _, rt := range p.RichText {
		if rt.StyleRef != 0 {
			fn(rt.StyleRef)
		}
	}
}
