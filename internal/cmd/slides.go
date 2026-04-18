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
		if err := emitPendingStyles(enc, pending, collectStyleIDs(sd.Shapes[i:i+1], nil)); err != nil {
			return err
		}
		if err := enc.Encode(sd.Shapes[i]); err != nil {
			return fmt.Errorf("JSON出力エラー: %w", err)
		}
	}

	// ノートで使用するスタイル定義を出力
	if len(sd.Notes) > 0 {
		if err := emitPendingStyles(enc, pending, collectStyleIDs(nil, sd.Notes)); err != nil {
			return err
		}
		if err := enc.Encode(notesOutput{Notes: sd.Notes}); err != nil {
			return fmt.Errorf("JSON出力エラー: %w", err)
		}
	}

	return nil
}

// emitPendingStyles は ids で参照されるスタイル定義のうち未出力のものを出力する
func emitPendingStyles(enc *json.Encoder, pending map[int]pptx.StyleDef, ids []int) error {
	if len(pending) == 0 {
		return nil
	}
	for _, id := range ids {
		if s, ok := pending[id]; ok {
			if err := enc.Encode(s); err != nil {
				return fmt.Errorf("スタイル定義の出力エラー: %w", err)
			}
			delete(pending, id)
		}
	}
	return nil
}

// collectStyleIDs はスライド内の全スタイル参照を収集する（重複なし、出現順）
func collectStyleIDs(shapes []pptx.Shape, notes []pptx.Paragraph) []int {
	seen := make(map[int]bool)
	var ids []int
	add := func(id int) {
		if id != 0 && !seen[id] {
			seen[id] = true
			ids = append(ids, id)
		}
	}
	pptx.WalkSlideParagraphs(shapes, notes, func(p *pptx.Paragraph) {
		add(p.StyleRef)
		for _, rt := range p.RichText {
			add(rt.StyleRef)
		}
	})
	return ids
}
