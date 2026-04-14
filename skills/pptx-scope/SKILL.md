---
name: pptx-scope
description: |-
  PowerPointファイル(.pptx)を読み取りスライド/図形/テキスト/画像などの情報をJSONLで出力する。
  プレゼン資料・設計書・フローチャート・提案書の内容確認、データ抽出、図形構造の把握に使用する。
allowed-tools:
  - Bash(*/.claude/skills/pptx-scope/scripts/pptx-scope *)
  - Read
---

# pptx-scope

PowerPointファイル（.pptx）の内容をCLIから出力するツール。

実行ファイル: `bash ${CLAUDE_SKILL_DIR}/scripts/pptx-scope`（以降 `pptx-scope` と表記）

## 出力の読み取り方

全コマンドの出力は自動的に一時ファイル（プレフィックス `pptx-scope-tmp-`）に保存され、stdout にはファイルパスと行数のみが返る。

```bash
$ pptx-scope info example.pptx
{"file":"$TMPDIR/pptx-scope-tmp-abc123.jsonl","lines":1}

$ pptx-scope slides --slide 1,2,3 example.pptx
{"file":"$TMPDIR/pptx-scope-tmp-abc456.jsonl","lines":5}
```

返された `file` パスを Read で読む（offset: 0始まり行番号, limit: 読む行数）。読み終わったら都度 `cleanup` サブコマンドで削除する。

## 利用フロー

1. `info` でスライド一覧とタイトルを確認し対象を特定
2. 目的に応じてコマンドを選択:

   - **スライド内容を確認する** → `slides --slide` で対象スライドを取得
   - **全体を把握する** → `slides` で数枚ずつ分割取得（一括取得はトークン消費大）
   - **特定キーワードを探す** → `search` で該当スライドを特定 → `slides --slide` で詳細取得

3. `slides` 出力に `image_id` があれば `image` で取得し Read で確認（確認後は `cleanup` で削除）

図形の書式情報（フォント・色・枠線）は常に出力される。

## コマンドリファレンス

### info

`pptx-scope info <file>` — ファイルの概要（スライド一覧、スライドサイズ）を出力。

出力例:
```jsonl
{"file":"基本設計書.pptx","slide_size":{"width":720,"height":540}}
{"slide":1,"title":"基本設計書","has_notes":true}
{"slide":2,"title":"目次"}
{"slide":3,"title":"システム構成","has_notes":true}
{"slide":4}
{"slide":5,"title":"フロー図","has_images":true,"hidden":true}
```

1行目はファイルメタ情報、2行目以降はスライド行（slides/search と共通形式）。

- `slide_size`: スライドサイズ（pt単位。標準4:3=720x540, 16:9=960x540）
- `title`: タイトルプレースホルダーのテキスト。存在しない場合は省略
- `has_notes`: ノートにテキストがある場合のみ `true`
- `has_images`: 画像を含む場合のみ `true`（グループ内の画像も検出）
- `hidden`: 非表示スライドの場合のみ `true`

### slides

```
pptx-scope slides [options] <file>
```

| オプション | 説明 | デフォルト |
|-----------|------|-----------|
| `--slide <number,...>` | 対象スライド番号（1始まり、複数指定可: `--slide 1,3`） | 全スライド |
| `--notes` | ノートも出力する | OFF |

出力例:
```jsonl
{"slide":1,"title":"基本設計書","shapes":2,"has_notes":true}
{"style":1,"name":"メイリオ","size":36,"bold":true,"color":"#333333","highlight":"#FFFF00"}
{"shape":1,"type":"rect","placeholder":"ctrTitle","pos":{"x":54,"y":180,"w":612,"h":115.75},"z":0,"alignment":{"vertical":"center"},"paragraphs":[{"text":"基本設計書","s":1,"alignment":{"horizontal":"center"}}]}
{"shape":2,"type":"rect","placeholder":"subTitle","pos":{"x":108,"y":306,"w":504,"h":138},"z":1,"paragraphs":[{"text":"2025年4月版"}]}
{"slide":2,"title":"目次","shapes":2}
{"shape":1,"type":"rect","placeholder":"title","pos":{"x":36,"y":21.63,"w":648,"h":90},"z":0,"paragraphs":[{"text":"目次"}]}
{"shape":2,"type":"rect","placeholder":"body","pos":{"x":36,"y":126,"w":648,"h":356.38},"z":1,"paragraphs":[{"text":"システム概要","bullet":"1."},{"text":"機能一覧","bullet":"2."},{"text":"データフロー","bullet":"3."}]}
```

スライドヘッダ行（`shapes` は図形数）に続いて、図形を1つずつ個別の行として出力する。`--slide` 未指定時は全スライドを順番に出力する。

**図形種別:**

- シェイプ: `rect`, `roundRect`, `ellipse`, `flowChartProcess`, `flowChartDecision` 等（`a:prstGeom` の `prst` 属性値）
- コネクタ: `type` は常に `"connector"`。`from`/`to` で接続先の図形IDを参照。`connector_type` でコネクタ形状、`arrow` で矢印の位置
- グループ: `type` は `"group"`。`children` に子要素の配列。子要素の `pos` はスライド上の絶対座標に変換済み
- テーブル: `type` は `"table"`。`table` フィールドに `cols`（列数）と `rows`（行データ配列）。結合で吸収されたセルは `null`
- 画像: `type` は `"picture"`。`image_id` で `image` サブコマンドにより画像を取得可能

**図形の主なフィールド:**

- `shape`: スライド内の連番ID（1始まり）
- `placeholder`: プレースホルダー種別（`title`, `ctrTitle`, `subTitle`, `body` 等）。プレースホルダーでなければ省略
- `name`: 図形名。プレースホルダーの場合は省略
- `pos`: 位置とサイズ（`x`, `y`, `w`, `h`。pt単位）
- `z`: Z-order（0始まり、大きいほど前面）
- `rotation`: 回転角度（時計回り、度単位。0の場合は省略）
- `flip`: 反転（`"h"`, `"v"`, `"hv"`。反転なしの場合は省略）
- `fill`: 塗りつぶし色（`#RRGGBB`）。塗りつぶしなしの場合は省略
- `line`: 枠線情報（`color`, `style`, `width`。`width` はpt単位）。枠線なしの場合は省略
- `link`: ハイパーリンク（`url` で外部URL、`slide` でスライド内リンクのスライド番号）
- `alignment`: テキストの垂直配置（`vertical` フィールド）。デフォルトの場合は省略
- `text_margin`: テキストボディの内部マージン（`left`, `right`, `top`, `bottom`。pt単位。デフォルトの場合は省略）
- `paragraphs`: 段落の配列。各段落に `text`, `bullet`, `level`, `margin_left`, `indent`, `line_spacing`, `space_before`, `space_after`, `s`, `alignment`, `link`, `rich_text`。行間・段落前後スペースは `"150%"` / `"6pt"` 形式の文字列（デフォルト値は省略）。フォント情報は `style` 行（独立JSONL行）に定義を抽出し `s` で参照する
- `callout_pointer`: 吹き出しのポインタ位置（`x`, `y`。pt単位）

**コネクタの追加フィールド:**

- `from`/`to`: 接続元/先の図形ID
- `from_idx`/`to_idx`: 接続ポイントのインデックス。矩形の場合 0=上, 1=左, 2=下, 3=右
- `connector_type`: コネクタ形状。`straightConnector1`（直線）、`bentConnector2`（L字・1回屈曲）、`bentConnector3`（コの字・2回屈曲）、`curvedConnector3`（曲線）等
- `adj`: 屈曲・カーブの調整値（1/100000単位。bent/curvedコネクタで屈曲位置を制御）
- `arrow`: 矢印ヘッドの位置。`"start"` は `start` 座標側、`"end"` は `end` 座標側、`"both"` は両端。矢印なしの場合は省略
- `start`/`end`: 始点・終点座標（`x`, `y`。pt単位。`pos` と `flip` から算出）
- `label`: コネクタ上のテキスト

**テーブルの出力例:**

```json
{"shape":4,"type":"table","name":"表 1","pos":{"x":36,"y":126,"w":648,"h":236.22},"z":3,"table":{"cols":3,"rows":[[{"text":"項目","fill":"#4285F4","border_bottom":{"color":"#CCCCCC","style":"solid","width":0.75}},{"text":"説明","fill":"#4285F4"},{"text":"備考","fill":"#4285F4"}],[{"text":"機能A"},null,{"text":"必須"}]]}}
```

セルごとに `fill`（背景色）、`border_left` / `border_right` / `border_top` / `border_bottom`（罫線）を持つ。未指定の場合は省略。

**ノート（`--notes` 指定時）:**

図形行の後にノート行 `{"notes":[...]}` が出力される。`notes` は `paragraphs` と同じ構造。

### image

`pptx-scope image <file> <image_id>` — 画像を一時ファイルに保存。

slides 出力の `image_id` を指定する。stdout に `{"file":"$TMPDIR/pptx-scope-tmp-abc123.png"}` が返る。

```jsonl
{"shape":5,"type":"picture","name":"図 1","pos":{"x":78.74,"y":78.74,"w":393.7,"h":236.22},"z":4,"alt_text":"システム構成図","image_id":"ppt/media/image1.png"}
```

```bash
pptx-scope image example.pptx ppt/media/image1.png
```

### search

```
pptx-scope search [options] <file>
```

| オプション | 説明 | デフォルト |
|-----------|------|-----------|
| `--text <text>` | 検索文字列（部分一致、大文字小文字無視） | 必須 |
| `--slide <number,...>` | 対象スライド番号（複数指定可） | 全スライド |
| `--notes` | ノートも検索対象にする | OFF |

マッチしたスライドのヘッダ行のみ出力する（info/slides と共通形式）。詳細は `slides --slide` で取得する。結果なしでも正常終了（終了コード 0）する。

```bash
pptx-scope search --text "データ" example.pptx
```

```jsonl
{"slide":2,"title":"システム構成","has_images":true}
{"slide":5,"title":"データフロー"}
```

### cleanup

`pptx-scope cleanup <file> [file...]` — pptx-scope が生成した一時ファイルを削除する。

```bash
$ pptx-scope cleanup /tmp/pptx-scope-tmp-abc123.jsonl /tmp/pptx-scope-tmp-def456.png
{"deleted":2}
```

### version

`pptx-scope version` — バージョン情報を表示する。

```bash
$ pptx-scope version
v0.0.9
```
