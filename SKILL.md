---
name: pptx-scope
description: PowerPointファイル（.pptx）を読み取る。プレゼン資料、設計書、フローチャート、提案書の内容確認・データ抽出時に使用する。
user-invocable: false
allowed-tools:
  - Bash
  - Read
---

# pptx-scope

PowerPointファイル（.pptx）の内容をCLIから出力するツール。

実行ファイル: `bash ${CLAUDE_SKILL_DIR}/scripts/pptx-scope <command> [options] <file>`

## 出力の読み取り方

全コマンドの出力は自動的に一時ファイルに保存され、stdout にはファイルパスと行数のみが返る。

```bash
$ bash ${CLAUDE_SKILL_DIR}/scripts/pptx-scope info example.pptx
{"file":"$TMPDIR/pptx-scope-abc123.jsonl","lines":1}

$ bash ${CLAUDE_SKILL_DIR}/scripts/pptx-scope slides --slide 1,2,3 example.pptx
{"file":"$TMPDIR/pptx-scope-abc456.jsonl","lines":5}
```

返された `file` パスを Read で読む（offset: 0始まり行番号, limit: 読む行数）。読み終わったら都度削除する。

## 利用フロー

```
1. info   → スライド一覧を確認し対象スライドを特定
2. slides → スライドの内容を取得
3. image  → 必要な画像を個別に取得して確認
4. search → 特定テキストの検索（slides より効率的）
```

基本的には `info` → `slides` で内容を把握する。特定のキーワードを探す場合は `search` が効率的。
図形の書式情報（フォント・色・枠線）は常に出力される。

**大規模プレゼンテーションの取得戦略:**

スライド数が多い場合、全スライドを一括取得するとトークン消費が大きくなる。まず `info` でスライド一覧とタイトルを確認し、`--slide` で必要なスライドに絞って取得する。全体を把握する必要がある場合は、数枚ずつ分割して取得する。特定のキーワードを探す場合は `search` で該当スライドを特定してから `slides` で詳細を取得するのが効率的。

**画像の確認:** `info` で `has_images: true` のスライドには画像が含まれる。`slides` 出力に `image_id` がある場合、内容の把握に役立つ可能性が高いため積極的に確認する。

1. `image` サブコマンドで画像を取得: `pptx-scope image <file> <image_id>`
2. 返された `file` パスを Read で画像の内容を確認する
3. 確認が終わったら削除する

## コマンドリファレンス

### info

`pptx-scope info <file>` — ファイルの概要（スライド一覧、スライドサイズ）を出力。

出力例:
```json
{"file":"基本設計書.pptx","slide_size":{"width":9144000,"height":6858000},"slides":[{"number":1,"title":"基本設計書","has_notes":true},{"number":2,"title":"目次"},{"number":3,"title":"システム構成","has_images":true},{"number":4},{"number":5,"title":"フロー図","hidden":true}]}
```

- `slides[].title`: タイトルプレースホルダーのテキスト。存在しない場合は省略
- `slides[].has_notes`: ノートにテキストがある場合のみ `true`
- `slides[].has_images`: 画像を含む場合のみ `true`（グループ内の画像も検出）
- `slides[].hidden`: 非表示スライドの場合のみ `true`
- `slide_size`: スライドサイズ（EMU単位。標準4:3=9144000x6858000, 16:9=12192000x6858000）

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
{"slide":1,"title":"基本設計書","shapes":[{"id":1,"type":"rect","placeholder":"ctrTitle","pos":{"x":685800,"y":2286000,"w":7772400,"h":1470025},"z":0,"alignment":{"vertical":"center"},"paragraphs":[{"text":"基本設計書","font":{"name":"メイリオ","size":4572000,"bold":true,"color":"#333333"},"alignment":{"horizontal":"center"}}]}]}
{"slide":2,"title":"目次","shapes":[{"id":1,"type":"rect","placeholder":"title","z":0,"paragraphs":[{"text":"目次"}]},{"id":2,"type":"rect","placeholder":"body","z":1,"paragraphs":[{"text":"システム概要","bullet":"1."},{"text":"機能一覧","bullet":"2."}]}]}
```

1スライドにつき1行のJSONオブジェクト（JSONL形式）。`--slide` 未指定時は全スライドを順番に出力する。

**図形種別:**

- シェイプ: `rect`, `roundRect`, `ellipse`, `flowChartProcess`, `flowChartDecision` 等（`a:prstGeom` の `prst` 属性値）
- コネクタ: `type` は常に `"connector"`。`from`/`to` で接続先の図形IDを参照。`connector_type` でコネクタ形状、`arrow` で矢印の位置
- グループ: `type` は `"group"`。`children` に子要素の配列。子要素の `pos` はスライド上の絶対座標に変換済み
- テーブル: `type` は `"table"`。`table` フィールドに `cols`（列数）と `rows`（行データ配列）。結合で吸収されたセルは `null`
- 画像: `type` は `"picture"`。`image_id` で `image` サブコマンドにより画像を取得可能

**図形の主なフィールド:**

- `id`: スライド内の連番ID（1始まり）
- `placeholder`: プレースホルダー種別（`title`, `ctrTitle`, `subTitle`, `body` 等）。プレースホルダーでなければ省略
- `name`: 図形名。プレースホルダーの場合は省略
- `pos`: 位置とサイズ（`x`, `y`, `w`, `h`。EMU単位）
- `z`: Z-order（0始まり、大きいほど前面）
- `rotation`: 回転角度（時計回り、度単位。0の場合は省略）
- `fill`: 塗りつぶし色（`#RRGGBB`）
- `line`: 枠線情報（`color`, `style`, `width`。`width` はEMU単位）
- `link`: ハイパーリンク（`url` で外部URL、`slide` でスライド内リンクのスライド番号）
- `alignment`: テキストの垂直配置（`vertical` フィールド）。デフォルトの場合は省略
- `paragraphs`: 段落の配列。各段落に `text`, `bullet`, `level`, `font`/`s`, `alignment`, `link`, `rich_text`。`font.size` はEMU単位。複数回使われるフォントは `_styles` 行（独立JSONL行）に定義を抽出し `s` で参照する
- `callout_pointer`: 吹き出しのポインタ位置（`x`, `y`。EMU単位）

**コネクタの追加フィールド:**

- `from`/`to`: 接続元/先の図形ID
- `from_idx`/`to_idx`: 接続ポイントのインデックス（図形上の接続位置、形状依存）
- `connector_type`: `line`, `straightConnector1`, `bentConnector3`, `curvedConnector3` 等
- `adj`: 屈曲・カーブの調整値（1/100000単位。bent/curvedコネクタで屈曲位置を制御）
- `arrow`: `"start"`, `"end"`, `"both"`, `"none"`
- `start`/`end`: 始点・終点座標（`x`, `y`。EMU単位。`pos` と `flip` から算出）
- `label`: コネクタ上のテキスト

**テーブルの出力例:**

```json
{"id":4,"type":"table","name":"表 1","pos":{"x":457200,"y":1600200,"w":8229600,"h":3000000},"z":3,"table":{"cols":3,"rows":[["項目","説明","備考"],["機能A",null,"必須"]]}}
```

**画像の確認方法:**

出力の `image_id` を使い、`image` サブコマンドで画像のバイナリを取得できる。

```jsonl
{"id":5,"type":"picture","name":"図 1","pos":{"x":1000000,"y":1000000,"w":5000000,"h":3000000},"z":4,"alt_text":"システム構成図","image_id":"ppt/media/image1.png"}
```

```bash
pptx-scope image example.pptx ppt/media/image1.png
```

**ノート（`--notes` 指定時）:**

`notes` フィールドに段落の配列が追加される。`paragraphs` と同じ構造。

### image

`pptx-scope image <file> <image_id>` — 画像を一時ファイルに保存。

slides 出力の `image_id` を指定する。stdout に `{"file":"$TMPDIR/pptx-scope-abc123.png"}` が返る。返された `file` パスを Read で確認し、終わったら削除する。

### search

```
pptx-scope search [options] <file>
```

| オプション | 説明 | デフォルト |
|-----------|------|-----------|
| `--text <text>` | 検索文字列（部分一致、大文字小文字無視） | 必須 |
| `--slide <number,...>` | 対象スライド番号（複数指定可） | 全スライド |
| `--notes` | ノートも検索対象にする | OFF |

出力形式は `slides` と同じJSONL。マッチしたスライドのみ出力し、図形内ではマッチした段落のみを含める。テーブルはいずれかのセルにヒットした場合テーブル全体を出力する。結果なしでも正常終了（終了コード 0）する。

```bash
pptx-scope search --text "データ" example.pptx
```

```jsonl
{"slide":2,"title":"システム構成","shapes":[{"id":3,"type":"rect","name":"テキストボックス 1","pos":{"x":1000000,"y":2000000,"w":3000000,"h":500000},"z":2,"paragraphs":[{"text":"データフロー図"}]}]}
```
