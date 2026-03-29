---
name: cc-read-pptx
description: PowerPointファイル（.pptx）を読み取る。プレゼン資料、設計書、フローチャート、提案書の内容確認・データ抽出時に使用する。
user-invocable: false
allowed-tools: Bash(cc-read-pptx *)
---

# cc-read-pptx

PowerPointファイル（.pptx）の内容をCLIから出力するツール。

## 利用フロー

```
1. info   → スライド一覧を確認し対象スライドを特定
2. slides → スライドの内容を取得（図形・テキスト・テーブル・コネクタ・画像）
3. search → 特定テキストの検索（slides より効率的）
```

基本的には `info` → `slides` で内容を把握する。特定のキーワードを探す場合は `search` が効率的。
図形の書式情報（フォント・色・枠線）は常に出力される。画像を確認するには `--extract-images <dir>` で抽出し、出力の `image.path` を Read ツールで読む。

**重要:** `info` の結果で `has_images: true` のスライドがある場合、`slides` コマンドには必ず `--extract-images /tmp/pptx_images` を付けること。抽出後は出力の `image.path` を Read ツールで読み、画像の内容を確認すること。

## コマンドリファレンス

### info

```bash
cc-read-pptx info <file>
```

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

```bash
cc-read-pptx slides [options] <file>
```

| オプション | 説明 | デフォルト |
|-----------|------|-----------|
| `--slide <number>` | 対象スライド番号（1始まり） | 全スライド |
| `--notes` | ノートも出力する | OFF |
| `--extract-images <dir>` | 画像を指定ディレクトリに抽出 | OFF（画像スキップ） |

出力例:
```jsonl
{"slide":1,"title":"基本設計書","shapes":[{"id":1,"type":"rect","placeholder":"ctrTitle","position":{"x":685800,"y":2286000,"cx":7772400,"cy":1470025},"z":0,"font":{"name":"メイリオ","size":36,"bold":true,"color":"#333333"},"alignment":{"horizontal":"center","vertical":"center"},"paragraphs":[{"text":"基本設計書"}]}]}
{"slide":2,"title":"目次","shapes":[{"id":1,"type":"rect","placeholder":"title","z":0,"paragraphs":[{"text":"目次"}]},{"id":2,"type":"rect","placeholder":"body","z":1,"paragraphs":[{"text":"システム概要","bullet":"1."},{"text":"機能一覧","bullet":"2."}]}]}
```

1スライドにつき1行のJSONオブジェクト（JSONL形式）。`--slide` 未指定時は全スライドを順番に出力する。

**図形種別:**

- シェイプ: `rect`, `roundRect`, `ellipse`, `flowChartProcess`, `flowChartDecision` 等（`a:prstGeom` の `prst` 属性値）
- コネクタ: `type` は常に `"connector"`。`from`/`to` で接続先の図形IDを参照。`connector_type` でコネクタ形状、`arrow` で矢印の位置
- グループ: `type` は `"group"`。`children` に子要素の配列
- テーブル: `type` は `"table"`。`table` フィールドに `cols`（列数）と `rows`（行データ配列）。結合で吸収されたセルは `null`
- 画像: `type` は `"picture"`。`--extract-images` 未指定時はスキップされる

**図形の主なフィールド:**

- `id`: スライド内の連番ID（1始まり）
- `placeholder`: プレースホルダー種別（`title`, `ctrTitle`, `subTitle`, `body` 等）。プレースホルダーでなければ省略
- `name`: 図形名。プレースホルダーの場合は省略
- `position`: 位置とサイズ（`x`, `y`, `cx`, `cy`。EMU単位）
- `z`: Z-order（0始まり、大きいほど前面）
- `fill`: 塗りつぶし色（`#RRGGBB`）
- `line`: 枠線情報（`color`, `style`, `width`）
- `paragraphs`: 段落の配列。各段落に `text`, `bullet`, `level`, `font`, `alignment`, `rich_text`
- `callout_pointer`: 吹き出しのポインタ位置（`x`, `y`。EMU単位）

**コネクタの追加フィールド:**

- `from`/`to`: 接続元/先の図形ID
- `connector_type`: `line`, `straightConnector1`, `bentConnector3`, `curvedConnector3` 等
- `arrow`: `"start"`, `"end"`, `"both"`, `"none"`
- `label`: コネクタ上のテキスト

**テーブルの出力例:**

```json
{"id":4,"type":"table","name":"表 1","position":{"x":457200,"y":1600200,"cx":8229600,"cy":3000000},"z":3,"table":{"cols":3,"rows":[["項目","説明","備考"],["機能A",null,"必須"]]}}
```

**画像の確認方法:**

`--extract-images` で抽出後、出力の `image.path` を Read ツールで読むことで画像の中身を視覚的に確認できる。

```jsonl
{"id":5,"type":"picture","name":"図 1","position":{"x":1000000,"y":1000000,"cx":5000000,"cy":3000000},"z":4,"alt_text":"システム構成図","image":{"format":"png","width":640,"height":480,"size":45230,"path":"/tmp/imgs/image_1.png"}}
```

**ノート（`--notes` 指定時）:**

`notes` フィールドに段落の配列が追加される。`paragraphs` と同じ構造。

### search

```bash
cc-read-pptx search [options] <file>
```

| オプション | 説明 | デフォルト |
|-----------|------|-----------|
| `--text <text>` | 検索文字列（部分一致、大文字小文字無視） | 必須 |
| `--slide <number>` | 対象スライド番号 | 全スライド |
| `--notes` | ノートも検索対象にする | OFF |

出力形式は `slides` と同じJSONL。マッチしたスライドのみ出力し、図形内ではマッチした段落のみを含める。テーブルはいずれかのセルにヒットした場合テーブル全体を出力する。結果なしでも正常終了（終了コード 0）する。

```bash
cc-read-pptx search --text "データ" example.pptx
```

```jsonl
{"slide":2,"title":"システム構成","shapes":[{"id":3,"type":"rect","name":"テキストボックス 1","position":{"x":1000000,"y":2000000,"cx":3000000,"cy":500000},"z":2,"paragraphs":[{"text":"データフロー図"}]}]}
```
