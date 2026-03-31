# cc-read-pptx

PowerPoint ファイル（.pptx）の内容を CLI から読み取り、JSONL形式で出力するGoツール。
Claude Code 等の AI エージェントが PowerPoint 資料（プレゼン、設計書、フローチャート、提案書など）を構造的に読み取る用途を主眼とする。

## インストール

```bash
go install github.com/nobmurakita/cc-read-pptx@latest
```

### Claude Code スキルのインストール

GitHub から直接インストール:

```bash
mkdir -p ~/.claude/skills/cc-read-pptx
curl -fsSL https://raw.githubusercontent.com/nobmurakita/cc-read-pptx/main/SKILL.md -o ~/.claude/skills/cc-read-pptx/SKILL.md
```

またはローカルからコピー:

```bash
mkdir -p ~/.claude/skills/cc-read-pptx
cp SKILL.md ~/.claude/skills/cc-read-pptx/SKILL.md
```

インストール後、Claude Code が PowerPoint ファイルの読み取りが必要な場面で自動的に cc-read-pptx を使用する。

## コマンド

### info — ファイルの概要を表示

```bash
cc-read-pptx info 基本設計書.pptx
```

```json
{"file":"基本設計書.pptx","slide_size":{"width":9144000,"height":6858000},"slides":[{"number":1,"title":"基本設計書","has_notes":true},{"number":2,"title":"目次"},{"number":3,"title":"システム構成"}]}
```

- `slides[].title`: タイトルプレースホルダーのテキスト。存在しない場合は省略
- `slides[].has_notes`: ノートにテキストがある場合のみ `true`
- `slides[].has_images`: 画像を含む場合のみ `true`（グループ内も検出）
- `slides[].hidden`: 非表示スライドの場合のみ `true`
- `slide_size`: スライドサイズ（EMU単位）

### slides — スライドの内容を出力

```bash
cc-read-pptx slides --slide 1 基本設計書.pptx
```

```jsonl
{"slide":1,"title":"基本設計書","shapes":[{"id":1,"type":"rect","placeholder":"ctrTitle","pos":{"x":685800,"y":2286000,"w":7772400,"h":1470025},"z":0,"paragraphs":[{"text":"基本設計書","font":{"name":"メイリオ","size":4572000,"bold":true}}]},{"id":2,"type":"rect","placeholder":"subTitle","pos":{"x":1371600,"y":3886200,"w":6400800,"h":1752600},"z":1,"paragraphs":[{"text":"2025年4月版"}]}]}
```

1スライドにつき1行のJSONオブジェクト（JSONL形式）。`--slide` 未指定時は全スライドを順番に出力する。
プレースホルダーのフォント情報（名前・サイズ・色）、位置・サイズ、箇条書きスタイルはスライドマスター・レイアウトから自動的に継承される。テーマフォント参照も実フォント名に解決される。

テーブル:

```bash
cc-read-pptx slides --slide 2 進捗_20200108.pptx
```

```jsonl
{"slide":2,"shapes":[{"id":1,"type":"table","name":"表 1","pos":{"x":457200,"y":1600200,"w":8229600,"h":3000000},"z":0,"table":{"cols":3,"rows":[["項目","説明","備考"],["機能A","データ取得","必須"]]}}]}
```

コネクタ:

```jsonl
{"id":20,"type":"connector","name":"直線矢印コネクタ 52","pos":{"x":3530458,"y":2689356,"w":296133,"h":4992},"z":19,"line":{"color":"#007CD5","style":"solid","width":38100},"from":3,"to":4,"connector_type":"straightConnector1","arrow":"end"}
```

画像は `image_id` で識別され、`image` サブコマンドで個別に取得できる:

```bash
cc-read-pptx slides --slide 7 設計書.pptx
```

```jsonl
{"id":5,"type":"picture","name":"図 1","pos":{"x":1000000,"y":1000000,"w":5000000,"h":3000000},"z":4,"alt_text":"構成図","image_id":"ppt/media/image1.png"}
```

ハイパーリンク（外部URL・スライド内リンク）:

```jsonl
{"id":2,"type":"rect","placeholder":"body","z":1,"paragraphs":[{"text":"はじめに","link":{"slide":4}},{"text":"背景","link":{"slide":5}}]}
```

```jsonl
{"text":"詳細はこちら","rich_text":[{"text":"アーキテクチャ草案","link":{"url":"https://example.com"}},{"text":"を参照"}]}
```

ノート付き:

```bash
cc-read-pptx slides --slide 1 --notes 資料.pptx
```

```jsonl
{"slide":1,"title":"概要","shapes":[...],"notes":[{"text":"このスライドでは概要を説明する。"},{"text":"スコープの定義","bullet":"•"}]}
```

**オプション:**

| オプション | 説明 | デフォルト |
|-----------|------|-----------|
| `--slide` | 対象スライド番号（1始まり、複数指定可: `--slide 1,3`） | 全スライド |
| `--notes` | ノートも出力する | OFF |

### image — 画像を取得

```bash
cc-read-pptx image 設計書.pptx ppt/media/image1.png /tmp/img.png
```

`slides` 出力の `image_id` を指定して、画像をファイルに保存する。

### search — テキストを検索

```bash
cc-read-pptx search --text "データ" 基本設計書.pptx
```

```jsonl
{"slide":2,"title":"システム構成","shapes":[{"id":3,"type":"rect","name":"テキストボックス 1","pos":{"x":1000000,"y":2000000,"w":3000000,"h":500000},"z":2,"paragraphs":[{"text":"データフロー図"}]}]}
```

マッチしたスライドのみ出力し、図形内ではマッチした段落のみを含める。テーブルはいずれかのセルにヒットした場合テーブル全体を出力する。結果なしでも正常終了（終了コード 0）する。

**オプション:**

| オプション | 説明 | デフォルト |
|-----------|------|-----------|
| `--text` | 検索文字列（部分一致、大文字小文字無視） | 必須 |
| `--slide` | 対象スライド番号（複数指定可） | 全スライド |
| `--notes` | ノートも検索対象にする | OFF |

## 図形種別

| 種別 | `type` の値 | 説明 |
|------|------------|------|
| シェイプ | `rect`, `roundRect`, `ellipse`, `flowChartProcess` 等 | `a:prstGeom` の `prst` 属性値 |
| コネクタ | `connector` | `from`/`to` で接続先の図形IDを参照 |
| グループ | `group` | `children` に子要素の配列 |
| テーブル | `table` | `table` フィールドに `cols` と `rows` |
| 画像 | `picture` | `image_id` で `image` サブコマンドにより取得可能 |

出力フィールドの詳細は [DESIGN.md](DESIGN.md) を参照。
