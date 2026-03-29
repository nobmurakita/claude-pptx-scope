# cc-read-pptx 設計ドキュメント

PowerPoint ファイル（.pptx）の内容をCLIから出力するGoツール。
Claude CodeがPowerPoint資料（プレゼン、フローチャート、仕様書など）を読み取る用途を主眼とする。

## コマンド構成と利用フロー

Claude Code からの典型的な利用フローは以下の通り:

1. **`info`** — ファイルの全体像を把握する（スライド一覧、スライドサイズ）。どのスライドを読むべきか判断する
2. **`slides`** — スライドの内容を取得する。図形・テキスト・テーブル・コネクタ・画像をまとめて出力する
3. **`search`** — プレゼンテーション内のテキストを検索する。全スライドまたは指定スライドから条件に合うテキストを抽出する

基本的には `info` → `slides` で内容を把握する。特定のキーワードを探す場合は `search` が効率的。

### `cc-read-pptx info <file>`

**役割:** ファイルレベルの概要を把握する。スライド一覧から対象スライドを特定し、以降の `slides` / `search` に渡す `--slide` を決定する。

ファイルレベルの概要をJSON形式で出力する。

- スライド一覧（番号、タイトル、ノートの有無）
- スライドサイズ

**出力例:**

```json
{
  "file": "基本設計書.pptx",
  "slide_size": {"width": 9144000, "height": 6858000},
  "slides": [
    {"number": 1, "title": "基本設計書", "has_notes": true},
    {"number": 2, "title": "目次"},
    {"number": 3, "title": "システム構成", "has_notes": true},
    {"number": 4},
    {"number": 5, "title": "フロー図", "hidden": true}
  ]
}
```

**`slides` 配列の各要素:**

| フィールド | 型 | 説明 |
|-----------|-----|------|
| `number` | number | スライド番号（1始まり。presentation.xml の `sldIdLst` の順序に基づく） |
| `title` | string | スライドタイトル。タイトルプレースホルダー（`ph type="title"` または `ph type="ctrTitle"`）のテキストから取得。存在しない場合は省略 |
| `has_notes` | bool | ノートスライドが存在しテキストがある場合のみ `true` を出力 |
| `hidden` | bool | 非表示スライドの場合のみ `true` を出力（`p:sld` の `show="0"` 属性） |

**`slide_size` オブジェクト:**

| フィールド | 型 | 説明 |
|-----------|-----|------|
| `width` | number | スライド幅（EMU単位。標準4:3 = 9144000, 16:9 = 12192000） |
| `height` | number | スライド高さ（EMU単位。標準4:3 = 6858000, 16:9 = 6858000） |

タイトルの取得方法:
- スライドの `p:spTree` 内で `p:nvSpPr/p:nvPr/p:ph` の `type` 属性が `title` または `ctrTitle` のシェイプを探す
- 該当シェイプの `p:txBody` 内の全 `a:t` 要素を結合する
- タイトルプレースホルダーが存在しない場合は `title` フィールドを省略する

### `cc-read-pptx slides [options] <file>`

**役割:** スライドの内容を取得する。図形・テキスト・テーブル・コネクタをまとめて1スライド1JSONオブジェクトで出力する。

**オプション:**

| オプション | 説明 | デフォルト |
|-----------|------|-----------|
| `--slide <number>` | 対象スライド番号（1始まり） | 全スライド |
| `--notes` | ノートも出力する | OFF |
| `--extract-images <dir>` | 画像を指定ディレクトリに抽出する。未指定時は画像をスキップ | OFF（画像スキップ） |

- `--slide` 未指定時は全スライドを順番に出力する（1行1スライドのJSONL）
- `--extract-images` 未指定時は `p:pic` 要素をスキップする
- 書式情報は常に出力する

**出力形式:**

1スライドにつき1行のJSONオブジェクトを出力する（JSONL形式だがスライド単位）。

**出力例:**

```jsonl
{"slide":1,"title":"基本設計書","shapes":[{"id":1,"type":"rect","placeholder":"ctrTitle","position":{"x":685800,"y":2286000,"cx":7772400,"cy":1470025},"z":0,"font":{"name":"メイリオ","size":36,"bold":true,"color":"#333333"},"alignment":{"horizontal":"center","vertical":"center"},"paragraphs":[{"text":"基本設計書"}]},{"id":2,"type":"rect","placeholder":"subTitle","position":{"x":1371600,"y":3886200,"cx":6400800,"cy":1752600},"z":1,"paragraphs":[{"text":"2025年4月版"}]}]}
{"slide":2,"title":"目次","shapes":[{"id":1,"type":"rect","placeholder":"title","position":{"x":457200,"y":274638,"cx":8229600,"cy":1143000},"z":0,"paragraphs":[{"text":"目次"}]},{"id":2,"type":"rect","placeholder":"body","position":{"x":457200,"y":1600200,"cx":8229600,"cy":4525963},"z":1,"paragraphs":[{"text":"システム概要","bullet":"1."},{"text":"機能一覧","bullet":"2."},{"text":"データフロー","bullet":"3."}]}]}
```

整形すると以下のような構造:

```json
{
  "slide": 2,
  "title": "目次",
  "shapes": [
    {
      "id": 1,
      "type": "rect",
      "placeholder": "title",
      "position": {"x": 457200, "y": 274638, "cx": 8229600, "cy": 1143000},
      "z": 0,
      "paragraphs": [
        {"text": "目次"}
      ]
    },
    {
      "id": 2,
      "type": "rect",
      "placeholder": "body",
      "position": {"x": 457200, "y": 1600200, "cx": 8229600, "cy": 4525963},
      "z": 1,
      "paragraphs": [
        {"text": "システム概要", "bullet": "1."},
        {"text": "機能一覧", "bullet": "2."},
        {"text": "データフロー", "bullet": "3."}
      ]
    }
  ]
}
```

**スライドオブジェクトのフィールド:**

| フィールド | 型 | 説明 |
|-----------|-----|------|
| `slide` | number | スライド番号 |
| `title` | string | スライドタイトル（存在する場合のみ） |
| `shapes` | array | 図形の配列 |
| `notes` | array | ノートの段落配列（`--notes` 指定時のみ。後述） |

#### 図形フィールド

`shapes` 配列の各要素:

| フィールド | 型 | 説明 |
|-----------|-----|------|
| `id` | number | スライド内の図形ID。spTree 内の出現順で1始まりの連番を割り当てる |
| `type` | string | 図形種別。後述の図形種別を参照 |
| `name` | string | 図形名（`p:cNvPr` の `name` 属性）。プレースホルダーの場合は省略 |
| `placeholder` | string | プレースホルダー種別（`title`, `ctrTitle`, `subTitle`, `body`, `dt`, `ftr`, `sldNum` 等）。プレースホルダーでない場合は省略 |
| `position` | object | 位置とサイズ（`a:xfrm` の `a:off` と `a:ext`。EMU単位） |
| `z` | number | Z-order。spTree 内の出現順で0始まり。大きいほど前面 |
| `rotation` | number | 回転角度（度単位、時計回り）。`a:xfrm` の `rot` 属性を60000で除算。0の場合は省略 |
| `flip` | string | 反転。`"h"`, `"v"`, `"hv"`。なければ省略 |
| `fill` | string | 塗りつぶし色（`#RRGGBB` 形式）。塗りつぶしなしの場合は省略 |
| `line` | object | 枠線情報。枠線がない場合は省略 |
| `callout_pointer` | object | 吹き出しのポインタ位置（吹き出し図形の場合のみ。後述） |
| `paragraphs` | array | 段落の配列。テキストがない場合は省略 |
| `table` | object | テーブルデータ（テーブルの場合。`paragraphs` の代わりに使用） |

**`position` オブジェクト:**

| フィールド | 型 | 説明 |
|-----------|-----|------|
| `x` | number | 左上X座標（EMU） |
| `y` | number | 左上Y座標（EMU） |
| `cx` | number | 幅（EMU） |
| `cy` | number | 高さ（EMU） |

**`line` オブジェクト:**

| フィールド | 型 | 説明 |
|-----------|-----|------|
| `color` | string | 線の色（`#RRGGBB` 形式） |
| `style` | string | 線のスタイル。`solid`, `dash`, `dot`, `dashDot` 等 |
| `width` | number | 線幅（ポイント単位）。`a:ln` の `w` 属性をEMUからポイントに変換（÷ 12700） |

#### 段落フィールド

`paragraphs` 配列の各要素:

| フィールド | 型 | 説明 |
|-----------|-----|------|
| `text` | string | 段落のプレーンテキスト |
| `bullet` | string | 箇条書き記号。箇条書きでない段落は省略 |
| `level` | number | インデントレベル（0始まり）。0の場合は省略 |
| `font` | object | フォント情報。デフォルト値のフィールドは省略 |
| `alignment` | object | 配置情報（`horizontal`, `vertical`）。デフォルトの場合は省略 |
| `rich_text` | array | リッチテキストラン（段落内に書式の異なるランが存在する場合のみ） |

テキストが空の段落は出力しない（箇条書き間の空行等）。

**箇条書き:**

PowerPoint の箇条書きは段落レベルのプロパティ（`a:pPr`）として定義される。

| XML要素 | `bullet` の出力 |
|---------|-----------------|
| `a:buChar` | `char` 属性の文字をそのまま使用（例: `"•"`, `"–"`, `">"` ） |
| `a:buAutoNum` | 番号を計算して文字列化（例: `"1."`, `"2."`, `"a."`, `"i)"` ） |
| `a:buNone` / 箇条書きなし | `bullet` フィールドを省略 |

`a:buAutoNum` の `type` 属性と出力形式:

| type 値 | 出力例 |
|---------|-------|
| `arabicPeriod` | `1.`, `2.`, `3.` |
| `arabicParenR` | `1)`, `2)`, `3)` |
| `alphaLcPeriod` | `a.`, `b.`, `c.` |
| `alphaUcPeriod` | `A.`, `B.`, `C.` |
| `romanLcPeriod` | `i.`, `ii.`, `iii.` |
| `romanUcPeriod` | `I.`, `II.`, `III.` |

- 番号は同一図形内で同じ `level` の連続する `a:buAutoNum` 段落をカウントして算出する
- `a:buAutoNum` の `startAt` 属性がある場合はその値から開始する（デフォルトは1）
- レベルが変わるか箇条書きなしの段落で番号はリセットされる
- 上記以外の `type` 値は `arabicPeriod` と同じ形式で出力する

`level` は `a:pPr` の `lvl` 属性（0始まり）。未指定時は0として扱う。0の場合はフィールドを省略する。

**`font` オブジェクト:**

```json
{"name": "メイリオ", "size": 11, "bold": true, "italic": true, "strikethrough": true, "underline": "single", "color": "#FF0000"}
```

- デフォルト値のフィールドは省略する
- `size` は hundredths of point（`a:rPr` の `sz` 属性）を point に変換（÷ 100）

**リッチテキスト:**

段落内のテキストの一部が異なる書式を持つ場合、`text` にはプレーンテキストを格納し、`rich_text` に書式付きランの配列を格納する。

```json
{"text": "重要: この機能は必須です", "rich_text": [{"text": "重要: ", "font": {"bold": true, "color": "#FF0000"}}, {"text": "この機能は必須です"}]}
```

- `text` は常にプレーンテキスト（検索・概要把握用）
- `rich_text` は書式の異なるランが存在する場合のみ出力
- 各ランの `font` はデフォルト値のフィールドを省略する

#### 吹き出しポインタ

吹き出し図形（`type` が `wedgeRectCallout`, `wedgeRoundRectCallout`, `wedgeEllipseCallout`, `cloudCallout`, `borderCallout1`, `borderCallout2`, `borderCallout3` 等）の場合、ポインタの指す位置をスライド上の絶対座標に変換して出力する。

```json
{"id": 3, "type": "wedgeRoundRectCallout", "position": {"x": 3000000, "y": 1000000, "cx": 2000000, "cy": 800000}, "z": 2, "callout_pointer": {"x": 1500000, "y": 2500000}, "paragraphs": [{"text": "ここに注目"}]}
```

**`callout_pointer` オブジェクト:**

| フィールド | 型 | 説明 |
|-----------|-----|------|
| `x` | number | ポインタ先端のX座標（EMU。スライド上の絶対座標） |
| `y` | number | ポインタ先端のY座標（EMU。スライド上の絶対座標） |

**算出方法:**

調整ハンドル `adj1`（X方向）と `adj2`（Y方向）は図形の中心を原点とした1/100000単位の相対比率で表される。これを図形の位置・サイズからスライド上の絶対座標に変換する:

- `pointer_x = position.x + position.cx / 2 + adj1 * position.cx / 100000`
- `pointer_y = position.y + position.cy / 2 + adj2 * position.cy / 100000`

調整ハンドルが未指定（デフォルト値）の場合は `callout_pointer` フィールドを省略する。

#### コネクタフィールド

コネクタ（`p:cxnSp`）は `type` が `"connector"` となり、以下の追加フィールドを持つ。

| フィールド | 型 | 説明 |
|-----------|-----|------|
| `from` | number | 接続元の図形ID。接続情報がない場合は省略 |
| `to` | number | 接続先の図形ID。接続情報がない場合は省略 |
| `connector_type` | string | コネクタ形状。`a:prstGeom` の `prst` 属性値（`line`, `straightConnector1`, `bentConnector3`, `curvedConnector3` 等） |
| `arrow` | string | 矢印の位置。`"start"`, `"end"`, `"both"`, `"none"` |
| `label` | string | コネクタ上のテキスト。テキストがない場合は省略 |

コネクタの `from` / `to` は、`p:cxnSp` 内の `a:stCxn` / `a:endCxn` 要素の `id` 属性を参照する。この `id` は PowerPoint が付与する図形IDであり、`slides` コマンドが割り当てる連番 `id` とは異なる。パース時に PowerPoint 図形IDから連番IDへのマッピングを行い、出力時は連番IDで参照する。接続先がスライド内に見つからない場合は `from` / `to` を省略する。

#### 画像フィールド（`--extract-images` 指定時）

画像（`p:pic`）は `type` が `"picture"` となり、以下の追加フィールドを持つ。`--extract-images` 未指定時は画像要素自体がスキップされる。

```json
{"id": 5, "type": "picture", "name": "図 1", "position": {"x": 1000000, "y": 1000000, "cx": 5000000, "cy": 3000000}, "z": 4, "alt_text": "システム構成図", "image": {"format": "png", "width": 640, "height": 480, "size": 45230, "path": "/tmp/imgs/image_1.png"}}
```

| フィールド | 型 | 説明 |
|-----------|-----|------|
| `alt_text` | string | 代替テキスト（`cNvPr` の `descr` 属性）。設定されていない場合は省略 |
| `image` | object | 画像メタデータ |

**`image` オブジェクト:**

| フィールド | 型 | 説明 |
|-----------|-----|------|
| `format` | string | 画像形式（`png`, `jpeg`, `emf` 等。リレーション先のファイル拡張子から判定） |
| `width` | number | 画像の幅（ピクセル）。`a:ext` の `cx` をEMUからピクセルに変換（÷ 9525） |
| `height` | number | 画像の高さ（ピクセル）。`a:ext` の `cy` をEMUからピクセルに変換（÷ 9525） |
| `size` | number | ファイルサイズ（バイト）。ZIPエントリから取得 |
| `path` | string | 抽出先のファイルパス |

画像ファイルはスライドの `.rels` から `p:blipFill/a:blip` の `r:embed` 属性で参照されるリレーションIDを解決し、ZIP内の `ppt/media/` 配下から抽出する。

#### グループフィールド

グループ（`p:grpSp`）は `type` が `"group"` となり、以下の追加フィールドを持つ。

| フィールド | 型 | 説明 |
|-----------|-----|------|
| `children` | array | 子要素の配列。各子要素は図形フィールドと同じ構造 |

- グループの `children` 内の子要素の `z` はグループ内での相対順序（0始まり）
- グループ自体の `z` はスライドレベルでの重なり順
- ネストしたグループも同じ構造で再帰的に表現する
- グループの `position` はグループ全体の位置とサイズ

#### テーブルフィールド

テーブル（`p:graphicFrame` 内の `a:tbl`）は `type` が `"table"` となり、`paragraphs` の代わりに `table` フィールドを持つ。

```json
{"id": 4, "type": "table", "name": "表 1", "position": {"x": 457200, "y": 1600200, "cx": 8229600, "cy": 3000000}, "z": 3, "table": {"cols": 3, "rows": [["項目", "説明", "備考"], ["機能A", null, "必須"]]}}
```

| フィールド | 型 | 説明 |
|-----------|-----|------|
| `cols` | number | 列数 |
| `rows` | array | 行の配列。各行は `cols` と同じ長さの配列。セルの値は文字列または `null` |

- 各セルのテキストは `a:txBody` 内の全 `a:t` を結合したプレーンテキスト。セル内の複数段落は `\n` で結合する
- 結合セル（`vMerge`, `hMerge`）の被結合セルは `null` となる（結合元セルにテキストが格納される）
- テーブル要素にはセルレベルのフォント等の書式は出力しない（テーブル内のセル書式は複雑であり、テキスト取得が主目的のため）

#### 図形種別

`type` フィールドの値は以下のいずれか:

- **コネクタ**: 常に `"connector"`（`p:cxnSp` 要素に対応）
- **グループ**: 常に `"group"`（`p:grpSp` 要素に対応）
- **画像**: 常に `"picture"`（`p:pic` 要素に対応）
- **テーブル**: 常に `"table"`（`p:graphicFrame` + `a:tbl` に対応）
- **シェイプ**: `a:prstGeom` の `prst` 属性値をそのまま使用する（`p:sp` 要素に対応）

シェイプの `prst` 値の例はcc-read-excelの設計ドキュメントと同一（`rect`, `roundRect`, `ellipse`, `flowChartProcess`, `flowChartDecision` 等）。

`a:prstGeom` が存在しない場合（カスタムジオメトリ `a:custGeom`）は `type` を `"customShape"` とする。

#### 図形の出力順

1. プレースホルダー（タイトル → サブタイトル → ボディ → その他のプレースホルダー）
2. 非プレースホルダーの図形（spTree 内の出現順）

テキストが空かつ書式情報もない図形（空のプレースホルダー等）はスキップする。

#### ノート（`--notes` 指定時）

スライドオブジェクトに `notes` フィールドが追加される。段落の配列で、`paragraphs` 配列の要素と同じ構造。

```json
{"slide": 1, "title": "基本設計書", "shapes": [...], "notes": [{"text": "このスライドでは基本設計の概要を説明する。"}, {"text": "スコープの定義", "bullet": "•"}, {"text": "前提条件の確認", "bullet": "•"}]}
```

- ノートスライド（`ppt/notesSlides/`）内の `body` プレースホルダーのテキストを取得する
- ノートが存在しない、またはテキストが空のスライドでは `notes` フィールドを省略する

### `cc-read-pptx search [options] <file>`

**役割:** プレゼンテーション内のテキストを検索する。全スライドを `slides` で取得してフィルタするより効率的。「特定のキーワードを含むスライドと要素を特定する」という使い方ができる。

**オプション:**

| オプション | 説明 | デフォルト |
|-----------|------|-----------|
| `--text <text>` | 検索文字列（部分一致、大文字小文字無視） | 必須 |
| `--slide <number>` | 対象スライド番号 | 全スライド |
| `--notes` | ノートも検索対象にする | OFF |

- `--text` は必須パラメータ
- 全角・半角は区別する。正規表現には対応しない

**検索の挙動:**

- 各段落の `text` フィールドに対して大文字・小文字を区別しない部分一致検索を行う
- マッチした段落を含む図形のみを `shapes` 配列に出力する。図形内ではマッチした段落のみを `paragraphs` に含める
- テーブル内のセルテキストも検索対象にする（いずれかのセルにヒットした場合はテーブル要素全体を出力する）
- `--notes` 指定時はノートの各段落も検索対象にする
- マッチがないスライドは出力しない
- 結果なしでも正常終了（終了コード 0）する

**出力形式:**

`slides` コマンドと同じ1スライド1JSONの形式。マッチしたスライドのみ出力する。

```bash
cc-read-pptx search --text "データ" example.pptx
```

```jsonl
{"slide":2,"title":"システム構成","shapes":[{"id":3,"type":"rect","name":"テキストボックス 1","position":{"x":1000000,"y":2000000,"cx":3000000,"cy":500000},"z":2,"paragraphs":[{"text":"データフロー図"}]}]}
{"slide":4,"title":"処理フロー","shapes":[{"id":2,"type":"rect","placeholder":"body","position":{"x":457200,"y":1600200,"cx":8229600,"cy":4525963},"z":1,"paragraphs":[{"text":"データ取得処理","bullet":"•"}]}]}
```

## 技術選定

- **言語:** Go
- **PPTXパーサー:** 自前実装（ZIP + encoding/xml による直接パース）
- **CLIフレームワーク:** [cobra](https://github.com/spf13/cobra)

## エラーハンドリング

- ファイルが存在しない / 読み取れない → エラーメッセージを stderr に出力
- .pptx 以外のファイル → 「.pptx 形式のみ対応」のエラーメッセージ
- 存在しないスライド番号 → 利用可能なスライド番号の範囲をエラーメッセージに含める
- パスワード保護されたファイル → 非対応としてエラーメッセージを出力
- 破損したファイル（不正なzip構造等） → エラーメッセージを出力
- `search` で結果が0件の場合 → 空出力で正常終了する
- 終了コード: 0=成功（検索結果なしも含む）、1=エラー
- エラーメッセージは stderr に `cc-read-pptx: <メッセージ>` の形式で出力する。stdout には常にJSONのみを出力する

## 設計方針

- 対応形式は .pptx のみ（.ppt は非対応）
- ZIP 内の XML を自前で直接パースする（外部 PowerPoint パーサーは使用しない）
- スライド XML はDOMパースで処理する（Excelのワークシートと異なり、スライドのXMLは通常小さいためSAXパースの必要性が低い）
- 出力は1スライド1JSONオブジェクト（JSONL形式）。`info` のみファイル全体を1JSONで出力
- デフォルトでテキストが空の図形をスキップする
- 書式情報は常に出力する（PowerPoint は書式自体がコンテンツの一部であり、1スライドあたりの要素数も少ないため）。デフォルト値のフィールドは省略する
- 色は可能な限り `#RRGGBB` 形式で出力する。テーマカラーは theme1.xml から RGB に変換し、tint 値がある場合は HSL 色空間で明度を調整して適用する
- 出力は常にUTF-8エンコーディング
- テキスト内の制御文字（改行、タブ等）はJSON仕様に従いエスケープする
- 位置情報はEMU単位で出力する（PPTX内部の座標系をそのまま使用し、変換誤差を防ぐ）
- スライドマスター・レイアウトからの書式継承はベストエフォートとする（完全な継承チェーンの解決は複雑であり、直接指定された書式を優先する）
- グループの子要素は `children` 配列にネストする（フラットなID参照ではなく再帰構造）

## コマンド構成（その他）

### `cc-read-pptx version`

バージョン情報をプレーンテキストで出力する。

```
cc-read-pptx version 0.1.0
```

バージョン番号は `go build -ldflags` でビルド時に埋め込む。未設定の場合は `dev` を表示する。

## 対応しない機能

以下の機能は意図的に対応しない。理由とともに記録する。

| 機能 | 理由 |
|------|------|
| .ppt 形式の読み取り | OOXML（.pptx）のみ対応。旧形式のファイルは事前に .pptx へ変換して使用する |
| パスワード保護されたファイル | 非対応 |
| アニメーション・トランジション | 静的なコンテンツ読み取りが目的。`p:timing` 要素は無視する |
| 埋め込みチャート | 独自のXML体系（`c:chartSpace`）で複雑。別途対応を検討 |
| SmartArt | 独自のXML体系（`dgm:`）で複雑。PowerPoint保存時に描画キャッシュとして `grpSp` に展開されるため、通常はグループとして読み取り可能 |
| 埋め込み動画・音声 | メディアファイルの内容はCLIで扱いにくい。画像のみ `--extract-images` で対応する |
| スライドマスター・レイアウトの完全な書式継承 | 継承チェーンの完全な解決は複雑。直接指定された書式のみを出力する |
| 塗りつぶしのパターン・グラデーション | ソリッド塗りつぶし（単色）のみ対応 |
| PowerPoint書き込み・編集 | 読み取り専用ツールとして設計 |
| 複数ファイルの同時処理 | 1コマンド1ファイルの原則。複数ファイルを処理する場合はファイルごとにコマンドを実行する |
| メモリ使用量の上限 | 設けない。スライドXMLは通常小さいため問題にならない |
