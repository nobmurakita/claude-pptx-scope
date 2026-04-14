# pptx-scope 設計ドキュメント

PowerPoint ファイル（.pptx）の内容をCLIから出力するGoツール。
Claude CodeがPowerPoint資料（プレゼン、フローチャート、仕様書など）を読み取る用途を主眼とする。

## コマンド構成と利用フロー

Claude Code からの典型的な利用フローは以下の通り:

1. **`info`** — ファイルの全体像を把握する（スライド一覧、スライドサイズ）。どのスライドを読むべきか判断する
2. **`slides`** — スライドの内容を取得する。図形・テキスト・テーブル・コネクタ・画像をまとめて出力する
3. **`image`** — 画像を一時ファイルに保存する。`slides` 出力の `image_id` を指定して個別に抽出する
4. **`search`** — プレゼンテーション内のテキストを検索する。全スライドまたは指定スライドから条件に合うテキストを抽出する
5. **`version`** — バージョン情報を表示する

基本的には `info` → `slides` で内容を把握する。特定のキーワードを探す場合は `search` が効率的。

### `pptx-scope info <file>`

**役割:** ファイルレベルの概要を把握する。スライド一覧から対象スライドを特定し、以降の `slides` / `search` に渡す `--slide` を決定する。

ファイルレベルの概要をJSONL形式で出力する。メタ情報行に続いてスライド情報を1行ずつ出力する。

- スライド一覧（番号、タイトル、ノートの有無）
- スライドサイズ

**出力例:**

```jsonl
{"file":"基本設計書.pptx","slide_size":{"width":720,"height":540}}
{"slide":1,"title":"基本設計書","has_notes":true}
{"slide":2,"title":"目次"}
{"slide":3,"title":"システム構成","has_notes":true}
{"slide":4}
{"slide":5,"title":"フロー図","has_images":true,"hidden":true}
```

**スライド行の各フィールド:**

| フィールド | 型 | 説明 |
|-----------|-----|------|
| `slide` | number | スライド番号（1始まり。presentation.xml の `sldIdLst` の順序に基づく） |
| `title` | string | スライドタイトル。タイトルプレースホルダー（`ph type="title"` または `ph type="ctrTitle"`）のテキストから取得。存在しない場合は省略 |
| `has_notes` | bool | ノートスライドが存在しテキストがある場合のみ `true` を出力 |
| `has_images` | bool | 画像を含む場合のみ `true` を出力（グループ内の画像も再帰的に検出） |
| `hidden` | bool | 非表示スライドの場合のみ `true` を出力（`p:sld` の `show="0"` 属性） |

**`slide_size` オブジェクト:**

| フィールド | 型 | 説明 |
|-----------|-----|------|
| `width` | number | スライド幅（pt単位。標準4:3 = 720, 16:9 = 960） |
| `height` | number | スライド高さ（pt単位。標準4:3 = 540, 16:9 = 540） |

タイトルの取得方法:
- スライドの `p:spTree` 内で `p:nvSpPr/p:nvPr/p:ph` の `type` 属性が `title` または `ctrTitle` のシェイプを探す
- 該当シェイプの `p:txBody` 内の全 `a:t` 要素を結合する
- タイトルプレースホルダーが存在しない場合は `title` フィールドを省略する

### `pptx-scope slides [options] <file>`

**役割:** スライドの内容を取得する。スライドヘッダ行に続いて、図形を1つずつ個別のJSONL行として出力する。

**オプション:**

| オプション | 説明 | デフォルト |
|-----------|------|-----------|
| `--slide <number,...>` | 対象スライド番号（1始まり、複数指定可: `--slide 1,3`） | 全スライド |
| `--notes` | ノートも出力する | OFF |

- `--slide` 未指定時は全スライドを順番に出力する
- 画像は `image_id` フィールドにZIP内のメディアパスを出力する（`image` サブコマンドで個別に取得可能）
- 書式情報は常に出力する

**出力形式:**

スライドヘッダ行（`shapes` は図形数）に続いて、図形を1つずつ個別の行として出力する。

**出力例:**

```jsonl
{"slide":1,"title":"基本設計書","shapes":2,"has_notes":true}
{"style":1,"name":"メイリオ","size":36,"bold":true,"color":"#333333"}
{"shape":1,"type":"rect","placeholder":"ctrTitle","pos":{"x":54,"y":180,"w":612,"h":115.75},"z":0,"alignment":{"vertical":"center"},"paragraphs":[{"text":"基本設計書","s":1,"alignment":{"horizontal":"center"}}]}
{"shape":2,"type":"rect","placeholder":"subTitle","pos":{"x":108,"y":306,"w":504,"h":138},"z":1,"paragraphs":[{"text":"2025年4月版"}]}
{"slide":2,"title":"目次","shapes":2}
{"shape":1,"type":"rect","placeholder":"title","pos":{"x":36,"y":21.63,"w":648,"h":90},"z":0,"paragraphs":[{"text":"目次"}]}
{"shape":2,"type":"rect","placeholder":"body","pos":{"x":36,"y":126,"w":648,"h":356.38},"z":1,"paragraphs":[{"text":"システム概要","bullet":"1."},{"text":"機能一覧","bullet":"2."},{"text":"データフロー","bullet":"3."}]}
```

整形すると以下のような構造（1つの図形行）:

```json
{
  "shape": 2,
  "type": "rect",
  "placeholder": "body",
  "pos": {"x": 36, "y": 126, "w": 648, "h": 356.38},
  "z": 1,
  "paragraphs": [
    {"text": "システム概要", "bullet": "1."},
    {"text": "機能一覧", "bullet": "2."},
    {"text": "データフロー", "bullet": "3."}
  ]
}
```

**スライドヘッダ行のフィールド:**

| フィールド | 型 | 説明 |
|-----------|-----|------|
| `slide` | number | スライド番号 |
| `title` | string | スライドタイトル（存在する場合のみ） |
| `shapes` | number | スライド内の図形数 |
| `has_notes` | bool | ノートにテキストがある場合のみ `true` |
| `has_images` | bool | 画像を含む場合のみ `true` |
| `hidden` | bool | 非表示スライドの場合のみ `true` |

#### 図形フィールド

`shapes` 配列の各要素:

| フィールド | 型 | 説明 |
|-----------|-----|------|
| `shape` | number | スライド内の図形ID。spTree 内の出現順で1始まりの連番を割り当てる |
| `type` | string | 図形種別。後述の図形種別を参照 |
| `name` | string | 図形名（`p:cNvPr` の `name` 属性）。プレースホルダーの場合は省略 |
| `placeholder` | string | プレースホルダー種別（`title`, `ctrTitle`, `subTitle`, `body`, `dt`, `ftr`, `sldNum` 等）。プレースホルダーでない場合は省略 |
| `pos` | object | 位置とサイズ（`a:xfrm` の `a:off` と `a:ext`。pt単位）。プレースホルダーでスライド上に未指定の場合、レイアウト/マスターから継承する |
| `z` | number | Z-order。spTree 内の出現順で0始まり。大きいほど前面 |
| `rotation` | number | 回転角度（度単位、時計回り）。`a:xfrm` の `rot` 属性を60000で除算。0の場合は省略 |
| `flip` | string | 反転。`"h"`, `"v"`, `"hv"`。なければ省略 |
| `fill` | string | 塗りつぶし色（`#RRGGBB` 形式）。`a:solidFill` を優先し、なければ `a:gradFill` の最初のストップカラーを代表色として使用する。塗りつぶしなしの場合は省略 |
| `line` | object | 枠線情報。枠線がない場合は省略 |
| `callout_pointer` | object | 吹き出しのポインタ位置（吹き出し図形の場合のみ。後述） |
| `alignment` | object | テキストの垂直配置（`a:bodyPr` の `anchor` 属性）。デフォルトの上揃えは省略 |
| `text_margin` | object | テキストボディの内部マージン（`a:bodyPr` の `lIns/rIns/tIns/bIns`）。OOXMLデフォルト値と同じ場合は省略 |
| `link` | object | 図形全体に設定されたハイパーリンク（後述）。ない場合は省略 |
| `paragraphs` | array | 段落の配列。テキストがない場合は省略 |
| `table` | object | テーブルデータ（テーブルの場合。`paragraphs` の代わりに使用） |

**`pos` オブジェクト:**

| フィールド | 型 | 説明 |
|-----------|-----|------|
| `x` | number | 左上X座標（pt） |
| `y` | number | 左上Y座標（pt） |
| `w` | number | 幅（pt） |
| `h` | number | 高さ（pt） |

**`line` オブジェクト:**

| フィールド | 型 | 説明 |
|-----------|-----|------|
| `color` | string | 線の色（`#RRGGBB` 形式） |
| `style` | string | 線のスタイル。`solid`, `dash`, `dot`, `dashDot` 等 |
| `width` | number | 線幅（pt単位）。`a:ln` の `w` 属性値をEMUからptに変換 |

#### 段落フィールド

`paragraphs` 配列の各要素:

| フィールド | 型 | 説明 |
|-----------|-----|------|
| `text` | string | 段落のプレーンテキスト |
| `bullet` | string | 箇条書き記号。箇条書きでない段落は省略 |
| `level` | number | インデントレベル（0始まり）。0の場合は省略 |
| `margin_left` | number | 段落の左マージン（pt単位、`a:pPr` の `marL` 属性）。0または未指定の場合は省略 |
| `indent` | number | 段落の1行目インデント（pt単位、`a:pPr` の `indent` 属性）。負値でぶら下がりインデント。0または未指定の場合は省略 |
| `line_spacing` | string | 行間（`a:lnSpc`）。`"150%"` のようなパーセント、または `"12pt"` のようなポイント指定。デフォルト（100%）は省略 |
| `space_before` | string | 段落前のスペース（`a:spcBef`）。同上の書式。0pt は省略 |
| `space_after` | string | 段落後のスペース（`a:spcAft`）。同上の書式。0pt は省略 |
| `s` | number | スタイルID。`style` 行で定義されたフォント情報への参照 |
| `alignment` | object | 水平配置情報。デフォルトの左揃えは省略 |
| `link` | object | ハイパーリンク（段落内の全テキストが同一リンクの場合。後述） |
| `rich_text` | array | リッチテキストラン（段落内に書式やリンクの異なるランが存在する場合のみ） |

テキストが空の段落は出力しない（箇条書き間の空行等）。

**段落内の改行（`a:br`）:**

`a:br` 要素は段落内の改行（Shift+Enter）を表す。`text` フィールドでは `\n` として出力される。`rich_text` モードでは `{"text": "\n"}` というランとして出力される。

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

**`style` 行（フォント情報）:**

```jsonl
{"style":1,"name":"メイリオ","size":11,"bold":true,"italic":true,"strikethrough":true,"underline":"single","color":"#FF0000","highlight":"#FFFF00","baseline":"super","cap":"all"}
```

- すべてのフォント情報は `style` 行に抽出され、段落/リッチテキストランでは `s` フィールドで参照する（後述の「スタイル重複排除」を参照）
- デフォルト値のフィールドは省略する
- `size` は pt単位。hundredths of point（`a:rPr` の `sz` 属性）を pt に変換（÷ 100）
- `color` は文字色、`highlight` は文字の背景色（`a:highlight` 要素）。いずれも `#RRGGBB` 形式
- `baseline` は上付き/下付き文字（`a:rPr` の `baseline` 属性）。正値 → `"super"`、負値 → `"sub"`、0/未指定 → 省略
- `cap` は英字の大文字化（`a:rPr` の `cap` 属性）。`"all"`（すべて大文字）/`"small"`（スモールキャップ）。`"none"`/未指定は省略
- プレースホルダー図形では、`name`, `size`, `color`, `highlight`, `baseline`, `cap` はスライドマスター・レイアウトからの継承値で補完される（後述の「プレースホルダーの継承」を参照）。テーマフォント参照（`+mj-lt` 等）は実フォント名に解決される

**リッチテキスト:**

段落内のテキストの一部が異なる書式を持つ場合、`text` にはプレーンテキストを格納し、`rich_text` に書式付きランの配列を格納する。

```jsonl
{"style":1,"bold":true,"color":"#FF0000"}
{"shape":1,...,"paragraphs":[{"text":"重要: この機能は必須です","rich_text":[{"text":"重要: ","s":1},{"text":"この機能は必須です"}]}]}
```

- `text` は常にプレーンテキスト（検索・概要把握用）
- `rich_text` は書式またはリンクの異なるランが存在する場合のみ出力
- 各ランの `s` はスタイルIDへの参照（`style` 行で定義）
- 各ランの `link` はハイパーリンクがある場合のみ出力

#### ハイパーリンク

ハイパーリンクは図形レベル（`p:cNvPr` の `a:hlinkClick`）とテキストランレベル（`a:rPr` の `a:hlinkClick`）の2箇所に設定できる。

```json
{"shape": 1, "type": "rect", "pos": {"x": 0, "y": 0, "w": 393.7, "h": 39.37}, "z": 0, "paragraphs": [{"text": "詳細はこちら", "link": {"url": "https://example.com"}}]}
```

```json
{"text": "スライド3を参照", "link": {"slide": 3}}
```

**`link` オブジェクト:**

| フィールド | 型 | 説明 |
|-----------|-----|------|
| `url` | string | 外部URL（`http://`, `https://`, `mailto:` 等）。外部リンクの場合のみ |
| `slide` | number | リンク先のスライド番号。スライド内リンクの場合のみ |

- 図形レベルのリンクは `Shape.link` に格納。図形全体がクリック対象となるリンク
- テキストランレベルのリンクは、段落内の全テキストが同一リンクなら `Paragraph.link` に格納。異なるリンクが混在する場合は `rich_text` の各ランの `link` に格納
- リンクの種別は `.rels` のリレーションで判別する。`action="ppaction://hlinksldjump"` の場合はスライド内リンクとしてスライド番号に変換。それ以外は外部URLとして扱う
- `ppaction://media` 等のメディアアクションは無視する（`r:id` が空のため）

#### 吹き出しポインタ

吹き出し図形（`type` が `wedgeRectCallout`, `wedgeRoundRectCallout`, `wedgeEllipseCallout`, `cloudCallout`, `borderCallout1`, `borderCallout2`, `borderCallout3` 等）の場合、ポインタの指す位置をスライド上の絶対座標に変換して出力する。

```json
{"shape": 3, "type": "wedgeRoundRectCallout", "pos": {"x": 236.22, "y": 78.74, "w": 157.48, "h": 62.99}, "z": 2, "callout_pointer": {"x": 118.11, "y": 196.85}, "paragraphs": [{"text": "ここに注目"}]}
```

**`callout_pointer` オブジェクト:**

| フィールド | 型 | 説明 |
|-----------|-----|------|
| `x` | number | ポインタ先端のX座標（pt。スライド上の絶対座標） |
| `y` | number | ポインタ先端のY座標（pt。スライド上の絶対座標） |

**算出方法:**

調整ハンドル `adj1`（X方向）と `adj2`（Y方向）は図形の中心を原点とした1/100000単位の相対比率で表される。これを図形の位置・サイズからスライド上の絶対座標に変換する:

- `pointer_x = pos.x + pos.w / 2 + adj1 * pos.w / 100000`
- `pointer_y = pos.y + pos.h / 2 + adj2 * pos.h / 100000`

吹き出し図形でない場合は `callout_pointer` フィールドを省略する。調整ハンドルが未指定の場合はOOXML仕様のデフォルト値を使用する。

#### コネクタフィールド

コネクタ（`p:cxnSp`）は `type` が `"connector"` となり、以下の追加フィールドを持つ。

| フィールド | 型 | 説明 |
|-----------|-----|------|
| `from` | number | 接続元の図形ID。接続情報がない場合は省略 |
| `to` | number | 接続先の図形ID。接続情報がない場合は省略 |
| `from_idx` | number | 接続元の接続ポイントインデックス。矩形の場合 0=上, 1=左, 2=下, 3=右。接続情報がない場合は省略 |
| `to_idx` | number | 接続先の接続ポイントインデックス。`from_idx` と同様。接続情報がない場合は省略 |
| `connector_type` | string | コネクタ形状。`a:prstGeom` の `prst` 属性値。`straightConnector1`（直線）、`bentConnector2`（L字・1回屈曲）、`bentConnector3`（コの字・2回屈曲）、`curvedConnector3`（曲線）等 |
| `arrow` | string | 矢印ヘッドの位置。`"start"` は `start` 座標側、`"end"` は `end` 座標側、`"both"` は両端。矢印なしの場合は省略 |
| `label` | string | コネクタ上のテキスト。テキストがない場合は省略 |

コネクタの `from` / `to` は、`p:cxnSp` 内の `a:stCxn` / `a:endCxn` 要素の `id` 属性を参照する。この `id` は PowerPoint が付与する図形IDであり、`slides` コマンドが割り当てる連番 `shape` とは異なる。パース時に PowerPoint 図形IDから連番IDへのマッピングを行い、出力時は連番IDで参照する。接続先がスライド内に見つからない場合は `from` / `to` を省略する。

#### 画像フィールド

画像（`p:pic`）は `type` が `"picture"` となり、以下の追加フィールドを持つ。

```json
{"shape": 5, "type": "picture", "name": "図 1", "pos": {"x": 78.74, "y": 78.74, "w": 393.7, "h": 236.22}, "z": 4, "alt_text": "システム構成図", "image_id": "ppt/media/image1.png"}
```

| フィールド | 型 | 説明 |
|-----------|-----|------|
| `alt_text` | string | 代替テキスト（`cNvPr` の `descr` 属性）。設定されていない場合は省略 |
| `image_id` | string | ZIP内のメディアパス。`image` サブコマンドに渡して画像を取得する |

画像パスはスライドの `.rels` から `p:blipFill/a:blip` の `r:embed` 属性で参照されるリレーションIDを解決し、ZIP内のメディアパス（例: `ppt/media/image1.png`）として出力する。

#### グループフィールド

グループ（`p:grpSp`）は `type` が `"group"` となり、以下の追加フィールドを持つ。

| フィールド | 型 | 説明 |
|-----------|-----|------|
| `children` | array | 子要素の配列。各子要素は図形フィールドと同じ構造 |

- グループの `children` 内の子要素の `z` もスライド全体で一意の連番（グループ内でリセットされない）
- ネストしたグループも同じ構造で再帰的に表現する
- グループの `pos` はグループ全体の位置とサイズ

#### テーブルフィールド

テーブル（`p:graphicFrame` 内の `a:tbl`）は `type` が `"table"` となり、`paragraphs` の代わりに `table` フィールドを持つ。

```json
{"shape": 4, "type": "table", "name": "表 1", "pos": {"x": 36, "y": 126, "w": 648, "h": 236.22}, "z": 3, "table": {"cols": 3, "rows": [[{"text": "項目"}, {"text": "説明"}, {"text": "備考"}], [{"text": "機能A", "paragraphs": [{"text": "機能A", "font": {"bold": true}}]}, null, {"text": "必須"}]]}}
```

| フィールド | 型 | 説明 |
|-----------|-----|------|
| `cols` | number | 列数 |
| `rows` | array | 行の配列。各行は `cols` と同じ長さの配列。セルの値はオブジェクトまたは `null` |

セルオブジェクトのフィールド:

| フィールド | 型 | 説明 |
|-----------|-----|------|
| `text` | string | セル内の全テキストを結合したプレーンテキスト |
| `fill` | string | セルの背景色（`#RRGGBB`）。`a:tcPr` の `solidFill`（なければ `gradFill` の代表色）。未指定の場合は省略 |
| `border_left` | object | 左罫線情報（`color`, `style`, `width`。`line` オブジェクトと同じ構造）。未指定の場合は省略 |
| `border_right` | object | 右罫線情報。同上 |
| `border_top` | object | 上罫線情報。同上 |
| `border_bottom` | object | 下罫線情報。同上 |
| `paragraphs` | array | 段落情報（フォント・箇条書き等の書式情報がある場合のみ） |

- `text` は `a:txBody` 内の全 `a:t` を結合したプレーンテキスト。セル内の複数段落はスペースで結合する
- `paragraphs` はフォント・リッチテキスト・箇条書き等の書式情報がある場合のみ出力される（プレーンテキストのみの場合は省略）
- 結合セル（`vMerge`, `hMerge`）の被結合セルは `null` となる（結合元セルにテキストが格納される）
- 罫線は `a:tcPr` の `lnL`/`lnR`/`lnT`/`lnB` 要素から解決する。4辺個別に `border_left`/`border_right`/`border_top`/`border_bottom` として出力する

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
非表示図形（`p:cNvPr` の `hidden="1"`）もスキップする。

#### ノート（`--notes` 指定時）

スライドの図形行の後に、ノートを独立したJSONL行として出力する。段落の配列で、`paragraphs` 配列の要素と同じ構造。

```jsonl
{"slide":1,"title":"基本設計書","shapes":2}
{"shape":1,"type":"rect","placeholder":"ctrTitle",...}
{"shape":2,"type":"rect","placeholder":"subTitle",...}
{"notes":[{"text":"このスライドでは基本設計の概要を説明する。"},{"text":"スコープの定義","bullet":"•"},{"text":"前提条件の確認","bullet":"•"}]}
```

- ノートスライド（`ppt/notesSlides/`）内の `body` プレースホルダーのテキストを取得する
- ノートが存在しない、またはテキストが空のスライドではノート行を出力しない

### `pptx-scope image <file> <image_id>`

**役割:** 画像を一時ファイルに保存する。`slides` 出力の `image_id`（ZIP内のメディアパス）を指定して、必要な画像だけをオンデマンドで取得する。

**引数:**

| 引数 | 説明 |
|------|------|
| `<file>` | 対象のPowerPointファイル |
| `<image_id>` | `slides` 出力の `image_id` フィールド値（例: `ppt/media/image1.png`） |

**使用例:**

```bash
pptx-scope image example.pptx ppt/media/image1.png
# stdout: {"file":"/var/folders/.../pptx-scope-1234567.png"}
```

`image_id` はZIP内のメディアファイルパスそのものであり、ステートレスにZIPから直接読み出す。マッピングテーブルや事前のパース処理は不要。

### `pptx-scope search [options] <file>`

**役割:** プレゼンテーション内のテキストを検索する。全スライドを `slides` で取得してフィルタするより効率的。「特定のキーワードを含むスライドと要素を特定する」という使い方ができる。

**オプション:**

| オプション | 説明 | デフォルト |
|-----------|------|-----------|
| `--text <text>` | 検索文字列（部分一致、大文字小文字無視） | 必須 |
| `--slide <number,...>` | 対象スライド番号（複数指定可） | 全スライド |
| `--notes` | ノートも検索対象にする | OFF |

- `--text` は必須パラメータ
- 全角・半角は区別する。正規表現には対応しない

**検索の挙動:**

- 各図形のテキストに対して大文字・小文字を区別しない部分一致検索を行う
- テーブル内のセルテキストも検索対象にする
- コネクタのラベルテキストも検索対象にする
- `--notes` 指定時はノートのテキストも検索対象にする
- マッチがないスライドは出力しない
- 結果なしでも正常終了（終了コード 0）する

**出力形式:**

マッチしたスライドのヘッダ行のみを出力する（`info` と同じ形式）。詳細は `slides --slide` で取得する。

```bash
pptx-scope search --text "データ" example.pptx
```

```jsonl
{"slide":2,"title":"システム構成","has_images":true}
{"slide":4,"title":"処理フロー"}
```

### `pptx-scope version`

**役割:** バージョン情報を表示する。

**出力:**

```
v0.0.9
```

- バージョン文字列はビルド時に `-ldflags="-X github.com/nobmurakita/claude-pptx-scope/internal/cmd.Version=<tag>"` で埋め込む。リリースビルド（GitHub Actions）では `GITHUB_REF_NAME`（タグ名）が設定される
- ldflags 未指定時のデフォルトは `latest`
- 他コマンドと異なり、一時ファイルへの書き出しは行わない

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
- エラーメッセージは stderr に `pptx-scope: <メッセージ>` の形式で出力する。stdout には常にJSONのみを出力する

## 設計方針

- 対応形式は .pptx のみ（.ppt は非対応）
- ZIP 内の XML を自前で直接パースする（外部 PowerPoint パーサーは使用しない）
- スライド XML はDOMパースで処理する（Excelのワークシートと異なり、スライドのXMLは通常小さいためSAXパースの必要性が低い）
- 全コマンドの出力は一時ファイルに書き出され、stdout にはファイルパスと行数のJSON（`{"file":"...","lines":N}`）のみを出力する。`--stdout` フラグで標準出力に直接書き出すことも可能（デバッグ用）
- 出力は全コマンドでJSONL形式。`info` はメタ情報行＋スライド行、`slides` はスライドヘッダ行＋図形1つ1行
- デフォルトでテキストが空の図形をスキップする
- 書式情報は常に出力する（PowerPoint は書式自体がコンテンツの一部であり、1スライドあたりの要素数も少ないため）。デフォルト値のフィールドは省略する
- 色は可能な限り `#RRGGBB` 形式で出力する。テーマカラーは theme1.xml から RGB に変換し、tint 値がある場合は HSL 色空間で明度を調整して適用する。グラデーション塗りつぶし（`a:gradFill`）は最初のストップカラーを代表色として使用する（色変換も適用される）
- 出力は常にUTF-8エンコーディング
- テキスト内の制御文字（改行、タブ等）はJSON仕様に従いエスケープする
- 数値はすべてpt単位で出力する（位置・サイズ・フォントサイズ・線幅。1pt = 12700 EMU。回転角度のみ度単位）
- プレースホルダーのフォント・位置・箇条書きはスライドマスター・レイアウトから継承する（後述の「プレースホルダーの継承」を参照）
- テーマフォント参照（`+mj-lt`, `+mn-lt` 等）は theme1.xml の `a:majorFont` / `a:minorFont` から実フォント名に解決する
- グループの子要素は `children` 配列にネストする（フラットなID参照ではなく再帰構造）
- AI エージェントが Read ツールで分割読みする前提で、出力は一時ファイルに書き出す。Bash ツールの出力サイズ制限（30K文字）を回避しつつ CLI 実行回数を最小化する

## 対応しない機能

以下の機能は意図的に対応しない。理由とともに記録する。

| 機能 | 理由 |
|------|------|
| .ppt 形式の読み取り | OOXML（.pptx）のみ対応。旧形式のファイルは事前に .pptx へ変換して使用する |
| パスワード保護されたファイル | 非対応 |
| アニメーション・トランジション | 静的なコンテンツ読み取りが目的。`p:timing` 要素は無視する |
| 埋め込みチャート | 独自のXML体系（`c:chartSpace`）で複雑。別途対応を検討 |
| SmartArt | 独自のXML体系（`dgm:`）で複雑。PowerPoint保存時に描画キャッシュとして `grpSp` に展開されるため、通常はグループとして読み取り可能 |
| 埋め込み動画・音声 | メディアファイルの内容はCLIで扱いにくい。画像のみ `image` サブコマンドで対応する |
| 塗りつぶしのパターン | パターン塗りつぶしは非対応。グラデーション（`a:gradFill`）は最初のストップカラーを代表色として出力する |
| PowerPoint書き込み・編集 | 読み取り専用ツールとして設計 |
| 複数ファイルの同時処理 | 1コマンド1ファイルの原則。複数ファイルを処理する場合はファイルごとにコマンドを実行する |
| メモリ使用量の上限 | 設けない。スライドXMLは通常小さいため問題にならない |

## プレースホルダーの継承

PowerPointはスライドマスター → スライドレイアウト → スライドの3段階で書式を継承する。プレースホルダー図形に対して、スライド上で明示的に指定されていないプロパティを上位から補完する。

### 継承対象

| プロパティ | 説明 |
|-----------|------|
| 位置・サイズ（`a:xfrm`） | スライドのプレースホルダーに `a:xfrm` がない場合、レイアウト → マスターの順で取得する |
| フォント名 | `a:defRPr` の `a:latin` / `a:ea` から取得。テーマフォント参照（`+mj-lt` 等）は実フォント名に解決する |
| フォントサイズ | `a:defRPr` の `sz` 属性から取得 |
| フォント色 | `a:defRPr` の `a:solidFill` から取得 |
| 太字・斜体・下線・取り消し線 | `a:defRPr` の `b`, `i`, `u`, `strike` 属性から取得。テキストラン（`a:rPr`）で明示的に指定されていない場合のみ継承する（`a:rPr` の属性が空文字列 = 未指定の場合に継承、`"0"` / `"false"` 等の明示的な値がある場合は継承しない） |
| 箇条書きスタイル | `a:buChar` / `a:buAutoNum` / `a:buNone` を取得 |

### 継承チェーン

各プロパティは以下の順で検索し、最初に見つかった値を使用する。スライド上で明示的に指定された値は常に優先される。

1. スライド上のテキストラン（`a:rPr`）の直接指定
2. スライドのプレースホルダーの `txBody > lstStyle > lvlNpPr > defRPr`
3. スライドレイアウトの対応プレースホルダーの `txBody > lstStyle > lvlNpPr > defRPr`
4. スライドマスターの対応プレースホルダーの `txBody > lstStyle > lvlNpPr > defRPr`
5. スライドマスターの `p:txStyles`（`titleStyle` / `bodyStyle` / `otherStyle`）
6. `presentation.xml` の `a:defaultTextStyle`

### プレースホルダーのマッチング

レイアウト/マスター上のプレースホルダーとの対応は `p:ph` 要素の `type` と `idx` 属性で行う。

1. `type` + `idx` の完全一致を試みる
2. 完全一致が見つからない場合、`type` のみでマッチする

`type` 未指定のプレースホルダーは `body` として扱う。

### テーマフォントの解決

レイアウト/マスターの `defRPr` でフォント名がテーマ参照（`+mj-lt`, `+mn-lt`, `+mj-ea`, `+mn-ea` 等）の場合、`theme1.xml` の `a:fontScheme` から実フォント名に解決する。`a:ea`（東アジア）フォントを優先し、なければ `a:latin` を使用する。

### キャッシュ

レイアウトとマスターのデータは `File` 構造体にキャッシュされ、同じファイル内の複数スライドで再利用される。一般的なPPTXファイルは1つのマスターと数個〜十数個のレイアウトで構成されるため、メモリ影響は軽微。

### 制限事項

- 塗りつぶし・枠線はプレースホルダー間で継承しない（直接指定のみ）
- 非プレースホルダー図形には継承を適用しない

## スタイル重複排除

すべてのフォント情報はスタイル定義行（`style` 行）に抽出し、段落やリッチテキストランでは `s` フィールドで参照IDを指定する。同じフォント情報はスライド横断で同一IDを共有する。

```jsonl
{"slide":1,"title":"基本設計書","shapes":2}
{"style":1,"name":"Arial","size":14,"color":"#3F3F3F"}
{"shape":1,"type":"rect","z":0,"paragraphs":[{"text":"本文テキスト","s":1},{"text":"別の本文","s":1}]}
{"style":2,"name":"Arial","size":18,"bold":true}
{"shape":2,"type":"rect","z":1,"paragraphs":[{"text":"見出し","s":2}]}
{"slide":2,"shapes":1}
{"shape":1,"type":"rect","z":0,"paragraphs":[{"text":"続きの本文","s":1}]}
```

**`style` 行:**

スタイル定義は独立したJSONL行として、そのスタイルを初めて使用する図形の直前に出力される。

| フィールド | 型 | 説明 |
|-----------|-----|------|
| `style` | number | スタイルID（出力全体で一意の連番） |
| その他 | | `font` オブジェクトと同じフィールド（`name`, `size`, `bold`, `italic`, `strikethrough`, `underline`, `color`, `highlight`, `baseline`, `cap`） |

**段落/リッチテキストランの `s` フィールド:**

| フィールド | 型 | 説明 |
|-----------|-----|------|
| `s` | number | `style` のスタイルIDへの参照。`font` フィールドの代わりに使用する |

- フォント情報を持つ段落/ランはすべて `s` による参照に置き換えられる
- スタイルIDはスライド横断で共有される（同じフォントは全スライドで同じIDを参照する）
- 既出のスタイルには `style` 行を再出力しない
