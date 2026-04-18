# pptx-scope 仕様

PowerPoint ファイル（.pptx）の内容をCLIから出力するGoツール。
Claude CodeがPowerPoint資料（プレゼン、フローチャート、仕様書など）を読み取る用途を主眼とする。

## コマンド構成と利用フロー

サブコマンドは用途別に以下の通り。各コマンドの詳細は個別の節を参照。

- **探索**: `info`（ファイル概要・スライド一覧）
- **取得**: `slides`（スライド内容） / `image`（画像抽出） / `search`（検索）
- **管理**: `cleanup`（一時ファイル削除） / `version`（バージョン表示）

典型フローは `info` → `slides` で内容を把握する。特定のキーワードを探す場合は `search` が効率的。

- `slides` 出力に `image_id` があれば `image` で実ファイルを抽出する。
- `search` は該当スライドを特定するだけで、詳細取得は `slides --slide` で行う。

### `pptx-scope info <file>`

**役割:** ファイルレベルの概要を把握する。スライド一覧から対象スライドを特定し、以降の `slides` / `search` に渡す `--slide` を決定する。

メタ情報行に続いてスライド情報を1行ずつJSONL形式で出力する。

**出力例:**

```jsonl
{"file":"基本設計書.pptx","slide_size":{"width":720,"height":540}}
{"slide":1,"title":"基本設計書","has_notes":true}
{"slide":2,"title":"目次"}
{"slide":3,"title":"システム構成","has_notes":true}
{"slide":4}
{"slide":5,"title":"フロー図","hidden":true}
```

**メタ情報行のフィールド:**

| フィールド | 型 | 説明 |
|-----------|-----|------|
| `file` | string | ファイル名 |
| `slide_size` | object | スライドサイズ（`width`, `height`。pt単位。標準4:3 = 720x540, 16:9 = 960x540） |

**スライド行の各フィールド:**

| フィールド | 型 | 説明 |
|-----------|-----|------|
| `slide` | number | スライド番号（1始まり） |
| `title` | string | タイトルプレースホルダーのテキスト。存在しない場合は省略 |
| `has_notes` | bool | ノートにテキストがある場合のみ `true` を出力 |
| `hidden` | bool | 非表示スライドの場合のみ `true` を出力 |

### `pptx-scope slides [options] <file>`

**役割:** スライドの内容を取得する。スライドヘッダ行に続いて、図形を1つずつ個別のJSONL行として出力する。書式情報（フォント・色・枠線）は常に出力する。

**オプション:**

| オプション | 説明 | デフォルト |
|-----------|------|-----------|
| `--slide <number,...>` | 対象スライド番号（1始まり、複数指定可: `--slide 1,3`） | 全スライド |
| `--notes` | ノートも出力する | OFF |

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

#### スライドヘッダ行

| フィールド | 型 | 説明 |
|-----------|-----|------|
| `slide` | number | スライド番号 |
| `title` | string | スライドタイトル（存在する場合のみ） |
| `shapes` | number | スライド内の図形数 |
| `has_notes` | bool | ノートにテキストがある場合のみ `true` |
| `hidden` | bool | 非表示スライドの場合のみ `true` |

#### 図形の出力順

1. プレースホルダー（タイトル → サブタイトル → ボディ → その他）
2. 非プレースホルダーの図形（スライド上の出現順）

#### スキップ条件

- テキストが空かつ書式情報もない図形
- 非表示図形

#### ノート（`--notes` 指定時）

スライドの図形行の後に、ノートを独立したJSONL行として出力する。`paragraphs` 配列の要素と同じ構造。

```jsonl
{"slide":1,"title":"基本設計書","shapes":2,"has_notes":true}
{"shape":1,"type":"rect","placeholder":"ctrTitle",...}
{"shape":2,"type":"rect","placeholder":"subTitle",...}
{"notes":[{"text":"このスライドでは基本設計の概要を説明する。"},{"text":"スコープの定義","bullet":"•"},{"text":"前提条件の確認","bullet":"•"}]}
```

- ノートが存在しない、またはテキストが空のスライドではノート行を出力しない

### `pptx-scope image <file> <image_id>`

**役割:** `slides` が返した `image_id`（ZIP内のメディアパス）を受け取り、画像を一時ファイルに抽出する。抽出後のファイルは Read ツール等で参照でき、不要になったら削除する。

**引数:**

| 引数 | 説明 |
|------|------|
| `<file>` | 対象のPowerPointファイル |
| `<image_id>` | `slides` 出力の `image_id` フィールド値（例: `ppt/media/image1.png`） |

**出力:**

- `{"file":"..."}` 形式のJSON1行（抽出先のパス）
- 抽出先ファイルはプレフィックス `pptx-scope-tmp-` で `os.TempDir()` 直下に生成される。拡張子は `image_id` のものを維持する
- `--stdout` 指定時は抽出先パスをそのまま出力する

### `pptx-scope search [options] <file>`

**役割:** プレゼンテーション内のテキストを検索する。全スライドを `slides` で取得してフィルタするより効率的。マッチしたスライドを特定し、`slides --slide` で詳細を取得する運用を想定。

**オプション:**

| オプション | 説明 | デフォルト |
|-----------|------|-----------|
| `--text <text>` | 検索文字列（部分一致、大文字小文字無視） | 必須 |
| `--slide <number,...>` | 対象スライド番号（複数指定可） | 全スライド |
| `--notes` | ノートも検索対象にする | OFF |

**検索の挙動:**

- 各図形・テーブルセル・コネクタラベルのテキストに対して大文字・小文字を区別しない部分一致検索を行う
- `--notes` 指定時はノートのテキストも検索対象にする
- 全角・半角は区別する。正規表現には対応しない
- マッチがないスライドは出力しない
- 結果なしでも正常終了（終了コード 0）する

**出力形式:**

マッチしたスライドのヘッダ行のみを出力する（`info` のスライド行と同じ形式）。

```jsonl
{"slide":2,"title":"システム構成"}
{"slide":4,"title":"処理フロー"}
```

### `pptx-scope cleanup <file> [file...]`

**役割:** `pptx-scope` が生成した一時ファイル（プレフィックス `pptx-scope-tmp-`）を削除する。`info` / `slides` / `search` / `image` の出力ファイルを使い終わった後の後始末に使う。

**安全確認:**

- ファイル名のプレフィックスが `pptx-scope-tmp-` であること
- 親ディレクトリが `os.TempDir()` 直下であること

上記2条件を満たさないパスを指定するとエラーで停止する（誤って他のファイルを削除しないため）。

**出力:** stdout に削除件数のJSON1行。

```json
{"deleted":2}
```

- 既に存在しないファイルはスキップする（カウントしない）

### `pptx-scope version`

**役割:** バージョン情報を表示する。

**出力:** バージョン文字列を1行出力する（例: `v0.0.9`）。他コマンドと異なり一時ファイルへの書き出しは行わない。

- バージョン文字列はビルド時に `-ldflags="-X github.com/nobmurakita/claude-pptx-scope/internal/cmd.Version=<tag>"` で埋め込む
- ldflags 未指定時のデフォルトは `latest`

## 図形の出力構造

`slides` コマンドが出力する図形行・段落・スタイル定義の構造を定義する。

### 図形フィールド

| フィールド | 型 | 説明 |
|-----------|-----|------|
| `shape` | number | スライド内の図形ID（出現順の1始まり連番） |
| `type` | string | 図形種別。後述の「図形種別」を参照 |
| `name` | string | 図形名。プレースホルダーの場合は省略 |
| `placeholder` | string | プレースホルダー種別（`title`, `ctrTitle`, `subTitle`, `body`, `dt`, `ftr`, `sldNum` 等）。プレースホルダーでない場合は省略 |
| `pos` | object | 位置とサイズ（`x`, `y`, `w`, `h`。pt単位）。プレースホルダーでスライド上に未指定の場合、レイアウト・マスターから継承した値を出力する |
| `z` | number | Z-order（0始まり、大きいほど前面） |
| `rotation` | number | 回転角度（度単位、時計回り）。0の場合は省略 |
| `flip` | string | 反転。`"h"`, `"v"`, `"hv"`。なければ省略 |
| `fill` | string | 塗りつぶし色（`#RRGGBB` 形式）。塗りつぶしなしの場合は省略。グラデーションは最初のストップカラーを代表色として使用 |
| `line` | object | 枠線情報。枠線がない場合は省略 |
| `callout_pointer` | object | 吹き出し図形のポインタ位置。後述 |
| `alignment` | object | テキストの垂直配置（`vertical` フィールド）。デフォルト（上揃え）は省略 |
| `text_margin` | object | テキストボディの内部マージン（`left`, `right`, `top`, `bottom`。pt単位）。OOXMLデフォルト値と同じ場合は省略 |
| `link` | object | 図形全体に設定されたハイパーリンク。後述 |
| `paragraphs` | array | 段落の配列。テキストがない場合は省略 |
| `table` | object | テーブルデータ。テーブルの場合のみ出力し `paragraphs` の代わりに使用 |

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
| `width` | number | 線幅（pt単位） |

### 段落フィールド

`paragraphs` 配列の各要素:

| フィールド | 型 | 説明 |
|-----------|-----|------|
| `text` | string | 段落のプレーンテキスト |
| `bullet` | string | 箇条書き記号。箇条書きでない段落は省略 |
| `level` | number | インデントレベル（0始まり）。0の場合は省略 |
| `margin_left` | number | 段落の左マージン（pt単位）。0または未指定の場合は省略 |
| `indent` | number | 段落の1行目インデント（pt単位、負値でぶら下がりインデント）。0または未指定の場合は省略 |
| `line_spacing` | string | 行間。`"150%"` のようなパーセント、または `"12pt"` のようなポイント指定。デフォルト（100%）は省略 |
| `space_before` | string | 段落前のスペース。`line_spacing` と同じ書式。0pt は省略 |
| `space_after` | string | 段落後のスペース。`line_spacing` と同じ書式。0pt は省略 |
| `s` | number | スタイルID。`style` 行で定義されたフォント情報への参照 |
| `alignment` | object | 水平配置情報。デフォルト（左揃え）は省略 |
| `link` | object | ハイパーリンク（段落内の全テキストが同一リンクの場合） |
| `rich_text` | array | リッチテキストラン（段落内に書式やリンクの異なるランが存在する場合のみ） |

- テキストが空の段落は出力しない
- 段落内の改行は `text` フィールドで `\n` として出力される。`rich_text` 内では `{"text":"\n"}` というランとして出力される

**箇条書き (`bullet`):**

| 種類 | 出力 |
|------|------|
| 文字箇条書き | 指定された文字（例: `"•"`, `"–"`, `">"`） |
| 自動番号 | 番号を計算して文字列化（例: `"1."`, `"2."`, `"a."`, `"i)"`） |
| 箇条書きなし | `bullet` フィールドを省略 |

自動番号の形式:

| 種類 | 出力例 |
|------|-------|
| `arabicPeriod` | `1.`, `2.`, `3.` |
| `arabicParenR` | `1)`, `2)`, `3)` |
| `alphaLcPeriod` | `a.`, `b.`, `c.` |
| `alphaUcPeriod` | `A.`, `B.`, `C.` |
| `romanLcPeriod` | `i.`, `ii.`, `iii.` |
| `romanUcPeriod` | `I.`, `II.`, `III.` |

- 番号は同一図形内で同じ `level` の連続する自動番号段落をカウントして算出する
- `startAt` 指定がある場合はその値から開始する（デフォルトは1）
- レベルが変わるか箇条書きなしの段落で番号はリセットされる
- 上記以外の種類は `arabicPeriod` と同じ形式で出力する

### スタイル定義（`style` 行）

すべてのフォント情報はスタイル定義行に抽出し、段落やリッチテキストランでは `s` フィールドで参照する。同じフォント情報はスライド横断で同一IDを共有する。

```jsonl
{"style":1,"name":"メイリオ","size":11,"bold":true,"italic":true,"strikethrough":true,"underline":"single","color":"#FF0000","highlight":"#FFFF00","baseline":"super","cap":"all"}
```

- スタイル定義は独立したJSONL行として、そのスタイルを初めて使用する図形の直前に出力される
- 既出のスタイルには `style` 行を再出力しない

| フィールド | 型 | 説明 |
|-----------|-----|------|
| `style` | number | スタイルID（出力全体で一意の連番） |
| `name` | string | フォント名 |
| `size` | number | フォントサイズ（pt単位） |
| `bold` | bool | 太字の場合のみ `true` |
| `italic` | bool | 斜体の場合のみ `true` |
| `strikethrough` | bool | 取り消し線ありの場合のみ `true` |
| `underline` | string | 下線スタイル。`single`, `double`, `singleAccounting`, `doubleAccounting`。下線なしは省略 |
| `color` | string | フォント色（`#RRGGBB` 形式） |
| `highlight` | string | 文字の背景色（`#RRGGBB` 形式）。未指定は省略 |
| `baseline` | string | 上付き/下付き。`"super"`（上付き）/`"sub"`（下付き）。標準位置は省略 |
| `cap` | string | 英字の大文字化。`"all"`（すべて大文字）/`"small"`（スモールキャップ）。`none`/未指定は省略 |

- デフォルト値のフィールドは省略する
- テーマフォント参照（`+mj-lt` 等）は実フォント名に解決される

### リッチテキスト

段落内のテキストの一部が異なる書式またはリンクを持つ場合、`text` にはプレーンテキストを格納し、`rich_text` に書式付きランの配列を格納する。

```jsonl
{"style":1,"bold":true,"color":"#FF0000"}
{"shape":1,...,"paragraphs":[{"text":"重要: この機能は必須です","rich_text":[{"text":"重要: ","s":1},{"text":"この機能は必須です"}]}]}
```

- `text` は常にプレーンテキスト（検索・概要把握用）
- `rich_text` は書式またはリンクの異なるランが存在する場合のみ出力
- 各ランの `s` はスタイルIDへの参照、`link` はハイパーリンクがある場合のみ出力

### ハイパーリンク

ハイパーリンクは図形レベルとテキストランレベルの2箇所に設定できる。

```json
{"shape":1,"type":"rect","pos":{"x":0,"y":0,"w":393.7,"h":39.37},"z":0,"paragraphs":[{"text":"詳細はこちら","link":{"url":"https://example.com"}}]}
```

```json
{"text":"スライド3を参照","link":{"slide":3}}
```

**`link` オブジェクト:**

| フィールド | 型 | 説明 |
|-----------|-----|------|
| `url` | string | 外部URL（`http://`, `https://`, `mailto:` 等）。外部リンクの場合のみ |
| `slide` | number | リンク先のスライド番号。スライド内リンクの場合のみ |

- 図形全体がクリック対象のリンクは図形の `link` に格納
- 段落内の全テキストが同一リンクなら段落の `link` に格納。異なるリンクが混在する場合は `rich_text` の各ランの `link` に格納

### 吹き出しポインタ

吹き出し図形（`wedgeRectCallout`, `wedgeRoundRectCallout`, `wedgeEllipseCallout`, `cloudCallout`, `borderCallout1`, `borderCallout2`, `borderCallout3` 等）の場合、ポインタの指す位置をスライド上の絶対座標に変換して出力する。

```json
{"shape":3,"type":"wedgeRoundRectCallout","pos":{"x":236.22,"y":78.74,"w":157.48,"h":62.99},"z":2,"callout_pointer":{"x":118.11,"y":196.85},"paragraphs":[{"text":"ここに注目"}]}
```

| フィールド | 型 | 説明 |
|-----------|-----|------|
| `x` | number | ポインタ先端のX座標（pt。スライド上の絶対座標） |
| `y` | number | ポインタ先端のY座標（pt。スライド上の絶対座標） |

吹き出し図形でない場合は `callout_pointer` フィールドを省略する。

### コネクタフィールド

コネクタは `type` が `"connector"` となり、以下の追加フィールドを持つ。

| フィールド | 型 | 説明 |
|-----------|-----|------|
| `from` | number | 接続元の図形ID。接続情報がない場合は省略 |
| `to` | number | 接続先の図形ID。接続情報がない場合は省略 |
| `from_idx` | number | 接続元の接続ポイントインデックス。矩形の場合 0=上, 1=左, 2=下, 3=右。接続情報がない場合は省略 |
| `to_idx` | number | 接続先の接続ポイントインデックス。`from_idx` と同様。接続情報がない場合は省略 |
| `connector_type` | string | コネクタ形状。`straightConnector1`（直線）、`bentConnector2`（L字・1回屈曲）、`bentConnector3`（コの字・2回屈曲）、`curvedConnector3`（曲線）等 |
| `arrow` | string | 矢印ヘッドの位置。`"start"`, `"end"`, `"both"`。矢印なしの場合は省略 |
| `adj` | object | 屈曲・カーブの調整値（1/100000単位）。bent/curvedコネクタで屈曲位置を制御。デフォルト値の場合は省略 |
| `start` | object | コネクタの始点座標 `{x, y}`（pt） |
| `end` | object | コネクタの終点座標 `{x, y}`（pt） |
| `label` | string | コネクタ上のテキスト。テキストがない場合は省略 |

- `from`/`to` はパース時にPowerPoint図形IDから連番IDへのマッピングを行い、出力時は連番IDで参照する
- 接続先がスライド内に見つからない場合は `from`/`to` を省略する

### 画像フィールド

画像は `type` が `"picture"` となり、以下の追加フィールドを持つ。

```json
{"shape":5,"type":"picture","name":"図 1","pos":{"x":78.74,"y":78.74,"w":393.7,"h":236.22},"z":4,"alt_text":"システム構成図","image_id":"ppt/media/image1.png"}
```

| フィールド | 型 | 説明 |
|-----------|-----|------|
| `alt_text` | string | 代替テキスト。設定されていない場合は省略 |
| `image_id` | string | ZIP内のメディアパス。`image` サブコマンドに渡して画像を取得する |

### グループフィールド

グループは `type` が `"group"` となり、以下の追加フィールドを持つ。

| フィールド | 型 | 説明 |
|-----------|-----|------|
| `children` | array | 子要素の配列。各子要素は図形フィールドと同じ構造 |

- グループの `pos` はグループ全体の位置とサイズ
- `children` 内の子要素の `pos` はスライド上の絶対座標に変換済み
- 子要素の `z` もスライド全体で一意の連番（グループ内でリセットされない）
- ネストしたグループも同じ構造で再帰的に表現する

### テーブルフィールド

テーブルは `type` が `"table"` となり、`paragraphs` の代わりに `table` フィールドを持つ。

```json
{"shape":4,"type":"table","name":"表 1","pos":{"x":36,"y":126,"w":648,"h":236.22},"z":3,"table":{"cols":3,"rows":[[{"text":"項目"},{"text":"説明"},{"text":"備考"}],[{"text":"機能A","paragraphs":[{"text":"機能A","s":2}]},null,{"text":"必須"}]]}}
```

| フィールド | 型 | 説明 |
|-----------|-----|------|
| `cols` | number | 列数 |
| `rows` | array | 行の配列。各行は `cols` と同じ長さの配列。セルの値はオブジェクトまたは `null` |

**セルオブジェクト:**

| フィールド | 型 | 説明 |
|-----------|-----|------|
| `text` | string | セル内の全テキストを結合したプレーンテキスト |
| `fill` | string | セルの背景色（`#RRGGBB` 形式）。グラデーションは最初のストップカラーを代表色として使用。未指定は省略 |
| `border_left` | object | 左罫線情報。`line` オブジェクトと同じ構造（`color`, `style`, `width`）。未指定は省略 |
| `border_right` | object | 右罫線情報。`border_left` と同じ構造 |
| `border_top` | object | 上罫線情報。`border_left` と同じ構造 |
| `border_bottom` | object | 下罫線情報。`border_left` と同じ構造 |
| `paragraphs` | array | 段落情報（フォント・箇条書き等の書式情報がある場合のみ） |

- `text` は複数段落をスペースで結合したプレーンテキスト
- `paragraphs` はフォント・リッチテキスト・箇条書き等の書式情報がある場合のみ出力される
- 結合セルの被結合セルは `null` となる（結合元セルにテキストが格納される）

### 図形種別

`type` フィールドの値は以下のいずれか:

- **コネクタ**: 常に `"connector"`
- **グループ**: 常に `"group"`
- **画像**: 常に `"picture"`
- **テーブル**: 常に `"table"`
- **シェイプ**: 形状名（`rect`, `roundRect`, `ellipse`, `triangle`, `diamond`, `flowChartProcess`, `flowChartDecision`, `flowChartTerminator` 等）
- **カスタムシェイプ**: `"customShape"`（形状が定義済みの基本図形ではなくカスタムジオメトリの場合）

シェイプの形状名の例:

| 形状名 | PowerPoint上の名称 |
|--------|-------------------|
| `rect` | 四角形 |
| `roundRect` | 角丸四角形 |
| `ellipse` | 楕円 |
| `triangle` | 三角形 |
| `diamond` | ひし形 |
| `flowChartProcess` | 処理 |
| `flowChartDecision` | 判断 |
| `flowChartTerminator` | 端子 |
| `flowChartPredefinedProcess` | 定義済み処理 |
| `flowChartDocument` | 書類 |
| `flowChartConnector` | 結合子 |

## エラーハンドリング

- ファイルが存在しない / 読み取れない → エラーメッセージを stderr に出力
- .pptx 以外のファイル → 「.pptx 形式のみ対応」のエラーメッセージ
- 存在しないスライド番号 → 利用可能なスライド番号の範囲をエラーメッセージに含める
- パスワード保護されたファイル → 非対応としてエラーメッセージを出力
- 破損したファイル（不正なzip構造等） → エラーメッセージを出力
- `search` で結果が0件の場合:
  - 本体は0行のJSONL（空）で正常終了する
  - stdout には通常どおり `{"file":"...","lines":0}` を出力（`--stdout` 時は何も出力しない）
- 終了コード: 0=成功（検索結果なしも含む）、1=エラー
- エラーメッセージは stderr に `pptx-scope: <メッセージ>` の形式で出力する。stdout には常にJSONのみを出力する

## 出力に関する共通仕様

### 出力先

- `info` / `slides` / `search` の出力は一時ファイル（プレフィックス `pptx-scope-tmp-`）に書き出され、stdout にはファイルパスと行数のJSON（`{"file":"...","lines":N}`）のみを出力する。`--stdout` フラグで標準出力に直接書き出すことも可能（デバッグ用）
- 一時ファイルは `cleanup` サブコマンドで削除する
- `image` は一時ファイルに画像を抽出し、stdout にパスのJSON（`{"file":"..."}`）を出力する。`--stdout` 指定時はパスをそのまま出力する

### フォーマット

- 対応形式は .pptx のみ（.ppt は非対応）
- 出力形式はJSONL（1行1JSONオブジェクト）
- 出力は常にUTF-8エンコーディング
- テキスト内の制御文字（改行、タブ等）はJSON仕様に従いエスケープする（`\n`, `\t` 等）
- 数値はすべてpt単位で出力する（位置・サイズ・フォントサイズ・線幅。回転角度のみ度単位）
- 書式フィールドはデフォルト値と異なる場合のみ出力する
- 色は可能な限り `#RRGGBB` 形式で出力する。テーマカラーはテーマ定義からRGBに変換し、tint 値がある場合は明度を調整して適用する
- プレースホルダーのフォント・位置・箇条書きはスライドマスター・レイアウトから継承した値で補完して出力する
- フォント情報は `style` 行に抽出し、段落/リッチテキストランでは `s` による参照に置き換える（スライド横断で共有）

## 対応しない機能

以下の機能は意図的に対応しない。理由とともに記録する。

| 機能 | 理由 |
|------|------|
| .ppt 形式の読み取り | OOXML（.pptx）のみ対応。旧形式のファイルは事前に .pptx へ変換して使用する |
| パスワード保護されたファイル | 非対応 |
| アニメーション・トランジション | 静的なコンテンツ読み取りが目的 |
| 埋め込みチャート | 独自のXML体系で複雑。別途対応を検討 |
| SmartArt | 独自のXML体系で複雑。PowerPoint保存時に描画キャッシュとしてグループに展開されるため、通常はグループとして読み取り可能 |
| 埋め込み動画・音声 | メディアファイルの内容はCLIで扱いにくい。画像のみ `image` サブコマンドで対応する |
| 塗りつぶしのパターン | パターン塗りつぶしは非対応。グラデーションは最初のストップカラーを代表色として出力する |
| 条件付き書式の評価 | 静的な書式のみ適用する |
| 正規表現による検索 | `--text` は部分一致のみ。正規表現はシェル側の `grep` と組み合わせて実現する想定 |
| PowerPoint書き込み・編集 | 読み取り専用ツールとして設計 |
| 複数ファイルの同時処理 | 1コマンド1ファイルの原則。複数ファイルを処理する場合はファイルごとにコマンドを実行する |
| メモリ使用量の上限 | 設けない |
