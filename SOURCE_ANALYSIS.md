# SplitDocx ソース解析メモ

## 概要

このリポジトリは、Word の `.docx` ファイルをセクション単位に分割し、各セクションごとの `.docx` と `.pdf` を出力する PowerShell ツール群です。

主な入口は次の2つです。

- `bulk_split_docx.ps1`
  - フォルダ内の複数 `.docx` をまとめて処理する一括実行スクリプト。
- `split_docx_final2.ps1`
  - 1つの `.docx` を分割する実処理スクリプト。
  - `bulk_split_docx.ps1` の既定 worker として呼び出されます。

## 実行イメージ

単一ファイルを処理する場合:

```powershell
.\split_docx_final2.ps1 -InputFile .\input.docx -OutputDir .\split_output
```

フォルダ内の `.docx` をまとめて処理する場合:

```powershell
.\bulk_split_docx.ps1 -TargetDir .\docs -OutputRoot .\bulk_output
```

サブフォルダも含める場合:

```powershell
.\bulk_split_docx.ps1 -TargetDir .\docs -OutputRoot .\bulk_output -Recurse
```

画像削除版を使いたい場合:

```powershell
.\bulk_split_docx.ps1 -TargetDir .\docs -OutputRoot .\bulk_output -WorkerScript .\split_docx_noPic.ps1
```

## ファイル構成

| ファイル | 役割 |
| --- | --- |
| `bulk_split_docx.ps1` | 一括処理用。対象フォルダ内の `.docx` を列挙し、worker スクリプトを順番に呼び出します。 |
| `split_docx_final.ps1` | 基本版。DOCX 分割、ページ番号調整、PDF 出力を行います。 |
| `split_docx_final2.ps1` | 現在の既定版。`final` に PDF 空白ページ除去処理を追加したものです。 |
| `split_docx_noPic.ps1` | 画像削除版。画像を含む段落や画像 run を削除する処理が追加されています。 |

## `bulk_split_docx.ps1` の処理

`bulk_split_docx.ps1` は、指定フォルダから `.docx` を探し、1件ずつ worker スクリプトに渡します。

主なパラメータ:

| パラメータ | 既定値 | 内容 |
| --- | --- | --- |
| `TargetDir` | `.` | 入力 `.docx` を探すフォルダ。 |
| `WorkerScript` | `.\split_docx_final2.ps1` | 各ファイルを処理する実処理スクリプト。 |
| `OutputRoot` | 空文字 | 出力先ルート。未指定時は `TargetDir\bulk_output`。 |
| `Recurse` | なし | 指定時、サブフォルダも検索します。 |

出力先は、入力ファイル名ごとに以下のように分かれます。

```text
bulk_output\
  input_file_1\
    001.docx
    001.pdf
    002.docx
    002.pdf
  input_file_2\
    ...
```

## `split_docx_final2.ps1` の処理フロー

`split_docx_final2.ps1` は、DOCX を ZIP として展開し、WordprocessingML の XML を編集してセクションごとの文書を作ります。

処理の大きな流れ:

1. 入力 `.docx` の存在確認
2. 出力フォルダ作成
3. 入力ファイル名から開始番号を推定
   - ファイル名に `001-999` のような形式がある場合、先頭の3桁を開始番号として使います。
4. 一時フォルダ作成
5. `.docx` を ZIP 展開
6. `word/document.xml` を読み込み
7. セクション開始段落を検出
8. 元文書内で各セクションの開始ページを測定
9. セクションごとの `document.xml` を生成
10. フッターのページ番号表記を整形
11. 開始ページ番号を `sectPr/pgNumType@start` に設定
12. セクション DOCX として再圧縮
13. Microsoft Word COM で PDF 変換
14. Python + `pypdf` で空白 PDF ページを削除
15. 一時フォルダ削除

## セクション検出ロジック

セクション開始位置は `Get-SectionParas` で検出されます。

検出モードは3種類あります。

| モード | 内容 |
| --- | --- |
| Mode1 `plainText` | 段落テキストが `(001)` のような形式で始まるものを検出します。 |
| Mode2 `blueItalic` | 青系の色かつ斜体の数字 run を検出します。番号は開始番号から連番で付与されます。 |
| Mode3 `numberedParagraph` | Word の番号付き段落を検出します。`List Paragraph` スタイルや番号レベルを参考にします。 |

優先順位は Mode1、Mode2、Mode3 の順です。

## ページ番号測定

既定では `FooterSource` が `OriginalPage` です。

この場合、スクリプトは元 DOCX の各セクション開始位置に一時ブックマークを埋め込み、Microsoft Word COM で実際のページ番号を読み取ります。

測定結果は出力フォルダに次の CSV として保存されます。

```text
section_page_map.csv
```

CSV 形式:

```csv
SectionNumber,OriginalPage
1,3
2,7
```

`-FirstPageNumber` を指定した場合は、測定ページ番号に補正が入ります。

## フッター処理

フッターは `Fix-FooterToDynamicPPage` で調整されます。

主な目的:

- 既存のページ番号ラベルを整理
- `PAGE` フィールドの前に `P` を付ける
- 複雑な Word フィールドにも対応
- 出力セクションごとに開始ページ番号を設定

開始ページ番号は `Set-SectionPageStart` で `sectPr` 内の `pgNumType` に書き込まれます。

例:

```xml
<w:pgNumType w:start="12" />
```

## PDF 変換

PDF 変換は `Convert-ToPdf` で行われます。

内部では Microsoft Word COM を使用します。

```powershell
New-Object -ComObject Word.Application
$doc.ExportAsFixedFormat($PdfPath, 17)
```

そのため、実行環境には Microsoft Word が必要です。

## PDF 空白ページ削除

`split_docx_final2.ps1` は、PDF 生成後に Python スクリプトを一時生成し、`pypdf` で空白ページを削除します。

削除対象:

- テキストが空のページ
- `P1`、`P12` のようにページ番号だけに見えるページ

`pypdf` が見つからない場合、スクリプトは `pip install pypdf` を試みます。

ネットワーク不可、権限制限、Python 未インストールの場合は、この空白ページ削除はスキップまたは失敗します。

## `split_docx_noPic.ps1` の追加処理

`split_docx_noPic.ps1` は、`split_docx_final2.ps1` 系の処理に画像削除機能を追加した派生版です。

追加されている主な関数:

| 関数 | 内容 |
| --- | --- |
| `Test-HasImage` | 段落内に画像があるか判定します。 |
| `Test-IsPageBreakOnly` | 改ページだけの段落か判定します。 |
| `Test-IsEmptyPara` | 空段落か判定します。 |
| `Remove-ImageAndPageBreakRuns` | 画像 run や改ページ run を削除します。 |
| `Remove-ImageParagraphs` | 画像のみの段落や関連する空段落を削除します。 |

注意点として、`bulk_split_docx.ps1` の既定 worker は `split_docx_final2.ps1` です。

画像削除版を使うには、明示的に次のように指定する必要があります。

```powershell
.\bulk_split_docx.ps1 -WorkerScript .\split_docx_noPic.ps1
```

## 外部依存

| 依存 | 用途 |
| --- | --- |
| Microsoft Word | ページ番号測定、PDF 変換 |
| Python | PDF 空白ページ削除 |
| `pypdf` | PDF の読み書き、空白ページ削除 |
| .NET `System.IO.Compression` | DOCX の ZIP 展開・再圧縮 |

## 注意点

### コメントの文字化け

多くのコメントが文字化けしています。

コード本体は PowerShell と XML 操作として読めますが、保守性は落ちています。

可能であれば、コメントを UTF-8 で整理し直すと今後の修正がしやすくなります。

### Word COM 依存

このツールは Microsoft Word COM に強く依存しています。

Word が入っていない環境、Office の自動化が制限されている環境、サーバー実行環境では正常に動かない可能性があります。

### `pip install` の自動実行

`split_docx_final2.ps1` は `pypdf` が無い場合に `pip install pypdf` を試みます。

業務端末やオフライン環境では失敗する可能性があります。

安定運用するなら、事前に Python と `pypdf` をセットアップしておく方が安全です。

### 出力ファイルの上書き

出力先に同名の `.docx` や `.pdf` がある場合、削除して再作成されます。

既存成果物を残したい場合は、出力先フォルダを分けて実行してください。

### 画像削除版は既定ではない

画像を削除したい運用の場合、`split_docx_noPic.ps1` を worker として指定する必要があります。

## 改善候補

1. コメント文字化けの修正
2. `README.md` の追加
3. `split_docx_final.ps1`、`split_docx_final2.ps1`、`split_docx_noPic.ps1` の差分整理
4. 既定 worker を設定ファイル化
5. `pypdf` の自動インストールをオプション化
6. Word COM が利用できない場合のエラーメッセージ改善
7. ログファイル出力の追加
8. サンプル入力と期待出力を使った回帰テストの追加

