# OfficeFileToPdf

[中文](README.md) | [日本語](README.ja.md) | [English](README.en.md)

## 概要

これは Windows 専用の一括変換ツールです。

指定したフォルダを再帰的に走査し、対応している Office 文書と PDF を `res/` フォルダへ出力します。できるだけ元のフォルダ構造を維持します。

出力後のファイルはすべて PDF になります。

また、`res/` の直下に `source_to_pdf_map.json` を生成し、元ファイルと出力 PDF の対応関係を記録します。

## 主な機能

- フォルダの再帰スキャン
- 元のフォルダ構造を保ったまま `res/` へ出力
- Excel、Word、PowerPoint を PDF に変換
- 既存の PDF も再処理して `res/` に出力
- PDF の余白を自動トリミング
- 元ファイルと出力 PDF の対応 JSON を生成
- `ToPdf.bat` のダブルクリック実行に対応
- フォルダを `ToPdf.bat` にドラッグ＆ドロップして実行可能

## 実行条件

- OS は Windows である必要があります
- Microsoft Office がインストールされている必要があります
- Office 文書の変換は Microsoft Office COM のみを使用します
- LibreOffice は使用しません
- 初回実行時に Python がない場合、`ToPdf.bat` が Python の自動インストールを試みます
- 初回実行時に `.venv` を作成し、必要な Python パッケージをインストールします

## 対応入力形式

PDF へ変換対象となる形式：

- Excel：`.xls` `.xlsx` `.xlsm` `.xlsb` `.csv` `.ods`
- Word：`.doc` `.docx` `.docm` `.rtf` `.odt`
- PowerPoint：`.ppt` `.pptx` `.pptm` `.odp`
- PDF：`.pdf`

変換しない形式：

- 画像ファイル：`.png` `.jpg` `.jpeg` `.bmp` `.gif` `.svg` `.tif` `.tiff` `.webp` `.heic` `.ico`
- テキスト系ファイル：`.txt` `.md` `.json` `.yaml` `.yml` `.xml` `.toml` `.ini` `.cfg` `.log` `.rst`
- スクリプト・実行ファイル：`.py` `.bat` `.cmd` `.ps1` `.exe` `.dll` `.sh`
- 拡張子のないファイル
- 現時点で未対応のその他の形式

## 出力ルール

- 出力先は通常、対象フォルダ直下の `res/`
- できるだけ元のフォルダ構造を維持
- 出力ファイルはすべて `.pdf`
- 同名競合がある場合は自動的にリネームして上書きを防止

例：

元フォルダ：

```text
input/
  report.xlsx
  deck.pptx
  docs/
    plan.docx
    old.pdf
    note.txt
```

出力フォルダ：

```text
input/
  res/
    report.pdf
    deck.pdf
    docs/
      plan.pdf
      old.pdf
    source_to_pdf_map.json
```

説明：

- `note.txt` はスキップされます
- `old.pdf` は `res/docs/old.pdf` として再出力されます
- `source_to_pdf_map.json` に元ファイルと出力ファイルの対応が保存されます

## 実行方法

### 方法 1：ダブルクリック

`ToPdf.bat` をダブルクリックしてください。

デフォルトでは、`ToPdf.bat` が置かれているフォルダを処理し、同じ階層に `res/` を生成します。

### 方法 2：フォルダをドラッグ＆ドロップ

処理したいフォルダを `ToPdf.bat` にドラッグ＆ドロップしてください。

そのフォルダが処理対象になります。

### 方法 3：コマンドライン実行

Windows の `cmd` で実行：

```bat
cd /d D:\OfficeFileToPdf
ToPdf.bat "D:\YourInputFolder"
```

スクリプトのあるフォルダ自体を処理する場合：

```bat
cd /d D:\OfficeFileToPdf
ToPdf.bat
```

## マッピング JSON

`res/source_to_pdf_map.json` に対応表が出力されます。

構造例：

```json
{
  "source_root": "D:/input",
  "output_root": "D:/input/res",
  "mappings": [
    {
      "source_path": "D:/input/docs/plan.docx",
      "source_relative_path": "docs/plan.docx",
      "output_pdf_path": "D:/input/res/docs/plan.pdf",
      "output_relative_path": "docs/plan.pdf",
      "kind": "word"
    }
  ]
}
```

各フィールド：

- `source_root`：入力元フォルダ
- `output_root`：出力フォルダ
- `mappings`：正常に生成された対応一覧
- `source_path`：元ファイルの絶対パス
- `source_relative_path`：元ファイルの相対パス
- `output_pdf_path`：生成された PDF の絶対パス
- `output_relative_path`：`res/` から見た相対パス
- `kind`：`excel`、`word`、`powerpoint`、`pdf` のいずれか

## ログファイル

実行時に次のログが生成されます。

- `ToPdf_run.log`
- `ToPdf_python.log`

用途：

- `ToPdf_run.log`：バッチ起動部分の実行ログ
- `ToPdf_python.log`：Python 本体の出力とエラーログ

問題がある場合は、まずこの 2 つを確認してください。

## 注意事項

- Windows 専用ツールです
- Office 文書の変換は Microsoft Office がそのファイルを正常に開けることが前提です
- 拡張子が対応対象でも、Office で開けない場合は変換に失敗します
- PDF のトリミングに失敗した場合は、処理全体を止めずに PDF をそのままコピーします
- スキップされたファイルは `mappings` に入りません
- 現在、失敗したファイルも `mappings` には書き込みません

## よくある質問

### ダブルクリックするとすぐ閉じる

最新の `ToPdf.bat` を使用してください。

最新版では実行後に画面が止まり、ログファイルの場所も表示されます。

### `Excel.Application`、`Word.Application`、`PowerPoint.Application` を起動できない

主な原因：

- Microsoft Office が未インストール
- Office の COM 登録が壊れている
- 現在の Windows 環境で Office が利用できない

### 一部のファイルが出力されない

次を確認してください。

- そのファイルがスキップ対象ではないか
- `ToPdf_python.log` に失敗ログがないか
- `source_to_pdf_map.json` に対応レコードがあるか

## ファイル一覧

- `ToPdf.py`：メインプログラム
- `ToPdf.bat`：Windows 用起動スクリプト
- `requirements.txt`：Python 依存関係
- `pdfToPng.py`：旧来の PDF トリミング用スクリプト

## ライセンス

このプロジェクトは MIT License を採用しています。

詳細は [`LICENSE`](LICENSE) を参照してください。
