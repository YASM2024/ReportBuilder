# ReportGenerator

<p align="center">
  <img src="icon/logo.png" alt="ReportGenerator" width="200">
</p>

**Windows 向けツール**です。コマンドプロンプトまたは PowerShell 上で動作し、`run.bat` による起動を想定しています。

任意の CSV ファイルを、あらかじめ定義した設定に従って Excel レポートに成型します。

- **設定ファイル (YAML)** … 列マッピング、配置、ソート、罫線、入出力フォルダ、保存名を定義
- **Excel ひな形 (.xlsx)** … 印刷サイズ・向き、固定ヘッダ、書式を定義
- **Python ツール** … 上記を読み込み、CSV からレポートを自動生成

## 動作の流れ

```
CSV (import/)  +  設定 (config/*.yaml)  +  ひな形 (templates/*.xlsx)
                              ↓
                    ReportGenerator (main.py)
                              ↓
                    Excel レポート (export/)
```

## 必要環境

- **Windows 10 / 11**
- Python 3.10 以上

## セットアップ

```powershell
cd ReportGenerator
python -m venv .venv
.\.venv\Scripts\python.exe -m pip install -r requirements.txt
```

仮想環境の場所を変える場合は `config/app.yaml` の `python` を編集します。

```yaml
python: venv/Scripts/python.exe
```

## 使い方

### run.bat で実行（Windows）

```bat
run.bat
```

`config\app.yaml` の `python` で指定した Python で `config\report.yaml` を読み込み、`import\` 内の CSV を一括処理します。  
バッチファイルは **ANSI（Shift_JIS）** で保存されているため、日本語メッセージが Windows のコマンドプロンプトで正しく表示されます。

### コマンドラインで実行

```powershell
.\.venv\Scripts\python.exe main.py -c config/report.yaml
```

### 特定の CSV のみ処理

```powershell
.\.venv\Scripts\python.exe main.py -c config/report.yaml --csv import/sales_data.csv
```

## プロジェクト構成

```
ReportGenerator/
├── main.py                     # 実行エントリ
├── run.bat                     # Windows 用起動バッチ
├── icon/
│   └── logo.png                # ロゴ画像
├── report_builder/             # コアロジック
│   ├── config_loader.py        # 設定ファイル読み込み
│   ├── excel_writer.py         # CSV 読込・Excel 書き込み
│   ├── paths.py                # アプリルート・パス解決（exe 対応）
│   └── builder.py              # 処理のオーケストレーション
├── config/
│   ├── app.yaml                # Python パス・起動設定
│   └── report.yaml             # レポート設定サンプル
├── templates/
│   └── sales_report.xlsx       # Excel ひな形
├── import/                     # CSV 入力フォルダ
└── export/                     # Excel 書き出しフォルダ
```

## 設定ファイルの書き方

### app.yaml（起動設定）

`run.bat` が参照するアプリケーション設定です。

| 項目 | 説明 | 例 |
|------|------|-----|
| `python` | 使用する Python 実行ファイル（相対または絶対パス） | `.venv/Scripts/python.exe` |
| `report` | レポート設定 YAML のパス | `config/report.yaml` |

パスは **プロジェクトルート（exe 化時は exe と同じフォルダ）** 基準の相対パスで指定できます。

### report.yaml（レポート設定）

レポート設定内のパス（`paths` セクション）は、**プロジェクトルート（exe 化時は exe と同じフォルダ）** を基準にした相対パスで指定します。作業ディレクトリに依存しません。

```yaml
paths:
  csv_input: import
  excel_output: export
  template: templates/sales_report.xlsx
```

### paths（入出力）

| 項目 | 説明 |
|------|------|
| `csv_input` | CSV の入力フォルダ |
| `excel_output` | Excel の書き出しフォルダ（`export/`） |
| `template` | Excel ひな形のパス |

### output（保存名）

| 項目 | 説明 | 例 |
|------|------|-----|
| `base_name` | ファイル名のベース | `売上明細` |
| `include_timestamp` | タイムスタンプを付けるか | `true` |
| `timestamp_format` | 日時の形式 | `%Y%m%d_%H%M%S` |

`include_timestamp: true` の場合、`売上明細_20260626_071110.xlsx` のように出力されます。

複数 CSV を一括処理する場合は、ファイル名に CSV 名も含まれます（例: `売上明細_sales_data_20260626_071110.xlsx`）。

### csv（CSV 読み込み）

| 項目 | 説明 | 既定値 |
|------|------|--------|
| `encoding` | 文字コード（`auto` で自動判定） | `utf-8-sig` |
| `delimiter` | 区切り文字 | `,` |
| `pattern` | 入力フォルダ内のファイルパターン | `*.csv` |

### data（データ配置）

| 項目 | 説明 |
|------|------|
| `sheet_name` | 書き込み先シート名（省略時はアクティブシート） |
| `start_row` | データの開始行（ヘッダ行の直下など） |
| `columns` | CSV 列 → Excel 列のマッピング |
| `sort` | ソート条件 |

#### columns

```yaml
columns:
  - csv: 商品コード      # CSV の列名（または 0 始まりの列番号）
    excel: A              # Excel の列（A, B, C ...）
    number_format: "#,##0"  # 任意: 数値書式
    align: right            # 任意: left | center | right
```

| `align` | 意味 |
|---------|------|
| `left` | 左詰め |
| `center` | 中央揃え |
| `right` | 右詰め |

`align` を省略した列は、ひな形側の配置が維持されます。

#### sort

```yaml
sort:
  - csv: 商品名
    order: asc    # 昇順
  - csv: 数量
    order: desc   # 降順
```

複数指定した場合、上から順にソートキーとして適用されます。

### borders（罫線）

罫線は **領域指定（推奨）** と **個別指定（従来形式）** の両方に対応しています。  
適用順序は `regions` / `data_table` → `horizontal` → `vertical` です。

#### data_table（表全体のショートカット・推奨）

`data.columns` の列範囲と `header_row` から表領域を自動解決します。

```yaml
borders:
  data_table:
    header_row: 5
    mode: outline    # outline | horizontal | vertical | grid
    style: thin
```

| `mode` | 内容 | 向いている用途 |
|--------|------|----------------|
| `outline` | 外枠の4辺のみ | すっきりした帳票・報告書 |
| `horizontal` | 横罫線のみ（各行の区切り） | 行区切りの一覧 |
| `vertical` | 縦罫線のみ（各列の区切り） | 列区切りの一覧 |
| `grid` | 横罫線＋縦罫線（マス状） | セル単位で読みやすい表 |

`internal: false` を指定すると、`horizontal` / `vertical` は外枠の線だけになります。

#### regions（矩形領域を直接指定）

```yaml
borders:
  regions:
    - from_row: 5
      to_row: last_data_row
      from_col: A
      to_col: E
      mode: grid
      style: thin
```

#### horizontal / vertical（個別指定・従来形式）

```yaml
borders:
  horizontal:
    - row: 5              # 行番号、または last_data_row
      from_col: A
      to_col: E
      style: thin
  vertical:
    - col: A
      from_row: 5
      to_row: last_data_row
      style: thin
```

`last_data_row` は、書き込んだデータの最終行に置き換わります。

## Excel ひな形で設定する項目

以下はひな形ファイル側で事前に設定してください。ツールはこれらを変更しません。

- 印刷サイズ（A4 など）
- 印刷向き（縦 / 横）
- タイトル・固定ヘッダ・列見出し
- 列幅・フォント・背景色
- 印刷タイトル行、ウィンドウ枠の固定

サンプルひな形（`templates/sales_report.xlsx`）では、A4 横、5 行目をヘッダ、6 行目からデータ開始という構成になっています。

## 新しいレポートの追加手順

1. Excel でひな形（`templates/xxx.xlsx`）を作成する
2. 対応する設定ファイル（`config/xxx.yaml`）を作成する
3. CSV を `import/` に置く
4. `run.bat` または `main.py -c config/xxx.yaml` を実行する

## exe 化

`config` / `import` / `export` / `templates` は exe に同梱せず、実行ファイルと同じフォルダに配置します。パスは exe のあるディレクトリを基準に解決されます。

ビルドは別プロジェクト [BuildExe](../BuildExe) を使用します。

### 配布フォルダ構成

```
配布フォルダ/
├── ReportGenerator.exe   # または tranceformer.exe
├── run.bat               # 任意
├── config/
│   ├── app.yaml          # Python パス（run.bat 使用時）
│   └── report.yaml
├── import/
│   └── （CSV ファイル）
├── export/               # 自動作成される
└── templates/
    └── sales_report.xlsx
```

### exe での実行

```powershell
cd 配布フォルダ
.\ReportGenerator.exe -c config/report.yaml
.\ReportGenerator.exe -c config/report.yaml --csv import/sales_data.csv
```

`-c` や `--csv` に渡す相対パスも、exe のあるフォルダ基準で解決されます（カレントディレクトリは問いません）。

## ライセンス

このプロジェクトは社内・個人利用を想定したツールです。
