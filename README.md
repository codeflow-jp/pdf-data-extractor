# PDF Data Extractor

PDFファイルからテキストや表データを自動抽出し、Excelファイルに出力するツールです。

## 機能

- **テキスト抽出モード** — PDFのテキストを1行ずつExcelに出力
- **テーブル抽出モード** — PDFの表を構造を保ったままExcelに出力
- `config.json` で設定を管理（コード変更不要で異なるPDFに対応可能）
- コマンドライン引数による入出力ファイルの指定

## 使用技術

- Python 3.14
- pdfplumber（PDF解析）
- openpyxl（Excel出力）

## セットアップ
```bash
pip install pdfplumber openpyxl
```

## 使い方

### 基本（config.jsonの設定で実行）
```bash
python pdf_tool.py
```

### コマンドライン引数で指定
```bash
python pdf_tool.py input.pdf output.xlsx
```

## 設定ファイル（config.json）
```json
{
    "input_pdf": "sample.pdf",
    "output_excel": "output.xlsx",
    "mode": "table",
    "sheet_name": "抽出データ"
}
```

| 項目 | 説明 |
|------|------|
| `input_pdf` | 読み込むPDFファイルのパス |
| `output_excel` | 出力するExcelファイルのパス |
| `mode` | `text`（テキスト抽出）または `table`（表抽出） |
| `sheet_name` | 出力Excelのシート名 |

## ファイル構成
```
pdf-data-extractor/
├── pdf_tool.py          # メインスクリプト
├── config.json          # 設定ファイル
├── make_sample_pdf.py   # テスト用PDF生成
├── sample.pdf           # サンプルPDF
└── README.md
```

## 出力例

### テーブル抽出モード
PDFの表がExcelにそのまま再現されます。

### テキスト抽出モード
PDFの全テキストが1行ずつExcelに出力されます。