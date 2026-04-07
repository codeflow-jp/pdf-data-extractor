# pdf_tool.py
# PDFデータ抽出ツール（テキスト・表の両対応）

import json
import sys
import os
import pdfplumber
import openpyxl


# ① 設定ファイルを読み込む
def load_config(config_path="config.json"):
    with open(config_path, "r", encoding="utf-8") as f:
        return json.load(f)


# ② PDFからテキストを抽出する
def extract_text(pdf_path):
    """全ページのテキストを行単位でリストにして返す"""
    results = []
    with pdfplumber.open(pdf_path) as pdf:
        for i, page in enumerate(pdf.pages):
            text = page.extract_text()
            if text:
                # 改行で分割して1行ずつ格納
                for line in text.split("\n"):
                    line = line.strip()
                    if line:
                        results.append({
                            "page": i + 1,
                            "text": line
                        })
    return results


# ③ PDFから表を抽出する
def extract_tables(pdf_path):
    """全ページの表をリストで返す"""
    results = []
    with pdfplumber.open(pdf_path) as pdf:
        for i, page in enumerate(pdf.pages):
            tables = page.extract_tables()
            for table in tables:
                results.append({
                    "page": i + 1,
                    "table": table
                })
    return results


# ④ テキストをExcelに書き出す
def write_text_to_excel(text_data, output_path, sheet_name):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = sheet_name

    ws.append(["ページ", "テキスト"])
    for item in text_data:
        ws.append([item["page"], item["text"]])

    # 列幅を自動調整
    ws.column_dimensions["A"].width = 10
    ws.column_dimensions["B"].width = 80

    wb.save(output_path)


# ⑤ 表をExcelに書き出す
def write_tables_to_excel(table_data, output_path, sheet_name):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = sheet_name

    for item in table_data:
        ws.append([f"--- ページ {item['page']} ---"])
        for row in item["table"]:
            ws.append(row)
        ws.append([])

    wb.save(output_path)


# ⑥ メイン処理
def main():
    config = load_config()

    # コマンドライン引数があればconfig設定を上書き
    pdf_path = sys.argv[1] if len(sys.argv) > 1 else config["input_pdf"]
    output_path = sys.argv[2] if len(sys.argv) > 2 else config["output_excel"]
    mode = config["mode"]
    sheet_name = config["sheet_name"]

    # PDFの存在チェック
    if not os.path.exists(pdf_path):
        print(f"エラー: {pdf_path} が見つかりません")
        sys.exit(1)

    # PDFの読み込みテスト
    try:
        with pdfplumber.open(pdf_path) as pdf:
            if len(pdf.pages) == 0:
                print("エラー: PDFにページがありません")
                sys.exit(1)
            print(f"[情報] {pdf_path}（{len(pdf.pages)}ページ）")
    except Exception as e:
        print(f"エラー: PDFを開けません → {e}")
        sys.exit(1)

    # モードに応じて処理を分岐
    if mode == "text":
        print("[モード] テキスト抽出")
        data = extract_text(pdf_path)
        if not data:
            print("テキストが見つかりませんでした")
            sys.exit(1)
        write_text_to_excel(data, output_path, sheet_name)
        print(f"完了: {output_path} を作成しました（{len(data)}行抽出）")

    elif mode == "table":
        print("[モード] テーブル抽出")
        data = extract_tables(pdf_path)
        if not data:
            print("テーブルが見つかりませんでした")
            sys.exit(1)
        write_tables_to_excel(data, output_path, sheet_name)
        print(f"完了: {output_path} を作成しました（{len(data)}件抽出）")

    else:
        print(f"エラー: 不明なモード '{mode}'（text または table を指定してください）")
        sys.exit(1)


if __name__ == "__main__":
    main()