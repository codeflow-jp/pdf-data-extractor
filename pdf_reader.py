# pdf_reader.py
# pdfplumberの基本操作を体験するスクリプト

import pdfplumber

# ① PDFを開く
with pdfplumber.open("sample.pdf") as pdf:

    # ② 基本情報を表示
    print(f"ページ数: {len(pdf.pages)}")
    print("=" * 50)

    # ③ 各ページを処理
    for i, page in enumerate(pdf.pages):
        print(f"\n--- ページ {i + 1} ---")

        # ④ テキスト抽出
        text = page.extract_text()
        if text:
            print("\n【テキスト】")
            print(text)

        # ⑤ 表（テーブル）抽出
        tables = page.extract_tables()
        if tables:
            print(f"\n【テーブル】{len(tables)}個検出")
            for t_idx, table in enumerate(tables):
                print(f"\nテーブル {t_idx + 1}:")
                for row in table:
                    print(row)
        else:
            print("\n【テーブル】検出なし")