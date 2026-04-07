# make_sample_pdf.py
# テスト用のサンプルPDFを生成するスクリプト

from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

# ① PDFファイルを作成
c = canvas.Canvas("sample.pdf", pagesize=A4)
width, height = A4

# ② フォント設定（日本語対応）
font_path = "C:/Windows/Fonts/msgothic.ttc"
pdfmetrics.registerFont(TTFont("Gothic", font_path))
c.setFont("Gothic", 12)

# ③ タイトルを描画
c.setFont("Gothic", 18)
c.drawString(200, height - 50, "月次売上レポート")

# ④ テキスト情報を描画
c.setFont("Gothic", 12)
c.drawString(50, height - 100, "作成日: 2025年7月1日")
c.drawString(50, height - 120, "担当者: 田中太郎")
c.drawString(50, height - 140, "部署: 営業部")

# ⑤ 表を描画（線と文字で手動作成）
table_top = height - 200
row_height = 25
col_widths = [150, 100, 100, 100]
headers = ["商品名", "数量", "単価", "小計"]
data = [
    ["ノートPC", "10", "80,000", "800,000"],
    ["マウス", "50", "2,000", "100,000"],
    ["キーボード", "30", "5,000", "150,000"],
    ["モニター", "15", "35,000", "525,000"],
]

# ヘッダー行
x = 50
for i, header in enumerate(headers):
    c.rect(x, table_top - row_height, col_widths[i], row_height)
    c.drawString(x + 5, table_top - row_height + 8, header)
    x += col_widths[i]

# データ行
for row_idx, row in enumerate(data):
    x = 50
    y = table_top - row_height * (row_idx + 2)
    for col_idx, cell in enumerate(row):
        c.rect(x, y, col_widths[col_idx], row_height)
        c.drawString(x + 5, y + 8, cell)
        x += col_widths[col_idx]

# ⑥ 合計行
c.setFont("Gothic", 14)
c.drawString(50, table_top - row_height * 7, "合計: 1,575,000円")

# ⑦ 保存
c.save()
print("sample.pdf を作成しました")