import openpyxl
from datetime import datetime

# Excelファイルを開く
workbook = openpyxl.load_workbook('10158BONTE渡辺宏樹様納品書兼請求書.xlsx')

# 最初のシートを選択
sheet = workbook.active

# 特定のセルの値を取得
cell_value = sheet['A4'].value
print("A4セルの値:", cell_value)
today = datetime.today()
formatted_date = f"{today.year}年{today.month}月{today.day}日"
sheet['A1'].value = formatted_date
sheet['H2'].value = "" #請求番号
sheet['A8'].value = "上谷" + "　様" #請求先名
workbook.save("10158BONTE渡辺宏樹様納品書兼請求書01.xlsx")
workbook.close()
