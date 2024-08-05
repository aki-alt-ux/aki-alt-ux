import openpyxl

# Excelファイルを開く
workbook = openpyxl.load_workbook('10158BONTE渡辺宏樹様納品書兼請求書.xlsx')

# 最初のシートを選択
sheet = workbook.active

# 特定のセルの値を取得
cell_value = sheet['A4'].value
print("A4セルの値:", cell_value)
sheet['A4'].value = "納品書兼請求書_清水スペシャル"
workbook.save("10158BONTE渡辺宏樹様納品書兼請求書01.xlsx")
workbook.close()
