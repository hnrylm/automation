import openpyxl
from openpyxl.styles import Font

# 1. 建立 Excel 檔案
workbook = openpyxl.Workbook()
worksheet = workbook.active
worksheet.title = "Products"

# 2. 寫入標題
worksheet.append(["Product", "Quantity", "Price"])

# 3. 準備數據（以後你可以隨意增加這裡的內容）
products_data = [
    ["Product 1", 10, 124],
    ["Product 2", 10, 124],
    ["Product 3", 5, 50],   # 這是新增的，代碼會自動計算
]

# 將數據寫入表格
for row in products_data:
    worksheet.append(row)

# 4. 【重點】動態計算最後一行的位置
last_row = worksheet.max_row  # 自動偵測目前寫到了第幾行

# 5. 寫入總消費（放在數據的下一行）
total_row = last_row + 2
worksheet.cell(row=total_row, column=1, value="總消費 HKD").font = Font(name="Arial", size=14, bold=True)

# 使用動態公式：計算 SUM(B2:B最後一行) * SUM(C2:C最後一行)
sum_formula = f"=SUM(B2:B{last_row})*SUM(C2:C{last_row})"
worksheet.cell(row=total_row, column=2, value=sum_formula)

# 6. 存檔
workbook.save("product_list.xlsx")
print(f"Excel 已生成！共處理了 {len(products_data)} 項產品。")
