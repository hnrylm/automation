import openpyxl
from openpyxl.styles import Font

workbook = openpyxl.Workbook()
worksheet = workbook.active
worksheet.title = "Products"

# 1. 写入表头
worksheet.append(["Product", "Quantity", "Price"])

# 2. 数据资料
products_data = [
    ["盒装纸巾 A", 10, 124],
    ["盒装纸巾 B", 10, 124],
]

for row in products_data:
    worksheet.append(row)

# 3. 动态获取最后一行
last_row = worksheet.max_row

# 4. 写入总计
total_row = last_row + 2
worksheet.cell(row=total_row, column=1, value="总共消费").font = Font(name="Arial", size=14, bold=True)

# 【核心修正】：使用 SUMPRODUCT 代替 SUM*SUM
sumproduct_formula = f"=SUMPRODUCT(B2:B{last_row}, C2:C{last_row})"
worksheet.cell(row=total_row, column=2, value=sumproduct_formula)

# 5. 保存
workbook.save("product_list.xlsx")
print("Bug 已修复！现在使用的是 SUMPRODUCT 逻辑。")
