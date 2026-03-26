from openpyxl import Workbook
from openpyxl.styles import Font

# 第一版本，只用 openpyxl 功能 sum(),不用loop只看计算结果
workbook = Workbook()
workbook.properties.title = "产品目录"

worksheet = workbook.active
worksheet.title = "Jan"

# header, cell "A1", cell "B1", cell "C1"
worksheet.append(["Product (产品名称)","Quantity (数量)", "Price（价钱）"])

# 内容
worksheet.append(["盒裝紙巾無味18盒裝",10,124.00])
worksheet.append(["盒裝紙巾無味18盒裝",10,124.00])

# 计算 Total，显示结果，从最后一笔留空一行
worksheet['A5'] = "总共消费 HKD"
worksheet['A5'].font = Font(name="Arial",size=14,bold=True)

# 单一消费 = cellB x cellC
# 总数 = 所有 单一消费加起来
worksheet['B5'] = "=Sum(B2:B3)*Sum(c2:c3)" #
workbook.save("product_list.xlsx")