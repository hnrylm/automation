from main import generate_excel  # 假設你把 main.py 的內容寫成了函數
from send_email import send_product_report

# 1. 先執行生成 Excel 的邏輯
print("正在生成最新的產品報表...")
# 如果你的 main.py 是一段直接執行的代碼，可以直接 import 它
import main 

# 2. 緊接著發送這份報表
print("報表生成成功，正在發送郵件...")
RECEIVER = "learnubuntuandprogramming@gmail.com"
send_product_report(RECEIVER, "product_list.xlsx")

print("✨ 全自動流程已圓滿完成！")
