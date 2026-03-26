import smtplib
from email.message import EmailMessage
# 1. 郵件配置
EMAIL_ADDRESS = "buildknowldege@gmail.com"
EMAIL_PASSWORD = "viuzseydidxfgwld"  # 不要帶空格
# RECEIVER = "learnubuntuandprogramming@gmail.com"
RECEIVER = "home.hang@gmail.com"
def send_product_report(receiver_email, file_path):
    # 2. 建立郵件內容
    msg = EmailMessage()
    msg['Subject'] = "📊 自動化報表：最新產品清單"
    msg['From'] = EMAIL_ADDRESS
    msg['To'] = receiver_email
    msg.set_content("你好！附件是今天自動生成的產品清單報表，請查收。")
    # 3. 讀取並添加附件 (product_list.xlsx)
    with open(file_path, 'rb') as f:
        file_data = f.read()
        file_name = f.name
    msg.add_attachment(
        file_data,
        maintype='application',
        subtype='vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        filename=file_name
    )
    # 4. 連接 Gmail 伺服器並發送
    try:
        with smtplib.SMTP_SSL('smtp.gmail.com', 465) as smtp:
            smtp.login(EMAIL_ADDRESS, EMAIL_PASSWORD)
            smtp.send_message(msg)
        print(f"✅ 郵件已成功發送至 {receiver_email}！")
    except Exception as e:
        print(f"❌ 發送失敗：{e}")
# 測試運行
if __name__ == "__main__":
    #send_product_report("learnubuntuandprogramming@gmail.com", "product_list.xlsx")
    send_product_report("home.hang@gmail.com", "product_list.xlsx")
