import sys
import openpyxl
from pathlib import Path
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication

#送付リスの読み込み
wb = openpyxl.load_workbook("送付リスト.xlsx")
ws = wb["Sheet1"]

customer_list = []

for row in ws.iter_rows(min_row=2):
  if row[0].value is None:
    break
  value_list = []
  for c in row:
    value_list.append(c.value)
  customer_list.append(value_list)
  
#PDFのフォルダー
pdf_dir = "PDF"

#メール送付リスト
mailing_list = []

#フォルダーからCPDのPDFファイルを1つずつ取得する
for invoice in Path(pdf_dir).glob("*.pdf"):
  #「ID」は、PDFの拡張子を除いた部分
  customer_id = invoice.stem
  #該当する送付先データを「送付リスト」から検索
  for customer in customer_list:
    if customer_id == customer[0]:
      #メール送付リストに「送付リスト」とPDFファイルのパスを追加
      mailing_list.append([customer, invoice])

#メール本文を読み込む
text = open("message.txt", encoding="utf-8")
body_temp = text.read()
text.close

#SMTPサーバー設定
smtp_server = "smtp.gmail.com"
port_number = 587

#ログイン情報
f = open('gmail.txt', 'r', encoding='UTF-8')
account = f.read()
print(account)
f = open('pass.txt', 'r', encoding='UTF-8')
password = f.read()

#SMTPサーバーに接続
server = smtplib.SMTP(smtp_server, port_number)
server.starttls()
server.login(account, password)

#メール送付リストの顧客に1つずつメール送信
my_address = "Tomoko <tomokyuuu@gmail.com>"
for data in mailing_list:
  customer = data[0]
  pdf_file = data[1]
  
  #メッセージの準備
  msg = MIMEMultipart()
  
  #件名、メールアドレスの設定
  msg["Subject"] = "CPD受講証明書 応用生態工学会富山支部"
  msg["From"] = my_address
  msg["To"] = customer[4]
  
  #メール本文の追加
  body_text = body_temp.format(
    company=customer[1],
    department=customer[2],
    person=customer[3]
   )
  body = MIMEText(body_text)
  msg.attach(body)
  #添付ファイルの追加
  pdf = open(pdf_file, mode="rb")
  pdf_data = pdf.read()
  pdf.close
  attach_file = MIMEApplication(pdf_data)
  attach_file.add_header("Content-Disposition",
                         "attachement",filename=pdf_file.name)
  msg.attach(attach_file)
  
  #メール送信
  print("メール送信", customer[0], customer[1])
  server.send_message(msg)  
  
#SMTPサーバーとの接続を閉じる
server.quit
print("処理完了")
