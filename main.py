import search

print(search.output)

# # 내보내기 옵션1. 엑셀파일로 저장하고 자동으로 구동
#     ws = []
#     for i in output:
#         ws.append(i)
#
# wb.save("results.xlsx")
# path = os.getcwd()
# excel = win32com.client.Dispatch("Excel.Application")
# excel.Visible = True
# filename = "/results.xlsx"
# print(path+filename)
# wb = excel.Workbooks.Open(path+filename)

# # 내보내기 옵션2. 이메일 보내기
# HTML로 변환
html = """
<HTML>
    <head>
    <meta charset="UTF-8">
    </head>
    <body>
        <h1>나라장터 공고 검색결과</h1>
        <table>
            {0}
        </table></body>
</HTML>"""

items = search.output
tr = "<tr>{0}</tr>"
td = "<td>{0}</td>"
subitems = [tr.format(''.join([td.format(a) for a in item])) for item in items]
html = html.format("".join(subitems))

print(html)

# 이메일 보내기
import csv
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import smtplib

me = open('myemail.txt').readline()
password = open('pw.txt').readline()
server = 'smtp.daum.net'
port = 465
you = open('myemail.txt').readline()

message = MIMEMultipart("alternative", None, [MIMEText(html, 'html',_charset='utf-8')])

message['Subject'] = "나라장터 입찰공고 검색결과"
message['From'] = me
message['To'] = you
server = smtplib.SMTP_SSL(server, port)
# server.ehlo()
# server.starttls()
server.login(me, password)
server.sendmail(me, you, message.as_string())
server.quit()