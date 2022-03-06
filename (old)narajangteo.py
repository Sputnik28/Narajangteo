# 크롬 브라우저를 띄우기 위해, 웹드라이버를 가져오기
import openpyxl
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import Select
import win32com.client
import os

# 크롬 드라이버로 크롬을 실행한다.
driver = webdriver.Chrome('./chromedriver')

def txt_reader(name):
    with open(name+".txt",'rb') as f:
        line = f.readline()
        return line.decode('utf-8').split('/')

queries = txt_reader('queries')

try:
    # 검색어
    # queries = ['온실가스', '명세서']  # 파이썬 리스트는 대괄호 []로 써야함

    queries = txt_reader('queries')
    print(queries)
    from openpyxl import Workbook

    wb = Workbook()
    ws = wb.active

    table = []

    for query in queries:

        # 입찰정보 검색 페이지로 이동
        driver.get('https://www.g2b.go.kr:8101/ep/tbid/tbidFwd.do?bidSearchType=1&taskClCds=5')

        # id값이 bidNm인 태그 가져오기
        bidNm = driver.find_element_by_id('bidNm')
        # 내용을 삭제 (버릇처럼 사용할 것!)
        bidNm.clear()
        # 검색어 입력후 엔터
        bidNm.send_keys(query)
        bidNm.send_keys(Keys.RETURN)

        # 검색 조건 체크
        option_dict = {'검색기간 1달': 'setMonth1_1', '입찰마감건 제외': 'exceptEnd', '검색건수 표시': 'useTotalCount'}
        for option in option_dict.values():
            checkbox = driver.find_element_by_id(option)
            checkbox.click()

        # 목록수 100건 선택 (드롭다운)
        recordcountperpage = driver.find_element_by_name('recordCountPerPage')
        selector = Select(recordcountperpage)
        selector.select_by_value('100')

        # 검색 버튼 클릭
        search_button = driver.find_element_by_class_name('btn_mdl')
        search_button.click()

        # 검색 결과 확인
        elem = driver.find_element_by_class_name('results')
        div_list = elem.find_elements_by_tag_name('div')

        # 검색 결과 모두 긁어서 리스트로 저장
        results = []
        for div in div_list:
            results.append(div.text)
            a_tags = div.find_elements_by_tag_name('a')
            if a_tags:
                for a_tag in a_tags:
                    link = a_tag.get_attribute('href')
                    results.append(link)

        # 검색결과 모음 리스트를 12개씩 분할하여 새로운 리스트로 저장
        result = [results[i * 12:(i + 1) * 12] for i in range((len(results) + 12 - 1) // 12)]

        table.extend(result)
    #     # 엑셀파일로 저장하고 자동으로 구동---------------시작--------
    #     for i in result:
    #         ws.append(i)
    #
    # wb.save("results.xlsx")
    # path = os.getcwd()
    # excel = win32com.client.Dispatch("Excel.Application")
    # excel.Visible = True
    # filename = "/results.xlsx"
    # print(path+filename)
    # wb = excel.Workbooks.Open(path+filename)
    # # 엑셀파일로 저장하고 자동으로 구동---------------종료--------

        # HTML로 변환
        # HTML로 변환
        html = """
    <HTML>
        <body>
            <h1>나라장터 공고 검색결과</h1>
            <table>
                {0}
            </table></body>
    </HTML>"""

        items = table
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

    message = MIMEMultipart("alternative", None, [MIMEText(html, 'html')])

    message['Subject'] = "나라장터 입찰공고 검색결과"
    message['From'] = me
    message['To'] = you
    server = smtplib.SMTP_SSL(server, port)
    # server.ehlo()
    # server.starttls()
    server.login(me, password)
    server.sendmail(me, you, message.as_string())
    server.quit()

except Exception as e:
    # 위 코드에서 에러가 발생한 경우 출력
    print(e)
finally:
    # 에러와 관계없이 실행되고, 크롬 드라이버를 종료
    driver.quit()
