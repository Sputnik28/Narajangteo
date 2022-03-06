# 크롬 브라우저를 띄우기 위해, 웹드라이버를 가져오기
import openpyxl
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import Select
from operator import itemgetter

# 크롬 드라이버로 크롬을 실행한다.
driver = webdriver.Chrome('./chromedriver')

# 검색어를 읽어서 리스트로 저장하는 함수를 정의한다.
# 파이썬 리스트는 대괄호 []로 표기됨. 예) ['온실가스', '명세서']
# 검색어가 '/'로 구분되어 입력되어 있어야 한다.
def txt_reader(name):
    with open(name+".txt",'rb') as f:
        line = f.readline()
        return line.decode('utf-8').split('/')

try:
    # 검색어를 읽어들인다.
    queries = txt_reader('keywords')
    # 읽어들인 검색어를 화면에 보여준다.
    print(queries)
    output = []
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
        # 위 코드를 통해, 검색 결과에 있던 정보가 단순히 모두 하나의 줄에 나열된 리스트가 생성됨. ['용역','공고번호','하이퍼링크','분류',..., '용역','공고번호','분류',...]
        # 따라서 결과를 정리한 리스트(results)를 12개씩 분할하여 새로운 리스트(계층이 2개인)로 저장해서 보기 편하도록 만듦. [['용역','공고번호','하이퍼링크','분류',..., ],..., ['용역','공고번호','하이퍼링크','분류',...., ]]
        result = [results[i * 12:(i + 1) * 12] for i in range((len(results) + 12 - 1) // 12)]
        # 리스트에서 필요로 하는 정보는 5, 6, 8, 10이므로 해당 요소만 뽑는다.
        # 파이썬에서는 리스트 인덱스가 0부터 시작하므로 다음과 같이 인덱스 넘버를 지정한다.
        s = [4, 5, 7, 9]
        result = [list(itemgetter(*s)(el)) for el in result]
        output.extend(result)
except Exception as e:
    # 위 코드에서 에러가 발생한 경우 출력
    print(e)
finally:
    # 에러와 관계없이 실행되고, 크롬 드라이버를 종료
    driver.quit()
