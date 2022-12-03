from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
# from selenium.webdriver.common.action_chains import ActionChains

from webdriver_manager.chrome import ChromeDriverManager
# from bs4 import BeautifulSoup

import time
import pandas as pd

# Headless Chrome (백그라운드에서 크롬을 돌려준다.)
# options = webdriver.ChromeOptions()
# options.headless = True
# options.add_argument("window-size=1920x1080")
# options.add_argument(
#     "user-agent=Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/95.0.4638.69 Safari/537.36")
# driver = webdriver.Chrome(ChromeDriverManager().install(), options=options)

# 크롬 드라이버 최신 상태로 Install (따로 드라이버 파일 설치할 필요 없음)
options = webdriver.ChromeOptions()
options.add_experimental_option('excludeSwitches', ['enable-logging'])
driver = webdriver.Chrome(ChromeDriverManager().install(), options=options)


# 1. 네이버 이동
url = 'https://sell.smartstore.naver.com/#/login'
driver.get(url)

# 브라우저 창 최대화
# driver.maximize_window()
driver.set_window_size(1500, 980)

# 스마트스토어 계정
id = "비밀번호" 
pw = "비밀번호"

driver.implicitly_wait(time_to_wait=5)

# 로그인
elem = driver.find_element(By.ID, 'loginId')
elem.send_keys(id)
elem = driver.find_element(By.ID, 'loginPassword')
elem.send_keys(pw)
elem.send_keys(Keys.RETURN)

driver.implicitly_wait(time_to_wait=5)

# 상품관리 클릭
elem = driver.find_element(
    By.XPATH, '//*[@id="seller-lnb"]/div/div[1]/ul/li[1]/a')
elem.click()

# 상품 조회/수정 클릭
elem = driver.find_element(
    By.XPATH, '//*[@id="seller-lnb"]/div/div[1]/ul/li[1]/ul/li[1]/a')
elem.click()

# 등록할 상품 엑셀
df = pd.read_excel('./상품등록예시.xlsx',
                sheet_name=0, engine='openpyxl')
# 카테고리 및 태그 엑셀
category_tag = pd.read_excel('./카테고리태그규칙.xlsx',
                            sheet_name=0, engine='openpyxl')
# 브랜드 명 및 브랜드 ID 엑셀                         
brand = pd.read_excel('./브랜드.xlsx',
                            sheet_name=0, engine='openpyxl')


# 등록할 상품 엑셀 loop
try:
    for r in df.index:
        
        if r == 0:
            print('첫번째 구동')
        
        elif r % 20 == 0:
            driver.quit()
            
            time.sleep(2)
            
            # Headless Chrome (백그라운드에서 크롬을 돌려준다.)
            # options = webdriver.ChromeOptions()
            # options.headless = True
            # options.add_argument("window-size=1920x1080")
            # options.add_argument(
            #     "user-agent=Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/95.0.4638.69 Safari/537.36")
            # driver = webdriver.Chrome(ChromeDriverManager().install(), options=options)

            # 크롬 드라이버 최신 상태로 Install (따로 드라이버 파일 설치할 필요 없음)
            driver = webdriver.Chrome(ChromeDriverManager().install())

            # 1. 네이버 이동
            url = 'https://sell.smartstore.naver.com/#/login'
            driver.get(url)

            # 브라우저 창 최대화
            driver.maximize_window()
            # driver.set_window_size(1820, 980)

            # 스마트스토어 계정
            id = 'continewus'
            pw = 'conteenew1'

            driver.implicitly_wait(time_to_wait=5)

            # 로그인
            elem = driver.find_element(By.ID, 'loginId')
            elem.send_keys(id)
            elem = driver.find_element(By.ID, 'loginPassword')
            elem.send_keys(pw)
            elem.send_keys(Keys.RETURN)

            driver.implicitly_wait(time_to_wait=5)

            # 상품관리 클릭
            elem = driver.find_element(
                By.XPATH, '//*[@id="seller-lnb"]/div/div[1]/ul/li[1]/a')
            elem.click()

            # 상품 조회/수정 클릭
            elem = driver.find_element(
                By.XPATH, '//*[@id="seller-lnb"]/div/div[1]/ul/li[1]/ul/li[1]/a')
            elem.click()
        

        #  --------------------------------  상품 조회   --------------------------------
        # 상품번호
        product_no = int(df.loc[r, '상품번호(스마트스토어)'])
        print("=" * 50)
        print(f"{r}번째")
        print(product_no)

        # 상품 번호 input 지우기
        elem = driver.find_element(
            By.XPATH, '//*[@id="seller-content"]/ui-view/div/ui-view[1]/div[2]/form/div[1]/div/ul/li[1]/div/div/div[2]/textarea')
        elem.clear()

        # 상품번호 입력
        elem = driver.find_element(
            By.XPATH, '//*[@id="seller-content"]/ui-view/div/ui-view[1]/div[2]/form/div[1]/div/ul/li[1]/div/div/div[2]/textarea')
        elem.send_keys(product_no)

        # 검색 버튼 클릭
        elem = driver.find_element(
            By.XPATH, '//*[@id="seller-content"]/ui-view/div/ui-view[1]/div[2]/form/div[2]/div/button[1]')
        elem.click()

        # 판매자상품코드 변수에 저장 (브랜드ID 입력에 사용)
        # celler_code = WebDriverWait(driver, 5).until(EC.presence_of_element_located(
            # (By.XPATH, '//*[@id="seller-content"]/ui-view/div/ui-view[2]/div[1]/div[2]/div[3]/div/div/div/div/div[3]/div[1]/div/div[@col-id="sellerManagementCode"]')))
                        
        # celler_code = driver.find_element(
        #     By.XPATH, '//*[@id="seller-content"]/ui-view/div/ui-view[2]/div[1]/div[2]/div[3]/div/div/div/div/div[3]/div[@class="ag-pinned-left-cols-container"]/div/div[@col-id="sellerManagementCode"]').text
        
        celler_code = str(df.loc[r, '판매자상품코드'])
        print(celler_code)

        # 복사버튼 클릭
        elem = driver.find_element(
            By.XPATH, '//*[@id="seller-content"]/ui-view/div/ui-view[2]/div[1]/div[2]/div[3]/div/div/div/div/div[3]/div[1]/div/div[3]/span/button')
        elem.click()

        driver.implicitly_wait(10)

        #  --------------------------------  상품 페이지   --------------------------------

        # 안전기준 준수 모달창이 뜨는 태그들
        safe = ['쿠션커버', '무릎담요', '이불커버', '니트/스웨터', '티셔츠', '발매트', '일반쿠션']

        # 기존의 카테고리 정보 가져오기
        elem = driver.find_element(
            By.XPATH, '//*[@id="productForm"]/ng-include/ui-view[3]/div/div[2]/div/div[1]/div/p[1]')
        origin_category = str(elem.text)
        last_word = origin_category.split('>')[-1]
        print(last_word)

        # 안전기준 준수 대상 품목 배너 확인
        if last_word in safe:

            try:
                # 확인버튼
                elem = WebDriverWait(driver, 5).until(EC.presence_of_element_located(
                    (By.XPATH, '//div[@class="seller-btn-area"]/button[contains(text(), "확인")]')))
                elem.click()
                time.sleep(1.5)
            except:
                print('첫번째 안전기준 준수 대상 품목 창은 없습니다.')

        # 상품명
        title = driver.find_element(
            By.XPATH, '//*[@id="productForm"]/ng-include/ui-view[7]/div/div[2]/div/div/div/div/div/div/input').get_attribute("value")

        # --------------------------------  카테고리  --------------------------------

        # 아트윈도용 카테고리와 태그를 가지고있는 row를 찾아서 가져온다
        ct = category_tag[category_tag['스마트스토어 카테고리 키워드'] == last_word]

        # 스마트 스토어와 아트윈도에서 같은 카테고리를 사용하는 상품들
        same = ['아트포스터', '엽서', '여행소품케이스', '카드/명함지갑', '캘린더/달력',
                '기타주방잡화', '노트', '플라스틱케이스', '기타휴대폰액세서리', '무릎담요', '아트포스터', '티셔츠', '데스크패드', '일반쿠션']

        if last_word in same:
            print('같은 카테고리입니다. 수정 입력을 하지 않습니다.')

        else:
            time.sleep(1)

            # 아트윈도 등록용 카테고리 번호
            category_no = int(ct.iloc[0]['아트윈도 등록용 카테고리'])

            # 카테고리 INPUT 클릭
            category_box = WebDriverWait(driver, 5).until(EC.presence_of_element_located(
                (By.XPATH, '//*[@id="productForm"]/ng-include/ui-view[3]/div/div[2]/div/div[1]/div/category-search/div[2]/div/div/div[1]')))
            # category_box = driver.find_element(
            #     By.XPATH, '//*[@id="productForm"]/ng-include/ui-view[3]/div/div[2]/div/div[1]/div/category-search/div[2]/div/div/div[1]')
            category_box.click()

            # 키테고리 INPUT clear
            category_input = WebDriverWait(driver, 5).until(EC.presence_of_element_located(
                (By.XPATH, '//*[@id="productForm"]/ng-include/ui-view[3]/div/div[2]/div/div/div/category-search/div[2]/div/div/div[1]/input')))
            # category_input = driver.find_element(
            #     By.XPATH, '//*[@id="productForm"]/ng-include/ui-view[3]/div/div[2]/div/div/div/category-search/div[2]/div/div/div[1]/input')
            category_input.send_keys(Keys.BACKSPACE)

            # 카테고리 번호 입력
            time.sleep(1)
            category_input.send_keys(category_no)

            time.sleep(1)
            category_input.send_keys(Keys.RETURN)

            # 아트윈도용 카테고리 입력 후 안전기준 준수 모달창
            if last_word in safe:

                try:
                    # 확인버튼
                    elem = WebDriverWait(driver, 5).until(EC.presence_of_element_located(
                        (By.XPATH, '//div[@class="seller-btn-area"]/button[contains(text(), "확인")]')))
                    elem.click()

                except:
                    print('2번째 안전기준 준수 대상 품목 창이 뜨지 않았습니다')
            else:
                print('2번째 안전기준 준수 대상이 아닙니다.')

        # --------------------------------  상세설명   --------------------------------

        try:
            # 상세설명으로 이동
            detail_describe = driver.find_element(
                By.XPATH, '//*[@id="productForm"]/ng-include/ui-view[13]/div/div[1]')
            driver.execute_script(
                            "arguments[0].scrollIntoView(true);", detail_describe)

            # 스마트에디터로 변경 클릭
            trans_to_smart = driver.find_element(
                By.XPATH, '//*[@id="productForm"]/ng-include/ui-view[13]/div/div[2]/div/div/ncp-editor-form/div[4]/div/a')
            trans_to_smart.click()

            # 새탭의 등록창 기다림
            WebDriverWait(driver, 10).until(EC.number_of_windows_to_be(2))
            driver.implicitly_wait(time_to_wait=10)

            # 상품마다 로딩속도가 달라 너무 빨리 등록 버튼을 입력하는 씹히는 현상이 발생한다.
            time.sleep(2)

            # 새 탭으로 드라이버 포커스를 옮겨준다
            driver.switch_to.window(driver.window_handles[-1])

            # 등록 버튼클릭
            resist_btn = driver.find_element(
                By.XPATH, '/html/body/ui-view[1]/ncp-editor-launcher/div[1]/div')
            resist_btn.click()

            # 원래의 브라우저로 드라이버 포커스를 옮겨준다
            driver.switch_to.window(driver.window_handles[0])

            print('스마트에디터로 변경이 완료되었습니다.')

        except:
            # 기존에 스마트에디터로 등록된 경우
            print('스마트에디터가 이미 등록되어있습니다.')

        # --------------------------------  상품 주요 정보   --------------------------------
        # 변경하지 않는 카테고리
        none_attr = ['엽서', '기타주방잡화', '돗자리/매트', '발매트',
                     '바란스', '바스/비치타월', '머그', '텀블러', '데코스티커']

        # 카테고리 아트포스터(50006312)
        art_poster = ['퍼즐/그림/사진액자', '아트포스터', '퍼즐']

        # 상품 주요정보로 이동 및 클릭                  
        elem = driver.find_element(
            By.XPATH, '//*[@id="_prod-attr-section"]/div[1]/div/div/a')
        driver.execute_script("arguments[0].scrollIntoView(true);", elem)
        elem.click()

        # --------------------------------  브랜드 입력   --------------------------------
        brand_row = brand[brand['작가코드'] == celler_code]

        brand_id = str(brand_row['브랜드아이디'].values[0])

        # 브랜드 Input 클릭
        brand_box = WebDriverWait(driver, 5).until(EC.presence_of_element_located(
            (By.XPATH, '//*[@id="_prod-attr-section"]/div[2]/div/ncp-naver-shopping-search-info/div[2]/div/div[1]/div/ncp-brand-manufacturer-input/div/div/div/div/div/div[1]')))
        brand_box.click()
        time.sleep(2)

        # 브랜드 ID 입력
        brand_input = WebDriverWait(driver, 5).until(EC.presence_of_element_located(
            (By.XPATH, '//*[@id="_prod-attr-section"]/div[2]/div/ncp-naver-shopping-search-info/div[2]/div/div[1]/div/ncp-brand-manufacturer-input/div/div/div/div/div/div[1]/input')))
        if celler_code != 'c001' :
            brand_input.send_keys(brand_id)

        else:
            #설정 안함 클릭
            noneset_click = driver.find_element(
                By.XPATH, '//*[@id="_prod-attr-section"]/div[2]/div/ncp-naver-shopping-search-info/div[2]/div/div[1]/div/ncp-brand-manufacturer-input/div/div/div/span/button')
            noneset_click.click()

            # 자체제작 상품 선택
            selfmade_click = driver.find_element(
                By.XPATH, '//*[@id="_prod-attr-section"]/div[2]/div/ncp-naver-shopping-search-info/div[2]/div/div[1]/div/div/div/div/label')
            selfmade_click.click()


        time.sleep(2)

        brand_input.send_keys(Keys.RETURN)


        # --------------------------------  상품 속성 입력   --------------------------------
        if last_word in none_attr:
            print('속성 입력이 필요 없는 상품 카테고리 입니다.')

        else:
            # 상품 속성으로 이동
            elem = driver.find_element(
                By.XPATH, '//*[@id="_prod-attr-section"]/div[2]/div/div[1]')
            driver.execute_script("arguments[0].scrollIntoView(true);", elem)
            time.sleep(2)

            if last_word in art_poster:
                # 캔버스 액자, 종이 포스터, 천 포스터, 쉬폰천포스터, 직소퍼즐

                # 그림장르 선택
                genre = str(df.loc[r]['추가필드1'])

                genre_select = driver.find_element(
                    By.XPATH, '//*[@id="_prod-attr-section"]/div[2]/div/div[1]/div/ncp-product-attribute-table/div[2]/table/tbody/tr[2]/td/div/div/div/div/div[1]')
                genre_select.click()

                select_option = driver.find_element(
                    By.XPATH, f'//*[@id="_prod-attr-section"]/div[2]/div/div[1]/div/ncp-product-attribute-table/div[2]/table/tbody/tr[2]/td/div/div/div/div/div[2]/div/div[contains(text(), "{genre}")]')
                select_option.click()

                # 캔버스 호수 선택
                size_unit = driver.find_element(
                    By.XPATH, '//*[@id="_prod-attr-section"]/div[2]/div/div[1]/div/ncp-product-attribute-table/div[2]/table/tbody/tr[4]/td/div/div[1]/div/div/div[1]')
                size_unit.click()

                # 50호 ~ 80호 선택
                select_option = driver.find_element(
                    By.XPATH, '//*[@id="_prod-attr-section"]/div[2]/div/div[1]/div/ncp-product-attribute-table/div[2]/table/tbody/tr[4]/td/div/div[1]/div/div/div[2]/div/div[5]')
                select_option.click()

                # 캔버스크기 > 실제값 입력

                if "m" in str(title).split()[-1]:
                    real_size = str(title).split()[-1][:-3]
                else:
                    real_size = str(title).split()[-1][:-1]
                print(real_size)

                real_size_input = driver.find_element(
                    By.XPATH, '//*[@id="_prod-attr-section"]/div[2]/div/div[1]/div/ncp-product-attribute-table/div[2]/table/tbody/tr[4]/td/div/div[2]/div[1]/div/input')
                real_size_input.send_keys(real_size)

                # 캔버스크기 > 단위 선택
                unit_select = driver.find_element(
                    By.XPATH, '//*[@id="_prod-attr-section"]/div[2]/div/div[1]/div/ncp-product-attribute-table/div[2]/table/tbody/tr[4]/td/div/div[2]/div[2]/div/div[1]')
                unit_select.click()

                #  단위 선택: default "mm"
                measure = "A02036"  # mm

                # if "cm" in str(title).split()[-1]:
                    # measure = "A02034"  # cm

                unit_option = driver.find_element(
                    By.XPATH, f'//*[@id="_prod-attr-section"]/div[2]/div/div[1]/div/ncp-product-attribute-table/div[2]/table/tbody/tr[4]/td/div/div[2]/div[2]/div/div[2]/div/div[@data-value="{measure}"]')
                unit_option.click()

                # 사용추천장소 전체 클릭
                rec_place = driver.find_elements(
                    By.XPATH, '//*[@id="_prod-attr-section"]/div[2]/div/div[1]/div/ncp-product-attribute-table/div[2]/table/tbody/tr[6]/td/div/div/div[1]/label')

                for idx, label in enumerate(rec_place):
                    print(label.text)
                    label.click()

                print('캔버스 액자, 종이 포스터, 천 포스터, 직소퍼즐 주요정보가 입력되었습니다.')

            elif last_word == '플라스틱케이스':
                # 폰케이스

                # 재질 선택
                material_select = driver.find_element(
                    By.XPATH, '//*[@id="_prod-attr-section"]/div[2]/div/div[1]/div/ncp-product-attribute-table/div[2]/table/tbody[1]/tr[2]/td/div/div/div/div/div[1]')
                material_select.click()

                material_option = driver.find_element(
                    By.XPATH, '//*[@id="_prod-attr-section"]/div[2]/div/div[1]/div/ncp-product-attribute-table/div[2]/table/tbody[1]/tr[2]/td/div/div/div/div/div[2]/div/div[contains(text(), "플라스틱")]')
                material_option.click()

                # 형태 체크박스 클릭
                case_shape = ['다이어리형', '케이스형']

                for label in case_shape:
                    check_box = driver.find_element(
                        By.XPATH, f'//*[@id="_prod-attr-section"]/div[2]/div/div[1]/div/ncp-product-attribute-table/div[2]/table/tbody[1]/tr[3]/td/div/div/div[1]/label[contains(text(), "{label}")]')
                    check_box.click()

                # 품목 선택
                item_select = driver.find_element(
                    By.XPATH, '//*[@id="_prod-attr-section"]/div[2]/div/div[1]/div/ncp-product-attribute-table/div[2]/table/tbody[1]/tr[4]/td/div/div/div/div/div[1]')
                item_select.click()

                # 품목 > 폰케이스 선택
                item_option = driver.find_element(
                    By.XPATH, '//*[@id="_prod-attr-section"]/div[2]/div/div[1]/div/ncp-product-attribute-table/div[2]/table/tbody[1]/tr[4]/td/div/div/div/div/div[2]/div/div[10]')
                item_option.click()

                # 부가기능 > 잠금방식 선택
                sub_select = driver.find_element(
                    By.XPATH, '//*[@id="_prod-attr-section"]/div[2]/div/div[1]/div/ncp-product-attribute-table/div[2]/table/tbody[2]/tr[2]/td/div/div/div/div/div[1]')
                sub_select.click()

                # 잠금방식 상,하판결합 선택
                sub_option = driver.find_element(
                    By.XPATH, '//*[@id="_prod-attr-section"]/div[2]/div/div[1]/div/ncp-product-attribute-table/div[2]/table/tbody[2]/tr[2]/td/div/div/div/div/div[2]/div/div[contains(text(), "상,하판결합")]')
                sub_option.click()

                # 규격 > 색상 선택
                case_color = '혼합'

                color_select = driver.find_element(
                    By.XPATH, '//*[@id="_prod-attr-section"]/div[2]/div/div[1]/div/ncp-product-attribute-table/div[2]/table/tbody[3]/tr[2]/td/div/div[1]/div/div/div[1]')
                color_select.click()

                color_option = driver.find_element(
                    By.XPATH, '//*[@id="_prod-attr-section"]/div[2]/div/div[1]/div/ncp-product-attribute-table/div[2]/table/tbody[3]/tr[2]/td/div/div[1]/div/div/div[2]/div/div[contains(text(), "색상")]')
                color_option.click()

                # 잠금방식 상,하판결합 선택
                color_input = driver.find_element(
                    By.XPATH, '//*[@id="_prod-attr-section"]/div[2]/div/div[1]/div/ncp-product-attribute-table/div[2]/table/tbody[3]/tr[2]/td/div/div[2]/div/div/input')
                color_input.send_keys(case_color)

                print('스마트폰 주요 정보가 입력되었습니다.')

            elif last_word == '여행소품케이스':
                # 여권케이스

                # 용도
                use_select = driver.find_element(
                    By.XPATH, '//*[@id="_prod-attr-section"]/div[2]/div/div[1]/div/ncp-product-attribute-table/div[2]/table/tbody/tr[2]/td/div/div/div/div/div[1]')
                use_select.click()

                use_option = driver.find_element(
                    By.XPATH, '//*[@id="_prod-attr-section"]/div[2]/div/div[1]/div/ncp-product-attribute-table/div[2]/table/tbody/tr[2]/td/div/div/div/div/div[2]/div/div[contains(text(), "기타")]')
                use_option.click()

                # 주요소재
                material_select = driver.find_element(
                    By.XPATH, '//*[@id="_prod-attr-section"]/div[2]/div/div[1]/div/ncp-product-attribute-table/div[2]/table/tbody/tr[3]/td/div/div/div/div/div[1]')
                material_select.click()

                material_option = driver.find_element(
                    By.XPATH, '//*[@id="_prod-attr-section"]/div[2]/div/div[1]/div/ncp-product-attribute-table/div[2]/table/tbody/tr[3]/td/div/div/div/div/div[2]/div/div[contains(text(), "기타")]')
                material_option.click()

                # 크기 > 소형
                size_checkbox = driver.find_element(
                    By.XPATH, '//*[@id="_prod-attr-section"]/div[2]/div/div[1]/div/ncp-product-attribute-table/div[2]/table/tbody/tr[4]/td/div/div/div[1]/label[contains(text(), "소형")]')
                size_checkbox.click()

                print('여권케이스 주요 정보가 입력되었습니다.')

            elif last_word == '카드/명함지갑':
                # 목걸이지갑

                # 성별 선택
                gender_select = driver.find_element(
                    By.XPATH, '//*[@id="_prod-attr-section"]/div[2]/div/div[1]/div/ncp-product-attribute-table/div[2]/table/tbody[1]/tr[2]/td/div/div/div/div/div[1]')
                gender_select.click()

                gender_option =  driver.find_element(
                    By.XPATH, '//*[@id="_prod-attr-section"]/div[2]/div/div[1]/div/ncp-product-attribute-table/div[2]/table/tbody[1]/tr[2]/td/div/div/div/div/div[2]/div/div[contains(text(), "남녀공용")]')
                gender_option.click()


                # 주요소재 > 인조가죽 체크
                material = driver.find_element(
                    By.XPATH, '//*[@id="_prod-attr-section"]/div[2]/div/div[1]/div/ncp-product-attribute-table/div[2]/table/tbody[1]/tr[3]/td/div/div/div[1]/label[contains(text(), "인조가죽(합성피혁)")]')
                material.click()

                # 장식 > 장식없음 체크
                deck = driver.find_element(
                    By.XPATH, '//*[@id="_prod-attr-section"]/div[2]/div/div[1]/div/ncp-product-attribute-table/div[2]/table/tbody[2]/tr[2]/td/div/div/div[1]/label[contains(text(), "장식없음")]')
                deck.click()

                # 패턴 > 프린트 체크
                Patt = driver.find_element(
                    By.XPATH, '//*[@id="_prod-attr-section"]/div[2]/div/div[1]/div/ncp-product-attribute-table/div[2]/table/tbody[2]/tr[3]/td/div/div/div[1]/label[contains(text(), "프린트")]')
                Patt.click()

                # 제품특징 > 목걸이형 체크
                character = driver.find_element(
                    By.XPATH, '//*[@id="_prod-attr-section"]/div[2]/div/div[1]/div/ncp-product-attribute-table/div[2]/table/tbody[2]/tr[4]/td/div/div/div[1]/label[contains(text(), "목걸이형")]')
                character.click()

                # 잠금방식 > 오픈형
                lock_select = driver.find_element(
                    By.XPATH, '//*[@id="_prod-attr-section"]/div[2]/div/div[1]/div/ncp-product-attribute-table/div[2]/table/tbody[3]/tr[2]/td/div/div/div/div/div[1]')
                driver.execute_script(
                    "arguments[0].scrollIntoView(true);", lock_select)
                lock_select.click()

                lock_option = driver.find_element(
                    By.XPATH, '//*[@id="_prod-attr-section"]/div[2]/div/div[1]/div/ncp-product-attribute-table/div[2]/table/tbody[3]/tr[2]/td/div/div/div/div/div[2]/div/div[contains(text(), "오픈형")]')
                lock_option.click()

                print('목걸이지갑 주요 정보가 입력되었습니다.')

            elif last_word == '마우스패드':
                # 마우스패드/장패드

                # 재질 > 고무
                material_select = driver.find_element(
                    By.XPATH, '//*[@id="_prod-attr-section"]/div[2]/div/div[1]/div/ncp-product-attribute-table/div[2]/table/tbody/tr[2]/td/div/div/div/div/div[1]')
                material_select.click()

                # time.sleep(0.5)

                material_option = driver.find_element(
                    By.XPATH, '//*[@id="_prod-attr-section"]/div[2]/div/div[1]/div/ncp-product-attribute-table/div[2]/table/tbody/tr[2]/td/div/div/div/div/div[2]/div/div[@data-value="10030314"]')
                material_option.click()

                print('마우스패드 주요 정보가 입력되었습니다.')

            elif last_word == '캘린더/달력':
                # 탁상/벽걸이캘린더

                # 형태 > 벽걸이 or 탁상
                shape_select = driver.find_element(
                    By.XPATH, '//*[@id="_prod-attr-section"]/div[2]/div/div[1]/div/ncp-product-attribute-table/div[2]/table/tbody/tr[2]/td/div/div/div/div/div[1]')
                shape_select.click()

                if "탁상" in title:
                    shape = "109466"
                elif "벽걸이" in title:
                    shape = "11330"

                shape_option = driver.find_element(
                    By.XPATH, f'//*[@id="_prod-attr-section"]/div[2]/div/div[1]/div/ncp-product-attribute-table/div[2]/table/tbody/tr[2]/td/div/div/div/div/div[2]/div/div[@data-value="{shape}"]')
                shape_option.click()

                # 종류 > 연간
                type_select = driver.find_element(
                    By.XPATH, '//*[@id="_prod-attr-section"]/div[2]/div/div[1]/div/ncp-product-attribute-table/div[2]/table/tbody/tr[3]/td/div/div/div/div/div[1]')
                type_select.click()

                type_option = driver.find_element(
                    By.XPATH, '//*[@id="_prod-attr-section"]/div[2]/div/div[1]/div/ncp-product-attribute-table/div[2]/table/tbody/tr[3]/td/div/div/div/div/div[2]/div/div[contains(text(), "연간")]')
                type_option.click()

                print('탁상/벽걸이캘린더 주요 정보가 입력되었습니다.')

            elif last_word == '노트':
                # 레더커버 사철노트

                # 매수 > 50 ~ 100매 선택
                nop_select = driver.find_element(
                    By.XPATH, '//*[@id="_prod-attr-section"]/div[2]/div/div[1]/div/ncp-product-attribute-table/div[2]/table/tbody/tr[2]/td/div/div[1]/div/div/div[1]')
                nop_select.click()

                nop_option = driver.find_element(
                    By.XPATH, '//*[@id="_prod-attr-section"]/div[2]/div/div[1]/div/ncp-product-attribute-table/div[2]/table/tbody/tr[2]/td/div/div[1]/div/div/div[2]/div/div[@data-value="10459133"]')
                nop_option.click()

                # 매수입력 ex) 80
                nop_input = driver.find_element(
                    By.XPATH, '//*[@id="_prod-attr-section"]/div[2]/div/div[1]/div/ncp-product-attribute-table/div[2]/table/tbody/tr[2]/td/div/div[2]/div[1]/div/input')
                nop_input.send_keys('80')

                # 매수 > 단위
                measure_select = driver.find_element(
                    By.XPATH, '//*[@id="_prod-attr-section"]/div[2]/div/div[1]/div/ncp-product-attribute-table/div[2]/table/tbody/tr[2]/td/div/div[2]/div[2]/div/div[1]')
                measure_select.click()

                measure_option = driver.find_element(
                    By.XPATH, '//*[@id="_prod-attr-section"]/div[2]/div/div[1]/div/ncp-product-attribute-table/div[2]/table/tbody/tr[2]/td/div/div[2]/div[2]/div/div[2]/div/div[@data-value="A02109"]')
                measure_option.click()

                # 종류 > 제본노트 선택
                type_select = driver.find_element(
                    By.XPATH, '//*[@id="_prod-attr-section"]/div[2]/div/div[1]/div/ncp-product-attribute-table/div[2]/table/tbody/tr[3]/td/div/div/div/div/div[1]')
                type_select.click()

                type_option = driver.find_element(
                    By.XPATH, '//*[@id="_prod-attr-section"]/div[2]/div/div[1]/div/ncp-product-attribute-table/div[2]/table/tbody/tr[3]/td/div/div/div/div/div[2]/div/div[contains(text(), "제본노트")]')
                type_option.click()

                # 내지 > 라인(룰드), 백지(플레인) 체크

                # 형태 체크박스 클릭
                inner = ['라인(룰드)', '백지(플레인)']

                for label in inner:
                    check_box = driver.find_element(
                        By.XPATH, f'//*[@id="_prod-attr-section"]/div[2]/div/div[1]/div/ncp-product-attribute-table/div[2]/table/tbody/tr[4]/td/div/div/div[1]/label[contains(text(), "{label}")]')
                    check_box.click()

                # 종류 > 제본노트 선택
                cover_select = driver.find_element(
                    By.XPATH, '//*[@id="_prod-attr-section"]/div[2]/div/div[1]/div/ncp-product-attribute-table/div[2]/table/tbody/tr[5]/td/div/div/div/div/div[1]')
                cover_select.click()

                cover_option = driver.find_element(
                    By.XPATH, '//*[@id="_prod-attr-section"]/div[2]/div/div[1]/div/ncp-product-attribute-table/div[2]/table/tbody/tr[5]/td/div/div/div/div/div[2]/div/div[contains(text(), "가죽커버")]')
                cover_option.click()

                print('노트 주요 정보가 입력되었습니다.')

            elif last_word == '기타휴대폰액세서리':
                # 스마트톡

                # 기본사양 > 용도 > 스마트폰전용, 휴대폰용
                use = ['스마트폰전용', '휴대폰용']

                for label in use:
                    check_box = driver.find_element(
                        By.XPATH, f'//*[@id="_prod-attr-section"]/div[2]/div/div[1]/div/ncp-product-attribute-table/div[2]/table/tbody[1]/tr[2]/td/div/div/div[1]/label[contains(text(), "{label}")]')
                    check_box.click()

                print('스마트톡 주요 정보가 입력되었습니다.')

            elif last_word == '케이스/파우치':
                # 에어팟 / 버즈 케이스

                # 기본사양 > 용도 > 이어폰,마이트
                check_box = driver.find_element(
                    By.XPATH, '//*[@id="_prod-attr-section"]/div[2]/div/div[1]/div/ncp-product-attribute-table/div[2]/table/tbody[1]/tr[2]/td/div/div/div[1]/label[contains(text(), "이어폰,마이크")]')
                check_box.click()

                print('에어팟 / 버즈 케이스 주요 정보가 입력되었습니다.')

            elif last_word == '쿠션커버':
                # 쿠션커버

                # 형태
                shape_select = driver.find_element(
                    By.XPATH, '//*[@id="_prod-attr-section"]/div[2]/div/div[1]/div/ncp-product-attribute-table/div[2]/table/tbody/tr[2]/td/div/div/div/div/div[1]')
                shape_select.click()

                shape_option = driver.find_element(
                    By.XPATH, '//*[@id="_prod-attr-section"]/div[2]/div/div[1]/div/ncp-product-attribute-table/div[2]/table/tbody/tr[2]/td/div/div/div/div/div[2]/div/div[contains(text(), "사각형")]')
                shape_option.click()

                # 커버포함 여부 > 커버포함
                cover_select = driver.find_element(
                    By.XPATH, '//*[@id="_prod-attr-section"]/div[2]/div/div[1]/div/ncp-product-attribute-table/div[2]/table/tbody/tr[3]/td/div/div/div/div/div[1]')
                cover_select.click()

                cover_option = driver.find_element(
                    By.XPATH, '//*[@id="_prod-attr-section"]/div[2]/div/div[1]/div/ncp-product-attribute-table/div[2]/table/tbody/tr[3]/td/div/div/div/div/div[2]/div/div[contains(text(), "커버포함")]')
                cover_option.click()

                # 주요소재 > 폴리에스테르
                check_box = driver.find_element(
                    By.XPATH, '//*[@id="_prod-attr-section"]/div[2]/div/div[1]/div/ncp-product-attribute-table/div[2]/table/tbody/tr[4]/td/div/div/div[1]/label[contains(text(), "폴리에스테르")]')
                check_box.click()

                # 가로사이즈
                width_select = driver.find_element(
                    By.XPATH, '//*[@id="_prod-attr-section"]/div[2]/div/div[1]/div/ncp-product-attribute-table/div[2]/table/tbody/tr[5]/td/div/div[1]/div/div/div[1]')
                width_select.click()

                width_option = driver.find_element(
                    By.XPATH, '//*[@id="_prod-attr-section"]/div[2]/div/div[1]/div/ncp-product-attribute-table/div[2]/table/tbody/tr[5]/td/div/div[1]/div/div/div[2]/div/div[@data-value="10718535"]')
                driver.execute_script(
                    "arguments[0].scrollIntoView(true);", width_select)
                width_option.click()

                width_input = driver.find_element(
                    By.XPATH, '//*[@id="_prod-attr-section"]/div[2]/div/div[1]/div/ncp-product-attribute-table/div[2]/table/tbody/tr[5]/td/div/div[2]/div[1]/div/input')
                width_input.send_keys('45')

                # 세로사이즈
                height_select = driver.find_element(
                    By.XPATH, '//*[@id="_prod-attr-section"]/div[2]/div/div[1]/div/ncp-product-attribute-table/div[2]/table/tbody/tr[6]/td/div/div[1]/div/div/div[1]')
                height_select.click()

                height_option = driver.find_element(
                    By.XPATH, '//*[@id="_prod-attr-section"]/div[2]/div/div[1]/div/ncp-product-attribute-table/div[2]/table/tbody/tr[6]/td/div/div[1]/div/div/div[2]/div/div[@data-value="10718535"]')
                height_option.click()

                height_input = driver.find_element(
                    By.XPATH, '//*[@id="_prod-attr-section"]/div[2]/div/div[1]/div/ncp-product-attribute-table/div[2]/table/tbody/tr[6]/td/div/div[2]/div[1]/div/input')
                height_input.send_keys('45')

                print('쿠션커버 주요 정보가 입력되었습니다.')

            elif last_word == '무릎담요':
                # 무릎담요

                # 주요소재 체크
                material = ['기모', '폴리에스테르']

                for label in material:
                    check_box = driver.find_element(
                        By.XPATH, f'//*[@id="_prod-attr-section"]/div[2]/div/div[1]/div/ncp-product-attribute-table/div[2]/table/tbody[1]/tr[2]/td/div/div/div[1]/label[contains(text(), "{label}")]')
                    check_box.click()

                # 패턴 > 프린트
                patt_box = driver.find_element(
                    By.XPATH, '//*[@id="_prod-attr-section"]/div[2]/div/div[1]/div/ncp-product-attribute-table/div[2]/table/tbody[1]/tr[3]/td/div/div/div[1]/label[contains(text(), "프린트")]')
                patt_box.click()

                # 가로사이즈
                width_select = driver.find_element(
                    By.XPATH, '//*[@id="_prod-attr-section"]/div[2]/div/div[1]/div/ncp-product-attribute-table/div[2]/table/tbody[2]/tr[2]/td/div/div[1]/div/div/div[1]')
                driver.execute_script(
                    "arguments[0].scrollIntoView(true);", width_select)
                width_select.click()

                width_option = driver.find_element(
                    By.XPATH, '//*[@id="_prod-attr-section"]/div[2]/div/div[1]/div/ncp-product-attribute-table/div[2]/table/tbody[2]/tr[2]/td/div/div[1]/div/div/div[2]/div/div[@data-value="10708434"]')
                width_option.click()

                width_input = driver.find_element(
                    By.XPATH, '//*[@id="_prod-attr-section"]/div[2]/div/div[1]/div/ncp-product-attribute-table/div[2]/table/tbody[2]/tr[2]/td/div/div[2]/div[1]/div/input')
                width_input.send_keys('150')

                # 세로사이즈
                height_select = driver.find_element(
                    By.XPATH, '//*[@id="_prod-attr-section"]/div[2]/div/div[1]/div/ncp-product-attribute-table/div[2]/table/tbody[2]/tr[3]/td/div/div[1]/div/div/div[1]')
                height_select.click()

                height_option = driver.find_element(
                    By.XPATH, '//*[@id="_prod-attr-section"]/div[2]/div/div[1]/div/ncp-product-attribute-table/div[2]/table/tbody[2]/tr[3]/td/div/div[1]/div/div/div[2]/div/div[@data-value="10708434"]')
                height_option.click()

                height_input = driver.find_element(
                    By.XPATH, '//*[@id="_prod-attr-section"]/div[2]/div/div[1]/div/ncp-product-attribute-table/div[2]/table/tbody[2]/tr[3]/td/div/div[2]/div[1]/div/input')
                height_input.send_keys('100')

                print('무릎담요 주요 정보가 입력되었습니다.')

            elif last_word == '이불커버':
                # 여름이불

                # 주요소재 체크
                material = ['면', '폴리에스테르']

                for label in material:
                    check_box = driver.find_element(
                        By.XPATH, f'//*[@id="_prod-attr-section"]/div[2]/div/div[1]/div/ncp-product-attribute-table/div[2]/table/tbody[1]/tr[2]/td/div/div/div[1]/label[contains(text(), "{label}")]')
                    check_box.click()

                # 패턴 > 프린트
                patt_box = driver.find_element(
                    By.XPATH, '//*[@id="_prod-attr-section"]/div[2]/div/div[1]/div/ncp-product-attribute-table/div[2]/table/tbody[1]/tr[3]/td/div/div/div[1]/label[contains(text(), "기타")]')
                patt_box.click()

                # 가로사이즈
                width_select = driver.find_element(
                    By.XPATH, '//*[@id="_prod-attr-section"]/div[2]/div/div[1]/div/ncp-product-attribute-table/div[2]/table/tbody[2]/tr[2]/td/div/div[1]/div/div/div[1]')
                driver.execute_script(
                    "arguments[0].scrollIntoView(true);", width_select)
                width_select.click()

                width_option = driver.find_element(
                    By.XPATH, '//*[@id="_prod-attr-section"]/div[2]/div/div[1]/div/ncp-product-attribute-table/div[2]/table/tbody[2]/tr[2]/td/div/div[1]/div/div/div[2]/div/div[@data-value="10708435"]')
                width_option.click()

                width_input = driver.find_element(
                    By.XPATH, '//*[@id="_prod-attr-section"]/div[2]/div/div[1]/div/ncp-product-attribute-table/div[2]/table/tbody[2]/tr[2]/td/div/div[2]/div[1]/div/input')
                width_input.send_keys('158')

                # 세로사이즈
                height_select = driver.find_element(
                    By.XPATH, '//*[@id="_prod-attr-section"]/div[2]/div/div[1]/div/ncp-product-attribute-table/div[2]/table/tbody[2]/tr[3]/td/div/div[1]/div/div/div[1]')
                height_select.click()

                height_option = driver.find_element(
                    By.XPATH, '//*[@id="_prod-attr-section"]/div[2]/div/div[1]/div/ncp-product-attribute-table/div[2]/table/tbody[2]/tr[3]/td/div/div[1]/div/div/div[2]/div/div[@data-value="10708435"]')
                height_option.click()

                height_input = driver.find_element(
                    By.XPATH, '//*[@id="_prod-attr-section"]/div[2]/div/div[1]/div/ncp-product-attribute-table/div[2]/table/tbody[2]/tr[3]/td/div/div[2]/div[1]/div/input')
                height_input.send_keys('185')

                print('여름이불 주요 정보가 입력되었습니다.')

            elif last_word == '파우치':
                # 가죽 파우치

                # 주요소재 체크

                check_box = driver.find_element(
                    By.XPATH, '//*[@id="_prod-attr-section"]/div[2]/div/div[1]/div/ncp-product-attribute-table/div[2]/table/tbody[1]/tr[2]/td/div/div/div[1]/label[contains(text(), "인조가죽(합성피혁)")]')
                check_box.click()

                # 디자인 > 장식 > 장식없음
                deco_box = driver.find_element(
                    By.XPATH, '//*[@id="_prod-attr-section"]/div[2]/div/div[1]/div/ncp-product-attribute-table/div[2]/table/tbody[2]/tr[2]/td/div/div/div[1]/label[contains(text(), "장식없음")]')
                deco_box.click()

                # 디자인 > 패턴 > 프린트
                deco_box = driver.find_element(
                    By.XPATH, '//*[@id="_prod-attr-section"]/div[2]/div/div[1]/div/ncp-product-attribute-table/div[2]/table/tbody[2]/tr[3]/td/div/div/div[1]/label[contains(text(), "프린트")]')
                deco_box.click()

                print('여름이불 주요 정보가 입력되었습니다.')

            elif last_word == '알람/탁상시계':
                # 무프레임 벽시계

                # 상품속성 > 기타속송 > 종류 > 디자인시계
                type_select = driver.find_element(
                    By.XPATH, '//*[@id="_prod-attr-section"]/div[2]/div/div[1]/div/ncp-product-attribute-table/div[2]/table/tbody/tr[2]/td/div/div/div/div/div[1]')
                type_select.click()

                type_option = driver.find_element(
                    By.XPATH, '//*[@id="_prod-attr-section"]/div[2]/div/div[1]/div/ncp-product-attribute-table/div[2]/table/tbody/tr[2]/td/div/div/div/div/div[2]/div/div[@data-value="10773017"]')
                type_option.click()

                # 전원
                power_check = driver.find_element(
                    By.XPATH, '//*[@id="_prod-attr-section"]/div[2]/div/div[1]/div/ncp-product-attribute-table/div[2]/table/tbody/tr[3]/td/div/div/div[1]/label[contains(text(), "건전지식")]')
                power_check.click()

                # 부가기능
                sub_check = driver.find_element(
                    By.XPATH, '//*[@id="_prod-attr-section"]/div[2]/div/div[1]/div/ncp-product-attribute-table/div[2]/table/tbody/tr[4]/td/div/div/div[1]/label[contains(text(), "무소음")]')
                sub_check.click()

                print('벽시계 주요 정보가 입력되었습니다.')

            elif last_word == '니트/스웨터':
                # 맨투맨

                # 주요소재
                material_check = driver.find_element(
                    By.XPATH, '//*[@id="_prod-attr-section"]/div[2]/div/div[1]/div/ncp-product-attribute-table/div[2]/table/tbody[1]/tr[2]/td/div/div/div[1]/label[contains(text(), "면")]')
                material_check.click()

                # 총기장 > 기본/하프
                tot_length_select = driver.find_element(
                    By.XPATH, '//*[@id="_prod-attr-section"]/div[2]/div/div[1]/div/ncp-product-attribute-table/div[2]/table/tbody[1]/tr[3]/td/div/div/div/div/div[1]')
                tot_length_select.click()

                tot_length_option = driver.find_element(
                    By.XPATH, '//*[@id="_prod-attr-section"]/div[2]/div/div[1]/div/ncp-product-attribute-table/div[2]/table/tbody[1]/tr[3]/td/div/div/div/div/div[2]/div/div[@data-value="10904496"]')
                tot_length_option.click()

                # 소매기장 > 긴팔
                sleeve_length_select = driver.find_element(
                    By.XPATH, '//*[@id="_prod-attr-section"]/div[2]/div/div[1]/div/ncp-product-attribute-table/div[2]/table/tbody[1]/tr[4]/td/div/div/div/div/div[1]')
                sleeve_length_select.click()

                sleeve_length_option = driver.find_element(
                    By.XPATH, '//*[@id="_prod-attr-section"]/div[2]/div/div[1]/div/ncp-product-attribute-table/div[2]/table/tbody[1]/tr[4]/td/div/div/div/div/div[2]/div/div[@data-value="10574795"]')
                sleeve_length_option.click()

                # 핏 > 기본핏
                fit_select = driver.find_element(
                    By.XPATH, '//*[@id="_prod-attr-section"]/div[2]/div/div[1]/div/ncp-product-attribute-table/div[2]/table/tbody[1]/tr[5]/td/div/div/div/div/div[1]')
                fit_select.click()

                fit_option = driver.find_element(
                    By.XPATH, '//*[@id="_prod-attr-section"]/div[2]/div/div[1]/div/ncp-product-attribute-table/div[2]/table/tbody[1]/tr[5]/td/div/div/div/div/div[2]/div/div[@data-value="10040059"]')
                fit_option.click()

                # 패턴 > 무지
                patt_check = driver.find_element(
                    By.XPATH, '//*[@id="_prod-attr-section"]/div[2]/div/div[1]/div/ncp-product-attribute-table/div[2]/table/tbody[2]/tr[2]/td/div/div/div[1]/label[contains(text(), "무지")]')
                patt_check.click()

                print('니트/스웨터 주요 정보가 입력되었습니다.')

            elif last_word == '티셔츠':
                # 반팔 티셔츠

                # 주요소재
                material_check = driver.find_element(
                    By.XPATH, '//*[@id="_prod-attr-section"]/div[2]/div/div[1]/div/ncp-product-attribute-table/div[2]/table/tbody[1]/tr[2]/td/div/div/div[1]/label[contains(text(), "면")]')
                material_check.click()

                # 총기장 > 기본/하프
                tot_length_select = driver.find_element(
                    By.XPATH, '//*[@id="_prod-attr-section"]/div[2]/div/div[1]/div/ncp-product-attribute-table/div[2]/table/tbody[1]/tr[3]/td/div/div/div/div/div[1]')
                tot_length_select.click()

                tot_length_option = driver.find_element(
                    By.XPATH, '//*[@id="_prod-attr-section"]/div[2]/div/div[1]/div/ncp-product-attribute-table/div[2]/table/tbody[1]/tr[3]/td/div/div/div/div/div[2]/div/div[@data-value="10904496"]')
                tot_length_option.click()

                # 소매기장 > 반팔
                sleeve_length_select = driver.find_element(
                    By.XPATH, '//*[@id="_prod-attr-section"]/div[2]/div/div[1]/div/ncp-product-attribute-table/div[2]/table/tbody[1]/tr[4]/td/div/div/div/div/div[1]')
                sleeve_length_select.click()

                sleeve_length_option = driver.find_element(
                    By.XPATH, '//*[@id="_prod-attr-section"]/div[2]/div/div[1]/div/ncp-product-attribute-table/div[2]/table/tbody[1]/tr[4]/td/div/div/div/div/div[2]/div/div[@data-value="10574793"]')
                sleeve_length_option.click()

                # 네크라인 > 라운드넥
                neck_select = driver.find_element(
                    By.XPATH, '//*[@id="_prod-attr-section"]/div[2]/div/div[1]/div/ncp-product-attribute-table/div[2]/table/tbody[1]/tr[5]/td/div/div/div/div/div[1]')
                neck_select.click()

                neck_option = driver.find_element(
                    By.XPATH, '//*[@id="_prod-attr-section"]/div[2]/div/div[1]/div/ncp-product-attribute-table/div[2]/table/tbody[1]/tr[5]/td/div/div/div/div/div[2]/div/div[@data-value="10040041"]')
                neck_option.click()

                # 핏 > 기본핏
                fit_select = driver.find_element(
                    By.XPATH, '//*[@id="_prod-attr-section"]/div[2]/div/div[1]/div/ncp-product-attribute-table/div[2]/table/tbody[1]/tr[6]/td/div/div/div/div/div[1]')
                fit_select.click()

                fit_option = driver.find_element(
                    By.XPATH, '//*[@id="_prod-attr-section"]/div[2]/div/div[1]/div/ncp-product-attribute-table/div[2]/table/tbody[1]/tr[6]/td/div/div/div/div/div[2]/div/div[@data-value="10040059"]')
                fit_option.click()

                # 패턴 > 무지
                patt_check = driver.find_element(
                    By.XPATH, '//*[@id="_prod-attr-section"]/div[2]/div/div[1]/div/ncp-product-attribute-table/div[2]/table/tbody[2]/tr[2]/td/div/div/div[1]/label[contains(text(), "무지")]')
                patt_check.click()

                print('티셔츠 주요 정보가 입력되었습니다.')

            elif last_word == '일반쿠션':
                # 형태 > 직사각형
                shape_select = driver.find_element(
                    By.XPATH, '//*[@id="_prod-attr-section"]/div[2]/div/div[1]/div/ncp-product-attribute-table/div[2]/table/tbody/tr[2]/td/div/div/div/div/div[1]')
                shape_select.click()

                shape_option = driver.find_element(
                    By.XPATH, '//*[@id="_prod-attr-section"]/div[2]/div/div[1]/div/ncp-product-attribute-table/div[2]/table/tbody/tr[2]/td/div/div/div/div/div[2]/div/div[@data-value="10197422"]')
                shape_option.click()

                # 커버포함여부 > 커버포함
                cover_select = driver.find_element(
                    By.XPATH, '//*[@id="_prod-attr-section"]/div[2]/div/div[1]/div/ncp-product-attribute-table/div[2]/table/tbody/tr[3]/td/div/div/div/div/div[1]')
                cover_select.click()

                cover_option = driver.find_element(
                    By.XPATH, '//*[@id="_prod-attr-section"]/div[2]/div/div[1]/div/ncp-product-attribute-table/div[2]/table/tbody/tr[3]/td/div/div/div/div/div[2]/div/div[@data-value="10774173"]')
                cover_option.click()

                # 주요소재 > 폴리에스테르
                material_click = driver.find_element(
                    By.XPATH, '//*[@id="_prod-attr-section"]/div[2]/div/div[1]/div/ncp-product-attribute-table/div[2]/table/tbody/tr[4]/td/div/div/div[1]/label[contains(text(), "폴리에스테르")]')
                material_click.click()

                # 가로사이즈
                width_select = driver.find_element(
                    By.XPATH, '//*[@id="_prod-attr-section"]/div[2]/div/div[1]/div/ncp-product-attribute-table/div[2]/table/tbody/tr[5]/td/div/div[1]/div/div/div[1]')
                width_select.click()

                width_option = driver.find_element(
                    By.XPATH, '//*[@id="_prod-attr-section"]/div[2]/div/div[1]/div/ncp-product-attribute-table/div[2]/table/tbody/tr[5]/td/div/div[1]/div/div/div[2]/div/div[@data-value="10718536"]')
                driver.execute_script(
                    "arguments[0].scrollIntoView(true);", width_select)
                width_option.click()

                width_input = driver.find_element(
                    By.XPATH, '//*[@id="_prod-attr-section"]/div[2]/div/div[1]/div/ncp-product-attribute-table/div[2]/table/tbody/tr[5]/td/div/div[2]/div[1]/div/input')
                width_input.send_keys('60')

                # 세로사이즈
                height_select = driver.find_element(
                    By.XPATH, '//*[@id="_prod-attr-section"]/div[2]/div/div[1]/div/ncp-product-attribute-table/div[2]/table/tbody/tr[6]/td/div/div[1]/div/div/div[1]')
                height_select.click()

                height_option = driver.find_element(
                    By.XPATH, '//*[@id="_prod-attr-section"]/div[2]/div/div[1]/div/ncp-product-attribute-table/div[2]/table/tbody/tr[6]/td/div/div[1]/div/div/div[2]/div/div[@data-value="10718534"]')
                height_option.click()

                height_input = driver.find_element(
                    By.XPATH, '//*[@id="_prod-attr-section"]/div[2]/div/div[1]/div/ncp-product-attribute-table/div[2]/table/tbody/tr[6]/td/div/div[2]/div[1]/div/input')
                height_input.send_keys('40')

                print('쿠션커버 주요 정보가 입력되었습니다.')


        # else:
        #     print('어떠한 주요정보도 입력되지 않았습니다.')

        #  --------------------------------   태그   --------------------------------

        tags = str(ct.iloc[0]['태그'])

        # 캘린더 태그
        if last_word == "캘린더/달력":
            if "탁상" in title:
                tags += ", 탁상달력, 탁상캘린더, 책상달력"
            elif "벽걸이" in title:
                tags += ", 벽걸이달력, 벽걸이캘린더"

        # 검색설정 태그 입력
        elem = driver.find_element(
            By.XPATH, '//*[@id="productForm"]/ng-include/ui-view[22]/div/div[1]/div/div/a')
        driver.execute_script("arguments[0].scrollIntoView(true);", elem)
        elem.click()

        # 태그 직접 입력 선택
        elem = driver.find_element(
            By.XPATH, '//*[@id="productForm"]/ng-include/ui-view[22]/div/div[2]/div/div[1]/div/div[3]/div/div/label')
        elem.click()

        # 태그 입력
        elem = driver.find_element(
            By.XPATH, '//*[@id="productForm"]/ng-include/ui-view[22]/div/div[2]/div/div[1]/div/div[3]/div[2]/div/div/div[1]/input')

        elem.send_keys(tags)
        time.sleep(1)

        elem.send_keys(Keys.RETURN)
        time.sleep(1)

        print(tags)

        #  -------------------------------- 노출 채널  --------------------------------
        # 노출채널로 스크롤 이동
        exposureChannel = driver.find_element(
            By.XPATH, '//*[@id="productForm"]/ng-include/ui-view[@name="exposureChannel"]')
        driver.execute_script(
            "arguments[0].scrollIntoView(true);", exposureChannel)

        # 쇼핑윈도 선택
        elem = WebDriverWait(driver, 5).until(EC.presence_of_element_located(
            (By.XPATH, '//*[@id="productForm"]/ng-include/ui-view[@name="exposureChannel"]/div/div[2]/div/div[1]/div/div[1]/div[2]/div/label')))
        # elem = driver.find_element(
        #     By.XPATH, '//*[@id="productForm"]/ng-include/ui-view[22]/div/div[2]/div/div[1]/div/div[1]/div[2]/div/label')
        elem.click()
        time.sleep(0.5)  # 쇼핑윈도를 선택하자마자 버튼을 누르게 되면 상위 div로 잡혀서 에러가 나온다

        # 스마트 스토어 선택해제
        elem = WebDriverWait(driver, 5).until(EC.presence_of_element_located(
            (By.XPATH, '//*[@id="productForm"]/ng-include/ui-view[@name="exposureChannel"]/div/div[2]/div/div[1]/div/div[1]/div[1]/div/label')))
        # elem = driver.find_element(
        #     By.XPATH, '//*[@id="productForm"]/ng-include/ui-view[22]/div/div[2]/div/div[1]/div/div[1]/div[1]/div/label')
        elem.click()

        #  --------------------------------  저장   --------------------------------
        # save = driver.find_element(
        #     By.CSS_SELECTOR, '#seller-content > ui-view > div.pc-fixed-area.navbar-fixed-bottom.hidden-xs > div.btn-toolbar.pull-right > div:nth-child(1) > button.btn.btn-primary.progress-button.progress-button-dir-horizontal.progress-button-style-top-line')
        # save.click()

        # # 저장완료 배너
        # # elem = driver.find_element(
        # #     By.XPATH, '//div[@class="seller-btn-area"]/button[contains(text(), "상품관리")]')
        # elem = WebDriverWait(driver, 10).until(EC.presence_of_element_located(
        #     (By.XPATH, '//div[@class="modal-dialog modal-xs"]/div//div[@class="modal-footer"]/div[@class="seller-btn-area"]/button[contains(text(), "상품관리")]')))
        # elem.click()

        #  --------------------------------  취소   --------------------------------
        elem = driver.find_element(
            By.CSS_SELECTOR, '#seller-content > ui-view > div.pc-fixed-area.navbar-fixed-bottom.hidden-xs > div.btn-toolbar.pull-right > div:nth-child(2) > button')
        elem.click()

        WebDriverWait(driver, 3).until(EC.alert_is_present())
        alt = driver.switch_to.alert
        alt.accept()
        
except:
    print("업로드를 실패했습니다.")
    print(f"row번호: {r + 2}, 상품번호: {product_no}")
    driver.save_screenshot('err_screen.png')

    tt = str(title)

    f = open("실패.txt", 'w')
    f.write(f"row번호: {r + 2}, 상품번호: {product_no}, 상품명: {tt}")
    f.close()

    driver.quit()
