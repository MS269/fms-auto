import datetime
import keyboard
import pandas
from selenium import webdriver
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.alert import Alert
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait

# 액셀 읽기
receipts_df = pandas.read_excel("excel/receipts.xlsx",
                                names=["shop", "date", "amount", "cost"],
                                dtype={"shop": str, "date": str, "amount": int, "cost": int})

# 액셀 데이터 처리
receipts = receipts_df.values.tolist()

# 크롬 드라이버 실행
driver = webdriver.Chrome()
url = "https://nfms.foodbank1377.org/"
driver.get(url)
driver.implicitly_wait(5)
driver.maximize_window()

# 로그인 대기 (ESC 입력시 실행)
while True:
    if keyboard.is_pressed("esc"):
        break

# 기부물품관리 클릭
driver.find_element_by_xpath(
    "/html/body/div[1]/div/div/div[1]/div/div/div/div[11]").click()

# 접수등록 클릭
driver.find_element_by_xpath(
    "/html/body/div[1]/div/div/div[2]/div/div[1]/div/div/div/div[4]/div[1]/div[1]/div[1]/div[2]/div[1]/div/div[3]/div/div/div/div[1]/div/div/div[4]").click()

# 카운터
count = 0

# 등록 시작
for receipt in receipts:
    count += 1

    # 데이터 읽기
    shop, date, amount, cost = map(lambda ret: ret, receipt)

    # 기부처
    shop_keyword = shop

    # 기부날짜
    date_keyword = datetime.datetime(
        2021, int(date[0:2]), int(date[2:4])).strftime("%Y%m%d")

    # 유통기한
    exp_date_keyword = (datetime.datetime(2021, int(date[0:2]), int(
        date[2:4])) + datetime.timedelta(days=90)).strftime("%Y%m%d")

    # 음식
    food_keyword = "돼지고기/kg"
    food_check = 3

    # 로그
    print(f"{count}: {shop_keyword}, {date_keyword}, {exp_date_keyword}, {amount}kg, {cost}원")

    # 등록 클릭
    driver.find_element_by_xpath(
        "/html/body/div[1]/div/div/div[2]/div/div[2]/div/div[2]/div/div/div/div/div[1]/div[1]/div[3]/div[32]/div/div[12]").click()

    # 기부일자 입력
    driver.find_element_by_xpath(
        "/html/body/div[1]/div/div/div[2]/div/div[2]/div/div[2]/div/div/div/div/div[1]/div[1]/div[3]/div[53]/div[3]/div[8]").click()
    ActionChains(driver).key_down(Keys.CONTROL).send_keys(
        "a").key_up(Keys.CONTROL).send_keys(date_keyword).perform()

    # 기부자 클릭
    driver.find_element_by_xpath(
        "/html/body/div[1]/div/div/div[2]/div/div[2]/div/div[2]/div/div/div/div/div[1]/div[1]/div[3]/div[53]/div[3]/div[15]").click()

    # 기부자 입력 및 조회
    driver.find_element_by_xpath(
        "/html/body/div[2]/div/div[1]/div/div/div[1]/div[18]").click()
    ActionChains(driver).send_keys(shop_keyword).send_keys(Keys.F2).perform()

    # 기부자 확인
    driver.find_element_by_xpath(
        "/html/body/div[2]/div/div[1]/div/div/div[1]/div[24]").click()

    # 목록 추가
    driver.find_element_by_xpath(
        "/html/body/div[1]/div/div/div[2]/div/div[2]/div/div[2]/div/div/div/div/div[1]/div[1]/div[3]/div[53]/div[3]/div[28]/div/div[1]/div/div").click()

    # 기부물품 입력 및 조회
    driver.find_element_by_xpath(
        "/html/body/div[2]/div/div[1]/div/div/div/div[17]").click()
    ActionChains(driver).send_keys(food_keyword).send_keys(Keys.F2).perform()

    # 기부물품 선택
    driver.find_element_by_xpath(
        f"/html/body/div[2]/div/div[1]/div/div/div/div[9]/div[1]/div[2]/div[1]/div/div[{food_check}]/div[2]/div/div[2]/div/div").click()

    # 기부물품 확인
    driver.find_element_by_xpath(
        "/html/body/div[2]/div/div[1]/div/div/div/div[14]/div/div[2]").click()

    # 수량 입력
    ActionChains(driver).send_keys(amount).send_keys(Keys.TAB).send_keys(
        Keys.TAB).send_keys(Keys.TAB).send_keys(Keys.TAB).perform()

    # 유통기한 입력
    ActionChains(driver).key_down(Keys.CONTROL).send_keys("a").key_up(
        Keys.CONTROL).send_keys(exp_date_keyword).send_keys(Keys.TAB).perform()

    # 금액 입력
    ActionChains(driver).send_keys(cost).perform()

    # 저장 클릭
    driver.find_element_by_xpath(
        "/html/body/div[1]/div/div/div[2]/div/div[2]/div/div[2]/div/div/div/div/div[1]/div[1]/div[3]/div[53]/div[3]/div[10]/div/div[6]").click()

    # 저장 확인
    Alert(driver).accept()
    WebDriverWait(driver, 5).until(EC.alert_is_present())
    Alert(driver).accept()
