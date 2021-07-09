import datetime
import random
import keyboard
import pandas
import pyautogui
import time
from selenium import webdriver
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.keys import Keys

# 버퍼링
TIME_SLEEP = 2

# 액셀 읽기
shops_df = pandas.read_excel(
    "excel/shops.xlsx", names=["a", "b", "c"], dtype={"a": str, "b": str, "c": str})
list_df = pandas.read_excel("excel/list.xlsx",
                            names=["shop", "date", "amount", "cost"],
                            dtype={"shop": str, "date": str, "amount": str, "cost": str})

# 액셀 데이터 처리
shops = shops_df.to_dict()
lists = list_df.values.tolist()

# 음식
foods = (("우유식빵", 0), ("슈크림빵", 1), ("단팥빵", 5), ("소보루빵", 2),
         ("도넛츠", 6), ("냉동떡", 0), ("포장반찬", 0), ("족발류", 0))

# NFMS 실행
driver = webdriver.Chrome()
url = "https://nfms.foodbank1377.org/"
driver.get(url)
driver.maximize_window()

# 로그인 대기 (ESC 입력시 실행)
while True:
    if keyboard.is_pressed("esc"):
        break

# 기부물품관리 클릭
driver.find_element_by_xpath(
    "/html/body/div[1]/div/div/div[1]/div/div/div/div[10]").click()

# 접수등록 클릭
driver.find_element_by_xpath(
    "/html/body/div[1]/div/div/div[2]/div/div[1]/div/div/div/div[4]/div[1]/div[1]/div[1]/div[2]/div[1]/div/div[3]/div/div/div/div[1]/div/div/div[4]").click()
time.sleep(TIME_SLEEP)

# 카운터
count = 0

# 등록 시작
for i in lists:
    count += 1

    # 데이터 읽기
    shop, date, amount, cost = map(lambda r: r, i)

    # 기부처
    shop_keyword = shops[shop[0]][int(shop[1:]) - 1]

    # 기부날짜
    date_keyword = datetime.datetime(
        2021, int(date[0:2]), int(date[2:4])).strftime("%Y%m%d")

    # 유통기한
    exp_date_keyword = (datetime.datetime(2021, int(date[0:2]), int(
        date[2:4])) + datetime.timedelta(hours=24*30)).strftime("%Y%m%d")

    # 음식 종류 랜덤 (0: 우유식빵, 1: 슈크림빵, 2: 단팥빵, 3: 소보루빵)
    food_idx = random.randrange(0, 4)

    # 음식 종류 예외 처리 (4: 도넛츠, 5: 냉동떡, 6: 포장반찬, 7: 족발류)
    if shop == "a6":
        food_idx = 4
    elif shop == "a4":
        food_idx = 5
    elif shop == "a8":
        food_idx = 5
    elif shop == "b16":
        food_idx = 5
    elif shop == "b18":
        food_idx = 5
    elif shop == "c12":
        food_idx = 5
    elif shop == "a10":
        food_idx = 6
    elif shop == "c7":
        food_idx = 7

    # 음식
    food_keyword = foods[food_idx][0]
    food_check = 2 + foods[food_idx][1]

    # 로그
    print(f"{count}: {shop_keyword}, {date_keyword}, {food_keyword} {amount}개, {cost}원")

    # 등록 클릭
    driver.find_element_by_xpath(
        "/html/body/div[1]/div/div/div[2]/div/div[2]/div/div[2]/div/div/div/div/div[1]/div[1]/div[3]/div[32]/div/div[12]").click()
    time.sleep(TIME_SLEEP)

    # 기부일자 입력
    driver.find_element_by_xpath(
        "/html/body/div[1]/div/div/div[2]/div/div[2]/div/div[2]/div/div/div/div/div[1]/div[1]/div[3]/div[53]/div[3]/div[8]").click()
    pyautogui.hotkey("ctrl", "a")
    ActionChains(driver).send_keys(date_keyword).perform()

    # 기부자 클릭
    driver.find_element_by_xpath(
        "/html/body/div[1]/div/div/div[2]/div/div[2]/div/div[2]/div/div/div/div/div[1]/div[1]/div[3]/div[53]/div[3]/div[15]").click()
    time.sleep(TIME_SLEEP)

    # 기부자 입력 및 조회
    driver.find_element_by_xpath(
        "/html/body/div[2]/div/div[1]/div/div/div[1]/div[18]").click()
    ActionChains(driver).send_keys(shop_keyword).send_keys(Keys.F2).perform()
    time.sleep(TIME_SLEEP)

    # 기부자 확인
    driver.find_element_by_xpath(
        "/html/body/div[2]/div/div[1]/div/div/div[1]/div[24]").click()
    time.sleep(TIME_SLEEP)

    # 목록 추가
    driver.find_element_by_xpath(
        "/html/body/div[1]/div/div/div[2]/div/div[2]/div/div[2]/div/div/div/div/div[1]/div[1]/div[3]/div[53]/div[3]/div[28]/div/div[1]/div/div").click()
    time.sleep(TIME_SLEEP)

    # 기부물품 입력 및 조회
    driver.find_element_by_xpath(
        "/html/body/div[2]/div/div[1]/div/div/div/div[17]").click()
    ActionChains(driver).send_keys(food_keyword).send_keys(Keys.F2).perform()
    time.sleep(TIME_SLEEP)

    # 기부물품 선택
    driver.find_element_by_xpath(
        f"/html/body/div[2]/div/div[1]/div/div/div/div[9]/div[1]/div[2]/div[1]/div/div[{food_check}]/div[2]/div/div[2]/div/div").click()

    # 기부물품 확인
    driver.find_element_by_xpath(
        "/html/body/div[2]/div/div[1]/div/div/div/div[14]/div/div[2]").click()
    time.sleep(TIME_SLEEP)

    # 수량 입력
    ActionChains(driver).send_keys(amount).send_keys(Keys.TAB).send_keys(
        Keys.TAB).send_keys(Keys.TAB).send_keys(Keys.TAB).perform()

    # 유통기한 입력
    ActionChains(driver).send_keys(
        exp_date_keyword).send_keys(Keys.TAB).perform()

    # 금액 입력
    ActionChains(driver).send_keys(cost).perform()

    # 저장 클릭
    driver.find_element_by_xpath(
        "/html/body/div[1]/div/div/div[2]/div/div[2]/div/div[2]/div/div/div/div/div[1]/div[1]/div[3]/div[53]/div[3]/div[10]/div/div[6]").click()
    time.sleep(TIME_SLEEP)

    # 저장 확인
    pyautogui.press("enter")
    time.sleep(TIME_SLEEP)
    pyautogui.press("enter")
    time.sleep(TIME_SLEEP)

exit()
