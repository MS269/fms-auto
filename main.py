import datetime
import keyboard
import pandas
import pyautogui
import pyperclip
import time

# TODO: 시작 단축키, 멈추는 단축키
# TODO: 빵 선택 마우스 위치 수정

# 버퍼링
TIME_SLEEP = 1

# 엑셀 열 크기
N = 29
M = 1

# 매크로 마우스 좌표
F = 8
FOOD = ((535, 520),
        (535, 550),
        (535, 670),
        (535, 580),
        (535, 610),
        (535, 520),
        (535, 520),
        (535, 520))
POS = ((1340, 340),
       (520, 435),
       (885, 430),
       (1225, 315),
       (925, 905),
       (1850, 500),
       (1000, 430),
       (940, 825),
       (800, 600),
       (1190, 590),
       (1660, 595),
       (1850, 345))

# 음식 엑셀 & 업체 엑셀
shops_df = pandas.read_excel("shop_list.xlsx")
list_df = pandas.read_excel("list.xlsx",
                            names=["shop", "date", "amount", "cost", "food"],
                            dtype={"shop": str, "date": str, "amount": str, "cost": str, "food": str})

# 음식 리스트
food_list = ("우유식빵", "슈크림빵", "단팥빵", "소보루빵", "도넛츠", "냉동떡", "포장반찬", "족발류")

# 업체 리스트
shop_list = {"a": [None], "b": [None], "c": [None]}
for i in range(0, N):
    shop_list["a"].append(shops_df.iloc[i][0])
    shop_list["b"].append(shops_df.iloc[i][1])
    shop_list["c"].append(shops_df.iloc[i][2])

# 매크로 시작
print("-------------------- 매크로 --------------------")
for i in range(0, M):
    data = list_df.iloc[i]
    shop, date, amount, cost = map(lambda r: r, data)

    # 로그
    print(
        f"---------- {i+2}: {shop} | {date} | {amount} | {cost} ----------")

    # 데이터 가공
    shop_keyword = shop_list[shop[0]][int(shop[1:])]
    date_keyword = datetime.datetime(
        2021, int(date[0:2]), int(date[2:4])).strftime("%Y%m%d")
    food_keyword = food_list[i % 4]
    expiration_date_keyword = (datetime.datetime(2021, int(date[0:2]), int(
        date[2:4])) + datetime.timedelta(hours=24*30)).strftime("%Y%m%d")

    # 음식 종류 예외 (4: 도넛츠, 5: 냉동떡, 6: 포장반찬, 7: 족발류)
    if shop == "a6":
        food_keyword = food_list[4]
    elif shop == "a4":
        food_keyword = food_list[5]
    elif shop == "a8":
        food_keyword = food_list[5]
    elif shop == "b16":
        food_keyword = food_list[5]
    elif shop == "b18":
        food_keyword = food_list[5]
    elif shop == "c12":
        food_keyword = food_list[5]
    elif shop == "a10":
        food_keyword = food_list[6]
    elif shop == "c7":
        food_keyword = food_list[7]

        # 로그
    print(
        f"- {shop_keyword} | {date_keyword} | {food_keyword} | {expiration_date_keyword} -")

    # 등록
    pyautogui.click(POS[0][0], POS[0][1])
    time.sleep(TIME_SLEEP)

    # 기부일자
    pyautogui.click(POS[1][0], POS[1][1])
    pyautogui.hotkey("ctrl", "a")
    pyautogui.press("backspace")
    pyperclip.copy(date_keyword)
    pyautogui.hotkey("ctrl", "v")

    # 기부자 검색
    pyautogui.click(POS[2][0], POS[2][1])
    time.sleep(TIME_SLEEP)
    pyautogui.click(POS[3][0], POS[3][1])
    pyperclip.copy(shop_keyword)
    pyautogui.hotkey("ctrl", "v")
    pyautogui.press("f2")
    time.sleep(TIME_SLEEP)
    pyautogui.click(POS[4][0], POS[4][1])

    # 추가
    pyautogui.click(POS[5][0], POS[5][1])
    time.sleep(TIME_SLEEP)
    pyautogui.click(POS[6][0], POS[6][1])
    pyperclip.copy(food_keyword)
    pyautogui.hotkey("ctrl", "v")
    pyautogui.press("f2")
    time.sleep(TIME_SLEEP)

    # 음식 종류 선택
    for i in range(0, F):
        if food_keyword == food_list[F]:
            pyautogui.click(FOOD[F][0], FOOD[F][1])

    pyautogui.click(POS[7][0], POS[7][1])

    # 목록 작성
    pyautogui.click(POS[8][0], POS[8][1])
    pyperclip.copy(amount)
    pyautogui.hotkey("ctrl", "v")
    pyautogui.click(POS[9][0], POS[9][1])
    pyperclip.copy(expiration_date_keyword)
    pyautogui.hotkey("ctrl", "v")
    pyautogui.click(POS[10][0], POS[10][1])
    pyperclip.copy(cost)
    pyautogui.hotkey("ctrl", "v")
    pyautogui.click(POS[11][0], POS[11][1])

    # 엔터 대기
    while True:
        if keyboard.is_pressed("enter"):
            break
    pyautogui.press("enter")
    time.sleep(TIME_SLEEP)
