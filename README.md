# fms-auto

기부물품관리시스템 접수등록 자동화

1. <a href="https://github.com/MS269/fms-auto#chromedriver.eve-다운">chromedriver.eve 다운</a>

2. <a href="https://github.com/MS269/fms-auto#사용법">사용법</a>

3. <a href="https://github.com/MS269/fms-auto#기부업체-수정법">기부업체 수정법</a>

---

## chromedriver.eve 다운

크롬 버전에 따라 다운 >
https://chromedriver.chromium.org/downloads

---

## 사용법

1.  excel/list.xlsx 에 리스트 작성

    EX)
    |shop|date|amount|cost|
    |:--:|:--:|:----:|:--:|
    |a1|0101|10|10000|
    |b10|0111|100|100000|
    |c1|1101|100|100000

    > 유의: 업체 꼭 소문자로!

    > 유의: 날짜 꼭 4자리수로!

2.  main.py 실행

> 유의: chromedriver.exe 다운로드해야함

---

## 기부업체 수정법

- excel/shop_list.xlsx 에 추가

  EX)
  |a|b|c|
  |:-:|:-:|:-:|
  |뚜레쥬르 김포구래점||샹제르망|
  |...|...|...|
  |더원빵카페|||

- 예외인 경우에 main.py 에 추가

  EX)

  ```py
  # 음식 종류 예외 처리 (4: 도넛츠, 5: 냉동떡, 6: 포장반찬, 7:발류)
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
  ```
