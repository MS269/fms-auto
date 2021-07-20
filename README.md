# fms-auto

기부물품관리시스템 인수확인증 접수등록 자동화

> <strong>이 프로그램을 사용함으로 인해 발생한 문제는 사용한 당사자에게 있습니다.</strong>

1. <a href="https://github.com/MS269/fms-auto#chromedriver.eve-다운로드">chromedriver.eve 다운로드</a>

2. <a href="https://github.com/MS269/fms-auto#사용법">사용법</a>

3. <a href="https://github.com/MS269/fms-auto#기부업체-수정법">기부업체 수정법</a>

---

## chromedriver.eve 다운로드

크롬 버전에 따라 다운로드 >
https://chromedriver.chromium.org/downloads

> 유의: main.py 와 같은 폴더에 있어야함!

---

## 사용법

1.  excel/receipts.xlsx 에 리스트 작성

    EX)
    |업체|날짜|수량|가격|
    |:--:|:--:|:----:|:--:|
    |a1|0101|10|10000|
    |b10|0111|100|100000|
    |c1|1101|1000|1000000|
    |...|...|...|...|

    > 유의: 업체 꼭 소문자로!

    > 유의: 날짜 꼭 4자리수로!

2.  main.py 실행

> 유의: chromedriver.exe 가 main.py 와 같은 폴더에 있어야함!

3. 공동인증서 로그인

4. ESC 입력

> 유의: 어떤 공지도 떠 있어선 안됨!

---

## 기부업체 수정법

- excel/shops.xlsx 에 추가

  EX)
  |A|B|C|
  |:-:|:-:|:-:|
  |뚜레쥬르 김포구래점||샹제르망|
  |...|...|...|
  |더원빵카페|||

- 예외인 경우에 excel/exceptions.xlsx 에 추가

  EX)
  |업체|음식 (4: 도넛츠, 5: 냉동떡, 6: 포장반찬, 7: 족발류)|
  |:-:|:-:|
  |a4|5|
  |b16|5|
  |c7|7|
  |...|...|
