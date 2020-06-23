
> While I am doing an internship at SAP..


# KISLINE CRAWLER & Daily Alert Program

![img1](static/kisline.png)

인턴하면서 만든 각종 프로그램

- KISLINE CRAWLER
- Daily Alert Email


> **KISLINE**이란?

나이스 평가정보에서 운영하는 기업정보조회 서비스. 550만개의 기업정보와 1165개 산업분류코드 체계의 산업분석 정보, 국내 7만개 업체의 특허정보를 서비스하고 있다.


### Set Environment (환경 설정)

1. *Python 3.7 이상 & 아나콘다 설치 필요*
    - 참고 : <https://www.anaconda.com/products/individual>
2. **ChromeDriver** 및 `selenium` 설치 필요
    - 커맨드창에 아래 명령어 입력!
    - `pip install selenium`
    - Chrome Driver Download Url : <https://chromedriver.chromium.org/downloads>
    - 운영체제에 맞는 Driver 선택.
3. Windows64bit 환경
4. *python notion* 라이브러리 설치 필요.
    - 커맨드창에 아래 명령어 입력
    - `pip install notion`



## 기능 소개

### KISLINE CRAWLER

- 주요 파일. 
- 불러오는 기업 정보 : 기업명/기업영문명/주소/우편번호/영위하는 사업/산업분류코드/사업자등록번호/Contact Person 정보/
- 3가지 모드
    1. Normal Mode : 기본 모드.
    2. Massive Mode : `asset/static/massive.xlsx` 파일에서 search_list를 불러와서 한번에 massive search 실행.
- 크롤링 결과는 Excel 파일로 저장.

### How To Run
```bash
$ python dcr.py
```


### Crwawling News  & Upload it to Notion

매일경제, 더벨, 서울경제 사이트에서 아래 키워드 검색 결과를 엑셀로 저장하고 notion 사이트에 업로드

 > Keyword : IPO, 투자 유치, 해외 진출, 급성장, 선도기업, M&A, 인수합병, 스핀오프 등

- 개인 notion 계정 필요.
- `pip install notion` 필요. (notion 입출력을 제공하는 파이썬 라이브러리)
- tokenV2, Key 확인 필요.
- 크롤링 결과는 동일 폴더 excel 파일에도 저장된다.


### How To Run
```bash
$ python Daily_News_Crawling_with_notion.py
```



### Email Auto Generation

Daily News Mornitoring 메일을 데일리로 report 하기 위해 만들었던 프로그램

- **Windows** 환경 및 **Outlook** 설치 필수.
- 파이썬 라이브러리 `win32` 라이브러리 사용 필요. 
    - ex. `import win32.client`
- 매일 요악정리한 내용은 `asset\static\daily_search_list.xlsx` 파일에 입력.

### How To Run

```bash
$ python daily_email.py
```

---

### 결과 화면

1. 실행 결과

![img2](static/capture.PNG)


2. 저장되는 파일 결과

![img2](static/capture2.PNG)

