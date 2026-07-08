# 📊 Smart Review Crawling (v2)

틴트 예약 리스트(엑셀)를 넣으면, **API(개발자 모드)** 로 각 채널 리뷰를 긁어 예약번호로 매칭하고 지사·날짜·가이드 단위로 리뷰율/평점을 집계하는 데스크톱 프로그램입니다.

기존 `Review Search`(Selenium 클릭 방식)의 후속 버전이며, 클릭 대신 **로그인된 크롬 세션으로 각 채널 API를 직접 호출**(캡처-리플레이 방식)합니다.

![Python](https://img.shields.io/badge/Python-3.8+-blue.svg)
![Selenium](https://img.shields.io/badge/Library-Selenium-brightgreen.svg)
![Platform](https://img.shields.io/badge/Platform-Windows%20%7C%20Mac-lightgrey.svg)

---

## 🎛️ 실행 모드 (4가지)

상단 토글에서 선택하며, 아래에 선택한 모드의 **날짜 기준**이 표시됩니다.

| 모드 | 기간(용도) | 날짜 기준 | 수집 별점 | No Show | 스페셜 카테고리 | 출력물 |
|---|---|---|---|---|---|---|
| **PFP** (성과제) | 1주일 | **참여일** | 1~5점 전부 | 제외 | 제외 | 폴더트리 + zip + PDF, 화면 리포트 |
| **Monthly** | 1달 | **리뷰 작성일** | 1~5점 전부 | 포함 | 제외 | 단일 엑셀(지사 시트 + Guide) |
| **Quarterly** | 3달 | **리뷰 작성일** | 1~5점 전부 | 포함 | 제외 | 단일 엑셀(지사 시트 + Guide) |
| **New Year Party** | 1년 | **리뷰 작성일** | 1~5점 전부 | 포함 | 제외 | 단일 엑셀(지사 시트 + Guide) |

- **좋은 리뷰 = 4~5점, 나쁜 리뷰 = 1~3점.** 세 모드(Monthly/Quarterly/NYP)는 1~5점을 전부 수집한 뒤 Guide 시트에서 좋은/나쁜/종합으로 나눠 집계합니다.
- 스페셜 카테고리 = `MBC 스튜디오 / Dr. Petit / 마리엠헤어` (상품명 부분일치). 전 모드 제외.
- **Team/Tour Count(분모)는 리뷰 수집 가능한 5개 채널(L·KK·GG·TPC·MRT) 예약만** 셉니다.

---

## 🔗 조회 채널 (5개) & 날짜 필드

| 코드 | 채널 | 예약번호 필드 | 참여일 기준(PFP) | 리뷰작성일 기준(Monthly/Q/NYP) |
|---|---|---|---|---|
| **L** | Klook | `booking_no` | `date_type=ParticipantTime` | `date_type=ReviewTime` |
| **KK** | KKday | `orderMid` | `begGoDate/endGoDate` | `begRecDate/endRecDate` |
| **GG** | GetYourGuide | `bookingReference` | `travelDateFrom/To` (±1일) | `reviewDateFrom/To` |
| **TPC** | Trip.com/Ctrip | `orderId` | `orderInfo.departureTime` | `commentTime` |
| **MRT** | MyRealTrip | `reservationNo` | `travelStartDate` | `createdAt` |

- L·KK·GG는 **서버에서 날짜 필터**, TPC·MRT는 **클라이언트에서 필터**(전량/작성일 정렬 스캔)합니다.
- TPC 익명 리뷰(`orderId=0`)는 예약번호가 없어 매칭에서 제외됩니다.
- VI(Viator)·CV(Civitatis)·CRE(Creatrip)·VE(Veltra)·HO(Headout) 등은 리뷰에 예약번호가 없어 자동 매칭 대상이 아닙니다.

---

## 🚀 빠른 시작

### 1️⃣ 크롬을 디버그 모드로 실행 (필수 ⭐)

프로그램이 로그인 세션을 인식할 수 있도록 **기존에 열린 모든 크롬 창을 닫고** 아래 명령어로 실행하세요.

- **Windows:**
  ```cmd
  "C:\Program Files\Google\Chrome\Application\chrome.exe" --remote-debugging-port=9222 --user-data-dir="C:\Chrome_debug"
  ```
- **Mac:**
  ```bash
  /Applications/Google\ Chrome.app/Contents/MacOS/Google\ Chrome --remote-debugging-port=9222
  ```

### 2️⃣ 5개 채널 로그인

- **L (Klook):** https://merchant.klook.com/reviews
- **KK (KKday):** https://scm.kkday.com/v1/en/comment/index
- **GG (GetYourGuide):** https://supplier.getyourguide.com/performance/reviews
- **TPC (Ctrip):** https://vbooking.ctrip.com/tour/comment_manage/comment/list?bizScene=ACTIVITY
- **MRT (MyRealTrip):** https://partner.myrealtrip.com/reviews/touractivity

### 3️⃣ 프로그램 실행

`start_main.bat` 더블클릭(필요 라이브러리 자동 설치) 또는 직접 실행:

```bash
pip install pandas selenium openpyxl reportlab tkcalendar
python main.py
```

1. **🔌 크롬 연결**
2. **📁 틴트 리포트 엑셀 선택**
3. 모드 선택 (PFP / Monthly / Quarterly / New Year Party)
   - PFP: 날짜 체크박스(자동 감지) + 지사(국가) + 가이드 선택
   - Monthly/Q/NYP: 기간 시작~종료일(엑셀에서 자동 감지, 캘린더로 수정 가능) + 지사 선택
4. **▶️ 리뷰 조회 시작** → 엑셀/폴더 자동 저장
5. PFP는 하단 **지사별 Copy** 버튼으로 화면 결과 복사 가능

> 라이브러리: `reportlab`(PDF), `tkcalendar`(캘린더). 없으면 각각 텍스트/기본입력으로 폴백합니다.

---

## 📂 입력 엑셀 형식

프로그램에 업로드할 예약 리스트 엑셀 파일에는 반드시 아래 컬럼이 포함되어야 합니다.

필수 컬럼: `Date, Area, Product, Agency, Agency Code, Main Guide, People`

- `Area` = 지사 (Seoul/Busan/Tokyo/Osaka/Nagoya/Fukuoka/Sapporo/Sydney/London …)
- `Agency` = 채널 코드 (L/KK/GG/TPC/MRT/…)
- `Agency Code` = 예약번호
- `Main Guide` = 가이드 이름(여러 명은 쉼표로 구분)
- `No Show` 컬럼(또는 시트)의 `O` = No Show → **PFP에서만 제외**

---

## 📊 결과물

### PFP (성과제)

`Smart Review Crawling` 폴더(main.py 폴더)에 다음 트리로 저장 + **지사별 zip** 자동 생성:

```
Review Crawling Result 20260629-20260705/
  종합.xlsx                     (종합 시트 + 지사별 탭)
  서울/                         ← 지사 폴더(국가 폴더 없이 바로)
    0. 서울_전체.xlsx           (전체 시트 + 가이드 탭, 맨 위로 정렬)
    최복염.xlsx / 최복염.pdf     (가이드별 엑셀 + PDF)
    ...
  서울.zip                      ← 서울 담당 전달용
  부산/ ...  부산.zip
```

- 화면: 전체/지사/날짜/투어·가이드 요약(5채널 % + 전체 %), `[전체 리뷰 품질]`(익명 포함).
- 가이드 PDF: 표(날짜·투어·에이전시·예약번호·Review Status·Rating·Check) + 날짜별/전체 요약. 한글은 맑은 고딕 임베드, Check(✓/▲/✗)는 심볼 폰트로 렌더.
- 엑셀 헤더의 `Review Status`(구 Review_Status), `Rating`, `Check` 컬럼 포함.

### Monthly / Quarterly / New Year Party

단일 엑셀 `{모드} Review Crawling YYYYMMDD-YYYYMMDD.xlsx`:

- **지사 시트(Seoul/Busan …)**: 좋은 리뷰(4~5점)만 → `Tour Date · Review Date · Agency Code · Tour · Star · Review · Guide · Agency`
- **Guide 시트**: 지사 블록을 가로로 나란히 → `Guide Name · Good · Bad · Total Review · Tour Count · Team Count · Good % · Bad % · Total %` (Total Review = Good + Bad, Good % + Bad % = Total %)

**[Guide 통계 시트 예시]**

| Guide Name | Good | Bad | Total Review | Tour Count | Team Count | Good % | Bad % | Total % |
| :--- | :--- | :--- | :--- | :--- | :--- | :--- | :--- | :--- |
| 가이드 A | 12 | 3 | 15 | 10 | 20 | 60.00% | 15.00% | 75.00% |
| 가이드 B | 6 | 2 | 8 | 5 | 12 | 50.00% | 16.67% | 66.67% |

---

## 🔧 참고 / 튜닝 포인트

- 채널 코드가 바뀌면 `main.py` 상단 `CHANNELS`·`CHANNEL_META` 만 수정.
- 각 채널 요청은 페이지가 실제 보내는 요청을 **런타임 캡처 후 재요청**하므로 사이트 UI가 바뀌어도 비교적 안전합니다. 단 첫 실행 시 채널별 수집 로그(`→ [L] p1: n행 …`)로 정상 수집을 점검하세요.
- Klook은 요청 limit과 무관하게 페이지당 30개를 주므로 **total 기준으로 끝까지** 페이지네이션합니다.
- MRT 토큰은 자주 만료됨 → 조회 직전 MRT 페이지가 로그인 상태여야 함(도구가 localStorage 토큰 사용).
- 스페셜 카테고리 키워드는 `main.py`의 `SPECIAL_PRODUCT_KEYS`에서 관리.

---

## ⚠️ 주의 사항

1. 반드시 **크롬 디버깅 모드**를 먼저 실행한 후 프로그램 내 [🔌 크롬 연결] 버튼을 눌러야 합니다.
2. 각 채널(L, KK, GG, TPC, MRT) 사이트에 미리 **로그인**이 되어 있어야 크롤링이 가능합니다.
3. 엑셀 파일이 열려 있는 상태에서는 저장이 실패할 수 있으니 파일을 닫고 실행해 주세요.
