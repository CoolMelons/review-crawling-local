# 🕷️ Review Crawler - 통합 리뷰 수집 및 가이드 통계기

**L (KLOOK), KK (KKDAY), GG (GetYourGuide)** 파트너 페이지에서 리뷰를 자동으로 수집하고, 예약 리스트와 대조하여 가이드별 성과를 엑셀로 산출하는 데스크톱 프로그램입니다.

![Python](https://img.shields.io/badge/Python-3.8+-blue.svg)
![Selenium](https://img.shields.io/badge/Library-Selenium-brightgreen.svg)
![Platform](https://img.shields.io/badge/Platform-Windows%20%7C%20Mac-lightgrey.svg)

---

## ✨ 주요 기능

- ✅ **통합 크롤링** - KLOOK(L), KKDAY(KK), GYG(GG) 리뷰를 한 번에 수집 (별점 4~5점 타겟)
- ✅ **지역별 자동 분류** - 서울, 부산, 도쿄, 오사카, 후쿠오카, 삿포로, 시드니 등 지역별 시트 분리 생성
- ✅ **가이드별 통계** - 가이드별 리뷰 개수, 투어 횟수, 팀 수 및 리뷰 참여율($\%$) 자동 산출
- ✅ **가이드 시트 재생성(Regenerate)** - 엑셀에서 리뷰 오타 수정 후 통계 시트만 즉시 업데이트
- ✅ **스마트 필터링** - 월별 데이터 분할 수집으로 대량 데이터 누락 방지

---

## 🚀 빠른 시작

아래 두 가지 방법 중 편한 방식으로 실행하세요.

### ✅ 방법 A) Release에서 EXE 다운로드 (가장 쉬움 / 추천)

1️⃣ GitHub 저장소 오른쪽 **Releases** 메뉴로 이동  
2️⃣ 최신 버전에서 실행 파일(`ReviewCrawler.exe`) 다운로드  
3️⃣ 다운로드한 파일 실행  

> ⚠️ **Windows에서 “알 수 없는 앱” 경고가 뜨는 경우**
> **[추가 정보]** 클릭 → **[실행]** 버튼을 눌러 진행하세요.

---

### ✅ 방법 B) GitHub에서 ZIP 다운로드 후 Python으로 실행

#### 1️⃣ 다운로드 및 설치
1. GitHub 페이지 상단 **Code → Download ZIP** 클릭 후 압축 해제
2. 해당 폴더의 터미널(CMD)에서 라이브러리 설치:
   ```bash
   pip install -r requirements.txt
   ```

#### 2️⃣ 크롬 디버깅 모드 실행 (필수 ⭐)
프로그램이 로그인 세션을 인식할 수 있도록 **기존에 열린 모든 크롬 창을 닫고** 아래 명령어로 실행하세요.

- **Windows:**
  ```cmd
  "C:\Program Files\Google\Chrome\Application\chrome.exe" --remote-debugging-port=9222 --user-data-dir="C:\Chrome_debug_temp"
  ```
- **Mac:**
  ```bash
  /Applications/Google\ Chrome.app/Contents/MacOS/Google\ Chrome --remote-debugging-port=9222
  ```

#### 3️⃣ 프로그램 실행
```bash
python review_crawler.py
```

---

## 📂 예약 파일 양식
프로그램에 업로드할 예약 리스트 엑셀 파일에는 반드시 아래 컬럼이 포함되어야 합니다.
- `Date`, `Area`, `Product`, `Agency`, `Agency Code`, `Main Guide`

---

## 📊 출력 결과 예시

프로그램 실행이 완료되면 `Review_시작일_종료일.xlsx` 파일이 생성됩니다.

**[Guide 통계 시트]**
| Guide Name | Review Count | Tour Count | Team Count | Review % |
| :--- | :--- | :--- | :--- | :--- |
| 가이드 A | 15 | 10 | 20 | 75.00% |
| 가이드 B | 8 | 5 | 12 | 66.67% |

---

## ⚠️ 주의 사항
1. 반드시 **크롬 디버깅 모드**를 먼저 실행한 후 프로그램 내 [🔌 크롬 연결] 버튼을 눌러야 합니다.
2. 각 에이전시(KLOOK, KKDAY, GYG) 사이트에 미리 **로그인**이 되어 있어야 크롤링이 가능합니다.
3. 엑셀 파일이 열려 있는 상태에서는 저장이 실패할 수 있으니 파일을 닫고 실행해 주세요.

---

## 🛠 기술 스택
- **Language**: Python 3.8+
- **Library**: Selenium, Pandas, Tkinter, Openpyxl
- **Automation**: Chrome DevTools Protocol (9222 Port)
