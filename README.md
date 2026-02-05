# 🕷️ Review Crawler - 통합 리뷰 수집 및 가이드 통계 도구

**L (KLOOK), KK (KKDAY), GG (GetYourGuide)** 파트너 센터의 리뷰 데이터를 자동으로 크롤링(Crawling)하고, 예약 리스트와 대조하여 가이드별 성과를 분석하는 자동화 툴입니다.

![Python](https://img.shields.io/badge/Python-3.8+-blue.svg)
![Selenium](https://img.shields.io/badge/Library-Selenium-brightgreen.svg)

---

## 🚀 핵심 기능
- **통합 크롤링**: 플랫폼별로 흩어진 리뷰를 한곳으로 수집 (KLOOK, KKDAY, GYG)
- **가이드 매칭**: 예약 데이터 기반으로 리뷰를 가이드 이름별로 자동 분류
- **성과 지표 산출**: 가이드별 리뷰 개수, 투어 횟수, 팀 수 대비 리뷰 참여율($\%$) 계산
- **재생성 모드**: 엑셀에서 리뷰 오타 수정 후, 통계 시트만 즉시 재계산

---

## ⚙️ 실행 전 필수 설정

이 프로그램은 보안 정책상 사용자의 **로그인 세션이 유지된 크롬 브라우저**를 제어해야 합니다. 반드시 아래 절차를 따라주세요.

### 1️⃣ 크롬 디버깅 모드 실행
기존에 열려 있는 모든 크롬 창을 완전히 종료한 후, 터미널(또는 CMD)에 다음 명령어를 입력하여 크롬을 실행하세요.

**Windows:**
```cmd
"C:\Program Files\Google\Chrome\Application\chrome.exe" --remote-debugging-port=9222 --user-data-dir="C:\Chrome_debug_temp"
