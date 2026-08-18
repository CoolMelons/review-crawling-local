# -*- coding: utf-8 -*-
"""
Smart Review Crawling
================================================================
기존 'Review Search'의 후속 버전.

핵심 변경점
- Selenium 클릭 스크래핑 ❌  →  로그인된 세션으로 API(개발자 모드) 호출 ✅
- 채널 5개: L(Klook), KK(KKday), GG(GetYourGuide), TPC(Ctrip/Trip.com), MRT(MyRealTrip)
- 1~5점 리뷰 전부 수집
- 전 지사(Area) 대상  (Seoul/Busan/Tokyo ...)
- 결과: 화면 통계(지사→날짜 계층) + 엑셀 결과 저장 + 전체 Copy

동작 방식(캡처-리플레이)
- 각 채널 리뷰 페이지가 실제로 보내는 요청을 CDP 프리로드 훅으로 가로채고,
  페이지네이션/날짜만 바꿔 그대로 재요청해서 리뷰를 수집한다.
- 예약번호(Agency Code)로 매칭한다.
  L=booking_no, KK=orderMid, GG=bookingReference, TPC=orderId, MRT=reservationNo

실행 준비
- 크롬을 디버그 모드로 실행 후 5개 채널에 로그인 (하단 안내 참고)
================================================================
"""

import os
import re
import json
import time
import traceback
from decimal import Decimal, InvalidOperation
from datetime import datetime, timedelta

import pandas as pd
from tkinter import (
    Tk, filedialog, Label, Button, Toplevel, StringVar, messagebox,
    Frame, Scrollbar, Canvas, Checkbutton, BooleanVar, Text, Entry, Radiobutton
)
from tkinter.ttk import Progressbar, Checkbutton as TtkCheckbutton

from selenium import webdriver
from selenium.webdriver.chrome.options import Options

# 캘린더 위젯(선택): 있으면 DateEntry, 없으면 일반 Entry로 폴백
try:
    from tkcalendar import DateEntry
    HAS_TKCAL = True
except Exception:
    DateEntry = None
    HAS_TKCAL = False


# =========================================================
# 설정
# =========================================================
REQUIRED_COLS = ["Date", "Area", "Product", "Agency", "Agency Code", "Main Guide", "People"]

# 조회 대상 채널(= 엑셀 Agency 컬럼 값).  값이 다르면 여기만 바꾸면 됨.
CHANNELS = ["L", "KK", "GG", "TPC", "MRT"]

CHANNEL_META = {
    "L":   {"name": "KLOOK",        "page": "https://merchant.klook.com/reviews",
            "endpoint": "review_list"},
    "KK":  {"name": "KKDAY",        "page": "https://scm.kkday.com/v1/en/comment/index",
            "endpoint": "get_comment_list"},
    "GG":  {"name": "GetYourGuide", "page": "https://supplier.getyourguide.com/performance/reviews",
            "endpoint": "/graphql", "body_contains": "bookingReference"},
    "TPC": {"name": "Trip.com",     "page": "https://vbooking.ctrip.com/tour/comment_manage/comment/list?bizScene=ACTIVITY",
            "endpoint": "listOrderComments"},
    "MRT": {"name": "MyRealTrip",   "page": "https://partner.myrealtrip.com/reviews/touractivity",
            "endpoint": "reviews/search"},
}

DEBUG_PORT = "127.0.0.1:9222"

# 결과 저장 기준 = main.py 가 있는 폴더 (돌린 폴더)
try:
    SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
except NameError:
    SCRIPT_DIR = os.getcwd()

# 결과 저장 위치 = 내 컴퓨터 다운로드 폴더 (없으면 스크립트 폴더로 폴백)
DOWNLOAD_DIR = os.path.join(os.path.expanduser("~"), "Downloads")
if not os.path.isdir(DOWNLOAD_DIR):
    DOWNLOAD_DIR = SCRIPT_DIR

# 지사(Area)의 국가 그룹 & 표시 순서
COUNTRY_GROUPS = [
    ("한국", ["Seoul", "Busan"]),
    ("일본", ["Tokyo", "Osaka", "Nagoya", "Fukuoka", "Sapporo"]),
    ("호주", ["Sydney"]),
    ("영국", ["London"]),
]

# 지사 영문 → 한글 시트명
AREA_KR = {
    "Seoul": "서울", "Busan": "부산", "Tokyo": "도쿄", "Osaka": "오사카",
    "Nagoya": "나고야", "Fukuoka": "후쿠오카", "Sapporo": "삿포로",
    "Sydney": "시드니", "London": "런던",
}

# 스페셜 카테고리 상품(리뷰 조회 대상 아님 → 결과/엑셀에서 제외). 공백제거·소문자 부분일치.
SPECIAL_PRODUCT_KEYS = ["mbc스튜디오", "dr.petit", "마리엠헤어"]

# 국가 영문명 (Monthly/Q/NYP 파일명용)
COUNTRY_EN = {"한국": "Korea", "일본": "Japan", "호주": "Australia", "영국": "UK", "기타": "Etc"}


def _region_tag(areas, english=False):
    """선택 지역들의 국가명을 순서대로 이어 붙임 (파일명용)."""
    order = [c for c, _ in COUNTRY_GROUPS]
    cs = []
    for a in areas:
        c = area_country(a)
        name = COUNTRY_EN.get(c, c) if english else c
        if name not in cs:
            cs.append(name)
    cs.sort(key=lambda n: order.index(n) if (not english and n in order) else
            ([COUNTRY_EN.get(k, k) for k in order].index(n) if english and n in [COUNTRY_EN.get(k, k) for k in order] else 99))
    return "_".join(cs)



def area_rank(area):
    a = str(area).strip().lower()
    for ci, (c, areas) in enumerate(COUNTRY_GROUPS):
        for ai, name in enumerate(areas):
            if name.lower() == a:
                return (ci, ai)
    return (99, str(area))  # 미등록 지사는 맨 뒤, 이름순


def area_country(area):
    a = str(area).strip().lower()
    for c, areas in COUNTRY_GROUPS:
        if any(name.lower() == a for name in areas):
            return c
    return "기타"


# =========================================================
# 공통 유틸
# =========================================================
def norm_code(x):
    """예약번호/코드를 안정적으로 문자열화 (큰 숫자/과학표기/끝 .0 처리) + 대문자/공백제거."""
    s = str(x).strip()
    if s.lower() in ["nan", "none", ""]:
        return ""
    s = s.lstrip("'")
    m = re.match(r"^(\d+)\.0$", s)
    if m:
        s = m.group(1)
    elif re.match(r"^\d+(\.\d+)?e[+-]?\d+$", s, flags=re.IGNORECASE):
        try:
            s = format(Decimal(s), "f").split(".")[0]
        except (InvalidOperation, ValueError):
            pass
    return re.sub(r"\s+", "", s).upper()


def canonicalize_guides(df, col="Main Guide"):
    """같은 사람인데 대소문자/공백만 다른 가이드 표기를 하나로 통일.
    표시 이름은 (1)대소문자 섞인 표기 우선 (2)많이 쓰인 표기 순. 전부 대문자면 그대로 둔다."""
    def _split(v):
        return [n.strip() for n in str(v).split(",") if n.strip()]

    def _key(n):
        return re.sub(r"\s+", " ", str(n).strip()).lower()

    variants = {}
    for raw in df[col].dropna():
        for n in _split(raw):
            k = _key(n)
            variants.setdefault(k, {})
            variants[k][n] = variants[k].get(n, 0) + 1

    disp = {}
    for k, cands in variants.items():
        def _low(x):
            return sum(1 for ch in x if ch.islower())
        best = sorted(cands.items(), key=lambda kv: (-_low(kv[0]), -kv[1], kv[0]))[0][0]
        disp[k] = re.sub(r"\s+", " ", str(best).strip())

    def _fix(v):
        if pd.isna(v):
            return v
        names = [disp.get(_key(n), re.sub(r"\s+", " ", n.strip())) for n in _split(v)]
        seen, out = set(), []
        for n in names:                      # 중복 제거(같은 사람 두 번 적힌 경우)
            if n not in seen:
                seen.add(n); out.append(n)
        return ", ".join(out)

    merged = sum(1 for k, c in variants.items() if len(c) > 1)
    df[col] = df[col].apply(_fix)
    return df, merged


def to_epoch_ms(dt):
    return int(pd.Timestamp(dt).timestamp() * 1000)


# =========================================================
# 메인 앱
# =========================================================
class PowerReviewApp:
    def __init__(self):
        self.root = Tk()
        self.root.title("📋 Smart Review Crawling")
        self.root.geometry("840x1340")

        self.driver = None
        self.df = None
        self.df_simple = None   # Monthly/NYP용 (스페셜+No Show 모두 포함)
        self.mode_var = StringVar(value="PFP")   # PFP / MONTHLY / NYP
        self._preload_id = None

        # No Show
        self.noshow_codes = set()
        self.noshow_teams = 0
        self.noshow_people = 0

        # 선택 체크박스 (날짜/지사/가이드)
        self.date_vars = {}          # {Timestamp: BooleanVar}
        self.select_all_dates = BooleanVar(value=True)
        self.branch_vars = {}        # {area: BooleanVar}
        self.select_all_branches = BooleanVar(value=True)
        self.country_vars = {}     # {country: BooleanVar} (지사=국가 단위 선택)
        self.guide_vars = {}         # {guide: BooleanVar}
        self.select_all_guides = BooleanVar(value=True)

        self.detail_lines = []
        self.last_report_text = ""
        self._branch_texts = {}

        self.setup_ui()

    # ---------------------------------------------------------
    # UI
    # ---------------------------------------------------------
    def setup_ui(self):
        Label(self.root, text="📋 Smart Review Crawling", font=("Arial", 18, "bold")).pack(pady=12)
        Label(self.root, text="API(개발자 모드) 방식 · L / KK / GG / TPC / MRT · 전 지사",
              font=("Arial", 9), fg="#555").pack()

        # 모드 선택 (PFP 성과제 / Monthly / New Year Party)
        mf = Frame(self.root, relief="solid", borderwidth=1, padx=10, pady=6)
        mf.pack(fill="x", padx=20, pady=(8, 0))
        Label(mf, text="모드:", font=("Arial", 11, "bold")).pack(side="left", padx=(0, 6))
        for _val, _txt in [("PFP", "PFP (1주일)"),
                           ("MONTHLY", "Monthly (1달)"),
                           ("QUARTERLY", "Quarterly (3달)"),
                           ("NYP", "New Year Party (1년)")]:
            Radiobutton(mf, text=_txt, variable=self.mode_var, value=_val,
                        command=self._on_mode_change).pack(side="left", padx=8)
        # 선택한 모드의 날짜 기준 표시 (지사→도시 표시처럼 동적)
        _bf = Frame(self.root)
        _bf.pack(fill="x", padx=22, pady=(0, 2))
        self.mode_basis_var = StringVar(value="")
        Label(_bf, textvariable=self.mode_basis_var, font=("Arial", 9), fg="#2196F3").pack(anchor="w")

        # 1) 크롬 연결
        f1 = Frame(self.root, relief="solid", borderwidth=1, padx=10, pady=10)
        f1.pack(fill="x", padx=20, pady=6)
        Label(f1, text="1️⃣ 크롬 연결 (디버그 모드)", font=("Arial", 12, "bold")).pack(anchor="w")
        Label(f1, text="⚠️ L, KK, GG, TPC, MRT 모두 로그인 필요", font=("Arial", 9), fg="red").pack(anchor="w")
        self.chrome_status = StringVar(value="🔴 크롬 미연결")
        Label(f1, textvariable=self.chrome_status, font=("Arial", 10)).pack(anchor="w", pady=4)
        Button(f1, text="🔌 크롬 연결", command=self.connect_chrome,
               width=20, bg="#4CAF50", fg="white").pack(anchor="w")

        # 2) 엑셀 선택
        f2 = Frame(self.root, relief="solid", borderwidth=1, padx=10, pady=10)
        f2.pack(fill="x", padx=20, pady=6)
        Label(f2, text="2️⃣ 틴트 리포트 엑셀 선택", font=("Arial", 12, "bold")).pack(anchor="w")
        self.file_status = StringVar(value="📁 파일 미선택")
        Label(f2, textvariable=self.file_status, font=("Arial", 10)).pack(anchor="w", pady=4)
        Button(f2, text="📁 파일 선택", command=self.select_file,
               width=20, bg="#2196F3", fg="white").pack(anchor="w")

        # 3) 날짜  (PFP=체크박스 / Monthly·NYP=기간 캘린더)  4) 지사  5) 가이드
        def _scroll_section(parent, title, height):
            fr = Frame(parent, relief="solid", borderwidth=1, padx=10, pady=6)
            fr.pack(fill="both", expand=True, padx=0, pady=0)
            top = Frame(fr); top.pack(fill="x")
            Label(top, text=title, font=("Arial", 11, "bold")).pack(side="left", anchor="w")
            cf = Frame(fr); cf.pack(fill="both", expand=True)
            cv = Canvas(cf, height=height)
            sb = Scrollbar(cf, orient="vertical", command=cv.yview)
            inner = Frame(cv)
            inner.bind("<Configure>", lambda e, c=cv: c.configure(scrollregion=c.bbox("all")))
            cv.create_window((0, 0), window=inner, anchor="nw")
            cv.configure(yscrollcommand=sb.set)
            cv.pack(side="left", fill="both", expand=True)
            sb.pack(side="right", fill="y")
            return fr, top, inner

        # 날짜 컨테이너 (모드에 따라 체크박스/기간 스왑)
        self.date_container = Frame(self.root)
        self.date_container.pack(fill="x", padx=20, pady=4)

        # (A) PFP용: 날짜 체크박스
        self.date_cb_section, d_top, self.date_frame = _scroll_section(self.date_container, "3️⃣ 날짜", 165)
        Checkbutton(d_top, text="전체", variable=self.select_all_dates,
                    command=self.toggle_all_dates).pack(side="right")

        # (B) Monthly/NYP용: 기간(시작일~종료일, 캘린더 선택)
        self.date_range_section = Frame(self.date_container, relief="solid", borderwidth=1, padx=10, pady=8)
        Label(self.date_range_section, text="3️⃣ 기간 (엑셀에서 자동 감지 · 수정 가능)",
              font=("Arial", 11, "bold")).pack(anchor="w")
        _rr = Frame(self.date_range_section); _rr.pack(fill="x", pady=6)
        Label(_rr, text="시작일:", font=("Arial", 10)).pack(side="left", padx=(0, 4))
        self.start_date_widget = self._mk_date(_rr); self.start_date_widget.pack(side="left", padx=(0, 14))
        Label(_rr, text="종료일:", font=("Arial", 10)).pack(side="left", padx=(0, 4))
        self.end_date_widget = self._mk_date(_rr); self.end_date_widget.pack(side="left")
        if not HAS_TKCAL:
            Label(self.date_range_section,
                  text="※ 달력 위젯이 없어 텍스트 입력(YYYY-MM-DD)입니다. 'pip install tkcalendar' 하면 달력이 떠요.",
                  font=("Arial", 8), fg="#888").pack(anchor="w")

        # 4) 지사 = 국가 단위 가로 선택
        bf = Frame(self.root, relief="solid", borderwidth=1, padx=10, pady=8)
        bf.pack(fill="x", padx=20, pady=4)
        Label(bf, text="4️⃣ 지사 (국가 선택)", font=("Arial", 11, "bold")).pack(anchor="w")
        self.country_row = Frame(bf)
        self.country_row.pack(fill="x", pady=3)
        self.branch_cities_var = StringVar(value="")
        Label(bf, textvariable=self.branch_cities_var, font=("Arial", 9), fg="#555",
              justify="left", wraplength=760).pack(anchor="w")

        # 5) 가이드 (PFP 전용 · Monthly/NYP는 전체 가이드라 숨김)
        self.guide_section, g_top, self.guide_frame = _scroll_section(self.root, "5️⃣ 가이드 (체크한 가이드만 · 팀/명)", 150)
        Checkbutton(g_top, text="전체", variable=self.select_all_guides,
                    command=self.toggle_all_guides).pack(side="right")

        # 실행
        self.run_button = Button(self.root, text="▶️ 리뷰 조회 시작",
               command=self._run_dispatch, height=2,
               bg="#FF9800", fg="white", font=("Arial", 11, "bold"))
        self.run_button.pack(fill="x", padx=20, pady=8)

        # 결과
        rf = Frame(self.root, relief="solid", borderwidth=1, padx=10, pady=10)
        rf.pack(fill="both", expand=True, padx=20, pady=6)
        hdr = Frame(rf)
        hdr.pack(fill="x")
        Label(hdr, text="📊 조회 결과 (지사 → 날짜)", font=("Arial", 12, "bold")).pack(side="left", anchor="w")
        Button(hdr, text="전체", command=self._search_clear, width=5).pack(side="right", padx=2)
        Button(hdr, text="🔍 검색", command=self._search_guide, width=7).pack(side="right", padx=2)
        self.search_var = StringVar()
        _se = Entry(hdr, textvariable=self.search_var, width=16)
        _se.pack(side="right", padx=2)
        _se.bind("<Return>", lambda e: self._search_guide())
        _se.bind("<KeyRelease>", lambda e: self._search_guide())
        Label(hdr, text="가이드 검색:", font=("Arial", 9)).pack(side="right", padx=(0, 3))
        rsf = Frame(rf)
        rsf.pack(fill="both", expand=True)
        rsb = Scrollbar(rsf)
        rsb.pack(side="right", fill="y")
        self.result_text = Text(rsf, height=13, width=70, yscrollcommand=rsb.set,
                                font=("Consolas", 9), wrap="none")
        self.result_text.pack(side="left", fill="both", expand=True)
        rsb.config(command=self.result_text.yview)

        self.progress_var = StringVar(value="")
        Label(self.root, textvariable=self.progress_var, font=("Arial", 9)).pack(pady=3)

        Label(self.root, text="💾 엑셀 자동 저장됨", font=("Arial", 9), fg="#4CAF50").pack()
        self.copy_btn_frame = Frame(self.root)
        self.copy_btn_frame.pack(pady=5)
        self._build_copy_buttons()

        self._result_df = None  # 마지막 결과 df (엑셀 저장용)
        self._on_mode_change()   # 초기 모드(PFP) 화면 반영

    def log(self, msg=""):
        try:
            line = str(msg)
            print(line)
            self.detail_lines.append(line)
            if getattr(self, "result_text", None) is not None:
                self.result_text.insert("end", line + "\n")
                self.result_text.see("end")
                self.root.update_idletasks()
        except Exception:
            pass

    # ---------------------------------------------------------
    # 크롬 연결
    # ---------------------------------------------------------
    def connect_chrome(self):
        try:
            options = Options()
            options.add_experimental_option("debuggerAddress", DEBUG_PORT)
            self.driver = webdriver.Chrome(options=options)
            self.driver.set_script_timeout(60)
            self.chrome_status.set("🟢 크롬 연결됨")
            messagebox.showinfo("성공", "크롬 연결 성공!\n\nL, KK, GG, TPC, MRT 로그인 상태를 확인하세요.")
        except Exception as e:
            self.chrome_status.set("🔴 크롬 연결 실패")
            messagebox.showerror(
                "연결 실패",
                f"크롬 연결 실패: {e}\n\n크롬을 디버그 모드로 먼저 실행하세요:\n\n"
                'Windows:\n"C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe" '
                '--remote-debugging-port=9222 --user-data-dir="C:\\Chrome_debug"\n\n'
                'Mac:\n/Applications/Google\\ Chrome.app/Contents/MacOS/Google\\ Chrome '
                '--remote-debugging-port=9222'
            )

    # ---------------------------------------------------------
    # 엑셀 로딩 + No Show
    # ---------------------------------------------------------
    def load_excel_with_noshow(self, file_path):
        xls = pd.read_excel(file_path, sheet_name=None)

        # No Show 시트 탐색
        noshow_sheet = None
        for name in xls.keys():
            nm = str(name).strip().lower()
            if nm in ["no show", "noshow", "no_show", "no-show"] or "no show" in nm:
                noshow_sheet = name
                break

        # 메인 시트 = No Show 제외 첫 시트
        main_sheet = None
        for name in xls.keys():
            if name == noshow_sheet:
                continue
            main_sheet = name
            break
        if main_sheet is None:
            raise ValueError("메인 데이터 시트를 찾을 수 없습니다.")

        df = xls[main_sheet].copy()
        df.columns = [str(c).strip() for c in df.columns]
        missing = [c for c in REQUIRED_COLS if c not in df.columns]
        if missing:
            raise ValueError(f"필수 컬럼 누락: {missing}")

        df["Agency Code"] = df["Agency Code"].apply(norm_code)

        # No Show 코드 수집
        noshow_codes = set()
        if noshow_sheet is not None:
            ns = xls[noshow_sheet].copy()
            ns.columns = [str(c).strip() for c in ns.columns]
            code_col = None
            for c in ns.columns:
                lc = c.lower()
                if lc in ["agency code", "booking code", "booking", "order", "order id",
                          "reservation", "reservation code"] or ("code" in lc and code_col is None):
                    code_col = c
            flag_col = None
            for c in ns.columns:
                lc = c.lower().replace(" ", "")
                if lc in ["noshow", "no_show", "no-show"] or ("show" in lc and flag_col is None):
                    flag_col = c
            if code_col is not None:
                for _, r in ns.iterrows():
                    code = norm_code(r.get(code_col, ""))
                    if not code:
                        continue
                    if flag_col is not None:
                        if str(r.get(flag_col, "")).strip().upper() == "O":
                            noshow_codes.add(code)
                    else:
                        row_text = " ".join(str(v) for v in r.values).upper()
                        if " O " in f" {row_text} " or row_text.strip() == "O":
                            noshow_codes.add(code)

        # 메인 시트에 No Show 컬럼이 있으면 함께 반영
        noshow_main_col = None
        for c in df.columns:
            norm = re.sub(r"[\s\-_]", "", str(c).strip().lower())
            if norm in ["noshow", "noshow(o)"] or ("no" in norm and "show" in norm):
                noshow_main_col = c
                break
        if noshow_main_col is not None:
            flags = df[noshow_main_col].astype(str).str.strip().str.upper()
            add = set(df.loc[flags == "O", "Agency Code"].apply(norm_code).tolist())
            noshow_codes |= {c for c in add if c}

        return df, noshow_codes, main_sheet, noshow_sheet

    def select_file(self):
        if not self.driver:
            messagebox.showerror("오류", "먼저 크롬을 연결하세요!")
            return
        path = filedialog.askopenfilename(
            title="틴트 리포트 엑셀 선택",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")])
        if not path:
            return
        try:
            df, noshow_codes, main_sheet, noshow_sheet = self.load_excel_with_noshow(path)

            df = df[df["Main Guide"].notna() & (df["Main Guide"].astype(str).str.strip() != "")].copy()
            df["Date"] = pd.to_datetime(df["Date"], errors="coerce")
            df = df[df["Date"].notna()].copy()
            df["Area"] = df["Area"].astype(str).str.strip()
            df["Agency"] = df["Agency"].astype(str).str.strip().str.upper()
            df["Agency Code"] = df["Agency Code"].astype(str).str.strip()
            df["People"] = pd.to_numeric(df["People"], errors="coerce").fillna(0).astype(int)

            # 가이드 이름 표기 통일 (대소문자/공백만 다른 동일인 병합)
            df, _merged_guides = canonicalize_guides(df)

            # 스페셜 카테고리 마스크 (MBC 스튜디오/Dr.Petit/마리엠헤어)
            _pn = df["Product"].astype(str).str.replace(r"\s+", "", regex=True).str.lower()
            special_mask = _pn.apply(lambda x: any(k in x for k in SPECIAL_PRODUCT_KEYS))

            # Monthly/NYP용: No Show 포함 · 스페셜 카테고리 제외
            self.df_simple = df[~special_mask].copy()

            # No Show 제외 (PFP 성과제 전용)
            self.noshow_codes = {str(c).strip() for c in noshow_codes if str(c).strip()}
            if self.noshow_codes:
                mask = df["Agency Code"].isin(self.noshow_codes)
                self.noshow_teams = int(mask.sum())
                self.noshow_people = int(df.loc[mask, "People"].sum())
                df = df[~mask].copy()
            else:
                self.noshow_teams = 0
                self.noshow_people = 0

            # 스페셜 카테고리 상품 제외 (PFP 성과제용)
            self.df = df[~special_mask.loc[df.index]].copy()
            self.input_path = path

            branches = sorted(df["Area"].unique().tolist())
            ns_msg = ""
            if _merged_guides:
                ns_msg += f" | 가이드 표기 통일 {_merged_guides}명"
            if self.noshow_teams:
                ns_msg = f" | No Show(O) {self.noshow_teams}팀 {self.noshow_people}명 (PFP만 제외 · Monthly/NYP는 포함)"
            self.file_status.set(
                f"✅ {len(df)}건 · 지사 {len(branches)}개({', '.join(branches[:6])}{'...' if len(branches)>6 else ''}){ns_msg}")

            self.build_date_checkboxes()
            self.build_branch_checkboxes()
            self.build_guide_checkboxes()
            # Monthly/NYP 기간(캘린더) 자동 채움
            try:
                self._set_date(self.start_date_widget, self.df_simple["Date"].min())
                self._set_date(self.end_date_widget, self.df_simple["Date"].max())
            except Exception:
                pass
        except Exception as e:
            messagebox.showerror("오류", f"파일 읽기 실패:\n{e}")

    def build_date_checkboxes(self):
        for w in self.date_frame.winfo_children():
            w.destroy()
        self.date_vars = {}
        for d in sorted(self.df["Date"].dt.normalize().unique()):
            ts = pd.Timestamp(d)
            sub = self.df[self.df["Date"].dt.normalize() == ts]
            var = BooleanVar(value=True)
            self.date_vars[ts] = var
            TtkCheckbutton(
                self.date_frame,
                text=f"{ts.strftime('%Y-%m-%d (%a)')}  ·  {len(sub)}팀 {int(sub['People'].sum())}명",
                variable=var, command=self._on_filter_change
            ).pack(anchor="w", padx=5, pady=1)

    def build_branch_checkboxes(self):
        for w in self.country_row.winfo_children():
            w.destroy()
        self.country_vars = {}
        order = [c for c, _ in COUNTRY_GROUPS] + ["기타"]
        present = []
        for area in self.df["Area"].unique():
            c = area_country(area)
            if c not in present:
                present.append(c)
        present.sort(key=lambda c: order.index(c) if c in order else 99)
        for c in present:
            n = self.df[self.df["Area"].apply(area_country) == c]
            var = BooleanVar(value=True)
            self.country_vars[c] = var
            TtkCheckbutton(self.country_row, text=f"{c} ({len(n)}팀)",
                           variable=var, command=self._on_branch_change).pack(side="left", padx=10)
        self._update_branch_cities()

    def _update_branch_cities(self):
        areas = sorted(self._selected_branches(), key=area_rank)
        self.branch_cities_var.set("지사: " + (", ".join(areas) if areas else "(국가를 선택하세요)"))

    def _on_branch_change(self):
        self._update_branch_cities()
        self.build_guide_checkboxes()

    def build_guide_checkboxes(self):
        if not hasattr(self, "guide_frame"):
            return
        for w in self.guide_frame.winfo_children():
            w.destroy()
        prev = {g: v.get() for g, v in self.guide_vars.items()}
        self.guide_vars = {}
        dsel = self._selected_dates()
        bsel = self._selected_branches()
        if not dsel or not bsel:
            Label(self.guide_frame, text="(날짜·지사(국가)를 먼저 선택하세요)", fg="#888").pack(anchor="w", padx=5)
            return
        sub = self.df[self.df["Date"].dt.normalize().isin(dsel) & self.df["Area"].isin(bsel)]
        if sub.empty:
            Label(self.guide_frame, text="(선택 조건에 가이드 없음)", fg="#888").pack(anchor="w", padx=5)
            return
        rows = []
        for guide, g in sub.groupby("Main Guide"):
            areas = sorted(g["Area"].unique(), key=area_rank)
            pa = areas[0]
            rows.append((area_rank(pa), pa, str(guide), len(g), int(g["People"].sum())))
        rows.sort(key=lambda t: (t[0], t[2]))  # 지사(국가)순 → 가이드 ㄱㄴㄷ
        for _rank, pa, guide, teams, people in rows:
            var = BooleanVar(value=prev.get(guide, True))
            self.guide_vars[guide] = var
            TtkCheckbutton(
                self.guide_frame,
                text=f"{guide}  ·  {teams}팀 {people}명   [{pa}]",
                variable=var
            ).pack(anchor="w", padx=5, pady=1)

    def _selected_dates(self):
        return [ts for ts, var in self.date_vars.items() if var.get()]

    def _selected_branches(self):
        if self.df is None or not self.country_vars:
            return []
        sel_c = [c for c, v in self.country_vars.items() if v.get()]
        return [a for a in self.df["Area"].unique() if area_country(a) in sel_c]

    def _on_filter_change(self):
        # 날짜/지사 선택이 바뀌면 가이드 목록/카운트 갱신
        self.build_guide_checkboxes()

    def toggle_all_dates(self):
        v = self.select_all_dates.get()
        for var in self.date_vars.values():
            var.set(v)
        self.build_guide_checkboxes()

    def toggle_all_branches(self):
        for var in self.country_vars.values():
            var.set(self.select_all_branches.get())
        self._on_branch_change()

    def toggle_all_guides(self):
        v = self.select_all_guides.get()
        for var in self.guide_vars.values():
            var.set(v)

    # ---------------------------------------------------------
    # 모드 전환 & 날짜 위젯 (Monthly/NYP 기간 캘린더)
    # ---------------------------------------------------------
    def _on_mode_change(self):
        m = self.mode_var.get()
        if getattr(self, "mode_basis_var", None) is not None:
            self.mode_basis_var.set("↳ 참여일 기준" if m == "PFP" else "↳ 리뷰 작성일 기준")
        self.date_cb_section.pack_forget()
        self.date_range_section.pack_forget()
        self.guide_section.pack_forget()
        if m == "PFP":
            self.date_cb_section.pack(fill="both", expand=True)
            self.guide_section.pack(fill="both", expand=True, padx=20, pady=4, before=self.run_button)
        else:
            self.date_range_section.pack(fill="x")

    def _mk_date(self, parent):
        if HAS_TKCAL:
            return DateEntry(parent, width=12, date_pattern="yyyy-mm-dd",
                             background="#2196F3", foreground="white", borderwidth=2)
        return Entry(parent, width=14)

    def _set_date(self, w, value):
        try:
            d = pd.Timestamp(value).date()
        except Exception:
            return
        if HAS_TKCAL and hasattr(w, "set_date"):
            try:
                w.set_date(d)
            except Exception:
                pass
        else:
            try:
                w.delete(0, "end"); w.insert(0, d.strftime("%Y-%m-%d"))
            except Exception:
                pass

    def _get_date(self, w):
        if HAS_TKCAL and hasattr(w, "get_date"):
            return w.get_date()
        try:
            return pd.Timestamp(str(w.get()).strip()).date()
        except Exception:
            return None

    # ---------------------------------------------------------
    # 캡처-리플레이 엔진
    # ---------------------------------------------------------
    def _clear_preload(self):
        if self._preload_id is not None:
            try:
                self.driver.execute_cdp_cmd(
                    "Page.removeScriptToEvaluateOnNewDocument",
                    {"identifier": self._preload_id})
            except Exception:
                pass
            self._preload_id = None

    def capture_request(self, page_url, endpoint, body_contains="", wait=8):
        """CDP 프리로드 훅으로 페이지가 보내는 실제 요청(url/method/headers/body)을 캡처."""
        self._clear_preload()
        hook = """
        (function(){
          window.__CAP__ = null;
          var WANT = %s;
          var NEED = %s;
          function save(url, method, headers, body){
            try{
              if(String(url).indexOf(WANT) >= 0
                 && (!NEED || (body && String(body).indexOf(NEED) >= 0))
                 && !window.__CAP__){
                window.__CAP__ = {url:String(url), method:(method||'GET'), headers:(headers||{}), body:(body||null)};
              }
            }catch(e){}
          }
          var of = window.fetch;
          window.fetch = function(){
            var a = arguments;
            try{
              var url = (a[0] && a[0].url) || a[0];
              var m = (a[1] && a[1].method) || (a[0] && a[0].method) || 'GET';
              var h = {};
              var hs = (a[1] && a[1].headers) || (a[0] && a[0].headers);
              if(hs){ if(hs instanceof Headers){ hs.forEach(function(v,k){h[k]=v;}); } else { for(var k in hs){h[k]=hs[k];} } }
              var b = (a[1] && a[1].body) || null;
              if(a[0] instanceof Request && !b){ a[0].clone().text().then(function(t){ save(url,m,h,t); }); }
              else { save(url,m,h,b); }
            }catch(e){}
            return of.apply(this, a);
          };
          var O = XMLHttpRequest.prototype.open,
              S = XMLHttpRequest.prototype.send,
              SR = XMLHttpRequest.prototype.setRequestHeader;
          XMLHttpRequest.prototype.open = function(m,u){ this.__m=m; this.__u=u; this.__h={}; return O.apply(this, arguments); };
          XMLHttpRequest.prototype.setRequestHeader = function(k,v){ this.__h[k]=v; return SR.apply(this, arguments); };
          XMLHttpRequest.prototype.send = function(b){ save(this.__u, this.__m, this.__h, b); return S.apply(this, arguments); };
        })();
        """ % (json.dumps(endpoint), json.dumps(body_contains or ""))

        res = self.driver.execute_cdp_cmd(
            "Page.addScriptToEvaluateOnNewDocument", {"source": hook})
        self._preload_id = res.get("identifier")

        self.driver.get(page_url)
        deadline = time.time() + wait
        cap = None
        while time.time() < deadline:
            cap = self.driver.execute_script("return window.__CAP__ || null;")
            if cap:
                break
            time.sleep(0.4)
        return cap

    def _fetch_page(self, js, *args):
        """execute_async_script 래퍼 (한 페이지 요청).
        주의: 내부 async 함수 안에서는 바깥 arguments 를 볼 수 없으므로,
        인자를 P[] 배열로, 콜백을 done 으로 클로저 캡처해서 넘긴다.
        각 채널 js 는 P[0], P[1] ... 로 인자에 접근한다."""
        wrapper = ("var done = arguments[arguments.length-1];\n"
                   "var P = Array.prototype.slice.call(arguments, 0, arguments.length-1);\n"
                   "(async function(){ try{\n" + js +
                   "\n}catch(e){ done({error:String(e)}); } })();")
        return self.driver.execute_async_script(wrapper, *args)

    # ---------------------------------------------------------
    # 채널별 수집기  →  {정규화 예약번호: rating(str)}
    # ---------------------------------------------------------
    def collect_klook(self, min_ms, date_set=None):
        reviews = {}
        cap = self.capture_request(CHANNEL_META["L"]["page"], "review_list", wait=10)
        if not cap:
            self.log("  ⚠ [L] 요청 캡처 실패 (로그인/페이지 확인)")
            return reviews
        js = """
        var cap=P[0], page=P[1], limit=P[2];
        var u=new URL(cap.url, location.origin);
        u.searchParams.set('page', String(page));
        u.searchParams.set('limit', String(limit));
        var res=await fetch(u.toString(), {credentials:'include', headers:(cap.headers||{})});
        var j=await res.json();
        var list=(j.result&&j.result.review_list)||j.review_list||[];
        var total=(j.result&&j.result.total)||null;
        done({rows:list.map(function(x){return {code:x.booking_no, rating:x.stars, rdate:x.review_time, pdate:x.participant_time};}), total:total});
        """
        page, limit, guard = 1, 30, 0
        while page <= 200:
            r = self._fetch_page(js, cap, page, limit)
            if not r or r.get("error"):
                self.log(f"  ⚠ [L] p{page} 오류: {r.get('error') if r else 'no-resp'}")
                break
            rows = r.get("rows") or []
            if not rows:
                break
            oldest_ok = True
            for x in rows:
                code = norm_code(x.get("code"))
                d = (x.get("pdate") or "")[:10]
                if code and (not date_set or d in date_set):
                    reviews[code] = self._fmt_rating(x.get("rating"))
                if not self._within(x.get("rdate"), min_ms):
                    oldest_ok = False
            self.log(f"  → [L] p{page}: {len(rows)}건 (기간내 누적 {len(reviews)})")
            if not oldest_ok:   # 리뷰작성일이 기간보다 과거 → 그만
                self.log("  → [L] 종료: 기간 이전 리뷰 도달")
                break
            page += 1
            guard += 1
            if guard > 200:
                break
            time.sleep(0.15)
        return reviews

    def collect_kkday(self, min_dt, max_dt):
        reviews = {}
        cap = self.capture_request(CHANNEL_META["KK"]["page"], "get_comment_list", wait=10)
        if not cap or not cap.get("body"):
            self.log("  ⚠ [KK] 요청 캡처 실패 (로그인/페이지 확인)")
            return reviews
        beg = pd.Timestamp(min_dt).strftime("%Y-%m-%d")
        end = pd.Timestamp(max_dt).strftime("%Y-%m-%d")
        js = """
        var cap=P[0], page=P[1], size=P[2], beg=P[3], end=P[4];
        var body={};
        try{ body=JSON.parse(cap.body)||{}; }catch(e){}
        body.begGoDate=beg; body.endGoDate=end;
        body.begRecDate=''; body.endRecDate='';   // release date 비움 (필수)
        body.currentPage=page; body.pageSize=size;
        var st=0, j={};
        try{
          var res=await fetch(cap.url, {method:'POST', credentials:'include',
              headers:Object.assign({}, cap.headers||{}, {'Content-Type':'application/json'}),
              body:JSON.stringify(body)});
          st=res.status; try{ j=await res.json(); }catch(e){ j={__parse:String(e)}; }
        }catch(e){ j={__fetch:String(e)}; }
        var rl=(j&&j.data&&j.data.recommandList)||[];
        var total=(j&&j.data&&(j.data.size!==undefined?j.data.size:null));
        var dbg=null;
        if(!rl.length){ dbg={status:st, keys:Object.keys(j||{}),
            msg:(j&&(j.msg||j.__parse||j.__fetch))||null,
            dataKeys:(j&&j.data)?Object.keys(j.data):null}; }
        done({rows:rl.map(function(x){return {code:x.orderMid, rating:x.recScore, id:x.recOid};}),
              total:total, dbg:dbg});
        """
        page, size, seen, total = 1, 50, set(), None
        while page <= 300:
            r = self._fetch_page(js, cap, page, size, beg, end)
            if not r or r.get("error"):
                self.log(f"  ⚠ [KK] p{page} 오류: {r.get('error') if r else 'no-resp'}")
                break
            if total is None:
                total = r.get("total")
            rows = r.get("rows") or []
            if not rows:
                if page == 1 and r.get("dbg"):
                    self.log(f"  🔎 [KK] 응답: {r.get('dbg')}")
                break
            for x in rows:
                key = x.get("id") or x.get("code")
                if key in seen:
                    continue
                seen.add(key)
                code = norm_code(x.get("code"))
                if code:
                    reviews[code] = self._fmt_rating(x.get("rating"))
            self.log(f"  → [KK] p{page}: {len(rows)}건 (누적 {len(reviews)}"
                     + (f" / {total}" if total else "") + ")")
            if total and len(seen) >= total:
                self.log("  → [KK] 종료: 전체 수집 완료")
                break
            if len(rows) < size:
                self.log("  → [KK] 종료: 페이지 끝")
                break
            page += 1
            time.sleep(0.15)
        return reviews

    def collect_ctrip(self, min_ms, wfrom, wto):
        reviews = {}
        cap = self.capture_request(CHANNEL_META["TPC"]["page"], "listOrderComments", wait=10)
        if not cap:
            self.log("  ⚠ [TPC] 요청 캡처 실패")
            return reviews
        js = """
        var cap=P[0], page=P[1], size=P[2];
        var body={sceneInfo:{bizScene:'ACTIVITY'}, paging:{pageNo:page, pageSize:size},
                  sorting:{orderBy:'COMMENT_TIME', desc:true}};
        try{ var b=JSON.parse(cap.body); if(b&&b.sceneInfo){ body.sceneInfo=b.sceneInfo; } }catch(e){}
        var res=await fetch(cap.url, {method:'POST', credentials:'include',
            headers:Object.assign({}, cap.headers||{}, {'Content-Type':'application/json'}),
            body:JSON.stringify(body)});
        var j=await res.json();
        var list=j.comments||[];
        done({rows:list.map(function(x){
              var cd=''; if(x.commentTime){ try{ cd=new Date(x.commentTime).toISOString().slice(0,10); }catch(e){} }
              return {code:String(x.orderId), rating:x.score, rdate:x.commentTime, cdate:cd};
        }), total:j.totalCount||null});
        """
        page, size, anon_ratings = 1, 50, []
        while page <= 300:
            r = self._fetch_page(js, cap, page, size)
            if not r or r.get("error"):
                self.log(f"  ⚠ [TPC] p{page} 오류: {r.get('error') if r else 'no-resp'}")
                break
            rows = r.get("rows") or []
            if not rows:
                break
            oldest_ok = True
            for x in rows:
                code = norm_code(x.get("code"))
                d = (x.get("cdate") or "")[:10]   # 리뷰작성일 기준
                keep = bool(d) and (wfrom <= d <= wto)   # 투어시작일 ~ 조회일
                if keep and code == "0":
                    anon_ratings.append(self._fmt_rating(x.get("rating")))  # 익명(주문번호 없음)
                if code and code != "0" and keep:
                    reviews[code] = self._fmt_rating(x.get("rating"))
                if not self._within(x.get("rdate"), min_ms):
                    oldest_ok = False
            self.log(f"  → [TPC] p{page}: {len(rows)}건 (기간내 매칭가능 {len(reviews)})")
            if not oldest_ok:
                self.log("  → [TPC] 종료: 기간 이전 리뷰 도달")
                break
            page += 1
            time.sleep(0.12)
        if anon_ratings:
            self.log(f"  · [TPC] 익명 리뷰(주문번호 없음) {len(anon_ratings)}건 (매칭 제외, 전체품질엔 포함)")
        return reviews, anon_ratings

    def collect_mrt(self, min_ms, date_set=None):
        reviews = {}
        # 페이지 이동으로 세션 토큰 확보
        self.driver.get(CHANNEL_META["MRT"]["page"])
        time.sleep(3)
        js = """
        var page=P[0], size=P[1];
        var tok=localStorage.getItem('accessToken');
        var res=await fetch('https://api3-backoffice.myrealtrip.com/review/partner/reviews/search',
            {method:'POST', credentials:'include',
             headers:{'Content-Type':'application/json','partner-access-token':tok},
             body:JSON.stringify({page:page, pageSize:size})});
        var j=await res.json();
        var list=Array.isArray(j.data)?j.data:[];
        done({rows:list.map(function(x){return {code:x.reservationNo, rating:x.score,
              rdate:x.createdAt, tdate:x.travelStartDate};}),
              total:(j.meta&&j.meta.totalCount)||null});
        """
        page, size = 1, 50
        while page <= 200:
            r = self._fetch_page(js, page, size)
            if not r or r.get("error"):
                self.log(f"  ⚠ [MRT] p{page} 오류: {r.get('error') if r else 'no-resp'}")
                break
            rows = r.get("rows") or []
            if not rows:
                break
            oldest_ok = True
            for x in rows:
                code = norm_code(x.get("code"))
                d = (x.get("tdate") or "")[:10]
                if code and (not date_set or d in date_set):
                    reviews[code] = self._fmt_rating(x.get("rating"))
                if not self._within_iso(x.get("rdate"), min_ms):
                    oldest_ok = False
            self.log(f"  → [MRT] p{page}: {len(rows)}건 (기간내 누적 {len(reviews)})")
            if len(rows) < size:
                self.log("  → [MRT] 종료: 페이지 끝(전량 수집)")
                break
            page += 1
            time.sleep(0.15)
        return reviews

    def collect_gg(self, gg_from, gg_to):
        reviews = {}
        cap = self.capture_request(CHANNEL_META["GG"]["page"], "/graphql",
                                   body_contains="bookingReference", wait=12)
        if not cap or not cap.get("body"):
            self.log("  ⚠ [GG] GraphQL 요청 캡처 실패")
            return reviews
        # GG는 리뷰에 정확한 투어날짜가 없어 활동일(travelDate) 범위(기간±1)로 서버 필터 후 전부 스캔
        js = """
        var cap=P[0], off=P[1], size=P[2], dfrom=P[3], dto=P[4];
        var payload={};
        try{ payload=JSON.parse(cap.body); }catch(e){}
        payload.variables = payload.variables || {};
        var inp = payload.variables.input || {};
        inp.travelDateFrom = dfrom;
        inp.travelDateTo = dto;
        inp.limit = size;
        inp.offset = off;
        payload.variables.input = inp;
        var st=0, txt="";
        try{
          var res=await fetch(cap.url, {method:'POST', credentials:'include',
              headers:Object.assign({}, cap.headers||{}, {'Content-Type':'application/json',
                  'apollo-require-preflight':'true',
                  'x-apollo-operation-name':(payload.operationName||'Reviews_ReviewSearch')}),
              body:JSON.stringify(payload)});
          st=res.status; txt=await res.text();
        }catch(e){ txt=''; }
        var j=null; try{ j=JSON.parse(txt); }catch(e){}
        function findArr(o,d){ if(d>12||!o||typeof o!=='object') return null;
          if(Array.isArray(o)&&o.length&&o[0]&&(o[0].bookingReference!==undefined||o[0].reviewId!==undefined)) return o;
          for(var k in o){ var r=findArr(o[k],d+1); if(r) return r; } return null; }
        var list=j?(findArr(j,0)||[]):[];
        var dbg=null;
        if(!list.length){ dbg={status:st, hasBody:(txt&&txt.length>0),
            errors:(j&&j.errors?JSON.stringify(j.errors).slice(0,200):null)}; }
        done({rows:list.map(function(x){return {code:x.bookingReference, rating:x.rating};}), dbg:dbg});
        """
        page, size = 1, 50
        while page <= 100:
            off = (page - 1) * size
            r = self._fetch_page(js, cap, off, size, gg_from, gg_to)
            if not r or r.get("error"):
                self.log(f"  ⚠ [GG] p{page} 오류: {r.get('error') if r else 'no-resp'}")
                break
            rows = r.get("rows") or []
            if not rows:
                if page == 1 and r.get("dbg"):
                    self.log(f"  🔎 [GG] 응답: {r.get('dbg')}")
                break
            for x in rows:
                code = norm_code(x.get("code"))
                if code:
                    reviews[code] = self._fmt_rating(x.get("rating"))
            self.log(f"  → [GG] p{page}: {len(rows)}건 (누적 {len(reviews)})")
            if len(rows) < size:
                self.log("  → [GG] 종료: 페이지 끝")
                break
            page += 1
            time.sleep(0.2)
        return reviews

    # ---------- 수집 보조 ----------
    @staticmethod
    def _fmt_rating(v):
        if v is None:
            return ""
        try:
            f = float(v)
            return str(int(f)) if f == int(f) else f"{f:.1f}"
        except (ValueError, TypeError):
            return str(v).strip()

    @staticmethod
    def _to_ms(v):
        """epoch(ms/s) 숫자 또는 날짜 문자열('YYYY-MM-DD ...', ISO, '... (GMT+9)') -> epoch ms."""
        if v is None or v == "":
            return None
        try:
            n = float(v)
            return int(n if n > 1e12 else n * 1000)
        except (ValueError, TypeError):
            pass
        try:
            t = str(v).split("(")[0].strip()
            return to_epoch_ms(pd.Timestamp(t))
        except Exception:
            return None

    @staticmethod
    def _within(value, min_ms):
        ms = PowerReviewApp._to_ms(value)
        return True if ms is None else ms >= min_ms

    @staticmethod
    def _within_iso(value, min_ms):
        return PowerReviewApp._within(value, min_ms)

    # ---------------------------------------------------------
    # 실행
    # ---------------------------------------------------------
    def start_processing(self):
        if not self.driver:
            messagebox.showerror("오류", "먼저 크롬을 연결하세요!")
            return
        if self.df is None:
            messagebox.showerror("오류", "먼저 엑셀을 선택하세요!")
            return
        sel_dates = self._selected_dates()
        sel_branches = self._selected_branches()
        sel_guides = [g for g, var in self.guide_vars.items() if var.get()]
        if not sel_dates:
            messagebox.showerror("오류", "최소 1개 날짜를 선택하세요!")
            return
        if not sel_branches:
            messagebox.showerror("오류", "최소 1개 지사를 선택하세요!")
            return
        if not sel_guides:
            messagebox.showerror("오류", "최소 1명 가이드를 선택하세요!")
            return

        df = self.df[
            self.df["Date"].dt.normalize().isin(sel_dates)
            & self.df["Area"].isin(sel_branches)
            & self.df["Main Guide"].isin(sel_guides)
        ].copy()
        if df.empty:
            messagebox.showerror("오류", "선택 조건에 해당하는 데이터가 없습니다.")
            return

        try:
            self.detail_lines = []
            self.result_text.delete(1.0, "end")
            self.log("📊 리뷰 조회 시작 (API)")
            self.log("=" * 70)

            df["Review_Status"] = ""
            df["Rating"] = ""
            df["Check"] = ""

            report_min = pd.Timestamp(min(sel_dates)).normalize()
            report_max = pd.Timestamp(max(sel_dates)).normalize()
            # 리뷰는 투어 이후 작성되므로 '작성일 >= 리포트 시작일 -1(GG ±1 버퍼)'까지 긁고 종료
            min_ms = to_epoch_ms(report_min - timedelta(days=1))
            date_set = {pd.Timestamp(ts).strftime("%Y-%m-%d") for ts in sel_dates}
            # GG는 활동일(travelDate) 범위 필터 = 리포트 기간 ±1일 (사용자 GG 로직)
            gg_from = (report_min - timedelta(days=1)).strftime("%Y-%m-%d")
            gg_to = (report_max + timedelta(days=1)).strftime("%Y-%m-%d")
            # TPC는 리뷰 '작성일' 기준 → 투어 시작일 ~ 조회일(오늘) 사이 작성 리뷰
            _today = pd.Timestamp.now().normalize()
            tpc_from = report_min.strftime("%Y-%m-%d")
            tpc_to = max(report_max, _today).strftime("%Y-%m-%d")

            self.log(f"📅 기간: {report_min.strftime('%Y-%m-%d')} ~ "
                     f"{report_max.strftime('%Y-%m-%d')}  "
                     f"(지사 {df['Area'].nunique()}개 · 예약 {len(df)}건)")
            self.log(f"   · TPC 리뷰작성일 검색: {tpc_from} ~ {tpc_to}")

            pw = self.create_progress_window()

            used = [c for c in CHANNELS if c in set(df["Agency"])]
            collected = {}
            extras = {}
            self.log("\n" + "=" * 70)
            self.log("1단계: 채널별 리뷰 수집")
            self.log("=" * 70)

            collectors = {
                "L": lambda: self.collect_klook(min_ms, date_set),
                "KK": lambda: self.collect_kkday(report_min, report_max),
                "GG": lambda: self.collect_gg(gg_from, gg_to),
                "TPC": lambda: self.collect_ctrip(min_ms, tpc_from, tpc_to),
                "MRT": lambda: self.collect_mrt(min_ms, date_set),
            }
            for i, ch in enumerate(used):
                pw.label.config(text=f"[{ch}] {CHANNEL_META[ch]['name']} 리뷰 수집 중...")
                pw.progress_bar["value"] = (i / max(len(used), 1)) * 100
                pw.window.update()
                self.log(f"\n🔍 [{ch}] {CHANNEL_META[ch]['name']}")
                try:
                    res = collectors[ch]()
                    if isinstance(res, tuple):
                        collected[ch], extras[ch] = res
                    else:
                        collected[ch], extras[ch] = res, []
                    self.log(f"  ✓ [{ch}] {len(collected[ch])}건 수집")
                except Exception as e:
                    collected[ch] = {}; extras[ch] = []
                    self.log(f"  ✗ [{ch}] 수집 실패: {e}")
                    traceback.print_exc()

            # 2단계: 매칭
            self.log("\n" + "=" * 70)
            self.log("2단계: 예약번호 매칭")
            self.log("=" * 70)
            def _is_good(rt):
                try:
                    return float(rt) >= 4
                except (ValueError, TypeError):
                    return False
            for idx, row in df.iterrows():
                ch = row["Agency"]
                code = norm_code(row["Agency Code"])
                if ch in collected:
                    if code in collected[ch]:
                        rt = collected[ch][code]
                        df.at[idx, "Rating"] = rt
                        if _is_good(rt):
                            df.at[idx, "Review_Status"] = "GOOD"
                            df.at[idx, "Check"] = "✓"          # 좋은리뷰(4-5)
                        else:
                            df.at[idx, "Review_Status"] = "LOW"
                            df.at[idx, "Check"] = "▲"          # 낮은평점(1-3) → 집계 제외
                    else:
                        df.at[idx, "Review_Status"] = "NO"
                        df.at[idx, "Check"] = "✗"
                else:
                    df.at[idx, "Review_Status"] = "SKIP"

            pw.window.destroy()

            self._result_df = df
            report = self.build_report(df, used, collected, extras)
            self.last_report_text = report
            self._build_copy_buttons()
            self.result_text.delete(1.0, "end")
            self.result_text.insert("end", report)
            self.result_text.see("1.0")

            # 자동 엑셀 저장
            saved = self.save_excel(df, ask=False)
            self.progress_var.set("✅ 완료" + (f" · 저장: {os.path.basename(saved)}" if saved else ""))
            messagebox.showinfo("완료", "리뷰 조회 완료!\n\n엑셀은 폴더에 자동 저장됐어요.\n아래 '전체 Copy' 또는 '지사 Copy' 버튼으로 복사해 보낼 수 있어요.")
        except Exception as e:
            self.progress_var.set(f"❌ 오류: {e}")
            self.log(f"오류: {e}")
            traceback.print_exc()
            messagebox.showerror("오류", f"처리 중 오류:\n{e}")

    def create_progress_window(self):
        w = Toplevel(self.root)
        w.title("처리 중...")
        w.geometry("420x110")
        lbl = Label(w, text="시작 중...", font=("Arial", 10))
        lbl.pack(pady=10)
        pb = Progressbar(w, length=380, mode="determinate")
        pb.pack(pady=10)
        w.protocol("WM_DELETE_WINDOW", lambda: None)
        w.progress_bar = pb
        w.label = lbl
        w.window = w
        return w

    # ---------------------------------------------------------
    # 리포트 (지사 → 날짜 → 투어/가이드)
    # ---------------------------------------------------------
    def _scope_label(self, df):
        by_country = {}
        for area in sorted(df["Area"].unique(), key=area_rank):
            by_country.setdefault(area_country(area), []).append(AREA_KR.get(area, str(area)))
        parts = []
        for c, _ in COUNTRY_GROUPS:
            if c in by_country:
                parts.append(f"{c}({', '.join(by_country[c])})")
        if "기타" in by_country:
            parts.append(f"기타({', '.join(by_country['기타'])})")
        return " · ".join(parts) if parts else "전체"

    def _summary_lines(self, sub, title):
        def ravg(vals):
            v = [float(x) for x in vals if str(x).replace('.', '', 1).isdigit()]
            return (sum(v) / len(v)) if v else 0
        out = []
        out.append("=" * 68)
        out.append(f"📈 {title}")
        out.append("=" * 68)
        tt = len(sub); tp = int(sub["People"].sum())
        rev = sub[sub["Agency"].isin(CHANNELS)]; nch = len(rev)
        chk = int((sub["Check"] == "✓").sum()); bad = int((sub["Check"] == "▲").sum())
        out.append(f"👥 총 예약: {tt}팀 {tp}명")
        out.append(f"✓ 좋은리뷰(4-5점): {chk}팀")
        if nch:
            out.append(f"   └ 5개 채널: {chk}/{nch}팀 ({chk/nch*100:.1f}%)")
        if tt:
            out.append(f"   └ 전체:     {chk}/{tt}팀 ({chk/tt*100:.1f}%)")
        if bad:
            out.append(f"▲ 낮은평점(1-3점): {bad}팀 (집계 제외)")
        avg = ravg(sub.loc[sub["Check"].isin(["✓", "▲"]), "Rating"].tolist())
        out.append(f"⭐ 평균 별점(전체 1-5): {avg:.1f}점" if avg else "⭐ 평균 별점: N/A")
        out.append("\n[채널별]")
        for ch in CHANNELS:
            cs = sub[sub["Agency"] == ch]
            if len(cs) == 0:
                continue
            c = int((cs["Check"] == "✓").sum())
            a = ravg(cs.loc[cs["Check"].isin(["✓", "▲"]), "Rating"].tolist())
            line = f"  {ch:4} {CHANNEL_META[ch]['name']:12} {c:3}/{len(cs):3}팀 ({c/len(cs)*100:5.1f}%)"
            if a:
                line += f"  평균 {a:.1f}"
            out.append(line)
        return out

    def build_report(self, df, used, collected, extras=None):
        L = []
        extras = extras or {}
        def rate_avg(vals):
            v = [float(x) for x in vals if str(x).replace('.', '').isdigit()]
            return (sum(v) / len(v)) if v else 0.0

        # ---- 전체 요약 ----
        total_teams = len(df)
        total_people = int(df["People"].sum())
        rev_mask = df["Agency"].isin(CHANNELS)
        rev_teams = int(rev_mask.sum())
        rev_people = int(df.loc[rev_mask, "People"].sum())
        checked = int((df["Check"] == "✓").sum())
        bad = int((df["Check"] == "▲").sum())
        all_ratings = df.loc[df["Check"].isin(["✓", "▲"]), "Rating"].tolist()

        L.append("=" * 68)
        L.append(f"📈 전체 요약  [{self._scope_label(df)}]")
        L.append("=" * 68)
        if self.noshow_teams:
            L.append(f"🚫 No Show(O) 제외: {self.noshow_teams}팀 {self.noshow_people}명")
        L.append(f"👥 총 예약: {total_teams}팀 {total_people}명")
        L.append(f"   └ 리뷰 조회대상(L/KK/GG/TPC/MRT): {rev_teams}팀 {rev_people}명")
        L.append(f"✓ 좋은리뷰(4-5점): {checked}팀")
        if rev_teams:
            L.append(f"   └ 5개 채널 기준: {checked}/{rev_teams}팀 ({checked/rev_teams*100:.1f}%)")
        if total_teams:
            L.append(f"   └ 전체 기준:   {checked}/{total_teams}팀 ({checked/total_teams*100:.1f}%)")
        if bad:
            L.append(f"▲ 낮은평점(1-3점): {bad}팀 (집계 제외)")
        avg = rate_avg(all_ratings)
        L.append(f"⭐ 평균 별점(전체 1-5): {avg:.1f}점" if avg else "⭐ 평균 별점: N/A")

        # 채널별
        L.append("\n[채널별]")
        for ch in CHANNELS:
            sub = df[df["Agency"] == ch]
            if len(sub) == 0:
                continue
            c = int((sub["Check"] == "✓").sum())
            a = rate_avg(sub.loc[sub["Check"].isin(["✓", "▲"]), "Rating"].tolist())
            line = f"  {ch:4} {CHANNEL_META[ch]['name']:12} {c:3}/{len(sub):3}팀 ({c/len(sub)*100:5.1f}%)"
            if a:
                line += f"  평균 {a:.1f}"
            L.append(line)

        # 전체 리뷰 품질 (채널 단위 · 예약매칭 무관 · 익명 포함)
        L.append("\n[전체 리뷰 품질 (채널 단위 · 익명 포함)]")
        for ch in CHANNELS:
            if ch not in collected:
                continue
            ratings = list(collected[ch].values()) + list(extras.get(ch, []))
            if not ratings:
                continue
            vals = [float(x) for x in ratings if str(x).replace('.', '', 1).isdigit()]
            n = len(ratings)
            qa = (sum(vals) / len(vals)) if vals else 0
            goodn = sum(1 for v in vals if v >= 4)
            badn = sum(1 for v in vals if v < 4)
            line = (f"  {ch:4} {CHANNEL_META[ch]['name']:12} 리뷰 {n:3}건 · 평균 {qa:.1f}"
                    f" · 좋은리뷰(4-5) {goodn} · 나쁜리뷰(1-3) {badn}")
            ex = len(extras.get(ch, []))
            if ex:
                line += f"  (익명 {ex} 포함)"
            L.append(line)

        # 조회 제외 에이전시
        others = df[~df["Agency"].isin(CHANNELS)]
        if len(others):
            L.append("\n[조회 제외 에이전시]")
            for ag, g in others.groupby("Agency"):
                L.append(f"  {ag:10} {len(g):3}팀 {int(g['People'].sum()):3}명 (개별 확인 필요)")

        # ---- 국가 → 지사 → 날짜 → 가이드/투어 ----
        self._branch_texts = {}
        cur_country = None
        for area in sorted(df["Area"].unique(), key=area_rank):
            adf = df[df["Area"] == area]
            country = area_country(area)
            if country != cur_country:
                L.append("\n" + "#" * 68)
                L.append(f"🌏 {country}")
                L.append("#" * 68)
                cur_country = country
            b_start = len(L)   # 지사 섹션 시작 (지사별 Copy용)
            a_rev = adf[adf["Agency"].isin(CHANNELS)]
            a_chk = int((adf["Check"] == "✓").sum())
            a_bad = int((adf["Check"] == "▲").sum())
            a_avg = rate_avg(adf.loc[adf["Check"].isin(["✓", "▲"]), "Rating"].tolist())
            n_all = len(adf); n_ch = len(a_rev)
            L.append("\n" + "=" * 68)
            head = f"🏢 [지사: {area}]  예약 {n_all}팀 {int(adf['People'].sum())}명"
            if n_ch:
                head += f"  ·  5채널 {a_chk}/{n_ch} ({a_chk/n_ch*100:.1f}%)"
            head += f"  ·  전체 {a_chk}/{n_all} ({a_chk/n_all*100:.1f}%)"
            if a_avg:
                head += f"  ·  평균 {a_avg:.1f}"
            if a_bad:
                head += f"  ·  낮은평점 {a_bad}"
            L.append(head)
            L.append("=" * 68)

            for date_val, ddf in adf.groupby(adf["Date"].dt.normalize()):
                L.append(f"\n  ── {pd.Timestamp(date_val).strftime('%Y-%m-%d (%a)')} ──")
                for (guide, product), g in ddf.groupby(["Main Guide", "Product"]):
                    teams = len(g)
                    people = int(g["People"].sum())
                    L.append(f"  • {guide} / {product} / {teams}팀 {people}명")
                    g5 = g[g["Agency"].isin(CHANNELS)]
                    n5 = len(g5); c5 = int((g5["Check"] == "✓").sum())
                    nAll = len(g); cAll = int((g["Check"] == "✓").sum())
                    tavg = rate_avg(g.loc[g["Check"].isin(["✓", "▲"]), "Rating"].tolist())
                    roll = "      ▶ "
                    if n5:
                        roll += f"5채널 {c5}/{n5} ({c5/n5*100:.0f}%) · "
                    roll += f"전체 {cAll}/{nAll} ({cAll/nAll*100:.0f}%)"
                    if tavg:
                        roll += f" · 평균 {tavg:.1f}"
                    L.append(roll)
                    for ch in CHANNELS:
                        cg = g[g["Agency"] == ch]
                        if len(cg) == 0:
                            continue
                        c = int((cg["Check"] == "✓").sum())
                        pct = f" ({c/len(cg)*100:.0f}%)" if len(cg) else ""
                        L.append(f"      [{ch}] {c}/{len(cg)}{pct}")
                        for _, r in cg.iterrows():
                            code = r["Agency Code"]; rt = r["Rating"]
                            if r["Check"] == "✓":
                                L.append(f"        ✓{code}({rt})")
                            elif r["Check"] == "▲":
                                L.append(f"        ▲{code}({rt}) 낮은평점")
                            else:
                                L.append(f"        ✗{code}")
                    # 조회 제외 에이전시
                    og = g[~g["Agency"].isin(CHANNELS)]
                    for ag, gg in og.groupby("Agency"):
                        codes = ", ".join(gg["Agency Code"].tolist())
                        L.append(f"      [{ag}] {len(gg)}팀 (개별확인): {codes}")

            _b_detail = "\n".join(L[b_start:])
            _b_summary = "\n".join(self._summary_lines(adf, f"지사: {AREA_KR.get(area, str(area))} 요약"))
            self._branch_texts[area] = _b_summary + "\n\n" + _b_detail   # 지사 Copy = 요약+상세

        L.append("\n" + "=" * 68)
        return "\n".join(L)

    # ---------------------------------------------------------
    # 엑셀 저장 / Copy
    # ---------------------------------------------------------
    @staticmethod
    def _autofit(ws):
        from openpyxl.utils import get_column_letter
        for col in ws.columns:
            m = 0
            letter = get_column_letter(col[0].column)
            for cell in col:
                if cell.value is None:
                    continue
                w = sum(2 if ord(ch) > 0x1100 else 1 for ch in str(cell.value))
                if w > m:
                    m = w
            ws.column_dimensions[letter].width = min(max(m + 2, 8), 60)

    @staticmethod
    def _safe_name(s):
        s = re.sub(r'[\\/:*?"<>|]', "_", str(s)).strip()
        return (s or "unnamed")[:80]

    def _write_sheet(self, path, df_out, sheet="Sheet1"):
        with pd.ExcelWriter(path, engine="openpyxl") as w:
            df_out.to_excel(w, index=False, sheet_name=sheet)
            self._autofit(w.sheets[sheet])

    def _guide_frame(self, sub, guide_cols):
        """가이드 1명 시트: 데이터 위, 요약(전체기간 → 날짜별) 아래에 모아서."""
        def _avg(vals):
            v = [float(x) for x in vals if str(x).replace(".", "", 1).isdigit()]
            return round(sum(v) / len(v), 1) if v else 0

        def _stat(d):
            s5 = d[d["Agency"].isin(CHANNELS)]
            y5 = len(s5); x5 = int((s5["Check"] == "✓").sum())
            a5 = _avg(s5.loc[s5["Check"].isin(["✓", "▲"]), "Rating"].tolist())
            p5 = round(x5 / y5 * 100, 1) if y5 else 0
            yA = len(d); xA = int((d["Check"] == "✓").sum())
            aA = _avg(d.loc[d["Check"].isin(["✓", "▲"]), "Rating"].tolist())
            pA = round(xA / yA * 100, 1) if yA else 0
            return (x5, y5, p5, a5, xA, yA, pA, aA)

        sub = sub.sort_values(by=["Date"])
        blank = {c: "" for c in guide_cols}
        rows = []
        # 1) 데이터 (날짜 바뀌면 빈 줄)
        prev = None
        for _, row in sub.iterrows():
            if prev is not None and row["Date"] != prev:
                rows.append(dict(blank))
            rows.append({c: row[c] for c in guide_cols})
            prev = row["Date"]
        # 2) 요약: 전체기간 먼저
        rows.append(dict(blank))
        rows.append(dict(blank))
        x5, y5, p5, a5, xA, yA, pA, aA = _stat(sub)
        rows.append({**blank, "Date": "[전체기간 5채널]", "Product": f"리뷰 {x5}/{y5} ({p5}%) · 평균 {a5}점"})
        rows.append({**blank, "Date": "[전체기간 전체]", "Product": f"리뷰 {xA}/{yA} ({pA}%) · 평균 {aA}점"})
        rows.append(dict(blank))
        # 3) 날짜별 소계
        for date_val, ddf in sub.groupby("Date"):
            x5, y5, p5, a5, xA, yA, pA, aA = _stat(ddf)
            rows.append({**blank, "Date": f"[{date_val} 5채널]", "Product": f"리뷰 {x5}/{y5} ({p5}%) · 평균 {a5}점"})
            rows.append({**blank, "Date": f"[{date_val} 전체]", "Product": f"리뷰 {xA}/{yA} ({pA}%) · 평균 {aA}점"})
            rows.append(dict(blank))
        return pd.DataFrame(rows, columns=guide_cols)

    def _ensure_pdf_font(self):
        """PDF용 한글 폰트 1회 등록. TTF(맑은고딕 등) 임베드 우선, 없으면 내장 CID 폴백."""
        if getattr(self, "_pdf_font", None) is not None:
            return self._pdf_font
        try:
            from reportlab.pdfbase import pdfmetrics
            from reportlab.pdfbase.ttfonts import TTFont
        except Exception:
            self._pdf_font = False
            return False
        ttf_candidates = [
            r"C:\Windows\Fonts\malgun.ttf",
            r"C:\Windows\Fonts\malgunbd.ttf",
            r"C:\Windows\Fonts\NanumGothic.ttf",
            "/System/Library/Fonts/Supplemental/AppleGothic.ttf",
            "/Library/Fonts/AppleGothic.ttf",
            "/usr/share/fonts/truetype/nanum/NanumGothic.ttf",
        ]
        for p in ttf_candidates:
            try:
                if os.path.exists(p):
                    pdfmetrics.registerFont(TTFont("KRFONT", p))
                    pdfmetrics.registerFontFamily("KRFONT", normal="KRFONT", bold="KRFONT",
                                                  italic="KRFONT", boldItalic="KRFONT")
                    self._pdf_font = "KRFONT"
                    return self._pdf_font
            except Exception:
                continue
        # 폴백: reportlab 내장 한국어 CID 폰트 (뷰어 한글팩 필요)
        try:
            from reportlab.pdfbase.cidfonts import UnicodeCIDFont
            pdfmetrics.registerFont(UnicodeCIDFont("HYSMyeongJo-Medium"))
            pdfmetrics.registerFontFamily("HYSMyeongJo-Medium", normal="HYSMyeongJo-Medium",
                                          bold="HYSMyeongJo-Medium", italic="HYSMyeongJo-Medium",
                                          boldItalic="HYSMyeongJo-Medium")
            self._pdf_font = "HYSMyeongJo-Medium"
            return self._pdf_font
        except Exception:
            self._pdf_font = False
            return False

    def _ensure_pdf_symbol_font(self):
        """✓/✗/▲ 표시용 심볼 폰트. Segoe UI Symbol 우선, 없으면 False(→O/X 치환)."""
        if getattr(self, "_pdf_sym", None) is not None:
            return self._pdf_sym
        try:
            from reportlab.pdfbase import pdfmetrics
            from reportlab.pdfbase.ttfonts import TTFont
        except Exception:
            self._pdf_sym = False
            return False
        cands = [
            r"C:\Windows\Fonts\seguisym.ttf",   # Windows: Segoe UI Symbol (✓✗▲ 포함)
            r"C:\Windows\Fonts\arialuni.ttf",
            "/System/Library/Fonts/Apple Symbols.ttf",
            "/usr/local/lib/python3.10/dist-packages/matplotlib/mpl-data/fonts/ttf/DejaVuSans.ttf",
        ]
        for pth in cands:
            try:
                if os.path.exists(pth):
                    pdfmetrics.registerFont(TTFont("SYMFONT", pth))
                    self._pdf_sym = "SYMFONT"
                    return self._pdf_sym
            except Exception:
                continue
        self._pdf_sym = False
        return False

    def _write_guide_pdf(self, path, gdf, title):
        """가이드 프레임을 표 PDF로 저장. reportlab/폰트 없으면 조용히 False."""
        font = self._ensure_pdf_font()
        if not font:
            return False
        try:
            from reportlab.lib.pagesizes import A4, landscape
            from reportlab.lib import colors
            from reportlab.lib.units import mm
            from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
            from reportlab.lib.styles import ParagraphStyle

            def esc(x):
                return str(x).replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")

            cols = list(gdf.columns)
            styN = ParagraphStyle("n", fontName=font, fontSize=7.5, leading=9.5)
            styT = ParagraphStyle("t", fontName=font, fontSize=13, leading=16)
            pidx = cols.index("Product") if "Product" in cols else -1
            cidx = cols.index("Check") if "Check" in cols else -1
            sym = self._ensure_pdf_symbol_font()
            _symmap = {"✓": "O", "✗": "X"}
            _hdr = ["Review Status" if _c == "Review_Status" else _c for _c in cols]
            data = [_hdr]
            sumrows = []
            for i, (_, r) in enumerate(gdf.iterrows(), start=1):
                vals = ["" if pd.isna(r[c]) else str(r[c]) for c in cols]
                if vals and str(vals[0]).startswith("["):
                    sumrows.append(i)
                if cidx >= 0 and not sym:
                    vals[cidx] = _symmap.get(vals[cidx], vals[cidx])
                if pidx >= 0:
                    vals[pidx] = Paragraph(esc(vals[pidx]), styN)
                data.append(vals)
            wmap = {"Date": 24, "Area": 15, "Product": 88, "Agency": 15, "Agency Code": 32,
                    "Review_Status": 28, "Rating": 16, "Check": 16}
            widths = [wmap.get(c, 20) * mm for c in cols]
            t = Table(data, colWidths=widths, repeatRows=1)
            st = [("FONTNAME", (0, 0), (-1, -1), font), ("FONTSIZE", (0, 0), (-1, -1), 7.5),
                  ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
                  ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#4472C4")),
                  ("GRID", (0, 0), (-1, -1), 0.4, colors.HexColor("#BBBBBB")),
                  ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                  ("ALIGN", (0, 0), (-1, -1), "CENTER"),
                  ("LEFTPADDING", (0, 0), (-1, -1), 3), ("RIGHTPADDING", (0, 0), (-1, -1), 3),
                  ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#F5F7FB")])]
            if pidx >= 0:
                st.append(("ALIGN", (pidx, 0), (pidx, -1), "LEFT"))   # Product(투어명)만 좌측
            if sym and cidx >= 0:
                st.append(("FONTNAME", (cidx, 1), (cidx, -1), sym))   # Check ✓/✗/▲ 심볼 폰트
            for ri in sumrows:
                st.append(("BACKGROUND", (0, ri), (-1, ri), colors.HexColor("#FFF2CC")))
            t.setStyle(TableStyle(st))
            doc = SimpleDocTemplate(path, pagesize=landscape(A4),
                                    leftMargin=10 * mm, rightMargin=10 * mm,
                                    topMargin=10 * mm, bottomMargin=10 * mm)
            doc.build([Paragraph(esc(title), styT), Spacer(1, 6), t])
            return True
        except Exception as e:
            self.log(f"  ⚠ PDF 생성 실패({os.path.basename(path)}): {e}")
            return False

    def save_excel(self, df, ask=False):
        try:
            base_dir = DOWNLOAD_DIR
            if ask:
                d = filedialog.askdirectory(title="저장 폴더 선택", initialdir=base_dir)
                if d:
                    base_dir = d
            out = df.copy()
            out["Date"] = pd.to_datetime(out["Date"]).dt.strftime("%Y-%m-%d")
            ranks = out["Area"].apply(area_rank)
            out["__c"] = ranks.apply(lambda t: t[0])
            out["__a"] = ranks.apply(lambda t: t[1] if isinstance(t[1], int) else 99)

            start = out["Date"].min().replace("-", "")
            end = out["Date"].max().replace("-", "")
            root = os.path.join(base_dir, f"Review Crawling Result {start}-{end}")
            os.makedirs(root, exist_ok=True)

            summary_cols = [c for c in ["Date", "Area", "Main Guide", "Product",
                                        "Agency", "Agency Code", "Review_Status", "Rating", "Check"]
                            if c in out.columns]
            branch_cols = [c for c in ["Date", "Main Guide", "Product",
                                       "Agency", "Agency Code", "Review_Status", "Rating", "Check"]
                           if c in out.columns]
            guide_cols = [c for c in ["Date", "Area", "Product", "Agency",
                                      "Agency Code", "Review_Status", "Rating", "Check"]
                          if c in out.columns]

            # 종합 (루트) : 종합 시트 + 지사별 탭 (가이드 탭 없음)
            comp = out.sort_values(by=["__c", "__a", "Main Guide", "Date"])
            _rg = _region_tag(out["Area"].unique())
            jpath = os.path.join(root, f"종합_{_rg}.xlsx" if _rg else "종합.xlsx")
            used_j = {"종합"}

            def _uqj(nm):
                nm = re.sub(r'[\\/*?:\[\]]', " ", str(nm)).strip()[:31] or "sheet"
                base = nm; k = 2
                while nm in used_j:
                    nm = f"{base[:28]}_{k}"; k += 1
                used_j.add(nm)
                return nm

            with pd.ExcelWriter(jpath, engine="openpyxl") as w:
                comp[summary_cols].rename(columns={"Review_Status": "Review Status"}).to_excel(w, sheet_name="종합", index=False)
                self._autofit(w.sheets["종합"])
                for area in sorted(out["Area"].unique(), key=area_rank):
                    bsub = out[out["Area"] == area].sort_values(by=["Main Guide", "Date"])
                    sh = _uqj(AREA_KR.get(area, str(area)))
                    bsub[branch_cols].rename(columns={"Review_Status": "Review Status"}).to_excel(w, sheet_name=sh, index=False)
                    self._autofit(w.sheets[sh])
            n = 1

            # 지사 폴더(서울/부산…)를 루트 바로 아래에 → 지사 전체 + 가이드별 함께
            for area in sorted(out["Area"].unique(), key=area_rank):
                bkr = AREA_KR.get(area, str(area))
                bdir = os.path.join(root, self._safe_name(bkr))   # 국가 폴더 없이 지사 폴더 바로
                os.makedirs(bdir, exist_ok=True)
                adf = out[out["Area"] == area].sort_values(by=["Main Guide", "Date"])
                # 지사 전체 파일 (전체 시트 + 가이드 탭) → 지사 폴더에 저장
                bpath = os.path.join(bdir, f"0. {self._safe_name(bkr)}_전체.xlsx")   # 맨 위로 오도록 0. 접두
                used_sheets = {"전체"}

                def _uq(nm):
                    nm = re.sub(r'[\\/*?:\[\]]', " ", str(nm)).strip()[:31] or "sheet"
                    base = nm; k = 2
                    while nm in used_sheets:
                        nm = f"{base[:28]}_{k}"; k += 1
                    used_sheets.add(nm)
                    return nm

                with pd.ExcelWriter(bpath, engine="openpyxl") as w:
                    adf[branch_cols].rename(columns={"Review_Status": "Review Status"}).to_excel(w, sheet_name="전체", index=False)
                    self._autofit(w.sheets["전체"])
                    for g in sorted(adf["Main Guide"].unique()):
                        gdf = self._guide_frame(adf[adf["Main Guide"] == g], guide_cols)
                        sh = _uq(g)
                        gdf.rename(columns={"Review_Status": "Review Status"}).to_excel(w, sheet_name=sh, index=False)
                        self._autofit(w.sheets[sh])
                n += 1
                # 가이드별 개별 파일 → 지사 폴더에 저장 (엑셀 + PDF)
                for g in sorted(adf["Main Guide"].unique()):
                    gdf = self._guide_frame(adf[adf["Main Guide"] == g], guide_cols)
                    self._write_sheet(os.path.join(bdir, f"{self._safe_name(g)}.xlsx"), gdf.rename(columns={"Review_Status": "Review Status"}), "리뷰")
                    n += 1
                    if self._write_guide_pdf(os.path.join(bdir, f"{self._safe_name(g)}.pdf"),
                                             gdf, f"가이드: {g}  ·  {bkr}"):
                        n += 1
                # 지사 폴더 압축 (바로 전달용): 결과 루트에 {지사}.zip 생성
                try:
                    import shutil
                    _zbase = os.path.join(root, self._safe_name(bkr))
                    if os.path.exists(_zbase + ".zip"):
                        os.remove(_zbase + ".zip")
                    shutil.make_archive(_zbase, "zip", root_dir=root, base_dir=self._safe_name(bkr))
                    self.log(f"  🗜 {bkr}.zip 생성")
                except Exception as _e:
                    self.log(f"  ⚠ {bkr} 압축 실패: {_e}")

            self.log(f"  💾 저장 완료: {root}  (엑셀 {n}개)")
            return root
        except Exception as e:
            self.log(f"  ⚠ 엑셀 저장 실패: {e}")
            traceback.print_exc()
            return None

    def save_excel_dialog(self):
        if self._result_df is None:
            messagebox.showwarning("경고", "저장할 결과가 없습니다. 먼저 조회를 실행하세요.")
            return
        p = self.save_excel(self._result_df, ask=True)
        if p:
            messagebox.showinfo("성공", f"저장 완료:\n{p}")

    def _build_copy_buttons(self):
        for w in self.copy_btn_frame.winfo_children():
            w.destroy()
        Button(self.copy_btn_frame, text="📋 전체 Copy",
               command=lambda: self._copy_text(self.last_report_text, "전체"),
               bg="#9C27B0", fg="white").pack(side="left", padx=4)
        for area in sorted(getattr(self, "_branch_texts", {}).keys(), key=area_rank):
            bkr = AREA_KR.get(area, str(area))
            Button(self.copy_btn_frame, text=f"{bkr} Copy",
                   command=lambda a=area: self._copy_text(
                       self._branch_texts.get(a, ""), AREA_KR.get(a, str(a)))
                   ).pack(side="left", padx=3)

    def _copy_text(self, text, label):
        if not text or not str(text).strip():
            messagebox.showwarning("경고", "복사할 내용이 없습니다. 먼저 조회하세요.")
            return
        self.root.clipboard_clear()
        self.root.clipboard_append(text)
        self.root.update()
        messagebox.showinfo("복사됨", f"✅ [{label}] 결과를 클립보드에 복사했습니다.")

    def _search_clear(self):
        if hasattr(self, "search_var"):
            self.search_var.set("")
        self.result_text.delete(1.0, "end")
        self.result_text.insert("end", self.last_report_text or "")
        self.result_text.see("1.0")

    def _search_guide(self):
        q = self.search_var.get().strip()
        self.result_text.delete(1.0, "end")
        if not q:
            self.result_text.insert("end", self.last_report_text or "")
            self.result_text.see("1.0")
            return
        if self._result_df is None:
            self.result_text.insert("end", "먼저 조회를 실행하세요.")
            return

        def _ravg(vals):
            v = [float(x) for x in vals if str(x).replace('.', '', 1).isdigit()]
            return (sum(v) / len(v)) if v else 0

        sub = self._result_df[self._result_df["Main Guide"].astype(str).str.contains(q, case=False, na=False)]
        if sub.empty:
            self.result_text.insert("end", f"🔍 '{q}' 와 일치하는 가이드가 없습니다.")
            return
        L = []
        for guide in sorted(sub["Main Guide"].unique()):
            gsub = sub[sub["Main Guide"] == guide]
            L.extend(self._summary_lines(gsub, f"가이드: {guide}"))
            for area in sorted(gsub["Area"].unique(), key=area_rank):
                asub = gsub[gsub["Area"] == area]
                for date_val, ddf in asub.groupby(asub["Date"].dt.normalize()):
                    L.append(f"\n  ── [{AREA_KR.get(area, str(area))}] "
                             f"{pd.Timestamp(date_val).strftime('%Y-%m-%d (%a)')} ──")
                    for product, g in ddf.groupby("Product"):
                        teams = len(g); people = int(g["People"].sum())
                        L.append(f"  • {product} / {teams}팀 {people}명")
                        g5 = g[g["Agency"].isin(CHANNELS)]
                        n5 = len(g5); c5 = int((g5["Check"] == "✓").sum())
                        nAll = len(g); cAll = int((g["Check"] == "✓").sum())
                        tavg = _ravg(g.loc[g["Check"].isin(["✓", "▲"]), "Rating"].tolist())
                        roll = "      ▶ "
                        if n5:
                            roll += f"5채널 {c5}/{n5} ({c5/n5*100:.0f}%) · "
                        roll += f"전체 {cAll}/{nAll} ({cAll/nAll*100:.0f}%)"
                        if tavg:
                            roll += f" · 평균 {tavg:.1f}"
                        L.append(roll)
                        for ch in CHANNELS:
                            cg = g[g["Agency"] == ch]
                            if len(cg) == 0:
                                continue
                            c = int((cg["Check"] == "✓").sum())
                            L.append(f"      [{ch}] {c}/{len(cg)} ({c/len(cg)*100:.0f}%)")
                            for _, r in cg.iterrows():
                                code = r["Agency Code"]; rt = r["Rating"]
                                if r["Check"] == "✓":
                                    L.append(f"        ✓{code}({rt})")
                                elif r["Check"] == "▲":
                                    L.append(f"        ▲{code}({rt}) 낮은평점")
                                else:
                                    L.append(f"        ✗{code}")
                        og = g[~g["Agency"].isin(CHANNELS)]
                        for ag, gg in og.groupby("Agency"):
                            L.append(f"      [{ag}] {len(gg)}팀 (개별확인): "
                                     + ", ".join(gg["Agency Code"].tolist()))
            L.append("")
        self.result_text.insert("end", "\n".join(L))
        self.result_text.see("1.0")

    def copy_results(self):
        txt = self.result_text.get(1.0, "end-1c")
        if not txt.strip():
            messagebox.showwarning("경고", "복사할 결과가 없습니다.")
            return
        self.root.clipboard_clear()
        self.root.clipboard_append(txt)
        self.root.update()
        messagebox.showinfo("성공", "✅ 전체 결과가 클립보드에 복사되었습니다.\n각 지사 섹션을 잘라 붙여넣기 하세요.")

    # =========================================================
    # Monthly / New Year Party  (간단 모드: 4~5점만, 리뷰내용 포함)
    # =========================================================
    def _run_dispatch(self):
        m = self.mode_var.get()
        if m == "PFP":
            self.start_processing()
        else:
            self.start_processing_simple(m)

    def collect_klook_simple(self, dfrom, dto, date_mode):
        """L: date_type 서버필터(ReviewTime/ParticipantTime) + 4~5점 + 내용."""
        reviews = {}
        dtype = "ReviewTime" if date_mode == "review_date" else "ParticipantTime"
        cap = self.capture_request(CHANNEL_META["L"]["page"], "review_list", wait=10)
        if not cap:
            self.log("  ⚠ [L] 캡처 실패"); return reviews
        js = """
        var cap=P[0], page=P[1], limit=P[2], dtype=P[3], sd=P[4], ed=P[5];
        var u=new URL(cap.url, location.origin);
        u.searchParams.set('date_type', dtype);
        u.searchParams.set('start_date', sd);
        u.searchParams.set('end_date', ed);
        u.searchParams.set('stars','0');
        u.searchParams.set('page', String(page));
        u.searchParams.set('limit', String(limit));
        var res=await fetch(u.toString(), {credentials:'include', headers:(cap.headers||{})});
        var j=await res.json();
        var list=(j.result&&j.result.review_list)||j.review_list||[];
        var total=(j.result&&j.result.total)||null;
        done({rows:list.map(function(x){return {code:x.booking_no, rating:x.stars, content:x.review||'', rdate:x.review_time};}), total:total});
        """
        # Klook은 요청 limit과 무관하게 페이지당 고정 개수를 주므로
        # 'len<limit 종료'가 아니라 total 기준으로 끝까지 페이지네이션한다.
        page, limit, total, got = 1, 50, None, 0
        while page <= 4000:
            r = self._fetch_page(js, cap, page, limit, dtype, dfrom, dto)
            if not r or r.get("error"):
                self.log(f"  ⚠ [L] p{page} 오류: {r.get('error') if r else 'no-resp'}"); break
            rows = r.get("rows") or []
            if total is None:
                total = r.get("total")
            if not rows:
                break
            for x in rows:
                try:
                    if float(x.get("rating")) < 0:
                        continue
                except (ValueError, TypeError):
                    continue
                code = norm_code(x.get("code"))
                if not code:
                    continue
                reviews[code] = {"rating": self._fmt_rating(x.get("rating")),
                                 "content": (x.get("content") or "").strip(),
                                 "rdate": (x.get("rdate") or "")[:10]}
            got += len(rows)
            self.log(f"  → [L] p{page}: {len(rows)}행 (전체 {got}" + (f"/{total}" if total else "") + f" · 수집 누적 {len(reviews)})")
            if total and got >= total:
                break
            page += 1
            time.sleep(0.1)
        return reviews

    def collect_kkday_simple(self, dfrom, dto, date_mode):
        """KK: begGoDate(참여)/begRecDate(리뷰작성) 서버필터 + recScores[4,5] + 내용."""
        reviews = {}
        cap = self.capture_request(CHANNEL_META["KK"]["page"], "get_comment_list", wait=10)
        if not cap or not cap.get("body"):
            self.log("  ⚠ [KK] 캡처 실패"); return reviews
        rev = (date_mode == "review_date")
        js = """
        var cap=P[0], page=P[1], size=P[2], sd=P[3], ed=P[4], rev=P[5];
        var body={}; try{ body=JSON.parse(cap.body)||{}; }catch(e){}
        if(rev){ body.begRecDate=sd; body.endRecDate=ed; body.begGoDate=''; body.endGoDate=''; }
        else { body.begGoDate=sd; body.endGoDate=ed; body.begRecDate=''; body.endRecDate=''; }
        body.recScores=[1,2,3,4,5];
        body.orderMid=''; body.prodOid=''; body.contactEmail=''; body.recImg='';
        body.currentPage=page; body.pageSize=size;
        var res=await fetch(cap.url,{method:'POST',credentials:'include',
            headers:Object.assign({}, cap.headers||{}, {'Content-Type':'application/json'}),
            body:JSON.stringify(body)});
        var j=await res.json();
        var rl=(j&&j.data&&j.data.recommandList)||[];
        var total=(j&&j.data&&(j.data.size!==undefined?j.data.size:null));
        done({rows:rl.map(function(x){return {code:x.orderMid, rating:x.recScore, title:x.recTitle||'', desc:x.recDesc||'', rdate:x.userRecDt||'', id:x.recOid};}), total:total});
        """
        page, size, seen, total = 1, 50, set(), None
        while page <= 3000:
            r = self._fetch_page(js, cap, page, size, dfrom, dto, rev)
            if not r or r.get("error"):
                self.log(f"  ⚠ [KK] p{page} 오류: {r.get('error') if r else 'no-resp'}"); break
            if total is None:
                total = r.get("total")
            rows = r.get("rows") or []
            if not rows:
                break
            for x in rows:
                key = x.get("id") or x.get("code")
                if key in seen:
                    continue
                seen.add(key)
                code = norm_code(x.get("code"))
                if not code:
                    continue
                content = ((x.get("title") or "").strip() + "\n" + (x.get("desc") or "").strip()).strip()
                reviews[code] = {"rating": self._fmt_rating(x.get("rating")),
                                 "content": content,
                                 "rdate": (x.get("rdate") or "")[:10]}
            self.log(f"  → [KK] p{page}: {len(rows)}행 (누적 {len(reviews)}" + (f"/{total}" if total else "") + ")")
            if total and len(seen) >= total:
                break
            if len(rows) < size:
                break
            page += 1
            time.sleep(0.15)
        return reviews

    def collect_gg_simple(self, dfrom, dto, date_mode):
        """GG: travelDate(참여)/reviewDate(작성) 서버필터 + ratings[4,5] + 내용."""
        reviews = {}
        cap = self.capture_request(CHANNEL_META["GG"]["page"], "/graphql",
                                   body_contains="bookingReference", wait=12)
        if not cap or not cap.get("body"):
            self.log("  ⚠ [GG] 캡처 실패"); return reviews
        rev = (date_mode == "review_date")
        # 참여일(활동일) 모드는 옛 도구와 동일하게 시작일 -1일 버퍼
        sd = dfrom if rev else (pd.Timestamp(dfrom) - timedelta(days=1)).strftime("%Y-%m-%d")
        ed = dto
        js = """
        var cap=P[0], off=P[1], size=P[2], sd=P[3], ed=P[4], rev=P[5];
        var payload={}; try{ payload=JSON.parse(cap.body); }catch(e){}
        payload.variables = payload.variables || {};
        var inp = payload.variables.input || {};
        if(rev){ inp.reviewDateFrom=sd; inp.reviewDateTo=ed; delete inp.travelDateFrom; delete inp.travelDateTo; }
        else { inp.travelDateFrom=sd; inp.travelDateTo=ed; delete inp.reviewDateFrom; delete inp.reviewDateTo; }
        inp.ratings=[1,2,3,4,5]; inp.limit=size; inp.offset=off;
        payload.variables.input = inp;
        var res=await fetch(cap.url,{method:'POST',credentials:'include',
            headers:Object.assign({}, cap.headers||{}, {'Content-Type':'application/json',
                'apollo-require-preflight':'true',
                'x-apollo-operation-name':(payload.operationName||'Reviews_ReviewSearch')}),
            body:JSON.stringify(payload)});
        var txt=await res.text(); var j=null; try{ j=JSON.parse(txt); }catch(e){}
        function findArr(o,d){ if(d>12||!o||typeof o!=='object')return null;
          if(Array.isArray(o)&&o.length&&o[0]&&(o[0].bookingReference!==undefined||o[0].reviewId!==undefined))return o;
          for(var k in o){var r=findArr(o[k],d+1); if(r)return r;} return null; }
        var list=j?(findArr(j,0)||[]):[];
        done({rows:list.map(function(x){return {code:x.bookingReference, rating:x.rating, content:x.comment||'', rdate:(x.createdAt||'').slice(0,10)};})});
        """
        page, size = 1, 50
        while page <= 2000:
            off = (page - 1) * size
            r = self._fetch_page(js, cap, off, size, sd, ed, rev)
            if not r or r.get("error"):
                self.log(f"  ⚠ [GG] p{page} 오류: {r.get('error') if r else 'no-resp'}"); break
            rows = r.get("rows") or []
            if not rows:
                break
            for x in rows:
                try:
                    if float(x.get("rating")) < 0:
                        continue
                except (ValueError, TypeError):
                    continue
                code = norm_code(x.get("code"))
                if not code:
                    continue
                reviews[code] = {"rating": self._fmt_rating(x.get("rating")),
                                 "content": (x.get("content") or "").strip(),
                                 "rdate": (x.get("rdate") or "")[:10]}
            self.log(f"  → [GG] p{page}: {len(rows)}행 (누적 {len(reviews)})")
            if len(rows) < size:
                break
            page += 1
            time.sleep(0.2)
        return reviews

    def collect_ctrip_simple(self, dfrom, dto, date_mode):
        """TPC: 작성일 정렬 페이지네이션 + 클라 필터(참여=departureTime / 작성=commentTime) + 4~5점."""
        reviews = {}
        cap = self.capture_request(CHANNEL_META["TPC"]["page"], "listOrderComments", wait=10)
        if not cap:
            self.log("  ⚠ [TPC] 캡처 실패"); return reviews
        rev = (date_mode == "review_date")
        js = """
        var cap=P[0], page=P[1], size=P[2];
        var body={sceneInfo:{bizScene:'ACTIVITY'}, paging:{pageNo:page, pageSize:size},
                  sorting:{orderBy:'COMMENT_TIME', desc:true}};
        try{ var b=JSON.parse(cap.body); if(b&&b.sceneInfo){ body.sceneInfo=b.sceneInfo; } }catch(e){}
        var res=await fetch(cap.url,{method:'POST',credentials:'include',
            headers:Object.assign({}, cap.headers||{}, {'Content-Type':'application/json'}),
            body:JSON.stringify(body)});
        var j=await res.json();
        var list=j.comments||[];
        done({rows:list.map(function(x){return {code:String(x.orderId), rating:x.score, content:x.content||'',
              ctime:x.commentTime, dtime:(x.orderInfo&&x.orderInfo.departureTime)||null};})});
        """
        lo = to_epoch_ms(pd.Timestamp(dfrom))
        hi = to_epoch_ms(pd.Timestamp(dto) + timedelta(days=1)) - 1
        stop_ms = to_epoch_ms(pd.Timestamp(dfrom) - timedelta(days=(1 if rev else 60)))
        page, size = 1, 50
        while page <= 3000:
            r = self._fetch_page(js, cap, page, size)
            if not r or r.get("error"):
                self.log(f"  ⚠ [TPC] p{page} 오류: {r.get('error') if r else 'no-resp'}"); break
            rows = r.get("rows") or []
            if not rows:
                break
            oldest_ok = True
            for x in rows:
                ct = self._to_ms(x.get("ctime"))
                dt = self._to_ms(x.get("dtime"))
                basis = ct if rev else dt
                if basis is not None and (lo <= basis <= hi):
                    try:
                        if float(x.get("rating")) >= 0:
                            code = norm_code(x.get("code"))
                            if code and code != "0":
                                rd = pd.Timestamp(ct, unit="ms").strftime("%Y-%m-%d") if ct else ""
                                reviews[code] = {"rating": self._fmt_rating(x.get("rating")),
                                                 "content": (x.get("content") or "").strip(),
                                                 "rdate": rd}
                    except (ValueError, TypeError):
                        pass
                if ct is not None and ct < stop_ms:
                    oldest_ok = False
            self.log(f"  → [TPC] p{page}: {len(rows)}행 (누적 {len(reviews)})")
            if not oldest_ok:
                break
            page += 1
            time.sleep(0.12)
        return reviews

    def collect_mrt_simple(self, dfrom, dto, date_mode):
        """MRT: 전체 페이지네이션 + 클라 필터(참여=travelStartDate / 작성=createdAt) + 4~5점."""
        reviews = {}
        self.driver.get(CHANNEL_META["MRT"]["page"])
        time.sleep(3)
        rev = (date_mode == "review_date")
        js = """
        var page=P[0], size=P[1];
        var tok=localStorage.getItem('accessToken');
        var res=await fetch('https://api3-backoffice.myrealtrip.com/review/partner/reviews/search',
            {method:'POST', credentials:'include',
             headers:{'Content-Type':'application/json','partner-access-token':tok},
             body:JSON.stringify({page:page, pageSize:size})});
        var j=await res.json();
        var list=Array.isArray(j.data)?j.data:[];
        done({rows:list.map(function(x){return {code:x.reservationNo, rating:x.score, content:x.comment||'',
              rdate:x.createdAt, tdate:x.travelStartDate};})});
        """
        page, size = 1, 50
        while page <= 3000:
            r = self._fetch_page(js, page, size)
            if not r or r.get("error"):
                self.log(f"  ⚠ [MRT] p{page} 오류: {r.get('error') if r else 'no-resp'}"); break
            rows = r.get("rows") or []
            if not rows:
                break
            for x in rows:
                basis = (x.get("rdate") or "")[:10] if rev else (x.get("tdate") or "")[:10]
                if not basis or not (dfrom <= basis <= dto):
                    continue
                try:
                    if float(x.get("rating")) < 0:
                        continue
                except (ValueError, TypeError):
                    continue
                code = norm_code(x.get("code"))
                if not code:
                    continue
                reviews[code] = {"rating": self._fmt_rating(x.get("rating")),
                                 "content": (x.get("content") or "").strip(),
                                 "rdate": (x.get("rdate") or "")[:10]}
            self.log(f"  → [MRT] p{page}: {len(rows)}행 (누적 {len(reviews)})")
            if len(rows) < size:
                break
            page += 1
            time.sleep(0.15)
        return reviews

    def start_processing_simple(self, mode):
        if not self.driver:
            messagebox.showerror("오류", "먼저 크롬을 연결하세요!"); return
        if getattr(self, "df_simple", None) is None:
            messagebox.showerror("오류", "먼저 엑셀을 선택하세요!"); return
        sel_branches = self._selected_branches()
        if not sel_branches:
            messagebox.showerror("오류", "최소 1개 지사를 선택하세요!"); return
        d1 = self._get_date(self.start_date_widget)
        d2 = self._get_date(self.end_date_widget)
        if d1 is None or d2 is None:
            messagebox.showerror("오류", "기간(시작일/종료일)을 확인하세요!"); return
        if d1 > d2:
            messagebox.showerror("오류", "시작일이 종료일보다 늦습니다!"); return

        date_mode = "review_date"   # Monthly·NYP 모두 리뷰 작성일 기준
        mode_name = {"MONTHLY": "Monthly", "QUARTERLY": "Quarterly", "NYP": "New Year Party"}[mode]
        report_min = pd.Timestamp(d1).normalize()
        report_max = pd.Timestamp(d2).normalize()
        dfrom = report_min.strftime("%Y-%m-%d")
        dto = report_max.strftime("%Y-%m-%d")

        # Monthly/NYP: No Show 포함 · 스페셜 제외 · 전체 가이드 · 기간 범위
        # 팀/투어 분모는 리뷰 수집 가능한 5개 채널 예약만 (L/KK/GG/TPC/MRT)
        _base = self.df_simple[
            (self.df_simple["Date"].dt.normalize() >= report_min)
            & (self.df_simple["Date"].dt.normalize() <= report_max)
            & (self.df_simple["Area"].isin(sel_branches))
        ].copy()
        # Team Count(리뷰율 분모) = 5채널 예약만 / Tour Count(실제 근무일) = 전체 채널
        res_all = _base
        res_df = _base[_base["Agency"].isin(CHANNELS)].copy()
        if res_df.empty:
            messagebox.showerror("오류", "선택 조건에 해당하는 예약이 없습니다."); return

        try:
            self.detail_lines = []
            self.result_text.delete(1.0, "end")
            basis_kr = "리뷰작성일" if date_mode == "review_date" else "참여일"
            self.log(f"📊 {mode_name} 리뷰 크롤링 시작 (1~5점 전체 수집 · {basis_kr} 기준)")
            self.log(f"📅 기간: {dfrom} ~ {dto} · 지사 {res_df['Area'].nunique()}개 · 예약 {len(res_df)}건")
            self.log("=" * 70)

            pw = self.create_progress_window()
            used = [c for c in CHANNELS if c in set(res_df["Agency"])]
            collectors = {
                "L": lambda: self.collect_klook_simple(dfrom, dto, date_mode),
                "KK": lambda: self.collect_kkday_simple(dfrom, dto, date_mode),
                "GG": lambda: self.collect_gg_simple(dfrom, dto, date_mode),
                "TPC": lambda: self.collect_ctrip_simple(dfrom, dto, date_mode),
                "MRT": lambda: self.collect_mrt_simple(dfrom, dto, date_mode),
            }
            collected = {}
            for i, ch in enumerate(used):
                pw.label.config(text=f"[{ch}] {CHANNEL_META[ch]['name']} 리뷰 수집 중...")
                pw.progress_bar["value"] = (i / max(len(used), 1)) * 100
                pw.window.update()
                self.log(f"\n🔍 [{ch}] {CHANNEL_META[ch]['name']}")
                try:
                    collected[ch] = collectors[ch]()
                    self.log(f"  ✓ [{ch}] {len(collected[ch])}건 수집")
                except Exception as e:
                    collected[ch] = {}
                    self.log(f"  ✗ [{ch}] 실패: {e}")
                    traceback.print_exc()
            pw.window.destroy()

            rows = []
            for _, r in res_df.iterrows():
                ch = r["Agency"]
                code = norm_code(r["Agency Code"])
                info = collected.get(ch, {}).get(code)
                if not info:
                    continue
                rows.append({
                    "Tour Date": pd.Timestamp(r["Date"]).strftime("%Y-%m-%d"),
                    "Review Date": info.get("rdate", ""),
                    "Agency Code": code,
                    "Tour": r.get("Product", ""),
                    "Star": info.get("rating", ""),
                    "Review": info.get("content", ""),
                    "Guide": r.get("Main Guide", ""),
                    "Area": r.get("Area", ""),
                    "Agency": ch,
                })
            matched_df = pd.DataFrame(rows)

            saved = self.save_excel_simple(matched_df, res_df, sel_branches, mode_name, dfrom, dto, res_all)

            self.log("\n" + "=" * 70)
            self.log(f"✅ {mode_name} 완료 · 매칭 {len(matched_df)}건")
            for area in sorted(set(sel_branches), key=area_rank):
                n = len(matched_df[matched_df["Area"] == area]) if not matched_df.empty else 0
                t = len(res_df[res_df["Area"] == area])
                self.log(f"  · {AREA_KR.get(area, area)}: 리뷰 {n} / 예약 {t}")
            self.progress_var.set("✅ 완료" + (f" · 저장: {os.path.basename(saved)}" if saved else ""))
            messagebox.showinfo("완료", f"{mode_name} 리뷰 크롤링 완료!\n\n엑셀 저장:\n{saved}")
        except Exception as e:
            self.progress_var.set(f"❌ 오류: {e}")
            self.log(f"오류: {e}")
            traceback.print_exc()
            messagebox.showerror("오류", f"처리 중 오류:\n{e}")

    @staticmethod
    def _autofit_simple(ws):
        from openpyxl.styles import Alignment
        from openpyxl.utils import get_column_letter
        for col in ws.columns:
            letter = get_column_letter(col[0].column)
            header = ws.cell(row=1, column=col[0].column).value
            if header == "Review":
                ws.column_dimensions[letter].width = 80
                for cell in col:
                    cell.alignment = Alignment(wrap_text=True, vertical="top")
            else:
                m = 0
                for cell in col:
                    v = cell.value
                    if v is None:
                        continue
                    w = sum(2 if ord(ch) > 127 else 1 for ch in str(v))
                    if w > m:
                        m = w
                ws.column_dimensions[letter].width = min(max(m + 2, 8), 40)

    def save_excel_simple(self, matched_df, res_df, areas, mode_name, dfrom, dto, res_all=None):
        from openpyxl import Workbook
        from openpyxl.styles import Font, PatternFill
        tag = dfrom.replace("-", "") + "-" + dto.replace("-", "")
        _rg = _region_tag(areas, english=True)
        fname = f"{mode_name} Review Crawling {_rg} {tag}.xlsx" if _rg else f"{mode_name} Review Crawling {tag}.xlsx"
        path = os.path.join(DOWNLOAD_DIR, fname)
        from openpyxl.styles import Alignment
        _center = Alignment(horizontal="center", vertical="center")
        wb = Workbook(); wb.remove(wb.active)
        cols = ["Tour Date", "Review Date", "Agency Code", "Tour", "Star", "Review", "Guide", "Agency"]

        def _arank(a):
            return {"L": 0, "KK": 1, "GG": 2, "TPC": 3, "MRT": 4}.get(str(a).strip(), 9)

        def _srank(s):
            try:
                return {5: 0, 4: 1}.get(int(float(s)), 9)
            except (ValueError, TypeError):
                return 9

        ordered = sorted(set(areas), key=area_rank)
        for area in ordered:
            ws = wb.create_sheet(self._safe_name(area))   # 영어 지사명 (Seoul/Busan...)
            for ci, h in enumerate(cols, 1):
                c = ws.cell(row=1, column=ci, value=h)
                c.font = Font(bold=True)
                c.fill = PatternFill(start_color="DDDDDD", end_color="DDDDDD", fill_type="solid")
            adf = matched_df[matched_df["Area"] == area].copy() if not matched_df.empty else None
            if adf is not None and not adf.empty:
                # 지사 시트는 좋은 리뷰(4~5점)만
                adf = adf[adf["Star"].apply(lambda s: str(s).strip() in ("4", "5"))].copy()
            if adf is not None and not adf.empty:
                adf["__a"] = adf["Agency"].apply(_arank)
                adf["__s"] = adf["Star"].apply(_srank)
                adf = adf.sort_values(by=["Tour Date", "__a", "__s", "Agency Code"]).drop(columns=["__a", "__s"])
                r = 2
                for _, row in adf.iterrows():
                    for ci, h in enumerate(cols, 1):
                        ws.cell(row=r, column=ci, value=row.get(h, ""))
                    r += 1
            self._autofit_simple(ws)
            for _ci in range(1, len(cols) + 1):
                ws.cell(row=1, column=_ci).alignment = _center

        self._build_guide_sheet_simple(wb, matched_df, res_df, ordered, res_all)
        wb.save(path)
        return path

    def _build_guide_sheet_simple(self, wb, matched_df, res_df, ordered, res_all=None):
        from openpyxl.styles import Font, PatternFill, Alignment
        from openpyxl.utils import get_column_letter
        ws = wb.create_sheet("Guide")
        center = Alignment(horizontal="center", vertical="center")

        def _split(v):
            if v is None:
                return []
            return [n.strip() for n in str(v).split(",") if n.strip()]

        def _isgood(sv):
            return str(sv).strip() in ("4", "5")

        headers = ["Guide Name", "Good", "Bad", "Total Review", "Tour Count",
                   "Team Count", "Good %", "Bad %", "Total %"]
        BW = len(headers)   # 8열
        if res_all is None:
            res_all = res_df
        start_col = 1
        for area in ordered:
            area_res = res_df[res_df["Area"] == area]          # Team(분모) = 5채널
            area_all = res_all[res_all["Area"] == area]        # Tour(근무일) = 전체 채널
            if area_res.empty:
                continue
            area_rev = matched_df[matched_df["Area"] == area] if not matched_df.empty else matched_df
            c = ws.cell(row=1, column=start_col, value=area)
            c.fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
            c.font = Font(bold=True, color="FFFFFF", size=12)
            c.alignment = center
            for i, h in enumerate(headers):
                cc = ws.cell(row=2, column=start_col + i, value=h)
                cc.font = Font(bold=True)
                cc.fill = PatternFill(start_color="DDDDDD", end_color="DDDDDD", fill_type="solid")
                cc.alignment = center
            # 가이드 이름 정규화(대소문자·공백 무시) → 같은 사람 합치기
            def _key(n):
                return re.sub(r"\s+", " ", str(n).strip()).lower()

            def _keys(v):
                return [_key(n) for n in _split(v)]

            # 표기 후보 수집 (5채널 + 전체 채널 모두에서)
            variants = {}
            for raw in list(area_res["Main Guide"].dropna()) + list(area_all["Main Guide"].dropna()):
                for n in _split(raw):
                    variants.setdefault(_key(n), {})
                    variants[_key(n)][n] = variants[_key(n)].get(n, 0) + 1

            def _display(k):
                cands = variants.get(k, {})
                if not cands:
                    return k
                # 대문자만인 표기보다 대소문자 섞인 표기 우선 → 그다음 많이 쓰인 순
                def _low(x):
                    return sum(1 for ch in x if ch.islower())
                best = sorted(cands.items(), key=lambda kv: (-_low(kv[0]), -kv[1], kv[0]))[0][0]
                return re.sub(r"\s+", " ", str(best).strip())   # 표시용 공백 정리

            # 대상 가이드 = 5채널 예약이 있는 사람 (분모 0 방지)
            names = set()
            for raw in area_res["Main Guide"].dropna():
                for n in _split(raw):
                    names.add(_key(n))

            stats = []
            for k in names:
                g = _display(k)
                if area_rev is not None and not area_rev.empty:
                    grev = area_rev[area_rev["Guide"].apply(lambda v, k=k: k in _keys(v))]
                    good = int(grev["Star"].apply(_isgood).sum())
                    total_rev = int(len(grev))
                else:
                    good = 0
                    total_rev = 0
                bad = total_rev - good
                gr = area_res[area_res["Main Guide"].apply(lambda v, k=k: k in _keys(v))]
                team = len(gr)   # 5채널 예약 수 (리뷰율 분모)
                ga = area_all[area_all["Main Guide"].apply(lambda v, k=k: k in _keys(v))]
                tour = ga["Date"].dt.normalize().nunique()   # 전체 채널 기준 실제 근무일
                gp = (good / team) if team > 0 else 0
                bp = (bad / team) if team > 0 else 0
                tp = (total_rev / team) if team > 0 else 0
                stats.append((g, good, bad, int(tour), int(team), gp, bp, tp))
            stats.sort(key=lambda x: (x[1], x[4]), reverse=True)
            r = 3
            for g, good, bad, tour, team, gp, bp, tp in stats:
                ws.cell(row=r, column=start_col, value=g)
                ws.cell(row=r, column=start_col + 1, value=good)
                ws.cell(row=r, column=start_col + 2, value=bad)
                ws.cell(row=r, column=start_col + 3, value=good + bad)   # Total Review = Good + Bad
                ws.cell(row=r, column=start_col + 4, value=tour)
                ws.cell(row=r, column=start_col + 5, value=team)
                for jj, val in [(6, gp), (7, bp), (8, tp)]:
                    pcell = ws.cell(row=r, column=start_col + jj, value=val)
                    pcell.number_format = '0.00%'
                r += 1
            for ci in range(start_col, start_col + BW):
                ws.column_dimensions[get_column_letter(ci)].width = 12
            start_col += BW + 1

    def quit_app(self):
        self._clear_preload()
        # 디버그 크롬 자체는 종료하지 않음 (사용자 세션 보존)
        try:
            self.root.quit()
            self.root.destroy()
        except Exception:
            pass

    def run(self):
        self.root.protocol("WM_DELETE_WINDOW", self.quit_app)
        self.root.mainloop()


if __name__ == "__main__":
    print("=" * 64)
    print("Smart Review Crawling 시작")
    print("=" * 64)
    print("\n⚠️  먼저 크롬을 디버그 모드로 실행하세요:")
    print("\nWindows:")
    print('  "C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe" '
          '--remote-debugging-port=9222 --user-data-dir="C:\\Chrome_debug"')
    print("\nMac:")
    print('  /Applications/Google\\ Chrome.app/Contents/MacOS/Google\\ Chrome '
          '--remote-debugging-port=9222')
    print("\n그 다음 아래 5개에 로그인:")
    print("  L  (Klook):  https://merchant.klook.com/reviews")
    print("  KK (KKday):  https://scm.kkday.com/v1/en/comment/index")
    print("  GG (GYG):    https://supplier.getyourguide.com/performance/reviews")
    print("  TPC(Trip):   https://vbooking.ctrip.com/tour/comment_manage/comment/list?bizScene=ACTIVITY")
    print("  MRT:         https://partner.myrealtrip.com/reviews/touractivity")
    print("=" * 64)
    PowerReviewApp().run()
