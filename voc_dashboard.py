import streamlit as st
import pandas as pd
import os

# 세션 초기화
if "login_type" not in st.session_state:
    st.session_state["login_type"] = None
if "login_user" not in st.session_state:
    st.session_state["login_user"] = None

ADMIN_CODE = "C3A"

# CONTACT_MAP 파일 경로
CONTACT_PATH = "contact_map.xlsx"

# 연락처 매핑 로드
def load_contact_map(path: str) -> dict:
    if not os.path.exists(path):
        st.warning(f"담당자 매핑 파일 '{path}' 을(를) 찾을 수 없습니다.")
        return {}
    df = pd.read_excel(path)
    # 컬럼명 자동 탐색 (예: 담당자, 이름, 구역담당자 등)
    name_col = None
    phone_col = None
    for col in df.columns:
        if any(k in col for k in ["구역담당자", "담당자", "성명", "이름"]):
            name_col = col
        if any(k in col for k in ["휴대폰", "전화", "연락처", "핸드폰"]):
            phone_col = col
    if name_col is None or phone_col is None:
        st.warning("담당자 매핑 파일에 이름 또는 전화번호 컬럼이 없습니다.")
        return {}
    contact = {}
    for _, row in df.iterrows():
        name = str(row[name_col]).strip()
        tel = str(row[phone_col]).strip()
        # 전화번호 정제 — 숫자만
        tel = "".join(ch for ch in tel if ch.isdigit())
        if len(tel) >= 4 and name:
            contact[name] = tel
    return contact

contacts_dict = load_contact_map(CONTACT_PATH)

def login_form():
    st.markdown("## 🔐 로그인")
    tab_admin, tab_user = st.tabs(["관리자 로그인", "사용자 로그인"])

    with tab_admin:
        pw = st.text_input("관리자 비밀번호", type="password", key="admin_pw")
        if st.button("관리자 로그인"):
            if pw == ADMIN_CODE:
                st.session_state["login_type"] = "admin"
                st.session_state["login_user"] = "ADMIN"
                st.success("관리자 로그인 성공")
                st.rerun()
            else:
                st.error("비밀번호가 올바르지 않습니다.")

    with tab_user:
        name = st.text_input("성명", key="user_name")
        input_pw = st.text_input("연락처 뒷 4자리", type="password", key="user_pw")

        if st.button("사용자 로그인"):
            real_tel = contacts_dict.get(name.strip())
            if real_tel:
                real_pw = real_tel[-4:]
                if input_pw == real_pw:
                    st.session_state["login_type"] = "user"
                    st.session_state["login_user"] = name.strip()
                    st.success(f"{name} 님 로그인 성공")
                    st.rerun()
                else:
                    st.error("로그인 실패: 비밀번호가 올바르지 않습니다.")
            else:
                st.error("등록된 사용자명이 아니거나 연락처 정보가 없습니다.")

if st.session_state["login_type"] is None:
    login_form()
    st.stop()

LOGIN_TYPE = st.session_state["login_type"]
LOGIN_USER = st.session_state["login_user"]

import os
from datetime import datetime, date
import smtplib
from email.message import EmailMessage

import numpy as np
import pandas as pd
import streamlit as st

# Plotly (고급 시각화, 없으면 자동 fallback)
try:
    import plotly.express as px
    HAS_PLOTLY = True
except Exception:
    HAS_PLOTLY = False

# ----------------------------------------------------
# 0. 기본 설정 & 라이트톤 / 반응형 레이아웃 CSS
# ----------------------------------------------------

st.set_page_config(page_title="해지 VOC 종합 대시보드", layout="wide")

st.markdown(
    """
    <style>
    /* 전체 배경 & 기본 폰트 (다크모드 무시, 항상 라이트톤 고정) */
    html, body {
        background-color: #f5f5f7 !important;
    }
    .stApp {
        background-color: #f5f5f7 !important;
        color: #111827 !important;
        font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif;
    }

    /* 본문 컨테이너 여백 (상단 잘림 방지 + 모바일 여백 보정) */
    .block-container {
        padding-top: 1.4rem !important;   /* 👈 상단 여백 조금 늘림 */
        padding-bottom: 3rem !important;
        padding-left: 1.0rem !important;
        padding-right: 1.0rem !important;
    }

    /* 헤더 영역 배경 */
    [data-testid="stHeader"] {
        background-color: #f5f5f7 !important;
    }

    /* 사이드바 스타일 */
    section[data-testid="stSidebar"] {
        background-color: #fafafa !important;
        border-right: 1px solid #e5e7eb;
    }
    section[data-testid="stSidebar"] .block-container {
        padding-top: 1.0rem;
    }

    /* 제목들 간격 */
    h1, h2, h3, h4 {
        margin-top: 0.4rem;
        margin-bottom: 0.35rem;
        font-weight: 600;
    }

    /* 데이터프레임 줄무늬 */
    .dataframe tbody tr:nth-child(odd) {
        background-color: #f9fafb;
    }
    .dataframe tbody tr:nth-child(even) {
        background-color: #eef2ff;
    }

    /* 입력창/셀렉트박스 공통 */
    textarea, input, select {
        border-radius: 8px !important;
    }

    /* 라디오 버튼 라벨 간격 */
    div[role="radiogroup"] > label {
        padding-right: 0.75rem;
    }

    /* 섹션 카드 공통 (피드백, 설명 등) */
    .section-card {
        background: #ffffff;
        border-radius: 16px;
        padding: 1.0rem 1.2rem;
        border: 1px solid #e5e7eb;
        box-shadow: 0 4px 8px rgba(15, 23, 42, 0.04);
        margin-bottom: 1.2rem;
    }
    .section-title {
        font-size: 1.05rem;
        font-weight: 600;
        margin-bottom: 0.6rem;
        display: flex;
        align-items: center;
        gap: 0.25rem;
    }

    /* 피드백 리스트 카드 */
    .feedback-item {
        background-color: #f9fafb;
        border-radius: 12px;
        padding: 0.7rem 0.9rem;
        margin-bottom: 0.6rem;
        border: 1px solid #e5e7eb;
    }
    .feedback-meta {
        font-size: 0.8rem;
        color: #6b7280;
        margin-top: 0.2rem;
    }
    .feedback-note {
        font-size: 0.85rem;
        color: #4b5563;
        margin-top: 0.2rem;
    }

    /* KPI 윗부분 여백 줄이기 */
    .element-container:has(> div[data-testid="stMetric"]) {
        padding-top: 0 !important;
        padding-bottom: 0.4rem !important;
    }

    /* 모바일 대응 — width 900px 이하면 자동 1열 레이아웃 */
    @media (max-width: 900px) {
        [data-testid="column"] {
            width: 100% !important;
            flex-direction: column !important;
        }
        .block-container {
            padding-left: 0.5rem !important;
            padding-right: 0.5rem !important;
        }
    }

    /* 표 overflow → 모바일 대응 */
    [data-testid="stDataFrame"] div {
        overflow-x: auto !important;
    }

    /* Plotly 차트 배경 투명 처리 */
    .js-plotly-plot .plotly {
        background-color: transparent !important;
    }
    </style>
    """,
    unsafe_allow_html=True,
)

# ----------------------------------------------------
# 1. 파일 경로 & SMTP 설정
# ----------------------------------------------------

if "SMTP_HOST" in st.secrets:
    SMTP_HOST = st.secrets["SMTP_HOST"]
    SMTP_PORT = int(st.secrets["SMTP_PORT"])
    SMTP_USER = st.secrets["SMTP_USER"]
    SMTP_PASSWORD = st.secrets["SMTP_PASSWORD"]
    SENDER_NAME = st.secrets["SENDER_NAME"]
else:
    # 로컬에서 dotenv 사용할 경우
    try:
        from dotenv import load_dotenv
        load_dotenv()
    except Exception:
        pass
    SMTP_HOST = os.getenv("SMTP_HOST", "")
    SMTP_PORT = int(os.getenv("SMTP_PORT", "587"))
    SMTP_USER = os.getenv("SMTP_USER", "")
    SMTP_PASSWORD = os.getenv("SMTP_PASSWORD", "")
    SENDER_NAME = os.getenv("SENDER_NAME", "해지VOC 관리자")

# ----------------------------------------------------
# 1-A. 파일 경로 설정
# ----------------------------------------------------
MERGED_PATH = "merged.xlsx"                 # VOC 통합파일
FEEDBACK_PATH = "feedback.csv"              # 처리내역 CSV 저장 경로
CONTACT_PATH = "contact_map.xlsx"           # 담당자 매핑 파일(확정)

# ----------------------------------------------------
# 2. 공통 유틸
# ----------------------------------------------------
def safe_str(x):
    if pd.isna(x):
        return ""
    return str(x).strip()


def detect_column(df: pd.DataFrame, keywords: list[str]) -> str | None:
    """담당자/이메일/휴대폰 컬럼 자동 탐색"""
    # 정확 일치 우선
    for k in keywords:
        if k in df.columns:
            return k
    # 부분 일치
    for col in df.columns:
        s = str(col)
        for k in keywords:
            if k.lower() in s.lower():
                return col
    return None

# ----------------------------------------------------
# 3. 데이터 로딩
# ----------------------------------------------------
@st.cache_data
def load_voc_data(path: str) -> pd.DataFrame:
    if not os.path.exists(path):
        st.error("❌ 'merged.xlsx' 파일이 존재하지 않습니다. 저장소 루트에 있는지 확인해주세요.")
        return pd.DataFrame()

    df = pd.read_excel(path)

    # 숫자형 컬럼(계약번호, 고객번호) 콤마 제거
    for col in ["계약번호", "고객번호"]:
        if col in df.columns:
            df[col] = (
                df[col]
                .astype(str)
                .str.replace(",", "", regex=False)
                .str.strip()
            )

    # 출처 정제 (고객리스트 → 해지시설)
    if "출처" in df.columns:
        df["출처"] = df["출처"].replace({"고객리스트": "해지시설"})

    # 계약번호 정제
    if "계약번호" in df.columns:
        df["계약번호_정제"] = (
            df["계약번호"]
            .astype(str)
            .str.replace(r"[^0-9A-Za-z]", "", regex=True)
            .str.strip()
        )
    else:
        df["계약번호_정제"] = ""

    # 접수일시 → datetime
    if "접수일시" in df.columns:
        df["접수일시"] = pd.to_datetime(df["접수일시"], errors="coerce")

    return df


@st.cache_data
def load_feedback(path: str) -> pd.DataFrame:
    """계약번호 단위 피드백 CSV"""
    if os.path.exists(path):
        try:
            fb = pd.read_csv(path, encoding="utf-8-sig")
        except Exception:
            fb = pd.read_csv(path)
    else:
        fb = pd.DataFrame(
            columns=["계약번호_정제", "고객대응내용", "등록자", "등록일자", "비고"]
        )
    return fb


def save_feedback(path: str, fb_df: pd.DataFrame) -> None:
    fb_df.to_csv(path, index=False, encoding="utf-8-sig")


@st.cache_data
def load_contact_map(path: str) -> tuple[pd.DataFrame, dict]:
    """영업구역 담당자 연락처 매핑"""
    if not os.path.exists(path):
        st.warning(
            f"❌ 담당자 매핑 파일 '{path}' 을(를) 찾을 수 없습니다. "
            "담당자 알림 탭에서는 직접 이메일 주소를 입력해서 사용해주세요."
        )
        return pd.DataFrame(), {}

    df_c = pd.read_excel(path)

    name_col = detect_column(df_c, ["구역담당자", "담당자", "처리자1", "성명", "이름"])
    email_col = detect_column(df_c, ["이메일", "메일", "E-MAIL"])
    phone_col = detect_column(df_c, ["휴대폰", "전화", "연락처", "핸드폰"])

    if not (name_col and email_col):
        st.warning(
            f"담당자 매핑 파일('{path}')에서 담당자/이메일 컬럼을 찾지 못했습니다. "
            "컬럼명을 확인해주세요."
        )
        return df_c, {}

    df_c = df_c[[name_col, email_col] + ([phone_col] if phone_col else [])].copy()
    df_c.rename(
        columns={
            name_col: "구역담당자_통합",
            email_col: "이메일",
            phone_col: "휴대폰" if phone_col else None,
        },
        inplace=True,
    )

    contact_dict: dict[str, dict] = {}
    for _, row in df_c.iterrows():
        name = safe_str(row["구역담당자_통합"])
        if not name:
            continue
        contact_dict[name] = {
            "email": safe_str(row.get("이메일", "")),
            "phone": safe_str(row.get("휴대폰", "")),
        }

    return df_c, contact_dict


# ---------- 실제 로딩 ----------
df = load_voc_data(MERGED_PATH)
if df.empty:
    st.stop()

if "feedback_df" not in st.session_state:
    st.session_state["feedback_df"] = load_feedback(FEEDBACK_PATH)

contact_df, manager_contacts = load_contact_map(CONTACT_PATH)

# ----------------------------------------------------
# 4. 지사명 축약 & 정렬 순서
# ----------------------------------------------------
if "관리지사" in df.columns:
    df["관리지사"] = df["관리지사"].replace(
        {
            "중앙지사": "중앙",
            "강북지사": "강북",
            "서대문지사": "서대문",
            "고양지사": "고양",
            "의정부지사": "의정부",
            "남양주지사": "남양주",
            "강릉지사": "강릉",
            "원주지사": "원주",
        }
    )
else:
    df["관리지사"] = ""

BRANCH_ORDER = ["중앙", "강북", "서대문", "고양", "의정부", "남양주", "강릉", "원주"]


def sort_branch(series):
    return sorted(
        [s for s in series if s in BRANCH_ORDER],
        key=lambda x: BRANCH_ORDER.index(x),
    )

# ----------------------------------------------------
# 5. 영업구역 / 담당자 통합 컬럼
# ----------------------------------------------------
def make_zone(row):
    if "영업구역번호" in row and pd.notna(row["영업구역번호"]):
        return row["영업구역번호"]
    if "담당상세" in row and pd.notna(row["담당상세"]):
        return row["담당상세"]
    if "영업구역정보" in row and pd.notna(row["영업구역정보"]):
        return row["영업구역정보"]
    return ""


df["영업구역_통합"] = df.apply(make_zone, axis=1)

mgr_priority = ["구역담당자", "담당자", "처리자"]


def pick_manager(row):
    for c in mgr_priority:
        if c in row and pd.notna(row[c]) and str(row[c]).strip() != "":
            return row[c]
    return ""


df["구역담당자_통합"] = df.apply(pick_manager, axis=1)

# 주소 컬럼 자동 탐색 (검색용)
address_cols = [c for c in df.columns if "주소" in str(c)]

# ----------------------------------------------------
# 6. 출처 분리 (해지VOC / 기타 출처) + 매칭여부
# ----------------------------------------------------
df_voc = df[df.get("출처") == "해지VOC"].copy()
df_other = df[df.get("출처") != "해지VOC"].copy()

other_sets = {
    src: set(df_other[df_other["출처"] == src]["계약번호_정제"].dropna())
    for src in ["해지시설", "해지요청", "설변", "정지", "해지파이프라인"]
    if "출처" in df_other.columns
}
other_union = set().union(*other_sets.values()) if other_sets else set()


# ------------------------------
# 로그인 폼
# ------------------------------
def login_form():

    st.markdown("## 🔐 로그인")

    tab_admin, tab_user = st.tabs(["관리자 로그인", "사용자 로그인"])

    # ---- 관리자 ----
    with tab_admin:
        pw = st.text_input("관리자 비밀번호", type="password", key="admin_pw")
        if st.button("관리자 로그인"):
            if pw == ADMIN_CODE:
                st.session_state["login_type"] = "admin"
                st.session_state["login_user"] = "ADMIN"
                st.success("관리자 로그인 성공")
                st.rerun()
            else:
                st.error("비밀번호가 올바르지 않습니다.")

    # ---- 사용자 ----
    with tab_user:
        name = st.text_input("성명", key="user_name")
        emp = st.text_input("5자리 사번", key="user_emp")

        if st.button("사용자 로그인"):
            if len(emp) == 5 and name.strip() != "":
                st.session_state["login_type"] = "user"
                st.session_state["login_user"] = name.strip()
                st.success(f"{name} 님 로그인 성공")
                st.rerun()
            else:
                st.error("성명 + 5자리 사번을 정확히 입력하세요.")

# ------------------------------
# 로그인 처리
# ------------------------------
if st.session_state["login_type"] is None:
    login_form()
    st.stop()

LOGIN_TYPE = st.session_state["login_type"]   # "admin" 또는 "user"
LOGIN_USER = st.session_state["login_user"]   # 관리자: ADMIN / 사용자: 성명


# ----------------------------------------------------
# 7. 설치주소 / 월정료 정제
# ----------------------------------------------------
def coalesce_cols(row, candidates):
    for c in candidates:
        if c in row.index:
            val = row[c]
            if pd.notna(val) and str(val).strip() not in ["", "None", "nan"]:
                return val
    return np.nan

df_voc["설치주소_표시"] = df_voc.apply(
    lambda r: coalesce_cols(r, ["시설_설치주소", "설치주소"]),
    axis=1,
)

fee_raw_col = "시설_KTT월정료(조정)" if "시설_KTT월정료(조정)" in df_voc.columns else None

def parse_fee(x: object) -> float:
    if pd.isna(x):
        return np.nan
    s = str(x).strip()
    if s == "" or s.lower() in ["nan", "none"]:
        return np.nan
    s = s.replace(",", "")
    digits = "".join(ch for ch in s if (ch.isdigit() or ch == "."))
    if digits == "":
        return np.nan
    try:
        v = float(digits)
    except Exception:
        return np.nan
    if v >= 200000:
        v = v / 10.0
    return v

if fee_raw_col is not None:
    df_voc["월정료_수치"] = df_voc[fee_raw_col].apply(parse_fee)

    def format_fee(v):
        if pd.isna(v):
            return ""
        return f"{int(round(v, 0)):,}"

    df_voc[fee_raw_col] = df_voc["월정료_수치"].apply(format_fee)

    def fee_band(v):
        if pd.isna(v):
            return "미기재"
        if v >= 100000:
            return "10만 이상"
        return "10만 미만"

    df_voc["월정료구간"] = df_voc["월정료_수치"].apply(fee_band)
else:
    df_voc["월정료_수치"] = np.nan
    df_voc["월정료구간"] = "미기재"


# ----------------------------------------------------
# 8. 리스크 등급/경과일 계산
# ----------------------------------------------------
today = date.today()

def compute_risk(row):
    dt = row.get("접수일시")
    if pd.isna(dt):
        return np.nan, "LOW"

    if not isinstance(dt, (pd.Timestamp, datetime)):
        try:
            dt = pd.to_datetime(dt, errors="coerce")
        except Exception:
            return np.nan, "LOW"

    if pd.isna(dt):
        return np.nan, "LOW"

    days = (today - dt.date()).days

    if days <= 3:
        level = "HIGH"
    elif days <= 10:
        level = "MEDIUM"
    else:
        level = "LOW"
    return days, level


# ----------------------------------------------------
# ⭐ 매칭여부 컬럼 생성 (로그인 분기보다 반드시 위)
# ----------------------------------------------------
df_voc["매칭여부"] = df_voc["계약번호_정제"].apply(
    lambda x: "매칭(O)" if x in other_union else "비매칭(X)"
)

# ----------------------------------------------------
# ⭐ 로그인 타입에 따른 unmatched_global 처리 (여기서 오류 안 남)
# ----------------------------------------------------
if LOGIN_TYPE == "user":
    df_user = df_voc[df_voc["구역담당자_통합"] == LOGIN_USER]
    unmatched_global = df_user[df_user["매칭여부"] == "비매칭(X)"]
else:
    unmatched_global = df_voc[df_voc["매칭여부"] == "비매칭(X)"]
    
# ----------------------------------------------------
# 9. 표시 컬럼 정의 & 스타일링
# ----------------------------------------------------
fixed_order = [
    "상호",
    "계약번호_정제",
    "매칭여부",
    "리스크등급",
    "경과일수",
    "출처",
    "관리지사",
    "영업구역번호",
    "영업구역_통합",
    "구역담당자_통합",
    "처리자",
    "담당유형",
    "처리유형",
    "처리내용",
    "접수일시",
    "서비스개시일",
    "계약종료일",
    "서비스중",
    "서비스소",
    "VOC유형",
    "VOC유형중",
    "VOC유형소",
    "해지상세",
    "등록내용",
    "설치주소_표시",
    "시설_KTT월정료(조정)",
    "계약상태(중)",
    "서비스(소)",
]
display_cols_raw = [c for c in fixed_order if c in df_voc.columns]


def filter_valid_columns(cols, df_base):
    valid_cols = []
    for c in cols:
        series = df_base[c]
        mask_valid = series.notna() & ~series.astype(str).str.strip().isin(
            ["", "None", "nan"]
        )
        if mask_valid.any():
            valid_cols.append(c)
    return valid_cols


display_cols = filter_valid_columns(display_cols_raw, df_voc)


def style_risk(df_view: pd.DataFrame):
    if "리스크등급" not in df_view.columns:
        return df_view

    def _row_style(row):
        level = row.get("리스크등급", "")
        if level == "HIGH":
            bg = "#fee2e2"
        elif level == "MEDIUM":
            bg = "#fef3c7"
        else:
            bg = "#e0f2fe"
        return [f"background-color: {bg};"] * len(row)

    return df_view.style.apply(_row_style, axis=1)

# ----------------------------------------------------
# 10. 사이드바 - 글로벌 필터
# ----------------------------------------------------
st.sidebar.title("🔧 글로벌 필터")

# 날짜 필터
if "접수일시" in df_voc.columns and df_voc["접수일시"].notna().any():
    min_d = df_voc["접수일시"].min().date()
    max_d = df_voc["접수일시"].max().date()
    dr = st.sidebar.date_input(
        "📅 접수일자 범위",
        value=(min_d, max_d),
        min_value=min_d,
        max_value=max_d,
    )
else:
    dr = None

# 지사 필터 → 버튼형 선택
branches_all = sort_branch(df_voc["관리지사"].dropna().unique())
sel_branches = st.sidebar.pills(
    "🏢 관리지사 선택",
    options=["전체"] + branches_all,
    selection_mode="multi",
    default=["전체"],          # 👈 기본값: 전체만 선택
    key="filter_branch_btn",
)

# 리스크 등급 필터
risk_all = ["HIGH", "MEDIUM", "LOW"]
sel_risk = st.sidebar.pills(
    "⚠ 리스크등급",
    options=risk_all,
    selection_mode="multi",
    default=risk_all,
    key="filter_risk_btn",
)

# 매칭여부
match_all = ["매칭(O)", "비매칭(X)"]
sel_match = st.sidebar.pills(
    "🔍 매칭여부",
    options=match_all,
    selection_mode="multi",
    default=["비매칭(X)"],     # 👈 기본값: 비매칭만 선택
    key="filter_match_btn",
)

# 월정료 구간 버튼 (10만원 단위)
fee_bands = [
    "전체",
    "0~10만",
    "10만~20만",
    "20만~30만",
    "30만~40만",
    "40만~50만",
    "50만 이상",
]

sel_fee_band = st.sidebar.radio(
    "💰 월정료 구간",
    options=fee_bands,
    index=0,
    key="filter_fee_band",
)

st.sidebar.markdown("---")
st.sidebar.caption(f"마지막 갱신: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")

# ----------------------------------------------------
# 11. 글로벌 필터 적용
# ----------------------------------------------------
voc_filtered_global = df_voc.copy()

# 날짜 필터
if dr and isinstance(dr, tuple) and len(dr) == 2:
    start_d, end_d = dr
    voc_filtered_global = voc_filtered_global[
        (voc_filtered_global["접수일시"] >= pd.to_datetime(start_d))
        & (voc_filtered_global["접수일시"] < pd.to_datetime(end_d) + pd.Timedelta(days=1))
    ]

# 지사 필터
if "전체" not in sel_branches:
    voc_filtered_global = voc_filtered_global[
        voc_filtered_global["관리지사"].isin(sel_branches)
    ]

# ⭐ 리스크 등급 필터 — 컬럼 존재 여부 체크 후 적용
if sel_risk and "리스크등급" in voc_filtered_global.columns:
    voc_filtered_global = voc_filtered_global[
        voc_filtered_global["리스크등급"].isin(sel_risk)
    ]

# 매칭여부 필터
if sel_match and "매칭여부" in voc_filtered_global.columns:
    voc_filtered_global = voc_filtered_global[
        voc_filtered_global["매칭여부"].isin(sel_match)
    ]

# 월정료 10만 단위 필터 적용
if sel_fee_band != "전체" and fee_raw_col is not None:
    fee = (
        voc_filtered_global[fee_raw_col]
        .astype(str)
        .str.replace(",", "", regex=False)
        .astype(float)
        .fillna(-1)
    )

    if sel_fee_band == "0~10만":
        voc_filtered_global = voc_filtered_global[(fee >= 0) & (fee < 100000)]
    elif sel_fee_band == "10만~20만":
        voc_filtered_global = voc_filtered_global[(fee >= 100000) & (fee < 200000)]
    elif sel_fee_band == "20만~30만":
        voc_filtered_global = voc_filtered_global[(fee >= 200000) & (fee < 300000)]
    elif sel_fee_band == "30만~40만":
        voc_filtered_global = voc_filtered_global[(fee >= 300000) & (fee < 400000)]
    elif sel_fee_band == "40만~50만":
        voc_filtered_global = voc_filtered_global[(fee >= 400000) & (fee < 500000)]
    elif sel_fee_band == "50만 이상":
        voc_filtered_global = voc_filtered_global[(fee >= 500000)]
# ------------------------------
# 🔐 로그인 타입에 따른 접근 제한
# ------------------------------
if LOGIN_TYPE == "user":
    # 사용자: 본인 담당 건만 보이게
    if "구역담당자_통합" in voc_filtered_global.columns:
        voc_filtered_global = voc_filtered_global[
            voc_filtered_global["구역담당자_통합"].astype(str) == str(LOGIN_USER)
        ]

unmatched_global = voc_filtered_global[
    voc_filtered_global["매칭여부"] == "비매칭(X)"
].copy()

# ----------------------------------------------------
# 12. 상단 KPI 카드
# ----------------------------------------------------
st.write("") # 빈공백추가
st.markdown("## 📊 해지 VOC 종합 대시보드")

total_voc_rows = len(voc_filtered_global)
unique_contracts = voc_filtered_global["계약번호_정제"].nunique()
unmatched_contracts = (
    voc_filtered_global[voc_filtered_global["매칭여부"] == "비매칭(X)"]["계약번호_정제"]
    .nunique()
)
matched_contracts = (
    voc_filtered_global[voc_filtered_global["매칭여부"] == "매칭(O)"]["계약번호_정제"]
    .nunique()
)

k1, k2, k3, k4 = st.columns(4)
k1.metric("VOC 접수건수(행)", f"{total_voc_rows:,}")
k2.metric("VOC 계약 수(유니크)", f"{unique_contracts:,}")
k3.metric("비매칭(X) 계약 수", f"{unmatched_contracts:,}")
k4.metric("매칭(O) 계약 수", f"{matched_contracts:,}")

st.markdown("---")

# ----------------------------------------------------
# 13. 탭 구성
# ----------------------------------------------------
tab_viz, tab_all, tab_unmatched, tab_drill, tab_filter, tab_alert = st.tabs(
    [
        "📊 지사/담당자 시각화",
        "📘 VOC 전체(계약 기준)",
        "🧯 해지방어 활동시설(비매칭)",
        "🔍 해지상담대상 활동등록",
        "🎯 해지방어 활동시설 정밀 필터(VOC유형소)",
        "📨 담당자 알림(베타)",
    ]
)

# ====================================================
# TAB VIZ — 지사 / 담당자 시각화 (개선 최종버전)
# ====================================================
with tab_viz:
    # 📌 viz_base 초기화 — unmatched_global 복사
    viz_base = unmatched_global.copy()

    # ✅ 리스크등급 컬럼 보호: 없으면 기본값 "LOW" 부여
    if "리스크등급" not in viz_base.columns:
        viz_base["리스크등급"] = "LOW"

    st.subheader("📊 지사 / 담당자별 비매칭 리스크 현황")

    if viz_base.empty:
        st.info("현재 조건에서 비매칭(X) 데이터가 없습니다.")
        st.stop()

    # -------------------------
    # 🔥 필터 UI (상단 고정)
    # -------------------------
    st.markdown("""
        <div style="
            background:#ffffff;
            border:1px solid #e5e7eb;
            padding:14px 20px;
            border-radius:12px;
            margin-bottom:18px;
            box-shadow:0 2px 6px rgba(0,0,0,0.05);
        ">
        <b>🎛️ 필터</b><br>
        지사와 담당자를 선택하면 아래 모든 시각화가 즉시 갱신됩니다.
        </div>
    """, unsafe_allow_html=True)

    colA, colB = st.columns(2)

    # -------------------------
    # 지사 선택
    # -------------------------
    b_opts = ["전체"] + sort_branch(viz_base["관리지사"].dropna().unique())
    sel_b_viz = colA.pills(
        "🏢 지사 선택",
        options=b_opts,
        selection_mode="single",
        default="전체",
        key="viz_branch",
    )
    sel_b_viz = sel_b_viz[0] if isinstance(sel_b_viz, list) else sel_b_viz

    # -------------------------
    # 담당자 선택
    # -------------------------
    tmp_mgr = viz_base.copy()
    if sel_b_viz != "전체":
        tmp_mgr = tmp_mgr[tmp_mgr["관리지사"] == sel_b_viz]

    mgr_list_viz = sorted([
        m for m in tmp_mgr["구역담당자_통합"].astype(str).unique().tolist()
        if m not in ["", "nan"]
    ])

    sel_mgr_viz = colB.selectbox(
        "👤 담당자 선택",
        options=["(전체)"] + mgr_list_viz,
        index=0,
        key="viz_mgr",
    )

    # -------------------------------------------------
    # 필터된 데이터 적용
    viz_filtered = viz_base.copy()
    if sel_b_viz != "전체":
        viz_filtered = viz_filtered[viz_filtered["관리지사"] == sel_b_viz]
    if sel_mgr_viz != "(전체)":
        viz_filtered = viz_filtered[viz_filtered["구역담당자_통합"].astype(str) == sel_mgr_viz]

    # ======================================================
    # 1) 지사별 비매칭 계약 수 (유니크 계약)
    # ======================================================
    st.markdown("### 🧱 지사별 비매칭 계약 수 (유니크 계약)")

    bc = (
        viz_filtered.groupby("관리지사")["계약번호_정제"]
        .nunique()
        .rename("비매칭계약수")
    )
    bc = bc.reindex(BRANCH_ORDER).fillna(0).astype(int)
    force_bar_chart(bc.reset_index(), "관리지사", "비매칭계약수", height=260)

    # ======================================================
    # 2) 담당자별 비매칭 TOP 15
    # ======================================================
    c2a, c2b = st.columns(2)
    with c2a:
        st.markdown("### 👤 담당자별 비매칭 TOP 15 (유니크 계약)")
        mc = (
            viz_filtered.groupby("구역담당자_통합")["계약번호_정제"]
            .nunique()
            .rename("비매칭계약수")
            .sort_values(ascending=False)
        )
        mc = mc[mc.index.astype(str).str.strip() != ""].head(15)
        mc_df = mc.reset_index()
        mc_df.columns = ["구역담당자_통합", "비매칭계약수"]
        force_bar_chart(mc_df, "구역담당자_통합", "비매칭계약수", height=300)

    # ======================================================
    # 3) 리스크 등급 분포 (계약 단위)
    # ======================================================
    with c2b:
        st.markdown("### 🔥 리스크 등급 분포 (계약 단위)")
        if "리스크등급" in viz_filtered.columns:
            rc = (
                viz_filtered["리스크등급"]
                .value_counts()
                .reindex(["HIGH", "MEDIUM", "LOW"])
                .fillna(0)
            )
            rc_df = pd.DataFrame({"리스크등급": rc.index, "건수": rc.values})
            force_bar_chart(rc_df, "리스크등급", "건수", height=300)
        else:
            st.info("리스크등급 데이터가 없어 분포를 표시할 수 없습니다.")

    # ======================================================
    # 4) 일별 비매칭 계약 추이 (유니크 계약)
    # ======================================================
    st.markdown("---")
    if "접수일시" in viz_filtered.columns and viz_filtered["접수일시"].notna().any():
        trend = (
            viz_filtered.assign(접수일=viz_filtered["접수일시"].dt.date)
            .groupby("접수일")["계약번호_정제"]
            .nunique()
            .rename("비매칭계약수")
            .sort_index()
        )
        if HAS_PLOTLY:
            fig4 = px.line(trend.reset_index(), x="접수일", y="비매칭계약수")
            fig4.update_layout(
                height=260,
                margin=dict(l=40, r=20, t=40, b=40),
                xaxis_title="접수일",
                yaxis_title="비매칭 계약 수",
            )
            st.plotly_chart(fig4, use_container_width=True)
        else:
            st.line_chart(trend, use_container_width=True, height=260)

    # ======================================================
    # 5) 담당자 레이더 차트 (선택된 담당자만)
    # ======================================================
    if sel_mgr_viz != "(전체)" and HAS_PLOTLY:
        mgr_data = viz_filtered[
            viz_filtered["구역담당자_통합"].astype(str) == sel_mgr_viz
        ]
        if not mgr_data.empty:
            radar = (
                mgr_data["리스크등급"]
                .value_counts()
                .reindex(["HIGH", "MEDIUM", "LOW"])
                .fillna(0)
            )
            radar_df = pd.DataFrame(
                {"리스크": ["HIGH", "MEDIUM", "LOW"], "계약수": radar.values}
            )
            fig_radar = px.line_polar(
                radar_df,
                r="계약수",
                theta="리스크",
                line_close=True,
            )
            fig_radar.update_layout(
                height=320,
                margin=dict(l=40, r=20, t=40, b=20),
                title=f"🌐 {sel_mgr_viz} 담당자의 리스크 프로파일",
            )
            st.plotly_chart(fig_radar, use_container_width=True)

# ====================================================
# TAB ALL — VOC 전체 (계약번호 기준 요약)
# ====================================================
with tab_all:
    st.subheader("📘 VOC 전체 (계약번호 기준 요약)")

    row1_col1, row1_col2 = st.columns([2, 3])

    branches_for_tab1 = ["전체"] + sort_branch(
        voc_filtered_global["관리지사"].dropna().unique()
    )
    selected_branch_tab1 = row1_col1.radio(
        "지사 선택",
        options=branches_for_tab1,
        horizontal=True,
        key="tab1_branch_radio",
    )

    temp_for_mgr = voc_filtered_global.copy()
    if selected_branch_tab1 != "전체":
        temp_for_mgr = temp_for_mgr[
            temp_for_mgr["관리지사"] == selected_branch_tab1
        ]

    mgr_options_tab1 = (
        ["전체"]
        + sorted(
            temp_for_mgr["구역담당자_통합"]
            .dropna()
            .astype(str)
            .unique()
            .tolist()
        )
        if "구역담당자_통합" in temp_for_mgr.columns
        else ["전체"]
    )

    selected_mgr_tab1 = row1_col2.radio(
        "담당자 선택",
        options=mgr_options_tab1,
        horizontal=True,
        key="tab1_mgr_radio",
    )

    s1, s2, s3 = st.columns(3)
    q_cn = s1.text_input("계약번호 검색(부분)", key="tab1_cn")
    q_name = s2.text_input("상호 검색(부분)", key="tab1_name")
    q_addr = s3.text_input("주소 검색(부분)", key="tab1_addr")

    temp = voc_filtered_global.copy()

    if selected_branch_tab1 != "전체":
        temp = temp[temp["관리지사"] == selected_branch_tab1]
    if selected_mgr_tab1 != "전체":
        temp = temp[temp["구역담당자_통합"].astype(str) == selected_mgr_tab1]

    if q_cn:
        temp = temp[
            temp["계약번호_정제"].astype(str).str.contains(q_cn.strip())
        ]
    if q_name and "상호" in temp.columns:
        temp = temp[
            temp["상호"].astype(str).str.contains(q_name.strip())
        ]
    if q_addr:
        cond = None
        if "설치주소_표시" in temp.columns:
            cond = temp["설치주소_표시"].astype(str).str.contains(q_addr.strip())
        else:
            for col in address_cols:
                if col in temp.columns:
                    series_cond = temp[col].astype(str).str.contains(q_addr.strip())
                    if cond is None:
                        cond = series_cond
                    else:
                        cond = cond | series_cond
        if cond is not None:
            temp = temp[cond]

    if temp.empty:
        st.info("조건에 맞는 VOC 데이터가 없습니다.")
    else:
        temp_sorted = temp.sort_values("접수일시", ascending=False)
        grp = temp_sorted.groupby("계약번호_정제")
        idx_latest = grp["접수일시"].idxmax()
        df_summary = temp_sorted.loc[idx_latest].copy()
        df_summary["접수건수"] = grp.size().reindex(df_summary["계약번호_정제"]).values

        summary_cols = [
            "계약번호_정제",
            "상호",
            "관리지사",
            "구역담당자_통합",
            "리스크등급",
            "경과일수",
            "매칭여부",
            "접수건수",
            "설치주소_표시",
            fee_raw_col if fee_raw_col is not None else None,
            "계약상태(중)",
            "서비스(소)",
        ]
        summary_cols = [c for c in summary_cols if c and c in df_summary.columns]
        summary_cols = filter_valid_columns(summary_cols, df_summary)

        st.markdown(f"📌 표시 계약 수: **{len(df_summary):,} 건**")
        st.dataframe(
            style_risk(df_summary[summary_cols]),
            use_container_width=True,
            height=480,
        )

# ====================================================
# TAB UNMATCHED — 해지방어 활동시설(비매칭)
# ====================================================
with tab_unmatched:
    st.subheader("🧯 해지방어 활동시설 (비매칭, 계약번호 기준)")

    st.caption("비매칭(X) = 해지 VOC 접수 후 시스템상 활동내역이 확인되지 않은 시설")

with st.expander("ℹ️ 해지방어 활동시설 안내", expanded=True):  # 👈 True로 변경
    st.write(
        "해지VOC 접수 후 **해지방어 활동내역이 시스템에 등록되지 않은 시설**입니다.\n"
        "- 실제 현장 대응 여부를 신속히 확인해 주세요.\n"
        "- 확인 후에는 반드시 `해지상담대상 활동등록` 탭에서 처리내역을 남겨주세요."
        )

    if unmatched_global.empty:
        st.info("현재 글로벌 필터 조건에서 비매칭(X) 계약이 없습니다.")
    else:
        # 🔎 필터 접기/펼치기
        with st.expander("🔎 지사 / 담당자 / 검색 필터", expanded=False):
            u_col1, u_col2 = st.columns([2, 3])

            branches_u = ["전체"] + sort_branch(
                unmatched_global["관리지사"].dropna().unique()
            )
            selected_branch_u = u_col1.radio(
                "지사 선택",
                options=branches_u,
                horizontal=True,
                key="tab2_branch_radio",
            )

            temp_u_for_mgr = unmatched_global.copy()
            if selected_branch_u != "전체":
                temp_u_for_mgr = temp_u_for_mgr[
                    temp_u_for_mgr["관리지사"] == selected_branch_u
                ]

            mgr_options_u = (
                ["전체"]
                + sorted(
                    temp_u_for_mgr["구역담당자_통합"]
                    .dropna()
                    .astype(str)
                    .unique()
                    .tolist()
                )
                if "구역담당자_통합" in temp_u_for_mgr.columns
                else ["전체"]
            )

            selected_mgr_u = u_col2.radio(
                "담당자 선택",
                options=mgr_options_u,
                horizontal=True,
                key="tab2_mgr_radio",
            )

            us1, us2 = st.columns(2)
            uq_cn = us1.text_input("계약번호 검색(부분)", key="tab2_cn")
            uq_name = us2.text_input("상호 검색(부분)", key="tab2_name")

        # ▶ 필터 적용
        temp_u = unmatched_global.copy()
        if selected_branch_u != "전체":
            temp_u = temp_u[temp_u["관리지사"] == selected_branch_u]
        if selected_mgr_u != "전체":
            temp_u = temp_u[temp_u["구역담당자_통합"].astype(str) == selected_mgr_u]

        if uq_cn:
            temp_u = temp_u[
                temp_u["계약번호_정제"].astype(str).str.contains(uq_cn.strip())
            ]
        if uq_name and "상호" in temp_u.columns:
            temp_u = temp_u[
                temp_u["상호"].astype(str).str.contains(uq_name.strip())
            ]

        if temp_u.empty:
            st.info("조건에 맞는 해지방어 활동시설(비매칭) 계약이 없습니다.")
        else:
            temp_u_sorted = temp_u.sort_values("접수일시", ascending=False)
            grp_u = temp_u_sorted.groupby("계약번호_정제")
            idx_latest_u = grp_u["접수일시"].idxmax()
            df_u_summary = temp_u_sorted.loc[idx_latest_u].copy()
            df_u_summary["접수건수"] = grp_u.size().reindex(
                df_u_summary["계약번호_정제"]
            ).values

            summary_cols_u = [
                "계약번호_정제",
                "상호",
                "관리지사",
                "구역담당자_통합",
                "리스크등급",
                "경과일수",
                "접수건수",
                "설치주소_표시",
                fee_raw_col if fee_raw_col is not None else None,
                "계약상태(중)",
                "서비스(소)",
            ]
            summary_cols_u = [
                c for c in summary_cols_u if c and c in df_u_summary.columns
            ]
            summary_cols_u = filter_valid_columns(summary_cols_u, df_u_summary)

            st.markdown(
                f"⚠ 해지방어 활동시설(비매칭) 계약 수: **{len(df_u_summary):,} 건**"
            )

            # ▶ 리스트(위쪽)
            st.data_editor(
                df_u_summary[summary_cols_u].reset_index(drop=True),
                use_container_width=True,
                height=420,
                hide_index=True,
                key="tab2_unmatched_editor",
            )

            # 행 선택 상태 읽기 → 선택된 계약번호 찾기
            selected_idx = None
            state = st.session_state.get("tab2_unmatched_editor", {})
            selected_rows = []
            if isinstance(state, dict):
                if "selected_rows" in state and state["selected_rows"]:
                    selected_rows = state["selected_rows"]
                elif "selection" in state and isinstance(
                    state["selection"], dict
                ):
                    rows_sel = state["selection"].get("rows")
                    if rows_sel:
                        selected_rows = rows_sel
            if selected_rows:
                selected_idx = selected_rows[0]

            u_contract_list = df_u_summary["계약번호_정제"].astype(str).tolist()
            default_index = 0
            if selected_idx is not None and 0 <= selected_idx < len(
                u_contract_list
            ):
                default_index = selected_idx + 1

            st.markdown("### 📂 선택한 계약번호 상세 VOC 이력")

            sel_u_contract = st.selectbox(
                "상세 VOC 이력을 볼 계약 선택 (표 행을 클릭하면 자동 선택됩니다)",
                options=["(선택)"] + u_contract_list,
                index=default_index,
                key="tab2_select_contract",
            )

            if sel_u_contract != "(선택)":
                voc_detail = temp_u[
                    temp_u["계약번호_정제"].astype(str) == sel_u_contract
                ].copy()
                voc_detail = voc_detail.sort_values("접수일시", ascending=False)

                # 🔹 상단 정보 (관리지사/담당자/계약번호/상호/월정료)
                latest = voc_detail.iloc[0]
                info_branch = latest.get("관리지사", "")
                info_mgr = latest.get("구역담당자_통합", "")
                info_name = latest.get("상호", "")
                info_fee = latest.get(fee_raw_col, "") if fee_raw_col else ""

                with st.expander("🔍 선택 계약 상세 정보 / VOC 이력", expanded=True):
                    st.markdown(
                        f"**관리지사:** {info_branch}  \n"
                        f"**구역담당자:** {info_mgr}  \n"
                        f"**계약번호:** {sel_u_contract}  \n"
                        f"**상호:** {info_name}  \n"
                        + (f"**{fee_raw_col}:** {info_fee}" if fee_raw_col else "")
                    )

                    st.markdown(
                        f"##### VOC 이력 ({len(voc_detail)}건)"
                    )
                    st.dataframe(
                        style_risk(voc_detail[display_cols]),
                        use_container_width=True,
                        height=350,
                    )

            st.download_button(
                "📥 해지방어 활동시설(비매칭) 원천 VOC 행 다운로드 (CSV)",
                temp_u.to_csv(index=False).encode("utf-8-sig"),
                file_name="해지방어_활동시설_원천행.csv",
                mime="text/csv",
            )
# ====================================================
# TAB DRILL — 해지상담대상 활동등록 (계약별 드릴다운)
# ====================================================
with tab_drill:
    st.subheader("🔍 해지상담대상 활동등록 (계약번호 기준 드릴다운)")

    base_all = voc_filtered_global.copy()

    match_choice = st.radio(
        "매칭여부 선택",
        options=["전체", "매칭(O)", "비매칭(X)"],
        horizontal=True,
        key="tab4_match_radio",
    )

    drill_base = base_all.copy()
    if match_choice == "매칭(O)":
        drill_base = drill_base[drill_base["매칭여부"] == "매칭(O)"]
    elif match_choice == "비매칭(X)":
        drill_base = drill_base[drill_base["매칭여부"] == "비매칭(X)"]

    # --------------------------------------
    # 🔎 필터 접기/펼치기 (정상 들여쓰기)
    # --------------------------------------
    with st.expander("🔎 지사 / 담당자 / 검색 필터", expanded=False):

        d1, d2 = st.columns([2, 3])

        branches_d = ["전체"] + sort_branch(
            drill_base["관리지사"].dropna().unique()
        )
        sel_branch_d = d1.radio(
            "지사 선택",
            options=branches_d,
            horizontal=True,
            key="tab4_branch_radio",
        )

        tmp_mgr_d = drill_base.copy()
        if sel_branch_d != "전체":
            tmp_mgr_d = tmp_mgr_d[tmp_mgr_d["관리지사"] == sel_branch_d]

        mgr_options_d = (
            ["전체"]
            + sorted(
                tmp_mgr_d["구역담당자_통합"]
                .dropna()
                .astype(str)
                .unique()
                .tolist()
            )
            if "구역담당자_통합" in tmp_mgr_d.columns
            else ["전체"]
        )

        sel_mgr_d = d2.radio(
            "담당자 선택",
            options=mgr_options_d,
            horizontal=True,
            key="tab4_mgr_radio",
        )

        dd1, dd2 = st.columns(2)
        dq_cn = dd1.text_input("계약번호 검색(부분)", key="tab4_cn")
        dq_name = dd2.text_input("상호 검색(부분)", key="tab4_name")

    # --------------------------------------
    # 🔍 필터 적용
    # --------------------------------------
    drill = drill_base.copy()
    if sel_branch_d != "전체":
        drill = drill[drill["관리지사"] == sel_branch_d]
    if sel_mgr_d != "전체":
        drill = drill[drill["구역담당자_통합"].astype(str) == sel_mgr_d]

    if dq_cn:
        drill = drill[
            drill["계약번호_정제"].astype(str).str.contains(dq_cn.strip())
        ]
    if dq_name and "상호" in drill.columns:
        drill = drill[
            drill["상호"].astype(str).str.contains(dq_name.strip())
        ]

    # --------------------------------------
    # 결과 없을 때
    # --------------------------------------
    if drill.empty:
        st.info("조건에 맞는 계약이 없습니다. 필터를 조정해보세요.")
        sel_cn = None
    else:
        drill_sorted = drill.sort_values("접수일시", ascending=False)
        g = drill_sorted.groupby("계약번호_정제")
        idx_latest_d = g["접수일시"].idxmax()
        df_d_summary = drill_sorted.loc[idx_latest_d].copy()
        df_d_summary["접수건수"] = g.size().reindex(
            df_d_summary["계약번호_정제"]
        ).values

        sum_cols_d = [
            "계약번호_정제",
            "상호",
            "관리지사",
            "구역담당자_통합",
            "리스크등급",
            "경과일수",
            "매칭여부",
            "접수건수",
            "설치주소_표시",
            fee_raw_col if fee_raw_col is not None else None,
            "계약상태(중)",
            "서비스(소)",
        ]
        sum_cols_d = [c for c in sum_cols_d if c and c in df_d_summary.columns]
        sum_cols_d = filter_valid_columns(sum_cols_d, df_d_summary)

        st.markdown("#### 📋 계약 요약 (최신 VOC 기준, 계약번호당 1행)")
        st.dataframe(
            style_risk(df_d_summary[sum_cols_d]),
            use_container_width=True,
            height=260,
        )

        cn_list = df_d_summary["계약번호_정제"].astype(str).tolist()

        def format_cn(cn_value: str) -> str:
            row = df_d_summary[
                df_d_summary["계약번호_정제"].astype(str) == str(cn_value)
            ].iloc[0]
            name = row.get("상호", "")
            branch = row.get("관리지사", "")
            cnt = row.get("접수건수", 0)
            return f"{cn_value} | {name} | {branch} | 접수 {int(cnt)}건"

        sel_cn = st.selectbox(
            "상세를 볼 계약 선택",
            options=cn_list,
            format_func=format_cn,
            key="tab4_cn_selectbox",
        )

        # --------------------------------------
        # 선택된 계약 상세 이력 표시
        # --------------------------------------
        if sel_cn:
            voc_hist = df_voc[
                df_voc["계약번호_정제"].astype(str) == str(sel_cn)
            ].copy()
            voc_hist = voc_hist.sort_values("접수일시", ascending=False)

            other_hist = df_other[
                df_other["계약번호_정제"].astype(str) == str(sel_cn)
            ].copy()

            base_info = voc_hist.iloc[0] if not voc_hist.empty else None

            st.markdown(f"### 🔎 선택된 계약번호: `{sel_cn}`")

            if base_info is not None:
                info_col1, info_col2, info_col3 = st.columns(3)
                info_col1.metric("상호", str(base_info.get("상호", "")))
                info_col2.metric("관리지사", str(base_info.get("관리지사", "")))
                info_col3.metric(
                    "구역담당자",
                    str(
                        base_info.get(
                            "구역담당자_통합", base_info.get("처리자", "")
                        )
                    ),
                )

                m2_1, m2_2, m2_3 = st.columns(3)
                m2_1.metric("VOC 접수건수", f"{len(voc_hist):,}건")
                m2_2.metric("리스크등급", str(base_info.get("리스크등급", "")))
                m2_3.metric("매칭여부", str(base_info.get("매칭여부", "")))

                st.caption(f"📍 설치주소: {str(base_info.get('설치주소_표시', ''))}")
                if fee_raw_col is not None:
                    st.caption(
                        f"💰 {fee_raw_col}: {str(base_info.get(fee_raw_col, ''))}"
                    )

            st.markdown("---")

            c_left, c_right = st.columns(2)

            with c_left:
                st.markdown("#### 📘 VOC 이력 (전체)")
                if voc_hist.empty:
                    st.info("VOC 이력이 없습니다.")
                else:
                    st.dataframe(
                        style_risk(voc_hist[display_cols]),
                        use_container_width=True,
                        height=320,
                    )

            with c_right:
                st.markdown("#### 📂 기타 출처 이력 (해지시설/요청/설변/정지/파이프라인)")
                if other_hist.empty:
                    st.info("기타 출처 데이터가 없습니다.")
                else:
                    st.dataframe(
                        other_hist,
                        use_container_width=True,
                        height=320,
                    )

# ----------------------------------------------------
# 글로벌 피드백 이력 & 입력 (선택된 sel_cn 기준)
# ----------------------------------------------------
st.markdown(
    '<div class="section-card"><div class="section-title">📝 해지상담대상 활동등록 (고객대응 / 현장 처리내역)</div>',
    unsafe_allow_html=True,
)

if "sel_cn" not in locals() or sel_cn is None:
    st.info("위의 '해지상담대상 활동등록' 탭에서 먼저 계약을 선택하면 처리내역을 관리할 수 있습니다.")
else:
    st.caption(f"선택된 계약번호: **{sel_cn}** 기준 처리내역 관리")

    fb_all = st.session_state["feedback_df"]
    fb_sel = fb_all[fb_all["계약번호_정제"].astype(str) == str(sel_cn)].copy()
    fb_sel = fb_sel.sort_values("등록일자", ascending=False)

    st.markdown("##### 📄 등록된 처리내역")
    if fb_sel.empty:
        st.info("등록된 처리 이력이 없습니다.")
    else:
        for idx, row in fb_sel.iterrows():
            with st.container():
                st.markdown('<div class="feedback-item">', unsafe_allow_html=True)
                col1, col2 = st.columns([6, 1])

                with col1:
                    st.write(f"**내용:** {row['고객대응내용']}")
                    st.markdown(
                        f"<div class='feedback-meta'>등록자: {row['등록자']} | 등록일: {row['등록일자']}</div>",
                        unsafe_allow_html=True,
                    )
                    if row.get("비고"):
                        st.markdown(
                            f"<div class='feedback-note'>비고: {row['비고']}</div>",
                            unsafe_allow_html=True
                        )

                with col2:
                    if LOGIN_TYPE == "admin":   # 관리자만 삭제 가능
                        if st.button("🗑 삭제", key=f"del_{idx}"):
                            fb_all = fb_all.drop(index=idx)
                            st.session_state["feedback_df"] = fb_all
                            save_feedback(FEEDBACK_PATH, fb_all)
                            st.success("삭제 완료!")
                            st.rerun()
                            st.markdown("</div>", unsafe_allow_html=True)

    # 새 처리내용 입력
    st.markdown("### ➕ 빠른 활동등록")

    if LOGIN_TYPE == "user":
        user_rows = unmatched_global.copy()
    else:
        user_rows = unmatched_global.copy()

    sel_quick = st.selectbox(
        "활동등록할 계약 선택",
        options=["(선택)"] + user_rows["계약번호_정제"].astype(str).tolist(),
        key="quick_cn",
     )

    if sel_quick != "(선택)":
        row = user_rows[user_rows["계약번호_정제"] == sel_quick].iloc[0]
        st.write(f"**계약번호:** {sel_quick}")
        st.write(f"**상호:** {row['상호']}")
        st.write(f"**설치주소:** {row['설치주소_표시']}")

        quick_content = st.text_area("활동내용 입력", key="quick_content")
        quick_writer = LOGIN_USER  # 사용자명 자동
        quick_note = st.text_input("비고", key="quick_note")

    if st.button("등록", key="quick_submit"):
        new_row = {
            "계약번호_정제": sel_quick,
            "고객대응내용": quick_content,
            "등록자": quick_writer,
            "등록일자": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "비고": quick_note,
        }
        fb_all = st.session_state["feedback_df"]
        fb_all = pd.concat([fb_all, pd.DataFrame([new_row])], ignore_index=True)
        st.session_state["feedback_df"] = fb_all
        save_feedback(FEEDBACK_PATH, fb_all)
        st.success("등록 완료되었습니다.")
        st.rerun()
st.markdown("</div>", unsafe_allow_html=True)

# ====================================================
# TAB FILTER — (정밀 필터 탭은 추후 확장용, 현재는 안내만)
# ====================================================
with tab_filter:
    st.subheader("🎯 해지방어 활동시설 정밀 필터 (VOC유형소 기준)")
    st.info(
        "현재 버전에서는 글로벌 필터 + 다른 탭에서 대부분 분석이 가능하도록 구성되어 있습니다.\n"
        "추후 필요 시 이 탭에 VOC유형소 중심의 추가 정밀 필터를 붙이면 됩니다."
    )

# ====================================================
# TAB ALERT — 담당자 알림(베타)
# ====================================================
with tab_alert:
    st.subheader("📨 담당자 알림 발송 (베타)")

    st.markdown(
        """
        담당자 파일(contact_map.xlsx)을 자동 매핑하여  
        비매칭(X) 계약 건을 **구역담당자별로 이메일로 발송**할 수 있습니다.
        """
    )

    if contact_df.empty:
        st.markdown(
            """
            <div style="
                background:#fff3cd;
                border-left:6px solid #ffca2c;
                padding:12px;
                border-radius:6px;
                margin-bottom:12px;
                font-size:0.95rem;
                line-height:1.45;
            ">
            <b>⚠ 담당자 매핑 파일을 찾을 수 없습니다.</b><br>
            'contact_map.xlsx' 파일이 저장소 루트(/) 위치에 있는지 확인하세요.<br>
            담당자 알림 탭에서는 이메일 주소를 직접 입력하여 사용할 수 있습니다.
            </div>
            """,
            unsafe_allow_html=True,
        )

        st.warning(
            "⚠ 담당자 매핑 파일이 업로드되지 않았습니다.\n"
            "contact_map.xlsx 파일을 저장소에 올려주세요."
        )
    else:
        st.success(f"담당자 매핑 파일 로드 완료 — 총 {len(contact_df)}명")

        unmatched_alert = unmatched_global.copy()
        grouped = unmatched_alert.groupby("구역담당자_통합")

        st.markdown("### 📧 알림 발송 대상(담당자별 비매칭 계약 수)")

        alert_list = []
        for mgr, g in grouped:
            mgr = safe_str(mgr)
            if not mgr:
                continue
            count = g["계약번호_정제"].nunique()
            email = manager_contacts.get(mgr, {}).get("email", "")
            alert_list.append([mgr, email, count])

        alert_df = pd.DataFrame(alert_list, columns=["담당자", "이메일", "비매칭 계약수"])
        st.dataframe(alert_df, use_container_width=True, height=300)

        st.markdown("---")

        st.markdown("### ✉ 개별 발송")

        sel_mgr = st.selectbox(
            "담당자 선택",
            options=["(선택)"] + alert_df["담당자"].tolist(),
            key="alert_mgr",
        )

        if sel_mgr != "(선택)":
            mgr_email = manager_contacts.get(sel_mgr, {}).get("email", "")
            st.write(f"📮 등록된 이메일: **{mgr_email or '(없음 — 직접 입력 필요)'}**")

            custom_email = st.text_input("이메일 주소(변경 또는 직접 입력)", value=mgr_email)

            df_mgr_rows = unmatched_alert[
                unmatched_alert["구역담당자_통합"].astype(str) == sel_mgr
            ]

            st.write(f"🔍 발송 데이터: **{len(df_mgr_rows)}건** 비매칭 VOC")

            if not df_mgr_rows.empty:
                st.dataframe(
                    df_mgr_rows[
                        ["계약번호_정제", "상호", "관리지사", "리스크등급", "경과일수"]
                    ],
                    use_container_width=True,
                    height=250,
                )
            else:
                st.info("해당 담당자에게 배정된 비매칭 계약이 없습니다.")

            subject = f"[해지VOC] {sel_mgr} 담당자 비매칭 계약 안내"
            body = (
                f"{sel_mgr} 담당자님,\n\n"
                f"아래 비매칭 해지 VOC 건이 확인되어 공유드립니다.\n"
                f"총 {len(df_mgr_rows)}건\n\n"
                "자세한 내용은 첨부 파일(CSV)을 확인해주세요.\n\n"
                "- 해지VOC 관리자 드림 -"
            )

            if st.button("📤 이메일 발송하기"):
                if not custom_email:
                    st.error("이메일 주소를 입력해주세요.")
                elif df_mgr_rows.empty:
                    st.error("발송할 비매칭 계약 데이터가 없습니다.")
                else:
                    try:
                        msg = EmailMessage()
                        msg["Subject"] = subject
                        msg["From"] = f"{SENDER_NAME} <{SMTP_USER}>"
                        msg["To"] = custom_email
                        msg.set_content(body)

                        # CSV 첨부
                        csv_bytes = df_mgr_rows.to_csv(index=False).encode("utf-8-sig")
                        msg.add_attachment(
                            csv_bytes,
                            maintype="application",
                            subtype="octet-stream",
                            filename=f"비매칭계약_{sel_mgr}.csv",
                        )

                        with smtplib.SMTP(SMTP_HOST, SMTP_PORT) as smtp:
                            smtp.starttls()
                            if SMTP_USER and SMTP_PASSWORD:
                                smtp.login(SMTP_USER, SMTP_PASSWORD)
                            smtp.send_message(msg)

                        st.success(f"✅ 이메일 발송 완료 → {custom_email}")

                    except Exception as e:
                        st.error(f"❌ 이메일 전송 실패: {e}")
