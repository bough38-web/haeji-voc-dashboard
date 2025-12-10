import streamlit as st
import pandas as pd
import numpy as np
import os
from datetime import datetime, date
import smtplib
from email.message import EmailMessage


# ==============================
# 0. 공통 설정 / 전역 상수
# ==============================

# 세션 초기화
if "login_type" not in st.session_state:
    st.session_state["login_type"] = None
if "login_user" not in st.session_state:
    st.session_state["login_user"] = None

ADMIN_CODE = "C3A"                 # 관리자 비밀번호
MERGED_PATH = "merged.xlsx"        # VOC 통합파일
FEEDBACK_PATH = "feedback.csv"     # 처리내역 CSV 저장 경로
CONTACT_PATH = "contact_map.xlsx"  # 담당자 매핑 파일

# Plotly 사용 여부
try:
    import plotly.express as px
    HAS_PLOTLY = True
except Exception:
    HAS_PLOTLY = False


# ------------------------------------------------
# 🔹 공통 막대그래프 (Plotly / 기본차트 자동 선택)
# ------------------------------------------------
def force_bar_chart(df: pd.DataFrame, x: str, y: str, height: int = 280):
    """Plotly가 있으면 Plotly, 없으면 기본 bar_chart 사용."""
    if df.empty:
        df = pd.DataFrame({x: ["데이터없음"], y: [0]})

    if HAS_PLOTLY:
        fig = px.bar(df, x=x, y=y, text=y)
        fig.update_traces(textposition="outside", textfont_size=11)
        max_y = df[y].max()
        fig.update_yaxes(range=[0, max_y * 1.3 if max_y > 0 else 1])
        fig.update_layout(
            height=height,
            margin=dict(l=40, r=20, t=60, b=40),
        )
        st.plotly_chart(fig, use_container_width=True)
    else:
        st.bar_chart(df.set_index(x)[y], height=height, use_container_width=True)


# ==============================
# 1. 페이지 기본 설정 & CSS
# ==============================
st.set_page_config(page_title="해지 VOC 종합 대시보드", layout="wide")

st.markdown(
    """
    <style>
    html, body {
        background-color: #f5f5f7 !important;
    }
    .stApp {
        background-color: #f5f5f7 !important;
        color: #111827 !important;
        font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif;
    }
    .block-container {
        padding-top: 1.4rem !important;
        padding-bottom: 3rem !important;
        padding-left: 1.0rem !important;
        padding-right: 1.0rem !important;
    }
    [data-testid="stHeader"] {
        background-color: #f5f5f7 !important;
    }
    section[data-testid="stSidebar"] {
        background-color: #fafafa !important;
        border-right: 1px solid #e5e7eb;
    }
    section[data-testid="stSidebar"] .block-container {
        padding-top: 1.0rem;
    }
    h1, h2, h3, h4 {
        margin-top: 0.4rem;
        margin-bottom: 0.35rem;
        font-weight: 600;
    }
    .dataframe tbody tr:nth-child(odd) {
        background-color: #f9fafb;
    }
    .dataframe tbody tr:nth-child(even) {
        background-color: #eef2ff;
    }
    textarea, input, select {
        border-radius: 8px !important;
    }
    div[role="radiogroup"] > label {
        padding-right: 0.75rem;
    }
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
    .element-container:has(> div[data-testid="stMetric"]) {
        padding-top: 0 !important;
        padding-bottom: 0.4rem !important;
    }
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
    [data-testid="stDataFrame"] div {
        overflow-x: auto !important;
    }
    .js-plotly-plot .plotly {
        background-color: transparent !important;
    }
    </style>
    """,
    unsafe_allow_html=True,
)

# ==============================
# 2. SMTP 설정
# ==============================
if "SMTP_HOST" in st.secrets:
    SMTP_HOST = st.secrets["SMTP_HOST"]
    SMTP_PORT = int(st.secrets["SMTP_PORT"])
    SMTP_USER = st.secrets["SMTP_USER"]
    SMTP_PASSWORD = st.secrets["SMTP_PASSWORD"]
    SENDER_NAME = st.secrets["SENDER_NAME"]
else:
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

# ==============================
# 3. 유틸 함수
# ==============================
def safe_str(x):
    if pd.isna(x):
        return ""
    return str(x).strip()


def detect_column(df: pd.DataFrame, keywords: list[str]) -> str | None:
    """컬럼명 자동 탐색."""
    for k in keywords:
        if k in df.columns:
            return k
    for col in df.columns:
        s = str(col)
        for k in keywords:
            if k.lower() in s.lower():
                return col
    return None

# ==============================
# 4. 데이터 로드 함수
# ==============================
@st.cache_data
def load_voc_data(path: str) -> pd.DataFrame:
    if not os.path.exists(path):
        st.error("❌ 'merged.xlsx' 파일이 존재하지 않습니다. 저장소 루트에 있는지 확인해주세요.")
        return pd.DataFrame()

    df = pd.read_excel(path)

    # 숫자형 문자열화
    for col in ["계약번호", "고객번호"]:
        if col in df.columns:
            df[col] = (
                df[col]
                .astype(str)
                .str.replace(",", "", regex=False)
                .str.strip()
            )

    # 출처 정제
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
def load_contact_map(path: str):
    """
    담당자 매핑 파일 로드.
    반환: (contact_df, manager_contacts_dict)
      - contact_df: 정제된 DataFrame
      - manager_contacts_dict: {담당자: {"email":..., "phone":...}}
    """
    if not os.path.exists(path):
        st.warning(
            f"❌ 담당자 매핑 파일 '{path}' 을(를) 찾을 수 없습니다. "
            "담당자 알림 탭에서는 직접 이메일 주소를 입력해서 사용해주세요."
        )
        return pd.DataFrame(), {}

    df_c = pd.read_excel(path)

    # 🔍 컬럼 자동 탐색 (연락처 오타 '연략처' 포함)
    name_col = detect_column(df_c, ["구역담당자", "담당자", "처리자1", "성명", "이름"])
    email_col = detect_column(df_c, ["이메일", "메일", "E-MAIL"])
    phone_col = detect_column(df_c, ["휴대폰", "전화", "연락처", "연략처", "핸드폰"])

    if not name_col:
        st.warning("담당자 매핑 파일에서 담당자 이름 컬럼을 찾지 못했습니다.")
        return df_c, {}

    # 사용할 컬럼만 선택
    cols = [name_col]
    if email_col:
        cols.append(email_col)
    if phone_col:
        cols.append(phone_col)

    df_c = df_c[cols].copy()

    # 표준 컬럼명으로 변경
    rename_map = {name_col: "구역담당자_통합"}
    if email_col:
        rename_map[email_col] = "이메일"
    if phone_col:
        rename_map[phone_col] = "휴대폰"
    df_c.rename(columns=rename_map, inplace=True)

    # 📱 휴대폰 컬럼 숫자만 남기고 정제 (뒷 4자리 로그인용)
    if "휴대폰" in df_c.columns:
        df_c["휴대폰"] = df_c["휴대폰"].apply(
            lambda x: "".join(ch for ch in safe_str(x) if ch.isdigit())
        )

    # 🔗 최종 매핑 딕셔너리 생성
    manager_contacts: dict[str, dict] = {}
    for _, row in df_c.iterrows():
        name = safe_str(row.get("구역담당자_통합", ""))
        if not name:
            continue

        email = safe_str(row.get("이메일", "")) if "이메일" in df_c.columns else ""
        phone = safe_str(row.get("휴대폰", "")) if "휴대폰" in df_c.columns else ""

        manager_contacts[name] = {
            "email": email,
            "phone": phone,
        }

    return df_c, manager_contacts


# ==============================
# 5. 실제 데이터 로딩
# ==============================
df = load_voc_data(MERGED_PATH)
if df.empty:
    st.stop()

if "feedback_df" not in st.session_state:
    st.session_state["feedback_df"] = load_feedback(FEEDBACK_PATH)

contact_df, manager_contacts = load_contact_map(CONTACT_PATH)

# 로그인용: 이름 -> 휴대폰 전체번호
contacts_phone = {
    name: info.get("phone", "")
    for name, info in manager_contacts.items()
    if info.get("phone", "")
}


# -----------------------------------------
# ⭐ 지사별 중간관리자 비밀번호 관리
# -----------------------------------------
BRANCH_ADMIN_PW = {
    "중앙": "C001",
    "강북": "C002",
    "서대문": "C003",
    "고양": "C004",
    "의정부": "C005",
    "남양주": "C006",
    "강릉": "C007",
    "원주": "C008",
}

# ==============================
# 로그인 스타일(CSS) 적용 ← ★ 이 부분이 정답!
# ==============================
login_css = """
<style>
.login-wrapper {
    position: fixed;
    top: 50%;
    left: 50%;
    transform: translate(-50%, -50%);
}
.login-card {
    width: 360px;
    padding: 35px;
    background: rgba(255,255,255,0.92);
    border-radius: 14px;
    box-shadow: 0 6px 20px rgba(0, 91, 172, 0.25);
}
.login-title {
    font-size: 26px;
    font-weight: 700;
    text-align: center;
    margin-bottom: 12px;
    color: #005BAC;
}
</style>
"""
st.markdown(login_css, unsafe_allow_html=True)

# ==============================
# 6. 로그인 폼 (연락처 뒷 4자리)
# ==============================
def login_form():

    st.markdown("## 🔐 로그인")

    tab_admin, tab_user, tab_branch_admin = st.tabs(
        ["관리자 로그인", "사용자 로그인", "중간관리자 로그인"]
    )

    # --------------------
    # 🔹 최고관리자 로그인
    # --------------------
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

    # --------------------
    # 🔹 사용자 로그인
    # --------------------
    with tab_user:
        name = st.text_input("성명", key="user_name")
        input_pw = st.text_input("연락처 뒷 4자리", type="password", key="user_pw")

        if st.button("사용자 로그인"):

            user_info = manager_contacts.get(name.strip())
            if not user_info:
                st.error("등록된 사용자명이 아닙니다.")
                return

            real_tel = user_info.get("phone", "")
            real_pw = real_tel[-4:] if len(real_tel) >= 4 else None

            if real_pw and input_pw == real_pw:
                st.session_state["login_type"] = "user"
                st.session_state["login_user"] = name.strip()
                st.success(f"{name} 님 로그인 성공")
                st.rerun()
            else:
                st.error("비밀번호가 올바르지 않습니다.")

    # --------------------
    # 🔹 지사 중간관리자 로그인
    # --------------------
    with tab_branch_admin:
        branch = st.selectbox("담당 지사 선택", list(BRANCH_ADMIN_PW.keys()), key="branch_select")
        name = st.text_input("중간관리자 성명", key="branch_admin_name")
        pw = st.text_input("중간관리자 비밀번호", type="password", key="branch_admin_pw")

        if st.button("중간관리자 로그인"):
            correct_pw = BRANCH_ADMIN_PW.get(branch)

            if pw == correct_pw:
                st.session_state["login_type"] = "branch_admin"
                st.session_state["login_user"] = name.strip()
                st.session_state["login_branch"] = branch
                st.success(f"{branch} 지사 중간관리자 로그인 성공!")
                st.rerun()
            else:
                st.error("비밀번호가 올바르지 않습니다.")



# 로그인 처리
if st.session_state["login_type"] is None:
    login_form()
    st.stop()

LOGIN_TYPE = st.session_state["login_type"]   # "admin" or "user"
LOGIN_USER = st.session_state["login_user"]   # 관리자: ADMIN / 사용자: 성명

# ==============================
# 7. 기본 전처리 (지사, 담당자, 출처 등)
# ==============================
# 지사 축약
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

# 주소 컬럼 자동 탐색
address_cols = [c for c in df.columns if "주소" in str(c)]

# 출처 분리
df_voc = df[df.get("출처") == "해지VOC"].copy()
df_other = df[df.get("출처") != "해지VOC"].copy()

# 👉 여기서 SP 필터 적용
if "담당유형" in df_voc.columns:
    df_voc = df_voc[df_voc["담당유형"].astype(str) == "SP"]

other_sets = {
    src: set(df_other[df_other["출처"] == src]["계약번호_정제"].dropna())
    for src in ["해지시설", "해지요청", "설변", "정지", "해지파이프라인"]
    if "출처" in df_other.columns
}
other_union = set().union(*other_sets.values()) if other_sets else set()

# 설치주소
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

# 월정료 정제
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

# 리스크/경과일 계산
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

df_voc[["경과일수", "리스크등급"]] = df_voc.apply(
    lambda r: pd.Series(compute_risk(r)), axis=1
)

def infer_cancel_reason(row):
    text_parts = []
    for col in ["해지상세", "VOC유형소", "등록내용"]:
        if col in row and pd.notna(row[col]):
            text_parts.append(str(row[col]))
    full_text = " ".join(text_parts)
    t = full_text.replace(" ", "").lower()

    # 경제적 사정
    econ = ["경제", "사정", "매출감소", "경영악화", "매출하락", "어려움", "고정비", "비용절감"]
    if any(k in t for k in econ):
        return "경제적 사정"

    # 품질/장애
    quality = ["장애", "고장", "불량", "끊김", "속도", "느림", "품질", "오류"]
    if any(k in t for k in quality):
        return "품질/장애 불만"

    # 가격/요금 불만
    price = ["비싸", "요금", "가격", "단가", "인상", "인하", "할인요청"]
    if any(k in t for k in price):
        return "요금/가격 불만"

    # 서비스/응대 불만
    svc = ["응대", "기사", "설치", "지연", "불친절", "안와요", "연락안옴"]
    if any(k in t for k in svc):
        return "서비스/응대 불만"

    # 경쟁사 이동
    comp = ["경쟁사", "타사", "다른회사", "이동", "옮김"]
    if any(k in t for k in comp):
        return "경쟁사/타사 이동"

    if full_text.strip():
        return "기타(텍스트 있음)"
    return "기타(정보 부족)"

def recommend_retention_policy(row):
    reason = row.get("AI_해지사유", "")
    risk = row.get("리스크등급", "LOW")
    fee = row.get("월정료_수치", np.nan)
    retp = row.get("리텐션P", np.nan)  # 없으면 NaN 유지

    # 월정료 티어
    if pd.notna(fee):
        if fee < 50000:
            fee_tier = "LOW"
        elif fee < 150000:
            fee_tier = "MID"
        else:
            fee_tier = "HIGH"
    else:
        fee_tier = "UNKNOWN"

    # 리텐션P 티어
    if pd.notna(retp):
        if retp >= 80:
            p_tier = "HIGH"
        elif retp >= 50:
            p_tier = "MID"
        else:
            p_tier = "LOW"
    else:
        p_tier = "UNKNOWN"

    primary = ""  # 추천1
    backup = ""   # 추천2
    comment = ""  # 상담 가이드

    # ----------------------
    # 경제적 사정
    # ----------------------
    if reason == "경제적 사정":
        if risk == "HIGH":
            if p_tier in ["HIGH", "MID"]:
                primary = "3개월간 월정료 30% 인하"
                backup = "2개월 유예 + 20% 인하"
                comment = "고객 재정 부담을 즉시 줄여줄 수 있는 인하/유예 정책을 우선 제안하세요."
            else:
                primary = "2개월간 20% 인하"
                backup = "1개월 유예 + 10% 인하"
                comment = "리텐션 여력이 낮아 무리한 인하보다는 중간 수준 인하를 제안하세요."
        elif risk == "MEDIUM":
            primary = "2개월간 10~20% 인하"
            backup = "1개월 유예"
            comment = "중간 리스크로, 단기간 인하와 유예 조합이 효과적입니다."
        else:
            primary = "1개월 유예 또는 10% 인하"
            backup = "서비스 혜택/가치 재설명 중심 설득"
            comment = "리스크가 낮으므로 소폭 혜택 + 설득 위주 접근이 적절합니다."

    # ----------------------
    # 품질/장애
    # ----------------------
    elif reason == "품질/장애 불만":
        primary = "무상 점검 + 1개월 요금감면"
        backup = "품질 모니터링 강화 및 장애 시 우선 출동 약속"
        comment = "장애 원인 설명과 함께 사후 관리 약속이 핵심입니다."

    # ----------------------
    # 가격/요금 불만
    # ----------------------
    elif reason == "요금/가격 불만":
        primary = "요금제 재구성(저가 요금안 제시) + 소폭 할인"
        backup = "옵션/부가서비스 정리로 총액 절감안 제시"
        comment = "가격 민감 고객에게는 상품 구조 변경 + 소폭 인하가 효과적입니다."

    # ----------------------
    # 서비스/응대 불만
    # ----------------------
    elif reason == "서비스/응대 불만":
        primary = "정식 사과 + 담당자 변경 + 소정의 보상(1개월 감면 등)"
        backup = "전담 관리 채널/담당자 지정"
        comment = "신뢰 회복과 응대 품질 개선 메시지를 중심으로 설득하세요."

    # ----------------------
    # 경쟁사 이동
    # ----------------------
    elif reason == "경쟁사/타사 이동":
        primary = "자사 강점/차별점 설명 + 적정 수준 혜택 제시"
        backup = "장기고객/충성고객 대상 추가 혜택 제안"
        comment = "과도한 할인보다는 차별화 포인트 + 적정 혜택 조합이 중요합니다."

    # ----------------------
    # 기타
    # ----------------------
    else:
        primary = "서비스 가치/필요성 설명 중심 유지 설득"
        backup = "고객 상황에 맞춘 맞춤형 조건 협의"
        comment = "사유가 뚜렷하지 않아, 대화를 통해 니즈를 다시 파악하는 것이 필요합니다."

    return {
        "primary_action": primary,
        "backup_action": backup,
        "comment": comment,
        "reason": reason,
        "risk": risk,
        "retp_tier": p_tier,
        "fee_tier": fee_tier,
    }

# 매칭여부
df_voc["매칭여부"] = df_voc["계약번호_정제"].apply(
    lambda x: "매칭(O)" if x in other_union else "비매칭(X)"
)

# 로그인 타입별 비매칭 풀 (unmatched_global) - 초기 버전
if LOGIN_TYPE == "user":
    df_user = df_voc[df_voc["구역담당자_통합"] == LOGIN_USER]
    unmatched_global = df_user[df_user["매칭여부"] == "비매칭(X)"].copy()
else:
    unmatched_global = df_voc[df_voc["매칭여부"] == "비매칭(X)"].copy()

# ==============================
# 8. 표시 컬럼 / 스타일링
# ==============================
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

# ==============================
# 9. 사이드바 글로벌 필터
# ==============================
st.sidebar.title("🔧 글로벌 필터")

# ==============================
# 📌 날짜 파싱 엔진(완성 버전)
# ==============================

def parse_date_safe(x):
    """모든 형태의 날짜를 강력하게 처리하는 통합 파서"""

    if pd.isna(x):
        return pd.NaT

    # 이미 datetime 형태
    if isinstance(x, (pd.Timestamp, datetime)):
        return x

    s = str(x).strip()

    if s in ["", "None", "nan", "NaN"]:
        return pd.NaT

    # 기본 자동 파싱 시도
    try:
        dt = pd.to_datetime(s, errors="coerce", infer_datetime_format=True)
        if pd.notna(dt):
            return dt
    except:
        pass

    # 사람이 입력한 다양한 패턴 수동 처리
    date_formats = [
        "%Y-%m-%d",
        "%Y/%m/%d",
        "%Y.%m.%d",
        "%Y-%m-%d %H:%M",
        "%Y/%m/%d %H:%M",
        "%Y.%m.%d %H:%M",
        "%Y-%m-%d %H:%M:%S",
        "%Y/%m/%d %H:%M:%S",
        "%Y.%m.%d %H:%M:%S",
        "%Y-%m-%d %p %I:%M",
        "%Y/%m/%d %p %I:%M",
    ]

    for fmt in date_formats:
        try:
            return datetime.strptime(s, fmt)
        except:
            continue

    # 정말 안 될 경우
    return pd.NaT


# 최종 적용
if "접수일시" in df_voc.columns:
    df_voc["접수일시"] = df_voc["접수일시"].apply(parse_date_safe)


# 지사 필터
branches_all = sort_branch(df_voc["관리지사"].dropna().unique())
sel_branches = st.sidebar.pills(
    "🏢 관리지사 선택",
    options=["전체"] + branches_all,
    selection_mode="multi",
    default=["전체"],
    key="filter_branch_btn",
)

# 리스크 필터
risk_all = ["HIGH", "MEDIUM", "LOW"]
sel_risk = st.sidebar.pills(
    "⚠ 리스크등급",
    options=risk_all,
    selection_mode="multi",
    default=risk_all,
    key="filter_risk_btn",
)

# 매칭여부 필터
match_all = ["매칭(O)", "비매칭(X)"]
sel_match = st.sidebar.pills(
    "🔍 매칭여부",
    options=match_all,
    selection_mode="multi",
    default=["비매칭(X)"],
    key="filter_match_btn",
)

# ---------------------------------------
# 💰 월정료 구간 (라디오 선택형 + 슬라이더 추가)
# ---------------------------------------

fee_bands = [
    "전체",
    "10만 이하",
    "10만~30만",
    "30만 이상",
]

# 라디오 필터
sel_fee_band_radio = st.sidebar.radio(
    "💰 월정료 구간",
    options=fee_bands,
    index=0,
    key="filter_fee_band_radio",
)

# 슬라이더 필터 (만원 단위)
fee_slider_min, fee_slider_max = st.sidebar.slider(
    "🔧 월정료 직접 범위 설정(만원)",
    min_value=0,
    max_value=100,
    value=(0, 100),
    step=1,
    key="filter_fee_band_slider",
)

st.sidebar.markdown("---")
st.sidebar.caption(f"마지막 갱신: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")

# ==============================
# 🔍 담당유형 필터 추가
# ==============================
if "담당유형" in df_voc.columns:
    담당유형_list = (
        ["전체"] 
        + sorted(df_voc["담당유형"].dropna().astype(str).unique().tolist())
    )
    sel_mgr_type = st.sidebar.selectbox(
        "👤 담당유형 선택",
        options=담당유형_list,
        index=담당유형_list.index("SP") if "SP" in 담당유형_list else 0,
        key="filter_mgr_type"
    )
else:
    sel_mgr_type = "전체"

# ==============================
# 🔍 VOC 유형 필터 추가
# ==============================

# VOC유형중 (중분류)
if "VOC유형중" in df_voc.columns:
    voc_mid_values = (
        ["전체"] 
        + sorted(df_voc["VOC유형중"].dropna().astype(str).unique().tolist())
    )
    sel_voc_mid = st.sidebar.selectbox(
        "📌 VOC유형중(중분류)",
        options=voc_mid_values,
        index=0,
        key="filter_voc_mid"
    )
else:
    sel_voc_mid = "전체"

# VOC유형소 (소분류)
if "VOC유형소" in df_voc.columns:
    voc_small_values = (
        ["전체"] 
        + sorted(df_voc["VOC유형소"].dropna().astype(str).unique().tolist())
    )
    sel_voc_small = st.sidebar.selectbox(
        "📌 VOC유형소(소분류)",
        options=voc_small_values,
        index=0,
        key="filter_voc_small"
    )
else:
    sel_voc_small = "전체"

# VOC유형 (소2)
if "VOC유형" in df_voc.columns:
    voc_type_values = (
        ["전체"] 
        + sorted(df_voc["VOC유형"].dropna().astype(str).unique().tolist())
    )
    sel_voc_type = st.sidebar.selectbox(
        "📌 VOC유형(대분류)",
        options=voc_type_values,
        index=0,
        key="filter_voc_type"
    )
else:
    sel_voc_type = "전체"


if st.sidebar.button("🔄 필터 초기화"):
    for key in list(st.session_state.keys()):
        if "filter" in key or "fee" in key:
            del st.session_state[key]
    st.success("글로벌 필터가 초기화되었습니다.")
    st.rerun()

# ---------------------------------------
# 🔐 로그인 타입별 데이터 접근 제어
# ---------------------------------------
voc_filtered_role = df_voc.copy()

# ➤ 일반 사용자: 본인 담당 데이터만
if LOGIN_TYPE == "user":
    voc_filtered_role = voc_filtered_role[
        voc_filtered_role["구역담당자_통합"].astype(str) == LOGIN_USER
    ]

# ➤ 중간관리자: 본인 지사 전체 데이터
elif LOGIN_TYPE == "branch_admin":
    branch = st.session_state.get("login_branch", "")
    voc_filtered_role = voc_filtered_role[
        voc_filtered_role["관리지사"].astype(str) == branch
    ]

# ➤ 최고관리자(admin): 모든 데이터 접근 가능

# 이후 글로벌 필터 적용
voc_filtered_global = voc_filtered_role.copy()

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

# 리스크 필터
if sel_risk and "리스크등급" in voc_filtered_global.columns:
    voc_filtered_global = voc_filtered_global[
        voc_filtered_global["리스크등급"].isin(sel_risk)
    ]

# 매칭여부 필터
if sel_match and "매칭여부" in voc_filtered_global.columns:
    voc_filtered_global = voc_filtered_global[
        voc_filtered_global["매칭여부"].isin(sel_match)
    ]

# 💰 월정료 필터 (라디오 + 슬라이더)
if fee_raw_col is not None and "월정료_수치" in voc_filtered_global.columns:
    fee_series = voc_filtered_global["월정료_수치"].fillna(-1)

    # ① 라디오 구간 필터
    if sel_fee_band_radio == "10만 이하":
        voc_filtered_global = voc_filtered_global[
            (fee_series >= 0) & (fee_series < 100000)
        ]
    elif sel_fee_band_radio == "10만~30만":
        voc_filtered_global = voc_filtered_global[
            (fee_series >= 100000) & (fee_series < 300000)
        ]
    elif sel_fee_band_radio == "30만 이상":
        voc_filtered_global = voc_filtered_global[
            (fee_series >= 300000)
        ]
    # "전체"는 패스

    # ② 슬라이더 추가 정밀 필터 (만원 → 원 단위 변환)
    slider_min_won = fee_slider_min * 10000
    slider_max_won = fee_slider_max * 10000

    fee_series2 = voc_filtered_global["월정료_수치"].fillna(-1)
    voc_filtered_global = voc_filtered_global[
        (fee_series2 >= slider_min_won) & (fee_series2 <= slider_max_won)
    ]

# 로그인 타입별 접근 제한 (사용자일 경우 한 번 더 안전하게)
if LOGIN_TYPE == "user":
    if "구역담당자_통합" in voc_filtered_global.columns:
        voc_filtered_global = voc_filtered_global[
            voc_filtered_global["구역담당자_통합"].astype(str) == str(LOGIN_USER)
        ]

# 비매칭 데이터
unmatched_global = voc_filtered_global[
    voc_filtered_global["매칭여부"] == "비매칭(X)"
].copy()

# ==============================
# 10. 상단 KPI
# ==============================
st.write("")
st.markdown("## 📊 해지 VOC 종합 대시보드")

total_voc_rows = len(voc_filtered_global)
unique_contracts = voc_filtered_global["계약번호_정제"].nunique()
unmatched_contracts = unmatched_global["계약번호_정제"].nunique()
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

# ==============================
# 11. 탭 구성
# ==============================
tab_viz, tab_all, tab_unmatched, tab_drill, tab_filter, tab_alert, tab_branch_admin_report = st.tabs(
    [
        "📊 지사/담당자 시각화",
        "📘 VOC 전체(계약 기준)",
        "🧯 해지방어 활동시설(비매칭)",
        "🔍 해지상담대상 활동등록",
        "🎯 정밀 필터",
        "📨 담당자 알림",
        "🏢 지사 관리자 전용",
    ]
)

# ----------------------------------------------------
# 🏢 지사 관리자 전용 대시보드
# ----------------------------------------------------
with tab_branch_admin_report:
    if LOGIN_TYPE != "branch_admin":
        st.info("이 탭은 지사 관리자만 접근할 수 있습니다.")
    else:
        branch = st.session_state.get("login_branch", "")
        st.subheader(f"🏢 {branch} 지사 관리자 대시보드")

        df_branch = df_voc[df_voc["관리지사"] == branch]

        st.metric("총 VOC 건수", len(df_branch))
        st.metric("비매칭 계약 수", df_branch[df_branch["매칭여부"] == "비매칭(X)"]["계약번호_정제"].nunique())

        st.markdown("### 🔥 리스크별 비매칭 구조")
        rc = (
            df_branch[df_branch["매칭여부"] == "비매칭(X)"]["리스크등급"]
            .value_counts()
            .reindex(["HIGH","MEDIUM","LOW"])
            .fillna(0)
        )
        st.bar_chart(rc)

        st.markdown("### 📋 지사 전체 비매칭 리스트")
        st.dataframe(
            df_branch[df_branch["매칭여부"]=="비매칭(X)"][display_cols],
            use_container_width=True,
            height=450,
        )

# ------------------------------------------------
# 🔹 적층 세로 막대그래프 (Plotly)
# ------------------------------------------------
def force_stacked_bar(df: pd.DataFrame, x: str, y_cols: list[str], height: int = 280):
    """
    Plotly 적용된 적층 세로 막대그래프
    df: DataFrame
    x: x축 컬럼명
    y_cols: 적층할 수치 컬럼 리스트 ["HIGH","MEDIUM","LOW"]
    """
    if df.empty or not y_cols:
        st.info("표시할 데이터가 없습니다.")
        return

    if HAS_PLOTLY:
        fig = px.bar(
            df,
            x=x,
            y=y_cols,
            barmode="stack",
            text_auto=True,
            height=height,
        )
        fig.update_layout(
            margin=dict(l=40, r=20, t=40, b=40),
        )
        st.plotly_chart(fig, use_container_width=True)
    else:
        st.warning("Plotly가 설치되어야 적층 막대그래프를 표시할 수 있습니다.")


# ----------------------------------------------------
# TAB VIZ — 지사 / 담당자 시각화
# ----------------------------------------------------
with tab_viz:
    viz_base = unmatched_global.copy()
    if "리스크등급" not in viz_base.columns:
        viz_base["리스크등급"] = "LOW"

    st.subheader("📊 지사 / 담당자별 비매칭 리스크 현황")

    if viz_base.empty:
        st.info("현재 조건에서 비매칭(X) 데이터가 없습니다.")
        st.stop()

    # UI 상단 안내 박스
    st.markdown(
        """
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
        """,
        unsafe_allow_html=True,
    )

    colA, colB = st.columns(2)

    # -----------------------------
    # 지사 선택
    # -----------------------------
    b_opts = ["전체"] + sort_branch(viz_base["관리지사"].dropna().unique())
    sel_b_viz = colA.pills(
        "🏢 지사 선택",
        options=b_opts,
        selection_mode="single",
        default="전체",
        key="viz_branch",
    )
    sel_b_viz = sel_b_viz[0] if isinstance(sel_b_viz, list) else sel_b_viz

    # -----------------------------
    # 담당자 선택
    # -----------------------------
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

    # -----------------------------
    # 필터 적용
    # -----------------------------
    viz_filtered = viz_base.copy()
    if sel_b_viz != "전체":
        viz_filtered = viz_filtered[viz_filtered["관리지사"] == sel_b_viz]
    if sel_mgr_viz != "(전체)":
        viz_filtered = viz_filtered[
            viz_filtered["구역담당자_통합"].astype(str) == sel_mgr_viz
        ]

    if viz_filtered.empty:
        st.info("선택한 조건에서 비매칭(X) 데이터가 없습니다.")
        st.stop()

    # ======================================================
    # 1) 지사별 비매칭 적층막대
    # ======================================================
    st.markdown("### 🧱 지사별 비매칭 계약 수 (유니크 계약, 리스크 적층)")

    branch_risk = (
        viz_filtered.groupby(["관리지사", "리스크등급"])["계약번호_정제"]
        .nunique()
        .reset_index(name="계약수")
    )

    if not branch_risk.empty:
        pivot_branch = branch_risk.pivot(
            index="관리지사", columns="리스크등급", values="계약수"
        ).fillna(0)

        pivot_branch = pivot_branch.reindex(BRANCH_ORDER).fillna(0)

        stack_cols = [c for c in ["HIGH", "MEDIUM", "LOW"] if c in pivot_branch.columns]

        force_stacked_bar(
            pivot_branch.reset_index(),
            x="관리지사",
            y_cols=stack_cols,
            height=260,
        )
    else:
        st.info("지사별 데이터가 없습니다.")

    # ======================================================
    # 2) 담당자 TOP 15 적층막대
    # ======================================================
    c2a, c2b = st.columns(2)

    with c2a:
        st.markdown("### 👤 담당자별 비매칭 TOP 15 (유니크 계약, 리스크 적층)")

        mgr_risk = (
            viz_filtered.groupby(["구역담당자_통합", "리스크등급"])["계약번호_정제"]
            .nunique()
            .reset_index(name="계약수")
        )

        if not mgr_risk.empty:
            pivot_mgr = mgr_risk.pivot(
                index="구역담당자_통합",
                columns="리스크등급",
                values="계약수"
            ).fillna(0)

            stack_cols_mgr = [c for c in ["HIGH", "MEDIUM", "LOW"] if c in pivot_mgr.columns]

            pivot_mgr["총계"] = pivot_mgr[stack_cols_mgr].sum(axis=1)
            pivot_mgr = pivot_mgr.sort_values("총계", ascending=False).head(15)
            pivot_mgr.drop(columns=["총계"], inplace=True)

            force_stacked_bar(
                pivot_mgr.reset_index(),
                x="구역담당자_통합",
                y_cols=stack_cols_mgr,
                height=300,
            )
        else:
            st.info("담당자 데이터가 없습니다.")

    # ======================================================
    # 3) 전체 리스크 등급 분포 적층 단일 막대
    # ======================================================
    with c2b:
        st.markdown("### 🔥 리스크 등급 분포 (계약 단위, 적층 막대)")

        rc = (
            viz_filtered["리스크등급"].value_counts()
            .reindex(["HIGH", "MEDIUM", "LOW"])
            .fillna(0)
        )

        risk_df = pd.DataFrame({
            "구분": ["전체"],
            "HIGH": [rc["HIGH"]],
            "MEDIUM": [rc["MEDIUM"]],
            "LOW": [rc["LOW"]],
        })

        force_stacked_bar(
            risk_df,
            x="구분",
            y_cols=["HIGH", "MEDIUM", "LOW"],
            height=300,
        )

    # ======================================================
    # 4) 일별 추이
    # ======================================================
    st.markdown("---")
    st.markdown("### 📈 일별 비매칭 계약 추이")

    if "접수일시" in viz_filtered.columns and viz_filtered["접수일시"].notna().any():
        trend = (
            viz_filtered.assign(접수일=viz_filtered["접수일시"].dt.date)
            .groupby("접수일")["계약번호_정제"]
            .nunique()
            .sort_index()
        )

        fig4 = px.line(trend.reset_index(), x="접수일", y="계약번호_정제")
        fig4.update_layout(height=260)
        st.plotly_chart(fig4, use_container_width=True)
    else:
        st.info("접수일시 데이터가 없습니다.")

    # ======================================================
    # 5) 담당자 리스크 레이더
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

            radar_df = pd.DataFrame({
                "리스크": ["HIGH", "MEDIUM", "LOW"],
                "계약수": radar.values,
            })

            fig_radar = px.line_polar(
                radar_df, r="계약수", theta="리스크", line_close=True
            )
            fig_radar.update_layout(height=320)
            st.plotly_chart(fig_radar, use_container_width=True)

# ======================================================
# 8) 추가 분석 그래프 (산점도 / 트리맵 / 히스토그램 / 박스플롯 / 도넛차트)
# ======================================================
st.markdown("---")
st.subheader("📐 추가 분석 그래프")

# ------------------------------------------------------
# 🔸 1. 산점도 (관리지사 / 담당자 / 계약건수 기반)
# ------------------------------------------------------
st.markdown("### 🔹 산점도 (관리지사 · 담당자 · 계약건수)")

if {"관리지사", "구역담당자_통합", "계약번호_정제"}.issubset(viz_filtered.columns):

    # 계약수 집계
    scatter_df = (
        viz_filtered.groupby(["관리지사", "구역담당자_통합", "리스크등급"])
        .agg(계약수=("계약번호_정제", "nunique"))
        .reset_index()
    )

    # 누락 표시 제거
    scatter_df["구역담당자_통합"] = scatter_df["구역담당자_통합"].fillna("(미배정)")

    fig_scat = px.scatter(
        scatter_df,
        x="관리지사",
        y="구역담당자_통합",
        size="계약수",
        color="리스크등급",
        hover_data=["계약수", "관리지사", "구역담당자_통합"],
        title="관리지사 · 담당자별 계약규모 산점도",
    )

    fig_scat.update_layout(height=450)
    st.plotly_chart(fig_scat, use_container_width=True)

else:
    st.info("관리지사 / 담당자 / 계약번호 정보를 찾을 수 없습니다.")
# ------------------------------------------------------
# 🔸 2. 트리맵 (지사 → 담당자 → 계약수)
# ------------------------------------------------------
if {"관리지사", "구역담당자_통합", "계약번호_정제"}.issubset(viz_filtered.columns):
    st.markdown("### 🔹 트리맵 (지사 → 담당자 → 계약수)")

    tree_df = (
        viz_filtered.groupby(["관리지사", "구역담당자_통합"])
        .agg(계약수=("계약번호_정제", "nunique"))
        .reset_index()
    )

    fig_tree = px.treemap(
        tree_df,
        path=["관리지사", "구역담당자_통합"],
        values="계약수",
        title="지사-담당자 구조 트리맵",
        color="계약수",
        color_continuous_scale="Blues",
    )
    st.plotly_chart(fig_tree, use_container_width=True)
# ------------------------------------------------------
# 🔸 3. 히스토그램 (월정료 / 경과일)
# ------------------------------------------------------
if "월정료_수치" in viz_filtered.columns:
    st.markdown("### 🔹 월정료 분포 (히스토그램)")
    fig_fee_hist = px.histogram(
        viz_filtered,
        x="월정료_수치",
        nbins=30,
        title="월정료 분포",
    )
    st.plotly_chart(fig_fee_hist, use_container_width=True)

if "경과일수" in viz_filtered.columns:
    st.markdown("### 🔹 경과일수 분포 (히스토그램)")
    fig_day_hist = px.histogram(
        viz_filtered,
        x="경과일수",
        nbins=30,
        title="VOC 경과일 분포",
    )
    st.plotly_chart(fig_day_hist, use_container_width=True)

# ------------------------------------------------------
# 🔸 4. 박스플롯 (지사별 월정료 / 경과일)
# ------------------------------------------------------
if "관리지사" in viz_filtered.columns and "월정료_수치" in viz_filtered.columns:
    st.markdown("### 🔹 박스플롯 — 지사별 월정료 비교")
    fig_fee_box = px.box(
        viz_filtered,
        x="관리지사",
        y="월정료_수치",
        points="all",
        color="관리지사",
    )
    st.plotly_chart(fig_fee_box, use_container_width=True)

if "관리지사" in viz_filtered.columns and "경과일" in viz_filtered.columns:
    st.markdown("### 🔹 박스플롯 — 지사별 VOC 경과일 비교")
    fig_day_box = px.box(
        viz_filtered,
        x="관리지사",
        y="경과일수",
        points="all",
        color="관리지사",
    )
    st.plotly_chart(fig_day_box, use_container_width=True)

# ------------------------------------------------------
# 🔸 5. 도넛 차트 (리스크 등급 비율)
# ------------------------------------------------------
if "리스크등급" in viz_filtered.columns:
    st.markdown("### 🔹 Risk 등급 비율 (도넛 차트)")
    rc = viz_filtered["리스크등급"].value_counts().reset_index()
    rc.columns = ["리스크등급", "건수"]

    fig_donut = px.pie(
        rc,
        names="리스크등급",
        values="건수",
        hole=0.5,
        title="리스크등급 비율",
    )
    st.plotly_chart(fig_donut, use_container_width=True)


    # ======================================================
    # 6) 텍스트 키워드 분석
    # ======================================================
    st.markdown("---")
    st.markdown("### 📝 텍스트 키워드 분석 (등록내용 + 처리내용 + 해지상세 + VOC유형소)")

    text_cols = ["등록내용", "처리내용", "해지상세", "VOC유형소"]
    available_cols = [c for c in text_cols if c in viz_filtered.columns]

    if available_cols:
        texts = []
        for col in available_cols:
            texts.extend(viz_filtered[col].dropna().astype(str).tolist())

        import re
        from collections import Counter

        words = re.findall(r"[가-힣A-Za-z]{2,}", " ".join(texts))
        freq_df = pd.DataFrame(Counter(words).most_common(50), columns=["단어", "빈도"])

        st.markdown("#### 🔍 최다 빈도 단어 TOP 50")
        force_bar_chart(freq_df, "단어", "빈도", height=350)

# ------------------------------------------------------------
# 공통: 지사 색상 테마 설정
# ------------------------------------------------------------
branch_color_map = {
    "강릉": "#1f77b4",
    "강북": "#ff7f0e",
    "고양": "#2ca02c",
    "남양주": "#d62728",
    "서대문": "#9467bd",
    "원주": "#8c564b",
    "의정부": "#e377c2",
    "중앙": "#7f7f7f",
    "기타": "#bcbd22",
}

st.markdown("### 🎛 산점도 필터 옵션")

risk_filter = st.multiselect(
    "리스크 등급 선택",
    ["HIGH", "MEDIUM", "LOW"],
    default=["HIGH", "MEDIUM", "LOW"],
)

mgr_search = st.text_input("담당자 검색어 입력 (부분검색 가능)")

scatter_df = viz_filtered.copy()
scatter_df = scatter_df[scatter_df["리스크등급"].isin(risk_filter)]

if mgr_search:
    scatter_df = scatter_df[
        scatter_df["구역담당자_통합"].astype(str).str.contains(mgr_search)
    ]

show_labels = st.checkbox("버블 위에 담당자 이름 표시", value=False)

st.markdown("### 🔵 고급 산점도 (지사 · 담당자 · 계약규모)")

# 버블 크기 선택
size_option = st.selectbox(
    "버블 크기 기준",
    ["계약건수", "월정료_수치", "경과일수"],
    index=0,
)

if size_option == "계약건수":
    temp = scatter_df.groupby(
        ["관리지사", "구역담당자_통합"]
    )["계약번호_정제"].nunique().reset_index()
    temp.rename(columns={"계약번호_정제": "bubble_size"}, inplace=True)
    scatter_df = scatter_df.merge(temp, on=["관리지사", "구역담당자_통합"], how="left")
    size_col = "bubble_size"
else:
    size_col = size_option

fig = px.scatter(
    scatter_df,
    x="관리지사",
    y="구역담당자_통합",
    size=size_col,
    color="관리지사",
    hover_data=["계약번호_정제", "월정료_수치", "경과일수", "리스크등급"],
    color_discrete_map=branch_color_map,
    opacity=0.8,
)

# 라벨 표시 옵션
if show_labels:
    fig.update_traces(text=scatter_df["구역담당자_통합"], textposition="top center")

fig.update_layout(
    height=600,
    title="📌 지사 · 담당자별 계약규모 산점도 (확장형)",
)
st.plotly_chart(fig, use_container_width=True)

st.markdown("### 📦 경과일수 박스플롯 (담당자별 지연 분석)")

if "경과일수" in viz_filtered.columns:
    fig_box = px.box(
        viz_filtered,
        x="구역담당자_통합",
        y="경과일수",
        color="관리지사",
        color_discrete_map=branch_color_map,
        title="담당자별 경과일수 분포 (지연 위험도 분석)",
    )
    fig_box.update_layout(height=550)
    st.plotly_chart(fig_box, use_container_width=True)
else:
    st.info("경과일수 데이터가 없어 박스플롯을 표시할 수 없습니다.")

st.markdown("### 🌳 Treemap (지사 → 담당자 → 리스크)")

tree_df = (
    viz_filtered
    .groupby(["관리지사", "구역담당자_통합", "리스크등급"])["계약번호_정제"]
    .nunique()
    .reset_index(name="계약수")
)

fig_tree = px.treemap(
    tree_df,
    path=["관리지사", "구역담당자_통합", "리스크등급"],
    values="계약수",
    color="관리지사",
    color_discrete_map=branch_color_map,
)
st.plotly_chart(fig_tree, use_container_width=True)

def ai_voc_risk_predict(row):
    text = " ".join([
        str(row.get("등록내용", "")),
        str(row.get("처리내용", "")),
        str(row.get("해지상세", "")),
    ]).lower()

    # 기본값
    reason = "미분류"
    risk = "LOW"

    # 규칙 기반 기본 분류
    if any(k in text for k in ["비싸", "요금", "부담", "가격"]):
        reason, risk = "경제적 사정", "MEDIUM"
    if any(k in text for k in ["불만", "항의", "문의 많음", "불친절"]):
        reason, risk = "서비스 불만", "HIGH"
    if any(k in text for k in ["타사", "경쟁사", "이동"]):
        reason, risk = "경쟁사 이동", "MEDIUM"

    # 고급 모델 확장 가능 부분 (OpenAI/LLM)
    # 여기서는 placeholder
    # ex) gpt_model.predict(text)

    return reason, risk

st.markdown("### 🤖 AI 기반 VOC 위험군 자동 분석")

ai_df = viz_filtered.copy()
ai_df["AI_사유"], ai_df["AI_리스크"] = zip(*ai_df.apply(ai_voc_risk_predict, axis=1))

ai_summary = ai_df["AI_리스크"].value_counts()

fig_ai = px.bar(
    ai_summary,
    title="AI 추론 리스크 분포",
    labels={"value": "건수", "index": "AI 리스크"},
    text_auto=True,
)
st.plotly_chart(fig_ai, use_container_width=True)


# ------------------------------------------------------
# 🔸 4. 경과일수 박스플롯 (지사/담당자 지연 분석, 개선 버전)
# ------------------------------------------------------
st.markdown("### 📦 경과일수 박스플롯 (지사/담당자 지연 분석)")

if "경과일수" in viz_filtered.columns:

    # --- 상단 컨트롤 ---
    c_box1, c_box2, c_box3 = st.columns([2, 2, 2])

    # 지사 필터
    branch_opts_box = ["전체"] + sort_branch(
        viz_filtered["관리지사"].dropna().unique()
    )
    sel_branch_box = c_box1.selectbox(
        "지사 선택",
        options=branch_opts_box,
        index=0,
        key="box_branch",
    )

    # 리스크 필터
    risk_opts_box = ["HIGH", "MEDIUM", "LOW"]
    sel_risk_box = c_box2.multiselect(
        "리스크등급 필터",
        options=risk_opts_box,
        default=risk_opts_box,
        key="box_risk",
    )

    # 상위 N명 (경과일수 긴 담당자만 추리기)
    top_n_mgr = c_box3.slider(
        "상위 담당자 N (경과일수 중앙값 기준)",
        min_value=5,
        max_value=50,
        value=20,
        step=5,
        key="box_top_n",
    )

    # --- 데이터 필터링 ---
    box_df = viz_filtered.copy()

    if sel_branch_box != "전체":
        box_df = box_df[box_df["관리지사"] == sel_branch_box]

    if sel_risk_box:
        box_df = box_df[box_df["리스크등급"].isin(sel_risk_box)]

    # 담당자/지사 라벨 생성: "지사 / 담당자"
    box_df["담당자_라벨"] = (
        box_df["관리지사"].fillna("미지정") + " / " +
        box_df["구역담당자_통합"].fillna("미지정")
    )

    # 데이터가 없으면 종료
    if box_df.empty:
        st.info("선택한 조건에서 표시할 데이터가 없습니다.")
    else:
        # --- 담당자별 경과일수 중앙값/건수 집계 ---
        agg_box = (
            box_df.groupby(["관리지사", "구역담당자_통합", "담당자_라벨"])
            .agg(
                경과일수_중앙값=("경과일수", "median"),
                계약건수=("계약번호_정제", "nunique"),
            )
            .reset_index()
        )

        # 경과일수 중앙값이 긴 담당자 상위 N명만 선택
        agg_box = agg_box.sort_values(
            "경과일수_중앙값", ascending=False
        ).head(top_n_mgr)

        top_labels = agg_box["담당자_라벨"].tolist()
        box_df_top = box_df[box_df["담당자_라벨"].isin(top_labels)].copy()

        st.caption(
            f"표시 대상 담당자 수: {len(top_labels)}명 "
            f"(경과일수 중앙값 상위 {top_n_mgr}명 기준)"
        )

        # --- 박스플롯 그리기 ---
        fig_box = px.box(
            box_df_top,
            x="담당자_라벨",
            y="경과일수",
            color="관리지사",
            points="outliers",  # 이상치만 점으로 표시
            hover_data=[
                "관리지사",
                "구역담당자_통합",
                "계약번호_정제",
                "상호",
                "리스크등급",
            ],
            title="담당자별 경과일수 분포 (지사 포함)",
        )

        # 전체 평균선 추가
        mean_days = box_df_top["경과일수"].mean()
        fig_box.add_hline(
            y=mean_days,
            line_dash="dash",
            annotation_text=f"전체 평균 {mean_days:.1f}일",
            annotation_position="top left",
        )

        # 레이아웃 튜닝 (라벨 회전/여백)
        fig_box.update_layout(
            xaxis_title="담당자 (지사 / 담당자명)",
            yaxis_title="경과일수",
            height=550,
            margin=dict(l=40, r=20, t=60, b=180),
            legend_title_text="관리지사",
        )
        fig_box.update_xaxes(
            tickangle=-45,
            tickfont=dict(size=10),
            categoryorder="array",
            categoryarray=top_labels,  # 중앙값 기준 정렬 유지
        )

        st.plotly_chart(fig_box, use_container_width=True)

else:
    st.info("경과일수 컬럼이 없습니다.")

# ----------------------------------------------------
# TAB ALL — VOC 전체 (계약번호 기준 요약)
# ----------------------------------------------------
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
            "AI_해지사유",
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

# ----------------------------------------------------
# TAB UNMATCHED — 해지방어 활동시설(비매칭)
# ----------------------------------------------------
with tab_unmatched:
    st.subheader("🧯 해지방어 활동시설 (비매칭, 계약번호 기준)")
    st.caption("비매칭(X) = 해지 VOC 접수 후 시스템상 활동내역이 확인되지 않은 시설")

with st.expander("ℹ️ 해지방어 활동시설 안내", expanded=True):
    st.write(
        "해지VOC 접수 후 **해지방어 활동내역이 시스템에 등록되지 않은 시설**입니다.\n"
        "- 실제 현장 대응 여부를 신속히 확인해 주세요.\n"
        "- 확인 후에는 반드시 `해지상담대상 활동등록` 탭에서 처리내역을 남겨주세요."
    )

    if unmatched_global.empty:
        st.info("현재 글로벌 필터 조건에서 비매칭(X) 계약이 없습니다.")
    else:
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

            st.data_editor(
                df_u_summary[summary_cols_u].reset_index(drop=True),
                use_container_width=True,
                height=420,
                hide_index=True,
                key="tab2_unmatched_editor",
            )

            # 행 선택 연계
            selected_idx = None
            state = st.session_state.get("tab2_unmatched_editor", {})
            selected_rows = []
            if isinstance(state, dict):
                if "selected_rows" in state and state["selected_rows"]:
                    selected_rows = state["selected_rows"]
                elif "selection" in state and isinstance(state["selection"], dict):
                    rows_sel = state["selection"].get("rows")
                    if rows_sel:
                        selected_rows = rows_sel
            if selected_rows:
                selected_idx = selected_rows[0]

            u_contract_list = df_u_summary["계약번호_정제"].astype(str).tolist()
            default_index = 0
            if selected_idx is not None and 0 <= selected_idx < len(u_contract_list):
                default_index = selected_idx + 1  # "(선택)" offset

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

                    st.markdown(f"##### VOC 이력 ({len(voc_detail)}건)")
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

# ----------------------------------------------------
# TAB DRILL — 해지상담대상 활동등록 (계약별 드릴다운)
# ----------------------------------------------------
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

        # 🔹 3번: AI 기반 방어 정책 추천 블록
        st.markdown("### 🤖 AI 기반 방어 정책 추천")

        # 방어정책 계산 (리텐션P 컬럼이 없으면 NaN으로 처리됨)
        rec = recommend_retention_policy(base_info)

        st.markdown(f"- **추론된 해지 사유:** `{rec['reason']}`")
        st.markdown(
            f"- **리스크 등급:** `{rec['risk']}` / "
            f"**리텐션P 티어:** `{rec['retp_tier']}` / "
            f"**월정료 티어:** `{rec['fee_tier']}`"
        )
        
        st.markdown("#### ✅ 1차 권장 정책")
        st.success(rec["primary_action"])
        
        st.markdown("#### 🔄 대안 정책")
        st.info(rec["backup_action"])
        
        st.markdown("#### 💬 상담 시 활용 가이드")
        st.write(rec["comment"])
        
        st.markdown("---")
        st.markdown("---")
        
        # LEFT / RIGHT 영역 구성
        c_left, c_right = st.columns(2)
        
        # ------------------------------------------------
        # LEFT : VOC 이력
        # ------------------------------------------------
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

# ------------------------------------------------
# RIGHT : 기타 출처 이력
# ------------------------------------------------
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
                            unsafe_allow_html=True,
                        )

                with col2:
                    if LOGIN_TYPE == "admin":
                        if st.button("🗑 삭제", key=f"del_{idx}"):
                            fb_all = fb_all.drop(index=idx)
                            st.session_state["feedback_df"] = fb_all
                            save_feedback(FEEDBACK_PATH, fb_all)
                            st.success("삭제 완료!")
                            st.rerun()
                            st.markdown("</div>", unsafe_allow_html=True)

    st.markdown("### ➕ 빠른 활동등록")

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
        quick_writer = LOGIN_USER
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

# ----------------------------------------------------
# TAB FILTER — 정밀 필터 (안내용)
# ----------------------------------------------------
with tab_filter:
    st.subheader("🎯 해지방어 활동시설 정밀 필터 (VOC유형소 기준)")
    st.info(
        "현재 버전에서는 글로벌 필터 + 다른 탭에서 대부분 분석이 가능하도록 구성되어 있습니다.\n"
        "추후 필요 시 이 탭에 VOC유형소 중심의 추가 정밀 필터를 붙이면 됩니다."
    )

# ----------------------------------------------------
# TAB ALERT — 담당자 알림(베타)
# ----------------------------------------------------
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
