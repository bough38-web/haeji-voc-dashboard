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
    .stApp {
        background: radial-gradient(circle at top left, #fdfbff 0, #f5f5f7 40%, #eef2ff 100%);
        color: #111827;
        font-family: -apple-system, BlinkMacSystemFont, "SF Pro Display", "Segoe UI", sans-serif;
    }

    /* 본문 컨테이너 여백 (상단 잘림 방지 + 모바일 여백 보정) */
    .block-container {
        padding-top: 0.8rem !important;
        padding-bottom: 3rem !important;
        padding-left: 1.0rem !important;
        padding-right: 1.0rem !important;
        max-width: 1400px;
    }

    /* 헤더 영역 배경 */
    [data-testid="stHeader"] {
        background-color: rgba(245,245,247,0.95);
        backdrop-filter: blur(18px);
        border-bottom: 1px solid rgba(148,163,184,0.2);
    }

    /* 사이드바 스타일 (유리 느낌) */
    section[data-testid="stSidebar"] {
        background: rgba(248,250,252,0.85);
        backdrop-filter: blur(18px);
        border-right: 1px solid rgba(148,163,184,0.25);
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
        border-radius: 10px !important;
    }

    /* 라디오 버튼 라벨 간격 */
    div[role="radiogroup"] > label {
        padding-right: 0.75rem;
    }

    /* 섹션 카드 공통 (피드백, 설명 등) */
    .section-card {
        background: rgba(255,255,255,0.9);
        border-radius: 18px;
        padding: 1.0rem 1.2rem;
        border: 1px solid rgba(148,163,184,0.3);
        box-shadow: 0 18px 40px rgba(15,23,42,0.08);
        margin-bottom: 1.2rem;
    }
    .section-title {
        font-size: 1.05rem;
        font-weight: 600;
        margin-bottom: 0.6rem;
        display: flex;
        align-items: center;
        gap: 0.3rem;
    }

    /* 피드백 리스트 카드 */
    .feedback-item {
        background: linear-gradient(135deg, #f9fafb, #eef2ff);
        border-radius: 14px;
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
# 1. SMTP / 파일 경로 설정
# ----------------------------------------------------
if "SMTP_HOST" in st.secrets:
    SMTP_HOST = st.secrets["SMTP_HOST"]
    SMTP_PORT = int(st.secrets["SMTP_PORT"])
    SMTP_USER = st.secrets["SMTP_USER"]
    SMTP_PASSWORD = st.secrets["SMTP_PASSWORD"]
    SENDER_NAME = st.secrets["SENDER_NAME"]
else:
    # 로컬 개발 환경(.env) 대응
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

MERGED_PATH = "merged.xlsx"                 # VOC 통합파일
FEEDBACK_PATH = "feedback.csv"              # 처리내역 CSV 저장 경로
CONTACT_PATH = "영업구역담당자_251204.xlsx"  # 담당자 매핑 파일

# ----------------------------------------------------
# 2. 공통 유틸 함수
# ----------------------------------------------------
def safe_str(x):
    if pd.isna(x):
        return ""
    return str(x).strip()


def detect_column(df: pd.DataFrame, keywords: list[str]) -> str | None:
    """담당자/이메일/휴대폰 컬럼 자동 탐색"""
    for k in keywords:
        if k in df.columns:
            return k
    for col in df.columns:
        s = str(col)
        for k in keywords:
            if k.lower() in s.lower():
                return col
    return None


def drop_all_empty_columns(df: pd.DataFrame) -> pd.DataFrame:
    """모든 값이 공란/NaN인 컬럼은 제거"""
    if df.empty:
        return df
    mask = df.apply(
        lambda col: col.notna() & (col.astype(str).str.strip() != ""),
        axis=0,
    )
    keep_cols = mask.any(axis=0)
    return df.loc[:, keep_cols]


def send_email_with_attachment(
    to_email: str,
    subject: str,
    body: str,
    df_attach: pd.DataFrame,
    filename: str = "attachment.csv",
):
    """CSV 첨부 이메일 발송 (빈 컬럼 제거 후 첨부)"""
    if not SMTP_HOST or not SMTP_USER:
        raise RuntimeError("SMTP 설정이 비어 있습니다. secrets 또는 .env를 확인해주세요.")

    df_clean = drop_all_empty_columns(df_attach)
    csv_bytes = df_clean.to_csv(index=False).encode("utf-8-sig")

    msg = EmailMessage()
    msg["Subject"] = subject
    msg["From"] = f"{SENDER_NAME} <{SMTP_USER}>"
    msg["To"] = to_email
    msg.set_content(body)

    msg.add_attachment(
        csv_bytes,
        maintype="application",
        subtype="octet-stream",
        filename=filename,
    )

    with smtplib.SMTP(SMTP_HOST, SMTP_PORT) as smtp:
        smtp.starttls()
        if SMTP_USER and SMTP_PASSWORD:
            smtp.login(SMTP_USER, SMTP_PASSWORD)
        smtp.send_message(msg)

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
    if not os.path.exists(path):
        st.warning(
            f"❌ 담당자 매핑 파일 '{path}' 을(를) 찾을 수 없습니다. "
            "담당자 알림 탭에서는 직접 이메일 주소를 입력해서 사용해주세요."
        )
        return pd.DataFrame(), {}

    df_c = pd.read_excel(path)

    name_col = detect_column(df_c, ["담당자", "구역담당자", "성명", "이름"])
    email_col = detect_column(df_c, ["이메일", "메일", "email"])
    phone_col = detect_column(df_c, ["휴대폰", "전화", "연락처", "핸드폰"])

    if not (name_col and email_col):
        st.warning(
            f"담당자 매핑 파일('{path}')에서 담당자/이메일 컬럼을 찾지 못했습니다. "
            "컬럼명을 확인해주세요."
        )
        return df_c, {}

    cols = [name_col, email_col]
    if phone_col:
        cols.append(phone_col)
    df_c = df_c[cols].copy()
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
# 4. 지사명 축약 & 영업구역/담당자 통합
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


def make_zone(row):
    if "영업구역번호" in row and pd.notna(row["영업구역번호"]):
        return row["영업구역번호"]
    if "담당상세" in row and pd.notna(row["담당상세"]):
        return row["담당상세"]
    if "영업구역정보" in row and pd.notna(row["영업구역정보"]):
        return row["영업구역정보"]
    return ""


df["영업구역_통합"] = df.apply(make_zone, axis=1)

mgr_priority = ["담당자", "구역담당자", "처리자"]


def pick_manager(row):
    for c in mgr_priority:
        if c in row and pd.notna(row[c]) and str(row[c]).strip() != "":
            return row[c]
    return ""


df["구역담당자_통합"] = df.apply(pick_manager, axis=1)

address_cols = [c for c in df.columns if "주소" in str(c)]

# ----------------------------------------------------
# 5. 출처 분리 + 매칭여부
# ----------------------------------------------------
df_voc = df[df.get("출처") == "해지VOC"].copy()
df_other = df[df.get("출처") != "해지VOC"].copy()

other_sets = {
    src: set(df_other[df_other["출처"] == src]["계약번호_정제"].dropna())
    for src in ["해지시설", "해지요청", "설변", "정지", "해지파이프라인"]
    if "출처" in df_other.columns
}
other_union = set().union(*other_sets.values()) if other_sets else set()

df_voc["매칭여부"] = df_voc["계약번호_정제"].apply(
    lambda x: "매칭(O)" if x in other_union else "비매칭(X)"
)

# ----------------------------------------------------
# 6. 설치주소 / 월정료 정제
# ----------------------------------------------------
def coalesce_cols(row, candidates):
    for c in candidates:
        if c in row.index:
            val = row[c]
            if pd.notna(val) and str(val).strip() not in ["", "None", "nan"]:
                return val
    return np.nan


df_voc["설치주소_표시"] = df_voc.apply(
    lambda r: coalesce_cols(r, ["시설_설치주소", "설치주소"]), axis=1
)

fee_raw_col = None
if "시설_KTT월정료(조정)" in df_voc.columns:
    fee_raw_col = "시설_KTT월정료(조정)"
elif "KTT월정료(조정)" in df_voc.columns:
    fee_raw_col = "KTT월정료(조정)"


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
    if v >= 200000:  # 10배 보정
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
# 7. 리스크 등급/경과일 계산 + 표시 컬럼
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


df_voc["경과일수"], df_voc["리스크등급"] = zip(
    *df_voc.apply(lambda r: compute_risk(r), axis=1)
)

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
# 8. 사이드바 글로벌 필터
# ----------------------------------------------------
st.sidebar.title("🔧 글로벌 필터")

if "접수일시" in df_voc.columns and df_voc["접수일시"].notna().any():
    min_d = df_voc["접수일시"].min().date()
    max_d = df_voc["접수일시"].max().date()
    dr = st.sidebar.date_input(
        "접수일자 범위",
        value=(min_d, max_d),
        min_value=min_d,
        max_value=max_d,
        key="global_date_range",
    )
else:
    dr = None

branches_all = sort_branch(df_voc["관리지사"].dropna().unique())
sel_branches = st.sidebar.multiselect(
    "관리지사(복수 선택)",
    options=branches_all,
    default=branches_all,
    key="global_branches",
)

risk_all = ["HIGH", "MEDIUM", "LOW"]
sel_risk = st.sidebar.multiselect(
    "리스크등급",
    options=risk_all,
    default=risk_all,
    key="global_risk",
)

match_all = ["매칭(O)", "비매칭(X)"]
sel_match = st.sidebar.multiselect(
    "매칭여부",
    options=match_all,
    default=match_all,
    key="global_match",
)

fee_filter_global = st.sidebar.radio(
    "월정료 구간(글로벌)",
    options=["전체", "10만 미만", "10만 이상"],
    index=0,
    key="global_fee_band",
)

st.sidebar.markdown("---")
st.sidebar.caption(
    f"마지막 갱신: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}"
)

# ----------------------------------------------------
# 9. 글로벌 필터 적용
# ----------------------------------------------------
voc_filtered = df_voc.copy()

if dr and isinstance(dr, tuple) and len(dr) == 2:
    start_d, end_d = dr
    if isinstance(start_d, date) and isinstance(end_d, date):
        voc_filtered = voc_filtered[
            (voc_filtered["접수일시"] >= pd.to_datetime(start_d))
            & (
                voc_filtered["접수일시"]
                < pd.to_datetime(end_d) + pd.Timedelta(days=1)
            )
        ]

if sel_branches:
    voc_filtered = voc_filtered[voc_filtered["관리지사"].isin(sel_branches)]

if sel_risk:
    voc_filtered = voc_filtered[voc_filtered["리스크등급"].isin(sel_risk)]

if sel_match:
    voc_filtered = voc_filtered[voc_filtered["매칭여부"].isin(sel_match)]

if fee_filter_global != "전체":
    if fee_filter_global == "10만 이상":
        voc_filtered = voc_filtered[voc_filtered["월정료_수치"] >= 100000]
    elif fee_filter_global == "10만 미만":
        voc_filtered = voc_filtered[
            (voc_filtered["월정료_수치"] < 100000)
            & voc_filtered["월정료_수치"].notna()
        ]

df_unmatched_filtered = voc_filtered[voc_filtered["매칭여부"] == "비매칭(X)"].copy()

# ----------------------------------------------------
# 10. 상단 KPI 카드
# ----------------------------------------------------
st.markdown("## 📊 해지 VOC 종합 대시보드")

total_voc_rows = len(voc_filtered)
unique_contracts = voc_filtered["계약번호_정제"].nunique()
unmatched_contracts = df_unmatched_filtered["계약번호_정제"].nunique()
matched_contracts = (
    voc_filtered[voc_filtered["매칭여부"] == "매칭(O)"]["계약번호_정제"].nunique()
)

c1, c2, c3, c4 = st.columns(4)
c1.metric("VOC 접수건수(행)", f"{total_voc_rows:,}")
c2.metric("VOC 계약 수(유니크)", f"{unique_contracts:,}")
c3.metric("비매칭(X) 계약 수", f"{unmatched_contracts:,}")
c4.metric("매칭(O) 계약 수", f"{matched_contracts:,}")

st.markdown("---")

# ----------------------------------------------------
# 11. 탭 구성
# ----------------------------------------------------
tab_viz, tab_all, tab_unmatched, tab_drill, tab_feedback, tab_alert = st.tabs(
    [
        "📊 지사/담당자 시각화",
        "📘 VOC 전체(계약 기준)",
        "🧯 해지방어 활동시설(비매칭)",
        "🔍 계약별 VOC 드릴다운",
        "📝 처리내역 관리",
        "📨 담당자 알림(이메일)",
    ]
)

# ====================================================
# TAB 1 — 지사 / 담당자 시각화
# ====================================================
with tab_viz:
    st.subheader("📊 지사 / 담당자별 비매칭 리스크 현황")

    if df_unmatched_filtered.empty:
        st.info("현재 조건에서 비매칭(X) 데이터가 없습니다.")
    else:
        left, right = st.columns([1, 2])

        with left:
            st.markdown("#### 🎛️ 시각화 필터")
            b_opts = ["전체"] + sort_branch(
                df_unmatched_filtered["관리지사"].dropna().unique()
            )
            sel_b_viz = st.radio(
                "지사",
                options=b_opts,
                index=0,
                key="viz_branch",
            )

            tmp = df_unmatched_filtered.copy()
            if sel_b_viz != "전체":
                tmp = tmp[tmp["관리지사"] == sel_b_viz]

            mgr_list_viz = (
                tmp["구역담당자_통합"]
                .dropna()
                .astype(str)
                .replace("nan", "")
                .unique()
                .tolist()
            )
            mgr_list_viz = sorted([m for m in mgr_list_viz if m])
            sel_mgr_viz = st.selectbox(
                "담당자(선택 시 레이더 차트 기준)",
                options=["(전체)"] + mgr_list_viz,
                index=0,
                key="viz_mgr",
            )

        with right:
            st.markdown("#### 🧱 지사별 비매칭 계약 수 (유니크 계약)")
            bc = (
                df_unmatched_filtered.groupby("관리지사")["계약번호_정제"]
                .nunique()
                .rename("비매칭계약수")
            )
            bc = bc[bc.index.isin(BRANCH_ORDER)].reindex(BRANCH_ORDER).dropna()

            if HAS_PLOTLY and not bc.empty:
                fig1 = px.bar(
                    bc.reset_index(),
                    x="관리지사",
                    y="비매칭계약수",
                    text="비매칭계약수",
                )
                fig1.update_traces(textposition="outside")
                fig1.update_layout(
                    height=260,
                    margin=dict(l=10, r=10, t=30, b=10),
                    xaxis_title="",
                    yaxis_title="계약 수",
                )
                st.plotly_chart(fig1, use_container_width=True)
            else:
                st.bar_chart(bc, use_container_width=True, height=260)

            c2a, c2b = st.columns(2)

            with c2a:
                st.markdown("#### 👤 담당자별 비매칭 TOP 15 (유니크 계약)")
                mc = (
                    df_unmatched_filtered.groupby("구역담당자_통합")["계약번호_정제"]
                    .nunique()
                    .rename("비매칭계약수")
                    .sort_values(ascending=False)
                )
                mc = mc[mc.index.astype(str).str.strip() != ""].head(15)

                if HAS_PLOTLY and not mc.empty:
                    fig2 = px.bar(
                        mc.reset_index(),
                        x="구역담당자_통합",
                        y="비매칭계약수",
                        text="비매칭계약수",
                    )
                    fig2.update_traces(textposition="outside")
                    fig2.update_layout(
                        height=300,
                        margin=dict(l=10, r=10, t=30, b=60),
                        xaxis_title="담당자",
                        yaxis_title="계약 수",
                        xaxis_tickangle=-45,
                    )
                    st.plotly_chart(fig2, use_container_width=True)
                else:
                    st.bar_chart(mc, use_container_width=True, height=300)

            with c2b:
                st.markdown("#### 🔥 리스크 등급 분포 (비매칭, 계약 단위)")
                rc = (
                    df_unmatched_filtered["리스크등급"]
                    .value_counts()
                    .reindex(["HIGH", "MEDIUM", "LOW"])
                    .fillna(0)
                )
                if HAS_PLOTLY and not rc.empty:
                    rc_df = rc.reset_index()
                    rc_df.columns = ["리스크등급", "건수"]
                    rc_df["건수"] = rc_df["건수"].astype(int)

                    fig3 = px.bar(
                        rc_df,
                        x="리스크등급",
                        y="건수",
                        text="건수",
                    )
                    fig3.update_traces(textposition="outside")
                    fig3.update_layout(
                        height=300,
                        margin=dict(l=10, r=10, t=30, b=10),
                        xaxis_title="리스크등급",
                        yaxis_title="계약 수",
                    )
                    st.plotly_chart(fig3, use_container_width=True)
                else:
                    st.bar_chart(rc, use_container_width=True, height=300)

            st.markdown("---")

            if "접수일시" in df_unmatched_filtered.columns:
                trend = (
                    df_unmatched_filtered.assign(
                        접수일=df_unmatched_filtered["접수일시"].dt.date
                    )
                    .groupby("접수일")["계약번호_정제"]
                    .nunique()
                    .rename("비매칭계약수")
                    .sort_index()
                )
                st.markdown("#### 📈 일별 비매칭 계약 추이 (유니크 계약)")
                if HAS_PLOTLY and not trend.empty:
                    fig4 = px.line(
                        trend.reset_index(),
                        x="접수일",
                        y="비매칭계약수",
                    )
                    fig4.update_layout(
                        height=260,
                        margin=dict(l=10, r=10, t=30, b=10),
                        xaxis_title="접수일",
                        yaxis_title="비매칭 계약 수",
                    )
                    st.plotly_chart(fig4, use_container_width=True)
                else:
                    st.line_chart(trend, use_container_width=True, height=260)

            # 선택한 담당자 레이더 차트 (HIGH/MEDIUM/LOW 비율)
            if sel_mgr_viz != "(전체)" and HAS_PLOTLY:
                mgr_data = df_unmatched_filtered[
                    df_unmatched_filtered["구역담당자_통합"].astype(str) == sel_mgr_viz
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
                        margin=dict(l=10, r=10, t=40, b=10),
                        title=f"🌐 {sel_mgr_viz} 담당자의 리스크 프로파일",
                    )
                    st.plotly_chart(fig_radar, use_container_width=True)

# ====================================================
# TAB 2 — VOC 전체 (계약번호 기준 요약)
# ====================================================
with tab_all:
    st.subheader("📘 VOC 전체 (계약번호 기준 요약)")

    # 최신 VOC 1건만 남기고 + 계약별 접수 건수
    temp = voc_filtered.sort_values("접수일시", ascending=False)
    grp_all = temp.groupby("계약번호_정제")
    idx_latest = grp_all["접수일시"].idxmax()
    df_summary = temp.loc[idx_latest].copy()
    df_summary["접수건수"] = grp_all.size().reindex(df_summary["계약번호_정제"]).values

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
# TAB 3 — 해지방어 활동시설(비매칭)
# ====================================================
with tab_unmatched:
    st.subheader("🧯 해지방어 활동시설 (비매칭, 계약번호 기준)")
    st.caption("비매칭(X) = 해지VOC 접수 후 시스템상 활동내역이 확인되지 않은 시설")

    if df_unmatched_filtered.empty:
        st.info("현재 글로벌 필터 조건에서 비매칭(X) 계약이 없습니다.")
    else:
        # 지사/담당자 필터
        u_col1, u_col2 = st.columns([2, 3])
        branches_u = ["전체"] + sort_branch(
            df_unmatched_filtered["관리지사"].dropna().unique()
        )
        selected_branch_u = u_col1.radio(
            "지사 선택",
            options=branches_u,
            horizontal=True,
            key="tab2_branch_radio",
        )

        temp_u_for_mgr = df_unmatched_filtered.copy()
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

        temp_u = df_unmatched_filtered.copy()
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

            st.dataframe(
                style_risk(df_u_summary[summary_cols_u]),
                use_container_width=True,
                height=420,
            )

            # 상세 이력 보기용 계약 선택
            st.markdown("### 📂 선택한 계약번호 상세 VOC 이력")
            sel_u_contract = st.selectbox(
                "상세 VOC 이력을 볼 계약 선택",
                options=["(선택)"] + df_u_summary["계약번호_정제"].astype(str).tolist(),
                key="tab2_select_contract",
            )

            if sel_u_contract != "(선택)":
                voc_detail = temp_u[
                    temp_u["계약번호_정제"].astype(str) == sel_u_contract
                ].copy()
                voc_detail = voc_detail.sort_values("접수일시", ascending=False)

                st.markdown(
                    f"#### 🔍 `{sel_u_contract}` VOC 상세 이력 ({len(voc_detail)}건)"
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
# TAB 4 — 계약별 VOC 드릴다운
# ====================================================
with tab_drill:
    st.subheader("🔍 계약별 VOC / 기타출처 드릴다운")

    base_all = voc_filtered.copy()

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

    d1, d2 = st.columns([2, 3])
    branches_d = ["전체"] + sort_branch(drill_base["관리지사"].dropna().unique())
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

            st.markdown("---")

            dcol1, dcol2 = st.columns(2)

            if not voc_hist.empty:
                dcol1.download_button(
                    "📥 선택 계약 VOC 이력 다운로드 (CSV)",
                    voc_hist.to_csv(index=False).encode("utf-8-sig"),
                    file_name=f"VOC이력_{sel_cn}.csv",
                    mime="text/csv",
                )

            export_frames = []
            if not voc_hist.empty:
                v_exp = voc_hist.copy()
                v_exp.insert(0, "구분", "VOC")
                export_frames.append(v_exp)

            if not other_hist.empty:
                o_exp = other_hist.copy()
                o_exp.insert(0, "구분", "기타출처")
                export_frames.append(o_exp)

            fb_all_for_export = st.session_state["feedback_df"]
            fb_sel_export = fb_all_for_export[
                fb_all_for_export["계약번호_정제"].astype(str) == str(sel_cn)
            ].copy()
            if not fb_sel_export.empty:
                f_exp = fb_sel_export.copy()
                f_exp.insert(0, "구분", "피드백")
                export_frames.append(f_exp)

            if export_frames:
                merged_export = pd.concat(export_frames, ignore_index=True)
                dcol2.download_button(
                    "📥 선택 계약 통합 이력 다운로드 (CSV)",
                    merged_export.to_csv(index=False).encode("utf-8-sig"),
                    file_name=f"통합이력_{sel_cn}.csv",
                    mime="text/csv",
                )

# ====================================================
# TAB 5 — 처리내역(피드백) 관리
# ====================================================
with tab_feedback:
    st.subheader("📝 해지상담대상 활동등록 (고객대응 / 현장 처리내역)")

    fb_all = st.session_state["feedback_df"]

    # 1) 계약 선택
    st.markdown("### 1) 계약 선택")
    all_contracts = sorted(
        set(
            fb_all["계약번호_정제"].dropna().astype(str).tolist()
            + voc_filtered["계약번호_정제"].dropna().astype(str).tolist()
        )
    )
    fb_cn = st.selectbox("피드백을 관리할 계약 선택", ["(선택)"] + all_contracts)

    if fb_cn != "(선택)":
        # 2) 기존 처리내역 표시
        fb_sel = fb_all[fb_all["계약번호_정제"].astype(str) == fb_cn].copy()
        fb_sel = fb_sel.sort_values("등록일자", ascending=False)

        st.markdown(f"### 2) 기존 처리내역 ({len(fb_sel)}건)")

        # 관리자 비밀번호 (사용자 지정: Q1=3, Q2=1, Q3=2 → "312" 예시)
        ADMIN_CODE = "312"
        admin_pw = st.text_input("관리자 비밀번호 (삭제 시 필요)", type="password")

        for idx, row in fb_sel.iterrows():
            with st.container():
                st.markdown(
                    "<div class='feedback-item'>",
                    unsafe_allow_html=True,
                )
                col1, col2 = st.columns([6, 1])

                with col1:
                    st.write(f"**내용:** {row['고객대응내용']}")
                    st.markdown(
                        f"<div class='feedback-meta'>등록자: {row['등록자']} "
                        f"| 등록일: {row['등록일자']}</div>",
                        unsafe_allow_html=True,
                    )
                    if safe_str(row.get("비고")):
                        st.markdown(
                            f"<div class='feedback-note'>비고: {row['비고']}</div>",
                            unsafe_allow_html=True,
                        )

                with col2:
                    if admin_pw == ADMIN_CODE:
                        if st.button("🗑 삭제", key=f"fb_del_{idx}"):
                            fb_all = fb_all.drop(index=idx)
                            st.session_state["feedback_df"] = fb_all
                            save_feedback(FEEDBACK_PATH, fb_all)
                            st.success("삭제 완료!")
                            st.rerun()
                st.markdown("</div>", unsafe_allow_html=True)

        # 3) 새 처리내역 입력
        st.markdown("### 3) 새 처리내역 등록")

        new_content = st.text_area("고객대응 / 현장 처리내용 입력")
        new_writer = st.text_input("등록자")
        new_note = st.text_input("비고")

        if st.button("등록하기", type="primary"):
            if not new_content or not new_writer:
                st.warning("내용과 등록자를 모두 입력해주세요.")
            else:
                new_row = {
                    "계약번호_정제": fb_cn,
                    "고객대응내용": new_content,
                    "등록자": new_writer,
                    "등록일자": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                    "비고": new_note,
                }
                fb_all = pd.concat(
                    [fb_all, pd.DataFrame([new_row])], ignore_index=True
                )
                st.session_state["feedback_df"] = fb_all
                save_feedback(FEEDBACK_PATH, fb_all)
                st.success("등록 완료!")
                st.rerun()

# ====================================================
# TAB 6 — 담당자 알림(이메일 발송)
# ====================================================
with tab_alert:
    st.subheader("📨 담당자 알림 발송 (비매칭 계약 이메일 공유)")

    st.markdown(
        """
        글로벌 필터(좌측)를 통해 기간, 지사, 리스크등급, 월정료 구간을 먼저 조정한 뒤,<br>
        현재 조건에서의 <b>비매칭(X) 계약</b>을 담당자별로 메일로 공유할 수 있습니다.
        """,
        unsafe_allow_html=True,
    )

    if contact_df.empty:
        st.markdown(
            """
            <div style="
                background:#fff3cd;
                border-left:6px solid #ffca2c;
                padding:12px;
                border-radius:6px;
                margin-top:12px;
                margin-bottom:12px;
                font-size:0.95rem;
                line-height:1.5;
            ">
            <b>⚠ 담당자 매핑 파일을 찾을 수 없습니다.</b><br>
            '영업구역담당자_251204.xlsx' 파일이 저장소 루트(/) 위치에 있는지 확인하세요.<br>
            그래도 이메일 주소를 직접 입력하면 발송 기능은 사용할 수 있습니다.
            </div>
            """,
            unsafe_allow_html=True,
        )

    if df_unmatched_filtered.empty:
        st.info("현재 글로벌 필터 조건에서 비매칭(X) 계약이 없어, 발송 대상이 없습니다.")
    else:
        # 담당자별 요약 생성 (계약단위)
        tmp = df_unmatched_filtered.sort_values("접수일시", ascending=False)
        grp = tmp.groupby(["구역담당자_통합", "계약번호_정제"])
        idx_latest_mgr = grp["접수일시"].idxmax()
        latest_per_contract = tmp.loc[idx_latest_mgr].copy()

        # 담당자별 집계
        summary_list = []
        for mgr, g in latest_per_contract.groupby("구역담당자_통합"):
            mgr_name = safe_str(mgr)
            if not mgr_name:
                continue
            total = g["계약번호_정제"].nunique()
            high = (g["리스크등급"] == "HIGH").sum()
            medium = (g["리스크등급"] == "MEDIUM").sum()
            low = (g["리스크등급"] == "LOW").sum()
            email = manager_contacts.get(mgr_name, {}).get("email", "")
            summary_list.append(
                {
                    "담당자": mgr_name,
                    "이메일": email,
                    "비매칭 계약수": total,
                    "HIGH": high,
                    "MEDIUM": medium,
                    "LOW": low,
                }
            )

        if not summary_list:
            st.info("담당자명이 비어있어 집계할 데이터가 없습니다.")
        else:
            summary_df = pd.DataFrame(summary_list).sort_values(
                ["비매칭 계약수", "HIGH"], ascending=[False, False]
            )
            st.markdown("### 👥 담당자별 비매칭 계약 요약")
            st.dataframe(summary_df, use_container_width=True, height=280)

            st.markdown("---")
            st.markdown("### ✉ 개별 담당자에게 이메일 발송")

            sel_mgr = st.selectbox(
                "담당자 선택",
                options=["(선택)"] + summary_df["담당자"].tolist(),
                key="alert_mgr",
            )

            if sel_mgr != "(선택)":
                # 해당 담당자 데이터 (계약번호 1행 + VOC 건수 포함)
                mgr_rows_all = df_unmatched_filtered[
                    df_unmatched_filtered["구역담당자_통합"].astype(str) == sel_mgr
                ].copy()

                if mgr_rows_all.empty:
                    st.info("선택한 담당자에 해당하는 비매칭 데이터가 없습니다.")
                else:
                    mgr_rows_all = mgr_rows_all.sort_values(
                        "접수일시", ascending=False
                    )
                    grp_mgr = mgr_rows_all.groupby("계약번호_정제")
                    idx_latest_m = grp_mgr["접수일시"].idxmax()
                    mgr_latest = mgr_rows_all.loc[idx_latest_m].copy()
                    mgr_latest["VOC_접수건수"] = grp_mgr.size().reindex(
                        mgr_latest["계약번호_정제"]
                    ).values

                    # 첨부용 컬럼 선택
                    attach_cols = [
                        "계약번호_정제",
                        "상호",
                        "관리지사",
                        "구역담당자_통합",
                        "리스크등급",
                        "경과일수",
                        "VOC_접수건수",
                        "설치주소_표시",
                        "VOC유형소",
                        "해지상세",
                        "월정료구간",
                        "접수일시",
                    ]
                    attach_cols = [c for c in attach_cols if c in mgr_latest.columns]
                    df_attach = mgr_latest[attach_cols].copy()
                    df_attach = drop_all_empty_columns(df_attach)

                    st.markdown(
                        f"🔍 **{sel_mgr}** 담당자 비매칭 계약: "
                        f"총 {len(df_attach):,}건 (계약 기준)"
                    )
                    st.dataframe(
                        df_attach,
                        use_container_width=True,
                        height=260,
                    )

                    # 이메일 주소
                    default_email = manager_contacts.get(sel_mgr, {}).get("email", "")
                    st.write(
                        f"📮 매핑된 이메일: **{default_email or '(없음 — 직접 입력 필요)'}**"
                    )
                    input_email = st.text_input(
                        "이메일 주소(변경 또는 직접 입력)", value=default_email
                    )

                    # 제목/내용
                    subject = f"[해지VOC] {sel_mgr} 담당자 비매칭 계약 안내 ({len(df_attach)}건)"
                    body = (
                        f"{sel_mgr} 담당자님,\n\n"
                        f"현재 기준 비매칭 해지 VOC 계약이 총 {len(df_attach)}건 확인되어 공유드립니다.\n"
                        f"계약별 최신 VOC 기준 1행 + VOC 접수건수 정보를 첨부 파일(CSV)로 전달드립니다.\n\n"
                        "현장 해지방어 활동 여부 및 향후 관리 계획 수립에 참고 부탁드립니다.\n\n"
                        "- 해지VOC 관리자 드림 -"
                    )

                    if st.button("📤 이메일 발송하기", key="btn_send_email"):
                        if not input_email:
                            st.error("이메일 주소를 입력해주세요.")
                        else:
                            try:
                                send_email_with_attachment(
                                    to_email=input_email,
                                    subject=subject,
                                    body=body,
                                    df_attach=df_attach,
                                    filename=f"비매칭계약_{sel_mgr}.csv",
                                )
                                st.success(f"✅ 이메일 발송 완료 → {input_email}")
                            except Exception as e:
                                st.error(f"❌ 이메일 전송 실패: {e}")
