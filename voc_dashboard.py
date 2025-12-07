# ===========================================================
# PART 1 — 기본 설정 / UI 스타일 / 데이터 로딩 및 전처리
# Google Material Glass 스타일 + 고가독성 설계
# ===========================================================

import os
from datetime import datetime, date
import smtplib
from email.message import EmailMessage

import numpy as np
import pandas as pd
import streamlit as st

# Plotly (없으면 자동 False)
try:
    import plotly.express as px
    HAS_PLOTLY = True
except:
    HAS_PLOTLY = False


# -----------------------------------------------------------
# 0. 페이지 설정 + Google Material Glass UI
# -----------------------------------------------------------

st.set_page_config(
    page_title="해지 VOC 종합 대시보드",
    layout="wide",
)

# Google Material Glass 스타일
st.markdown(
    """
    <style>
    /* 전체 배경 */
    html, body {
        background: #f3f4f6 !important;
    }
    .stApp {
        background: #f3f4f6 !important;
        font-family: "Inter", -apple-system, BlinkMacSystemFont, sans-serif;
        color: #111827;
    }
    
    /* Glass Card */
    .glass-card {
        background: rgba(255,255,255,0.65);
        backdrop-filter: blur(9px);
        -webkit-backdrop-filter: blur(9px);
        border-radius: 18px;
        padding: 1.1rem 1.3rem;
        border: 1px solid rgba(255,255,255,0.35);
        box-shadow: 0 6px 18px rgba(0,0,0,0.06);
        margin-bottom: 1rem;
    }

    /* 제목 */
    h1, h2, h3, h4 {
        font-weight: 650 !important;
        letter-spacing: -0.01em;
    }

    /* 표 라인 */
    .dataframe tbody tr:nth-child(odd) { background: #fafafa; }
    .dataframe tbody tr:nth-child(even) { background: #f0f6ff; }

    /* 좌측 사이드바 */
    section[data-testid="stSidebar"] {
        background: rgba(255,255,255,0.6) !important;
        backdrop-filter: blur(9px);
        border-right: 1px solid #e5e7eb;
    }

    /* Plotly 투명 배경 */
    .js-plotly-plot .plotly {
        background-color: transparent !important;
    }
    </style>
    """,
    unsafe_allow_html=True,
)

# -----------------------------------------------------------
# 1. 환경변수 (SMTP 미리 설정됨)
# -----------------------------------------------------------

if "SMTP_HOST" in st.secrets:
    SMTP_HOST = st.secrets["SMTP_HOST"]
    SMTP_PORT = int(st.secrets["SMTP_PORT"])
    SMTP_USER = st.secrets["SMTP_USER"]
    SMTP_PASSWORD = st.secrets["SMTP_PASSWORD"]
    SENDER_NAME = st.secrets["SENDER_NAME"]
else:
    # 로컬 개발용 dotenv
    try:
        from dotenv import load_dotenv
        load_dotenv()
    except:
        pass
    SMTP_HOST = os.getenv("SMTP_HOST", "")
    SMTP_PORT = int(os.getenv("SMTP_PORT", "587"))
    SMTP_USER = os.getenv("SMTP_USER", "")
    SMTP_PASSWORD = os.getenv("SMTP_PASSWORD", "")
    SENDER_NAME = os.getenv("SENDER_NAME", "해지VOC 관리자")

# -----------------------------------------------------------
# 2. 파일 경로
# -----------------------------------------------------------

MERGED_PATH = "merged.xlsx"                     # 통합 VOC
FEEDBACK_PATH = "feedback.csv"                 # 활동등록 CSV
CONTACT_PATH = "영업구역담당자_251204.xlsx"    # 담당자 매핑파일


# -----------------------------------------------------------
# 3. 유틸 함수
# -----------------------------------------------------------

def safe_str(x):
    if pd.isna(x):
        return ""
    return str(x).strip()


def detect_column(df: pd.DataFrame, keywords: list[str]):
    """담당자/메일 컬럼 자동 탐색"""
    # 완전 일치 우선
    for k in keywords:
        if k in df.columns:
            return k
    # 부분 일치
    for col in df.columns:
        for k in keywords:
            if k.lower() in str(col).lower():
                return col
    return None


# -----------------------------------------------------------
# 4. 데이터 로딩 함수
# -----------------------------------------------------------

@st.cache_data
def load_voc_data(path):
    if not os.path.exists(path):
        st.error("❌ merged.xlsx 파일이 없습니다.")
        return pd.DataFrame()

    df = pd.read_excel(path)

    # 계약번호 정제
    if "계약번호" in df.columns:
        df["계약번호"] = df["계약번호"].astype(str).str.replace(",", "").str.strip()
        df["계약번호_정제"] = df["계약번호"].str.replace(r"[^0-9A-Za-z]", "", regex=True)
    else:
        df["계약번호_정제"] = ""

    # 출처 정리
    if "출처" in df.columns:
        df["출처"] = df["출처"].replace({"고객리스트": "해지시설"})

    # 접수일 변환
    if "접수일시" in df.columns:
        df["접수일시"] = pd.to_datetime(df["접수일시"], errors="coerce")

    return df


@st.cache_data
def load_feedback(path):
    if os.path.exists(path):
        try:
            return pd.read_csv(path, encoding="utf-8-sig")
        except:
            return pd.read_csv(path)
    return pd.DataFrame(columns=["계약번호_정제", "고객대응내용", "등록자", "등록일자", "비고"])


def save_feedback(path, df):
    df.to_csv(path, index=False, encoding="utf-8-sig")


@st.cache_data
def load_contact_map(path):
    if not os.path.exists(path):
        st.warning("⚠ 담당자 매핑파일이 없습니다.")
        return pd.DataFrame(), {}

    df = pd.read_excel(path)

    name_col = detect_column(df, ["담당자", "구역담당자", "성명"])
    email_col = detect_column(df, ["이메일", "email"])

    if not name_col or not email_col:
        st.warning("⚠ 담당자/이메일 컬럼을 찾지 못했습니다.")
        return df, {}

    df = df[[name_col, email_col]].copy()
    df.columns = ["담당자", "이메일"]

    mapping = {}
    for _, row in df.iterrows():
        name = safe_str(row["담당자"])
        if name:
            mapping[name] = {"email": safe_str(row["이메일"])}

    return df, mapping


# -----------------------------------------------------------
# 5. 실제 데이터 로딩
# -----------------------------------------------------------

df = load_voc_data(MERGED_PATH)
if df.empty:
    st.stop()

if "feedback_df" not in st.session_state:
    st.session_state["feedback_df"] = load_feedback(FEEDBACK_PATH)

contact_df, manager_contacts = load_contact_map(CONTACT_PATH)

# ===========================================================
# PART 2 — VOC 전처리 / 주소·월정료 통합 / 매칭 판정 / 리스크 등급 계산
# ===========================================================

# -----------------------------------------------------------
# 6. 지사명 축약 + 정렬 우선순위
# -----------------------------------------------------------

if "관리지사" in df.columns:
    df["관리지사"] = df["관리지사"].replace({
        "중앙지사": "중앙",
        "강북지사": "강북",
        "서대문지사": "서대문",
        "고양지사": "고양",
        "의정부지사": "의정부",
        "남양주지사": "남양주",
        "강릉지사": "강릉",
        "원주지사": "원주",
    })
else:
    df["관리지사"] = ""

BRANCH_ORDER = ["중앙","강북","서대문","고양","의정부","남양주","강릉","원주"]


def sort_branch(list_values):
    ordered = [b for b in BRANCH_ORDER if b in list_values]
    return ordered


# -----------------------------------------------------------
# 7. 영업구역/담당자 통합 컬럼
# -----------------------------------------------------------

def pick_zone(row):
    for c in ["영업구역번호", "담당상세", "영업구역정보"]:
        if c in row and pd.notna(row[c]):
            return row[c]
    return ""

df["영업구역_통합"] = df.apply(pick_zone, axis=1)


def pick_manager(row):
    for c in ["구역담당자", "담당자", "처리자"]:
        if c in row and pd.notna(row[c]) and str(row[c]).strip() != "":
            return row[c]
    return ""

df["구역담당자_통합"] = df.apply(pick_manager, axis=1)


# -----------------------------------------------------------
# 8. 출처 분리 + 매칭여부 판정
# -----------------------------------------------------------

df_voc = df[df.get("출처") == "해지VOC"].copy()
df_other = df[df.get("출처") != "해지VOC"].copy()

other_contracts = set(df_other["계약번호_정제"].dropna().tolist())

df_voc["매칭여부"] = df_voc["계약번호_정제"].apply(
    lambda x: "매칭(O)" if x in other_contracts else "비매칭(X)"
)


# -----------------------------------------------------------
# 9. 설치주소 통합 (시설_설치주소 → 설치주소 우선, None 제거)
# -----------------------------------------------------------

def merge_address(row):
    for c in ["시설_설치주소", "설치주소"]:
        if c in row and pd.notna(row[c]) and str(row[c]).strip() not in ["","None","nan"]:
            return row[c]
    return np.nan

df_voc["설치주소_표시"] = df_voc.apply(merge_address, axis=1)

# 주소 컬럼 자동검색용
address_cols = [c for c in df.columns if "주소" in str(c)]


# -----------------------------------------------------------
# 10. 월정료 정제 (문자 제거 → 실수 변환 → 10배 오류 보정 → 천단위 콤마표기)
# -----------------------------------------------------------

# 원본 컬럼 탐색
fee_col = None
for cand in ["시설_KTT월정료(조정)", "KTT월정료(조정)", "월정료"]:
    if cand in df_voc.columns:
        fee_col = cand
        break

def parse_fee(v):
    if pd.isna(v): 
        return np.nan
    s = str(v).replace(",", "").strip()
    s = "".join(ch for ch in s if ch.isdigit())
    if s == "":
        return np.nan
    f = float(s)

    # 55,000 → 55000 정상
    # 550,000 또는 550,0000 → 오류 가능 → 10배 보정
    if f >= 200000:  # 20만↑이면 10배로 간주
        f = f / 10
    return f

if fee_col:
    df_voc["월정료_수치"] = df_voc[fee_col].apply(parse_fee)
else:
    df_voc["월정료_수치"] = np.nan
    fee_col = None


def fee_band(v):
    if pd.isna(v):
        return "미기재"
    return "10만 이상" if v >= 100000 else "10만 미만"

df_voc["월정료구간"] = df_voc["월정료_수치"].apply(fee_band)


# 표시용(천단위 콤마)
if fee_col:
    df_voc["월정료_표시"] = df_voc["월정료_수치"].apply(
        lambda x: "" if pd.isna(x) else f"{int(x):,}"
    )


# -----------------------------------------------------------
# 11. 리스크 등급 계산 (접수일로부터 경과일수 기준)
# -----------------------------------------------------------

today = date.today()

def risk_score(row):
    dt = row.get("접수일시")
    if pd.isna(dt):
        return np.nan, "LOW"

    if isinstance(dt, datetime):
        dt = dt.date()

    days = (today - dt).days

    # 규칙 HIGH: 3일 이하 / MEDIUM: 10일 이하 / 그 외 LOW
    if days <= 3:
        level = "HIGH"
    elif days <= 10:
        level = "MEDIUM"
    else:
        level = "LOW"

    return days, level


df_voc["경과일수"], df_voc["리스크등급"] = zip(*df_voc.apply(risk_score, axis=1))


# -----------------------------------------------------------
# 12. 비매칭 계약 DataFrame
# -----------------------------------------------------------

df_unmatched = df_voc[df_voc["매칭여부"] == "비매칭(X)"].copy()

# ============================================================
# PART 3 — 전체 UI / 글로벌 필터 / KPI 카드 / 고급 시각화 개선
# ============================================================

# ------------------------------------------------------------
# 13. Streamlit 기본 페이지 설정 + CSS (라이트톤 유지)
# ------------------------------------------------------------
st.set_page_config(
    page_title="해지 VOC 종합 대시보드",
    layout="wide",
)

# UI 개선용 CSS
st.markdown("""
<style>
html, body, .stApp { background-color:#f5f5f7 !important; }
.block-container { padding-top:0.6rem !important; }

/* 카드형 KPI UI */
.kpi-card {
    background:#ffffff;
    padding:1rem 1.2rem;
    border-radius:14px;
    border:1px solid #e5e7eb;
    box-shadow:0 4px 8px rgba(0,0,0,0.03);
}

/* 탭 제목 간격 */
h2,h3 { margin-top:0.4rem; }

/* 데이터프레임 stripe */
.dataframe tbody tr:nth-child(odd) { background:#fafafa; }
.dataframe tbody tr:nth-child(even) { background:#eef2ff; }

/* Plotly 배경 투명 */
.js-plotly-plot .plotly { background-color: transparent !important; }

/* 4개씩 행 형태의 지사 배치용 wrap UI */
.branch-grid {
    display: grid;
    grid-template-columns: repeat(4, minmax(0,1fr));
    gap: 12px;
}
@media (max-width: 1200px){
    .branch-grid { grid-template-columns: repeat(2, minmax(0,1fr)); }
}
@media (max-width: 700px){
    .branch-grid { grid-template-columns: repeat(1, minmax(0,1fr)); }
}
.branch-item {
    background:#ffffff;
    padding:0.9rem;
    border-radius:10px;
    border:1px solid #e5e7eb;
    text-align:center;
    font-weight:600;
}
</style>
""", unsafe_allow_html=True)

# ------------------------------------------------------------
# 14. 글로벌 필터 (모든 탭에 공통 적용)
# ------------------------------------------------------------
st.sidebar.title("🔧 글로벌 필터")

# 날짜 필터
if "접수일시" in df_voc.columns and df_voc["접수일시"].notna().any():
    mind = df_voc["접수일시"].min().date()
    maxd = df_voc["접수일시"].max().date()
    dr = st.sidebar.date_input(
        "접수일자 범위",
        value=(mind, maxd),
        min_value=mind,
        max_value=maxd
    )
else:
    dr = None

# 지사 선택
branch_all = sort_branch(df_voc["관리지사"].dropna().unique())
sel_branches = st.sidebar.multiselect(
    "관리지사 선택",
    options=branch_all,
    default=branch_all
)

# 리스크 필터
risk_opts = ["HIGH", "MEDIUM", "LOW"]
sel_risk = st.sidebar.multiselect(
    "리스크등급",
    options=risk_opts,
    default=risk_opts
)

# 매칭여부
sel_match = st.sidebar.multiselect(
    "매칭여부",
    options=["매칭(O)", "비매칭(X)"],
    default=["매칭(O)", "비매칭(X)"]
)

# 월정료
fee_global = st.sidebar.radio(
    "월정료 구간",
    ["전체", "10만 미만", "10만 이상"],
    index=0
)

st.sidebar.markdown("---")
st.sidebar.caption("필터 변경 시 모든 탭이 즉시 반영됩니다.")

# ------------------------------------------------------------
# 15. 글로벌 필터 적용된 VOC DataFrame 생성
# ------------------------------------------------------------
voc_filtered = df_voc.copy()

# 날짜 적용
if dr:
    start, end = dr
    voc_filtered = voc_filtered[
        (voc_filtered["접수일시"] >= pd.to_datetime(start)) &
        (voc_filtered["접수일시"] < pd.to_datetime(end) + pd.Timedelta(days=1))
    ]

# 지사 필터
voc_filtered = voc_filtered[voc_filtered["관리지사"].isin(sel_branches)]

# 리스크
voc_filtered = voc_filtered[voc_filtered["리스크등급"].isin(sel_risk)]

# 매칭여부
voc_filtered = voc_filtered[voc_filtered["매칭여부"].isin(sel_match)]

# 월정료
if fee_global == "10만 이상":
    voc_filtered = voc_filtered[voc_filtered["월정료_수치"] >= 100000]
elif fee_global == "10만 미만":
    voc_filtered = voc_filtered[
        voc_filtered["월정료_수치"].notna() &
        (voc_filtered["월정료_수치"] < 100000)
    ]

# 비매칭만 따로
unmatched_filtered = voc_filtered[voc_filtered["매칭여부"] == "비매칭(X)"]

# ------------------------------------------------------------
# 16. KPI 카드 (기업형 UI)
# ------------------------------------------------------------

st.markdown("## 📊 해지 VOC 종합 대시보드")

c1, c2, c3, c4 = st.columns(4)

with c1:
    st.markdown("<div class='kpi-card'>", unsafe_allow_html=True)
    st.metric("총 VOC 행", f"{len(voc_filtered):,}")
    st.markdown("</div>", unsafe_allow_html=True)

with c2:
    st.markdown("<div class='kpi-card'>", unsafe_allow_html=True)
    st.metric("계약 수(유니크)", f"{voc_filtered['계약번호_정제'].nunique():,}")
    st.markdown("</div>", unsafe_allow_html=True)

with c3:
    st.markdown("<div class='kpi-card'>", unsafe_allow_html=True)
    st.metric(
        "비매칭 계약 수",
        f"{unmatched_filtered['계약번호_정제'].nunique():,}"
    )
    st.markdown("</div>", unsafe_allow_html=True)

with c4:
    st.markdown("<div class='kpi-card'>", unsafe_allow_html=True)
    st.metric(
        "매칭 계약 수",
        f"{voc_filtered[voc_filtered['매칭여부']=='매칭(O)']['계약번호_정제'].nunique():,}"
    )
    st.markdown("</div>", unsafe_allow_html=True)

st.markdown("---")

# ------------------------------------------------------------
# 17. 탭 구성
# ------------------------------------------------------------
tab_viz, tab_all, tab_unmatched, tab_drill, tab_alert = st.tabs([
    "📊 지사/담당자 시각화",
    "📘 VOC 전체",
    "🧯 비매칭 시설",
    "🔍 계약별 상세",
    "📨 담당자 알림"
])

# ============================================================
# TAB 1 — 지사/담당자 시각화
# ============================================================
with tab_viz:
    st.subheader("📊 지사 / 담당자 비매칭 리스크 현황")

    if unmatched_filtered.empty:
        st.info("현재 조건에서 비매칭 데이터 없음")
    else:

        # ---------------------------
        # 1) 지사를 4개씩 자동 레이아웃 배치
        # ---------------------------
        st.markdown("### 🏢 지사 목록")

        branch_items = unmatched_filtered["관리지사"].dropna().unique().tolist()
        branch_items = sort_branch(branch_items)

        # 4개씩 카드형 표시
        branch_html = "<div class='branch-grid'>"
        for br in branch_items:
            branch_html += f"""
            <div class='branch-item'>
                {br}<br>
                <span style='font-size:0.85rem;color:#555;'>비매칭: {unmatched_filtered[unmatched_filtered['관리지사']==br]['계약번호_정제'].nunique()}</span>
            </div>"""
        branch_html += "</div>"

        st.markdown(branch_html, unsafe_allow_html=True)

        st.markdown("---")

        # ------------------------------
        # 2) 지사 선택 시 담당자 자동 드롭다운
        # ------------------------------
        st.markdown("### 👥 지사 선택 → 담당자 선택")

        col1, col2 = st.columns([1, 1])

        sel_branch_viz = col1.selectbox(
            "지사 선택",
            options=["전체"] + branch_items
        )

        temp_mgr = unmatched_filtered.copy()
        if sel_branch_viz != "전체":
            temp_mgr = temp_mgr[temp_mgr["관리지사"] == sel_branch_viz]

        mgr_list = sorted(
            temp_mgr["구역담당자_통합"].dropna().astype(str).unique().tolist()
        )

        sel_mgr_viz = col2.selectbox(
            "담당자 선택",
            options=["전체"] + mgr_list
        )

        # ------------------------------
        # 3) 그래프 사이즈 확대 + plotly 적용
        # ------------------------------
        st.markdown("### 📈 지사별 비매칭 계약 수")

        bc = (
            unmatched_filtered.groupby("관리지사")["계약번호_정제"]
            .nunique()
            .rename("건수")
            .reindex(branch_items)
        )

        fig1 = px.bar(
            bc.reset_index(),
            x="관리지사",
            y="건수",
            text="건수",
        )
        fig1.update_traces(textposition="outside")
        fig1.update_layout(
            height=380,   # 🔥 그래프 확대
            margin=dict(l=20, r=20, t=40, b=30),
            xaxis_title="지사",
            yaxis_title="비매칭 계약 수"
        )
        st.plotly_chart(fig1, use_container_width=True)

        # ------------------------------
        # 4) 담당자 TOP 그래프
        # ------------------------------
        st.markdown("### 👤 담당자별 비매칭 TOP 20")

        mc = (
            unmatched_filtered.groupby("구역담당자_통합")["계약번호_정제"]
            .nunique()
            .sort_values(ascending=False)
            .head(20)
        )

        fig2 = px.bar(
            mc.reset_index(),
            x="구역담당자_통합",
            y="계약번호_정제",
            text="계약번호_정제",
        )
        fig2.update_traces(textposition="outside")
        fig2.update_layout(
            height=420,
            margin=dict(l=20, r=20, t=40, b=120),
            xaxis_title="담당자",
            yaxis_title="비매칭 계약 수",
            xaxis_tickangle=-45
        )
        st.plotly_chart(fig2, use_container_width=True)

# ============================================================
# PART 4 — 계약번호 상세조회 / 비매칭 상세 / 피드백(활동등록) 시스템
# ============================================================

# ------------------------------------------------------------
# TAB 2 — VOC 전체 (계약 기준 정제·요약)
# ------------------------------------------------------------
with tab_all:
    st.subheader("📘 VOC 전체 (계약번호 기준)")

    # 🔽 지사 / 담당자 동적 필터
    colA, colB = st.columns([1, 1])

    branch_opts = ["전체"] + branch_all
    sel_branch_all = colA.selectbox("지사 선택", branch_opts)

    temp_mgr = voc_filtered.copy()
    if sel_branch_all != "전체":
        temp_mgr = temp_mgr[temp_mgr["관리지사"] == sel_branch_all]

    mgr_opts = ["전체"] + sorted(
        temp_mgr["구역담당자_통합"].dropna().astype(str).unique().tolist()
    )
    sel_mgr_all = colB.selectbox("담당자 선택", mgr_opts)

    # 🔽 검색
    c1, c2, c3 = st.columns(3)
    q_cn = c1.text_input("계약번호 검색(부분)")
    q_nm = c2.text_input("상호 검색(부분)")
    q_addr = c3.text_input("주소 검색(부분)")

    df_all = voc_filtered.copy()

    if sel_branch_all != "전체":
        df_all = df_all[df_all["관리지사"] == sel_branch_all]

    if sel_mgr_all != "전체":
        df_all = df_all[df_all["구역담당자_통합"].astype(str) == sel_mgr_all]

    if q_cn:
        df_all = df_all[df_all["계약번호_정제"].astype(str).str.contains(q_cn)]

    if q_nm and "상호" in df_all.columns:
        df_all = df_all[df_all["상호"].astype(str).str.contains(q_nm)]

    if q_addr:
        cond = None
        for col in address_cols:
            if col in df_all.columns:
                now = df_all[col].astype(str).str.contains(q_addr)
                cond = now if cond is None else (cond | now)
        if cond is not None:
            df_all = df_all[cond]

    # 🔽 계약번호 기준 1건으로 축약
    if not df_all.empty:
        df_sorted = df_all.sort_values("접수일시", ascending=False)
        grp = df_sorted.groupby("계약번호_정제")
        idx_latest = grp["접수일시"].idxmax()
        df_summary = df_sorted.loc[idx_latest].copy()
        df_summary["접수건수"] = grp.size().reindex(df_summary["계약번호_정제"]).values

        show_cols = [
            "계약번호_정제", "상호", "관리지사", "구역담당자_통합",
            "리스크등급", "경과일수", "매칭여부", "접수건수",
            "VOC유형", "VOC유형소", "등록내용", "설치주소_표시"
        ]
        show_cols = [c for c in show_cols if c in df_summary.columns]

        st.markdown(f"📌 요약 계약 수: **{len(df_summary):,} 건**")
        st.dataframe(style_risk(df_summary[show_cols]),
                     use_container_width=True, height=450)
    else:
        st.info("조건에 맞는 VOC 데이터가 없습니다.")

# ------------------------------------------------------------
# TAB 3 — 비매칭 시설 (계약 기준 정제 + UI 개선)
# ------------------------------------------------------------
with tab_unmatched:
    st.subheader("🧯 해지방어 활동시설 (비매칭)")

    df_u = unmatched_filtered.copy()

    if df_u.empty:
        st.info("현재 조건에서 비매칭 데이터 없음")
    else:
        colA, colB = st.columns([1, 1])

        sel_branch_u = colA.selectbox(
            "지사 선택",
            ["전체"] + branch_all
        )

        temp_mgr_u = df_u.copy()
        if sel_branch_u != "전체":
            temp_mgr_u = temp_mgr_u[temp_mgr_u["관리지사"] == sel_branch_u]

        mgr_opts_u = ["전체"] + sorted(
            temp_mgr_u["구역담당자_통합"].dropna().astype(str).unique().tolist()
        )
        sel_mgr_u = colB.selectbox("담당자 선택", mgr_opts_u)

        q1, q2 = st.columns(2)
        uq_cn = q1.text_input("계약번호 검색")
        uq_nm = q2.text_input("상호 검색")

        df_u2 = df_u.copy()
        if sel_branch_u != "전체":
            df_u2 = df_u2[df_u2["관리지사"] == sel_branch_u]

        if sel_mgr_u != "전체":
            df_u2 = df_u2[df_u2["구역담당자_통합"].astype(str) == sel_mgr_u]

        if uq_cn:
            df_u2 = df_u2[df_u2["계약번호_정제"].astype(str).str.contains(uq_cn)]

        if uq_nm and "상호" in df_u2.columns:
            df_u2 = df_u2[df_u2["상호"].astype(str).str.contains(uq_nm)]

        # 계약번호 기준 요약
        df_sorted_u = df_u2.sort_values("접수일시", ascending=False)
        grp_u = df_sorted_u.groupby("계약번호_정제")
        idx_u = grp_u["접수일시"].idxmax()
        df_u_sum = df_sorted_u.loc[idx_u].copy()
        df_u_sum["접수건수"] = grp_u.size().reindex(df_u_sum["계약번호_정제"]).values

        show_cols_u = [
            "계약번호_정제",
            "상호",
            "관리지사",
            "구역담당자_통합",
            "리스크등급",
            "경과일수",
            "접수건수",
            "VOC유형소",
            "등록내용",
            "설치주소_표시"
        ]
        show_cols_u = [c for c in show_cols_u if c in df_u_sum.columns]

        st.markdown(f"⚠ 비매칭 계약: **{len(df_u_sum):,} 건**")
        st.dataframe(style_risk(df_u_sum[show_cols_u]),
                     use_container_width=True, height=430)

        # 선택된 계약 상세 보기
        sel_cn_u = st.selectbox(
            "상세 이력 조회할 계약 선택",
            ["(선택)"] + df_u_sum["계약번호_정제"].astype(str).tolist()
        )

        if sel_cn_u != "(선택)":
            voc_detail = df_voc[df_voc["계약번호_정제"] == sel_cn_u].sort_values("접수일시", ascending=False)
            st.markdown(f"### 🔍 `{sel_cn_u}` 상세 VOC 이력")
            st.dataframe(style_risk(voc_detail[display_cols]),
                         use_container_width=True, height=350)

# ------------------------------------------------------------
# TAB 4 — 계약별 상세 + 피드백(활동등록)
# ------------------------------------------------------------
with tab_drill:
    st.subheader("🔍 계약별 상세 조회 + 활동등록(피드백)")

    df_d = voc_filtered.copy()

    # 🔽 계약번호 하나를 선택하도록
    cn_list = sorted(df_d["계약번호_정제"].dropna().unique().tolist())
    sel_cn = st.selectbox("계약번호 선택", ["(선택)"] + cn_list)

    if sel_cn != "(선택)":
        voc_hist = df_voc[df_voc["계약번호_정제"] == sel_cn].sort_values("접수일시", ascending=False)
        other_hist = df_other[df_other["계약번호_정제"] == sel_cn]

        # 기본 정보
        st.markdown("### 📌 요약 정보")
        base = voc_hist.iloc[0]

        c1, c2, c3, c4 = st.columns(4)
        c1.metric("상호", base.get("상호", ""))
        c2.metric("지사", base.get("관리지사", ""))
        c3.metric("담당자", base.get("구역담당자_통합", ""))
        c4.metric("VOC이력", f"{len(voc_hist)} 건")

        st.caption(f"📍 주소: {base.get('설치주소_표시','')}")
        if fee_raw_col:
            st.caption(f"💰 월정료: {base.get(fee_raw_col, '')}")

        st.markdown("---")

        # VOC 상세 이력
        st.markdown("### 📘 VOC 상세 이력")
        st.dataframe(style_risk(voc_hist[display_cols]),
                     use_container_width=True, height=350)

        # 기타 출처
        st.markdown("### 📂 기타 출처 이력")
        if other_hist.empty:
            st.info("기타 출처 없음")
        else:
            st.dataframe(other_hist, use_container_width=True, height=300)

        st.markdown("---")

        # --------------------------------------------------
        # 피드백(활동등록)
        # --------------------------------------------------
        st.markdown("## 📝 담당자 처리내역 등록")

        fb_all = st.session_state["feedback_df"]
        fb_sel = fb_all[fb_all["계약번호_정제"] == sel_cn].sort_values("등록일자", ascending=False)

        st.markdown("### 📄 기존 처리내역")
        if fb_sel.empty:
            st.info("등록된 처리내역 없음")
        else:
            st.dataframe(fb_sel, use_container_width=True, height=280)

        st.markdown("### ➕ 새 처리내역 등록")

        new_content = st.text_area("고객대응 내용")
        new_writer = st.text_input("등록자")
        new_note = st.text_input("비고(Optional)")

        if st.button("등록하기"):
            if not new_content or not new_writer:
                st.warning("내용 + 등록자 필수 입력")
            else:
                new_row = {
                    "계약번호_정제": sel_cn,
                    "고객대응내용": new_content,
                    "등록자": new_writer,
                    "등록일자": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                    "비고": new_note
                }
                fb_all = pd.concat([fb_all, pd.DataFrame([new_row])], ignore_index=True)
                st.session_state["feedback_df"] = fb_all
                save_feedback(FEEDBACK_PATH, fb_all)
                st.success("등록 완료")
                st.rerun()

# ============================================================
# PART 5 — 담당자 알림 발송(기업용 완성본)
# ============================================================

with tab_alert:
    st.subheader("📨 담당자 알림 발송(기업용 완성형)")

    st.markdown(
        """
        비매칭(X) 발생 시설을 담당자에게 이메일로 알림 전송합니다.<br>
        CSV에는 계약번호 중복 없이 **최신 VOC 기준 1행 요약 정제본**이 첨부됩니다.
        """,
        unsafe_allow_html=True
    )

    # 담당자 목록 생성
    unmatched_alert = unmatched_filtered.copy()
    alert_group = unmatched_alert.groupby("구역담당자_통합")

    alert_rows = []
    for mgr, g in alert_group:
        if not mgr:
            continue
        count = g["계약번호_정제"].nunique()
        email = manager_contacts.get(mgr, {}).get("email", "")
        alert_rows.append([mgr, email, count])

    alert_df = pd.DataFrame(alert_rows, columns=["담당자", "이메일", "비매칭 계약수"])
    st.dataframe(alert_df, use_container_width=True, height=260)

    st.markdown("---")

    # -----------------------------
    # 개별 담당자 선택
    # -----------------------------
    sel_mgr = st.selectbox(
        "담당자 선택",
        ["(선택)"] + alert_df["담당자"].tolist()
    )

    if sel_mgr != "(선택)":
        # 담당자의 이메일 자동 로드
        default_email = manager_contacts.get(sel_mgr, {}).get("email", "")
        email_input = st.text_input("이메일 주소", value=default_email)

        # 해당 담당자의 비매칭 데이터 필터
        df_mgr = unmatched_alert[
            unmatched_alert["구역담당자_통합"] == sel_mgr
        ]

        if df_mgr.empty:
            st.info("해당 담당자는 비매칭 시설이 없습니다.")
        else:
            st.markdown(f"### 🔍 {sel_mgr} 담당자 비매칭 계약 목록")
            st.dataframe(
                df_mgr[
                    [
                        "계약번호_정제",
                        "상호",
                        "관리지사",
                        "VOC유형",
                        "VOC유형소",
                        "등록내용",
                        "리스크등급",
                        "경과일수"
                    ]
                ],
                use_container_width=True,
                height=300
            )

            # -----------------------------
            # CSV 생성 (최신 VOC 1건 요약)
            # -----------------------------
            df_sorted = df_mgr.sort_values("접수일시", ascending=False)
            grp = df_sorted.groupby("계약번호_정제")
            idx = grp["접수일시"].idxmax()
            df_latest = df_sorted.loc[idx].copy()

            df_latest = df_latest[
                [
                    "계약번호_정제",
                    "상호",
                    "관리지사",
                    "구역담당자_통합",
                    "VOC유형",
                    "VOC유형소",
                    "등록내용",
                    "설치주소_표시",
                    "리스크등급",
                    "경과일수"
                ]
            ]

            csv_bytes = df_latest.to_csv(index=False).encode("utf-8-sig")

            # -----------------------------
            # 이메일 본문 자동 구성
            # -----------------------------
            subject = f"[해지VOC] {sel_mgr} 담당자 비매칭 시설 안내"
            body = (
                f"{sel_mgr} 담당자님,\n\n"
                f"총 {len(df_latest)}건의 비매칭 시설이 확인되었습니다.\n"
                f"첨부된 CSV 파일을 확인하시고 신속히 처리 부탁드립니다.\n\n"
                "- 해지VOC 관리자 드림 -"
            )

            # -----------------------------
            # 이메일 발송 버튼
            # -----------------------------
            if st.button("📤 이메일 발송하기"):
                try:
                    msg = EmailMessage()
                    msg["Subject"] = subject
                    msg["From"] = f"{SENDER_NAME} <{SMTP_USER}>"
                    msg["To"] = email_input
                    msg.set_content(body)

                    msg.add_attachment(
                        csv_bytes,
                        maintype="application",
                        subtype="octet-stream",
                        filename=f"비매칭계약_{sel_mgr}.csv",
                    )

                    with smtplib.SMTP(SMTP_HOST, SMTP_PORT) as smtp:
                        smtp.starttls()
                        smtp.login(SMTP_USER, SMTP_PASSWORD)
                        smtp.send_message(msg)

                    st.success(f"✅ 이메일 발송 완료 → {email_input}")

                except Exception as e:
                    st.error(f"❌ 이메일 전송 실패: {e}")

    st.markdown("---")
    st.caption("※ 비매칭: 해지VOC 접수 후 활동내역 미등록 시설")

