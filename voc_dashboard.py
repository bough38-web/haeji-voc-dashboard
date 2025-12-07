############################################################
#  PART 1 — Core Imports, CSS, Config, SMTP, File Loader  #
############################################################

import os
from datetime import datetime, date
import smtplib
from email.message import EmailMessage

import numpy as np
import pandas as pd
import streamlit as st

# Try Plotly, fallback if not installed
try:
    import plotly.express as px
    HAS_PLOTLY = True
except Exception:
    HAS_PLOTLY = False


# ============================================================
# 0. 기본 설정 & Apple-Style Light UI CSS
# ============================================================
st.set_page_config(page_title="해지 VOC 종합 대시보드", layout="wide")

st.markdown(
    """
    <style>
    html, body, .stApp {
        background-color: #f5f5f7 !important;
        color: #111827 !important;
        font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif;
    }

    .block-container {
        padding-top: 0.8rem !important;
        padding-bottom: 2rem !important;
    }

    [data-testid="stSidebar"] {
        background-color: #fafafa !important;
        border-right: 1px solid #e5e7eb;
    }

    h1,h2,h3,h4 {
        font-weight: 600;
    }

    .branch-grid {
        display: grid;
        grid-template-columns: repeat(4, minmax(0, 1fr));
        gap: 12px;
    }
    @media(max-width:1200px){
        .branch-grid { grid-template-columns: repeat(2, 1fr); }
    }
    @media(max-width:700px){
        .branch-grid { grid-template-columns: repeat(1, 1fr); }
    }

    .branch-card {
        background:#ffffff;
        border-radius:14px;
        padding:1rem;
        border:1px solid #e5e7eb;
        box-shadow:0 3px 6px rgba(0,0,0,0.05);
    }

    .badge-high {
        background:#ef4444; color:#fff; padding:2px 7px; border-radius:6px;
        font-size:0.75rem; font-weight:600;
    }
    .badge-medium {
        background:#f59e0b; color:#fff; padding:2px 7px; border-radius:6px;
        font-size:0.75rem; font-weight:600;
    }
    .badge-low {
        background:#3b82f6; color:#fff; padding:2px 7px; border-radius:6px;
        font-size:0.75rem; font-weight:600;
    }
    </style>
    """,
    unsafe_allow_html=True,
)

############################################################
# 1. SMTP 설정 + 파일 경로
############################################################

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
    except:
        pass
    SMTP_HOST = os.getenv("SMTP_HOST", "")
    SMTP_PORT = int(os.getenv("SMTP_PORT", "587"))
    SMTP_USER = os.getenv("SMTP_USER", "")
    SMTP_PASSWORD = os.getenv("SMTP_PASSWORD", "")
    SENDER_NAME = os.getenv("SENDER_NAME", "해지VOC 관리자")

MERGED_PATH = "merged.xlsx"
FEEDBACK_PATH = "feedback.csv"
CONTACT_PATH = "contact_map.xlsx"


############################################################
# 2. 유틸
############################################################

def safe_str(x):
    if pd.isna(x):
        return ""
    return str(x).strip()


def detect_column(df: pd.DataFrame, keywords: list[str]):
    """담당자/이메일/휴대폰 컬럼 자동 탐색"""
    for k in keywords:
        if k in df.columns:
            return k

    for col in df.columns:
        for k in keywords:
            if k.lower() in str(col).lower():
                return col
    return None


############################################################
# 3. 데이터 로더
############################################################

@st.cache_data
def load_voc_data(path: str) -> pd.DataFrame:
    if not os.path.exists(path):
        st.error("❌ merged.xlsx 파일이 없습니다.")
        return pd.DataFrame()

    df = pd.read_excel(path)

    # 계약번호 정제
    if "계약번호" in df.columns:
        df["계약번호"] = (
            df["계약번호"].astype(str)
            .str.replace(",", "", regex=False)
            .str.strip()
        )
        df["계약번호_정제"] = (
            df["계약번호"]
            .astype(str)
            .str.replace(r"[^0-9A-Za-z]", "", regex=True)
            .str.strip()
        )
    else:
        df["계약번호_정제"] = ""

    # 출처 정제
    if "출처" in df.columns:
        df["출처"] = df["출처"].replace({"고객리스트": "해지시설"})

    # 날짜 변환
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
    return pd.DataFrame(
        columns=["계약번호_정제", "고객대응내용", "등록자", "등록일자", "비고"]
    )


def save_feedback(path, df):
    df.to_csv(path, index=False, encoding="utf-8-sig")


@st.cache_data
def load_contact_map(path: str):
    """contact_map.xlsx 자동 매핑"""
    if not os.path.exists(path):
        st.warning("⚠ contact_map.xlsx 파일을 찾을 수 없습니다.")
        return pd.DataFrame(), {}

    df_c = pd.read_excel(path)

    name_col = detect_column(df_c, ["담당자", "구역담당자", "성명"])
    email_col = detect_column(df_c, ["이메일", "메일"])
    phone_col = detect_column(df_c, ["휴대폰", "전화"])

    if not name_col or not email_col:
        st.warning("⚠ 담당자/이메일 컬럼을 찾지 못했습니다. contact_map.xlsx 확인 필요.")
        return df_c, {}

    rename_map = {
        name_col: "구역담당자_통합",
        email_col: "이메일",
    }
    if phone_col:
        rename_map[phone_col] = "휴대폰"

    df_c = df_c.rename(columns=rename_map)

    info = {}
    for _, row in df_c.iterrows():
        nm = safe_str(row["구역담당자_통합"])
        if nm:
            info[nm] = {
                "email": safe_str(row.get("이메일", "")),
                "phone": safe_str(row.get("휴대폰", "")),
            }

    return df_c, info


# ---------------- 실 데이터 로드 ----------------
df = load_voc_data(MERGED_PATH)
if df.empty:
    st.stop()

if "feedback_df" not in st.session_state:
    st.session_state["feedback_df"] = load_feedback(FEEDBACK_PATH)

contact_df, manager_contacts = load_contact_map(CONTACT_PATH)

############################################################
# 4. 지사명 축약 및 정렬순서
############################################################

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

BRANCH_ORDER = ["중앙", "강북", "서대문", "고양", "의정부", "남양주", "강릉", "원주"]


def sort_branch(series):
    """지사명 정렬"""
    return sorted(
        [s for s in series if s in BRANCH_ORDER],
        key=lambda x: BRANCH_ORDER.index(x)
    )


############################################################
# 5. 영업구역 / 담당자 통합 컬럼 생성
############################################################

def make_zone(row):
    """영업구역 컬럼 자동 통합"""
    for c in ["영업구역번호", "담당상세", "영업구역정보"]:
        if c in row and pd.notna(row[c]) and str(row[c]).strip():
            return row[c]
    return ""

df["영업구역_통합"] = df.apply(make_zone, axis=1)

# 담당자 우선순위
mgr_priority = ["구역담당자", "담당자", "처리자"]

def pick_manager(row):
    """담당자 자동 통합"""
    for c in mgr_priority:
        if c in row and pd.notna(row[c]) and str(row[c]).strip() != "":
            return row[c]
    return ""

df["구역담당자_통합"] = df.apply(pick_manager, axis=1)

# 주소 컬럼 자동 탐색
address_cols = [c for c in df.columns if "주소" in str(c)]


############################################################
# 6. 출처 분리 + 매칭 여부 계산
############################################################

df_voc = df[df.get("출처") == "해지VOC"].copy()
df_other = df[df.get("출처") != "해지VOC"].copy()

# 기타 출처 계약번호 SET
other_sets = {
    src: set(df_other[df_other["출처"] == src]["계약번호_정제"].dropna())
    for src in ["해지시설", "해지요청", "설변", "정지", "해지파이프라인"]
    if "출처" in df_other.columns
}
other_union = set().union(*other_sets.values()) if len(other_sets) else set()

df_voc["매칭여부"] = df_voc["계약번호_정제"].apply(
    lambda x: "매칭(O)" if x in other_union else "비매칭(X)"
)


############################################################
# 7. 설치주소 / 월정료 정제 + 월정료 구간
############################################################

def coalesce_cols(row, cols):
    """여러 주소 후보 중 가장 먼저 나오는 유효값"""
    for c in cols:
        if c in row.index and pd.notna(row[c]) and str(row[c]).strip():
            return row[c]
    return np.nan

df_voc["설치주소_표시"] = df_voc.apply(
    lambda r: coalesce_cols(r, ["시설_설치주소", "설치주소"]),
    axis=1
)

# 월정료 원본 컬럼 탐색
fee_raw_col = None
for cand in ["시설_KTT월정료(조정)", "KTT월정료(조정)"]:
    if cand in df_voc.columns:
        fee_raw_col = cand
        break

def parse_fee(v):
    if pd.isna(v):
        return np.nan
    s = str(v).replace(",", "").strip()
    if not s:
        return np.nan
    digits = "".join([ch for ch in s if ch.isdigit() or ch == "."])
    if digits == "":
        return np.nan
    value = float(digits)
    # 데이터 오류(10배 이상) 보정
    if value >= 200000:
        value = value / 10
    return value

if fee_raw_col:
    df_voc["월정료_수치"] = df_voc[fee_raw_col].apply(parse_fee)

    def pretty_fee(v):
        if pd.isna(v):
            return ""
        return f"{int(round(v, 0)):,}"

    df_voc[fee_raw_col] = df_voc["월정료_수치"].apply(pretty_fee)

    # 월정료 구간 (사용자 요청: 10만 미만 / 20만 / 30만 / 40만 / 50만 이상)
    def fee_band(v):
        if pd.isna(v):
            return "미기재"
        if v >= 500000:
            return "50만 이상"
        if v >= 400000:
            return "40만 이상"
        if v >= 300000:
            return "30만 이상"
        if v >= 200000:
            return "20만 이상"
        if v >= 100000:
            return "10만 이상"
        return "10만 미만"

    df_voc["월정료구간"] = df_voc["월정료_수치"].apply(fee_band)

else:
    df_voc["월정료_수치"] = np.nan
    df_voc["월정료구간"] = "미기재"


############################################################
# 8. 리스크 등급 계산 (경과일 기반)
############################################################

today = date.today()

def compute_risk(row):
    dt = row.get("접수일시")
    if pd.isna(dt):
        return np.nan, "LOW"
    if not isinstance(dt, (pd.Timestamp, datetime)):
        dt = pd.to_datetime(dt, errors="coerce")
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

df_voc["경과일수"], df_voc["리스크등급"] = zip(*df_voc.apply(compute_risk, axis=1))

df_unmatched = df_voc[df_voc["매칭여부"] == "비매칭(X)"].copy()

############################################################
# 9. 표시 컬럼 정의 & 스타일링
############################################################

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
    fee_raw_col if fee_raw_col else None,
    "계약상태(중)",
    "서비스(소)",
]

display_cols_raw = [c for c in fixed_order if c in df_voc.columns]


def filter_valid_columns(cols, df_base):
    valid = []
    for c in cols:
        s = df_base[c]
        m = s.notna() & ~s.astype(str).str.strip().isin(["", "None", "nan"])
        if m.any():
            valid.append(c)
    return valid


display_cols = filter_valid_columns(display_cols_raw, df_voc)


def style_risk(df_view):
    if "리스크등급" not in df_view.columns:
        return df_view

    def _row(st_row):
        level = st_row.get("리스크등급", "")
        if level == "HIGH":
            bg = "#fee2e2"
        elif level == "MEDIUM":
            bg = "#fef3c7"
        else:
            bg = "#e0f2fe"
        return [f"background-color:{bg};"] * len(st_row)

    return df_view.style.apply(_row, axis=1)


############################################################
# 10. 사이드바 글로벌 필터
############################################################

st.sidebar.title("🔧 글로벌 필터")

# 날짜 범위
if "접수일시" in df_voc and df_voc["접수일시"].notna().any():
    min_d = df_voc["접수일시"].min().date()
    max_d = df_voc["접수일시"].max().date()
    dr = st.sidebar.date_input(
        "접수일자 범위",
        value=(min_d, max_d),
        min_value=min_d,
        max_value=max_d
    )
else:
    dr = None

# 관리지사
branches_all = sort_branch(df_voc["관리지사"].dropna().unique())
sel_branches = st.sidebar.multiselect(
    "관리지사(복수 선택)",
    options=branches_all,
    default=branches_all
)

# 리스크
risk_all = ["HIGH", "MEDIUM", "LOW"]
sel_risk = st.sidebar.multiselect(
    "리스크등급",
    options=risk_all,
    default=risk_all
)

# 매칭여부
match_all = ["매칭(O)", "비매칭(X)"]
sel_match = st.sidebar.multiselect(
    "매칭여부",
    options=match_all,
    default=match_all
)

# 월정료 구간 — 개선한 10만/20만/30만/40만/50만 이상 버전
fee_filter_global = st.sidebar.radio(
    "월정료 구간(글로벌)",
    ["전체", "10만 미만", "10만 이상", "20만 이상", "30만 이상", "40만 이상", "50만 이상"],
    index=0
)

st.sidebar.markdown("---")
st.sidebar.caption(f"마지막 갱신: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")


############################################################
# 11. 글로벌 필터 적용
############################################################

voc_filtered_global = df_voc.copy()

# 날짜
if dr and len(dr) == 2:
    sd, ed = dr
    voc_filtered_global = voc_filtered_global[
        (voc_filtered_global["접수일시"] >= pd.to_datetime(sd))
        & (voc_filtered_global["접수일시"] < pd.to_datetime(ed) + pd.Timedelta(days=1))
    ]

# 지사
if sel_branches:
    voc_filtered_global = voc_filtered_global[voc_filtered_global["관리지사"].isin(sel_branches)]

# 리스크
if sel_risk:
    voc_filtered_global = voc_filtered_global[voc_filtered_global["리스크등급"].isin(sel_risk)]

# 매칭
if sel_match:
    voc_filtered_global = voc_filtered_global[voc_filtered_global["매칭여부"].isin(sel_match)]

# 월정료 필터
if fee_filter_global != "전체":
    cond = voc_filtered_global["월정료_수치"]

    if fee_filter_global == "10만 미만":
        voc_filtered_global = voc_filtered_global[(cond < 100000) & cond.notna()]
    elif fee_filter_global == "10만 이상":
        voc_filtered_global = voc_filtered_global[cond >= 100000]
    elif fee_filter_global == "20만 이상":
        voc_filtered_global = voc_filtered_global[cond >= 200000]
    elif fee_filter_global == "30만 이상":
        voc_filtered_global = voc_filtered_global[cond >= 300000]
    elif fee_filter_global == "40만 이상":
        voc_filtered_global = voc_filtered_global[cond >= 400000]
    elif fee_filter_global == "50만 이상":
        voc_filtered_global = voc_filtered_global[cond >= 500000]

unmatched_global = voc_filtered_global[voc_filtered_global["매칭여부"] == "비매칭(X)"].copy()


############################################################
# 12. KPI 카드
############################################################

st.markdown("## 📊 해지 VOC 종합 대시보드")

total_voc_rows = len(voc_filtered_global)
unique_contracts = voc_filtered_global["계약번호_정제"].nunique()
unmatched_contracts = unmatched_global["계약번호_정제"].nunique()
matched_contracts = voc_filtered_global[voc_filtered_global["매칭여부"] == "매칭(O)"]["계약번호_정제"].nunique()

k1, k2, k3, k4 = st.columns(4)
k1.metric("VOC 접수건수(행)", f"{total_voc_rows:,}")
k2.metric("VOC 계약 수(유니크)", f"{unique_contracts:,}")
k3.metric("비매칭(X) 계약 수", f"{unmatched_contracts:,}")
k4.metric("매칭(O) 계약 수", f"{matched_contracts:,}")

st.markdown("---")

############################################################
# 13. 탭 구성
############################################################

tab_viz, tab_all, tab_unmatched, tab_drill, tab_filter, tab_alert = st.tabs(
    [
        "📊 지사/담당자 시각화",
        "📘 VOC 전체(계약 기준)",
        "🧯 해지방어 활동시설(비매칭)",
        "🔍 해지상담대상 활동등록",
        "🎯 정밀 필터(VOC유형소)",
        "📨 담당자 알림(베타)",
    ]
)

# ============================================================
# TAB VIZ — 지사 / 담당자 시각화
# ============================================================
with tab_viz:
    st.subheader("📊 지사 / 담당자별 비매칭 리스크 현황 (리뉴얼 개선판)")

    if unmatched_global.empty:
        st.info("현재 조건에서 비매칭(X) 데이터가 없습니다.")
        st.stop()

    # ------------------------------
    # 1) 지사 요약 카드
    # ------------------------------
    df_clean = unmatched_global.dropna(subset=["관리지사"])

    branch_stats = (
        df_clean.groupby("관리지사")
        .agg(
            계약수=("계약번호_정제", "nunique"),
            HIGH=("리스크등급", lambda s: (s == "HIGH").sum()),
            MEDIUM=("리스크등급", lambda s: (s == "MEDIUM").sum()),
            LOW=("리스크등급", lambda s: (s == "LOW").sum()),
        )
    ).reindex(BRANCH_ORDER).dropna(how="all")

    st.markdown("### 🏢 지사별 비매칭 요약")

    html = '<div class="branch-grid">'
    for branch, r in branch_stats.iterrows():
        html += f"""
        <div class="branch-card">
            <div class="branch-card-header">{branch}</div>
            <div class="branch-card-sub">계약 {int(r['계약수'])}건</div>

            <span class="badge-high">HIGH {int(r['HIGH'])}</span>&nbsp;
            <span class="badge-medium">MED {int(r['MEDIUM'])}</span>&nbsp;
            <span class="badge-low">LOW {int(r['LOW'])}</span>
        </div>
        """
    html += "</div>"
    st.markdown(html, unsafe_allow_html=True)

    st.markdown("---")

    # ------------------------------
    # 2) 지사 선택 → 담당자 동적 필터
    # ------------------------------
    f1, f2, f3 = st.columns([1.2, 1.2, 1])

    branch_opts = ["전체"] + sort_branch(df_clean["관리지사"].unique())
    sel_branch = f1.selectbox("지사 선택", branch_opts)

    df_mgr_scope = df_clean.copy()
    if sel_branch != "전체":
        df_mgr_scope = df_mgr_scope[df_mgr_scope["관리지사"] == sel_branch]

    mgr_options = (
        df_mgr_scope["구역담당자_통합"]
        .dropna()
        .astype(str)
        .replace("nan", "")
        .unique()
        .tolist()
    )
    mgr_options = sorted([m for m in mgr_options if m.strip() != ""])

    sel_mgr = f2.selectbox("담당자 선택 (상세 리스크)", ["(전체)"] + mgr_options)

    scope_df = df_mgr_scope.copy()
    if sel_mgr != "(전체)":
        scope_df = scope_df[scope_df["구역담당자_통합"].astype(str) == sel_mgr]

    f3.metric("선택 계약 수", f"{scope_df['계약번호_정제'].nunique():,}")
    st.caption(
        f"HIGH { (scope_df['리스크등급']=='HIGH').sum() } / "
        f"MEDIUM { (scope_df['리스크등급']=='MEDIUM').sum() } / "
        f"LOW { (scope_df['리스크등급']=='LOW').sum() }"
    )

    st.markdown("---")

    # ------------------------------
    # 3) 지사별 리스크 스택 바
    # ------------------------------
    st.markdown("### 🧱 지사별 리스크 분포 (STACKED BAR)")

    risk_branch = (
        df_clean.groupby(["관리지사", "리스크등급"])["계약번호_정제"]
        .nunique()
        .reset_index()
    )
    risk_branch["관리지사"] = pd.Categorical(
        risk_branch["관리지사"], categories=BRANCH_ORDER, ordered=True
    )

    if HAS_PLOTLY:
        fig1 = px.bar(
            risk_branch,
            x="관리지사",
            y="계약번호_정제",
            color="리스크등급",
            barmode="stack",
        )
        fig1.update_layout(height=360)
        st.plotly_chart(fig1, use_container_width=True)
    else:
        st.bar_chart(
            risk_branch.pivot(
                index="관리지사",
                columns="리스크등급",
                values="계약번호_정제",
            ).fillna(0)
        )

    st.markdown("---")

    # ------------------------------
    # 4) 담당자 TOP15 + 전체 리스크 도넛
    # ------------------------------
    g1, g2 = st.columns(2)

    # TOP15
    with g1:
        st.markdown("#### 👤 담당자별 비매칭 TOP 15")
        df_scope = df_clean if sel_branch == "전체" else df_clean[df_clean["관리지사"] == sel_branch]

        top15 = (
            df_scope.groupby("구역담당자_통합")["계약번호_정제"]
            .nunique()
            .sort_values(ascending=False)
            .head(15)
        )

        if HAS_PLOTLY:
            fig2 = px.bar(
                top15.reset_index(),
                x="구역담당자_통합",
                y="계약번호_정제",
                text="계약번호_정제",
            )
            fig2.update_traces(textposition="outside")
            fig2.update_layout(height=330, xaxis_tickangle=-35)
            st.plotly_chart(fig2, use_container_width=True)
        else:
            st.bar_chart(top15)

    # 도넛 차트
    with g2:
        st.markdown("#### 🍩 전체 비매칭 리스크 비율")

        rc = df_clean["리스크등급"].value_counts().reindex(["HIGH", "MEDIUM", "LOW"]).fillna(0)
        rc_df = rc.reset_index()
        rc_df.columns = ["리스크등급", "건수"]

        if HAS_PLOTLY:
            fig3 = px.pie(
                rc_df,
                names="리스크등급",
                values="건수",
                hole=0.45,
            )
            fig3.update_layout(height=330)
            st.plotly_chart(fig3, use_container_width=True)
        else:
            st.bar_chart(rc_df.set_index("리스크등급")["건수"])

    st.markdown("---")

    # ------------------------------
    # 5) 일자별 추이 + 선택 담당자 리스크막대
    # ------------------------------
    t1, t2 = st.columns(2)

    with t1:
        st.markdown("#### 📈 일별 비매칭 추이")
        trend = (
            df_clean.assign(접수일=df_clean["접수일시"].dt.date)
            .groupby("접수일")["계약번호_정제"]
            .nunique()
        )

        if HAS_PLOTLY:
            fig4 = px.line(trend.reset_index(), x="접수일", y="계약번호_정제")
            fig4.update_layout(height=260)
            st.plotly_chart(fig4, use_container_width=True)
        else:
            st.line_chart(trend)

    with t2:
        st.markdown("#### 👤 선택 담당자 리스크 분포")

        if sel_mgr == "(전체)":
            st.info("담당자를 선택하면 상세 리스크가 표시됩니다.")
        else:
            mgr_df = df_clean[df_clean["구역담당자_통합"].astype(str) == sel_mgr]
            rc2 = mgr_df["리스크등급"].value_counts().reindex(["HIGH","MEDIUM","LOW"]).fillna(0)

            if HAS_PLOTLY:
                fig5 = px.bar(
                    rc2.reset_index(),
                    x="index",
                    y="리스크등급",
                    text="리스크등급",
                )
                fig5.update_traces(textposition="outside")
                fig5.update_layout(height=260)
                st.plotly_chart(fig5, use_container_width=True)
            else:
                st.bar_chart(rc2)

# ============================================================
# TAB ALL (VOC 전체)
# ============================================================
with tab_all:
    st.subheader("📘 VOC 전체 (계약번호 기준 요약)")

    col1, col2 = st.columns([2, 3])

    branch_opts = ["전체"] + sort_branch(voc_filtered_global["관리지사"].dropna().unique())
    sel_branch_all = col1.radio("지사 선택", branch_opts, horizontal=True)

    df_mgr_temp = voc_filtered_global.copy()
    if sel_branch_all != "전체":
        df_mgr_temp = df_mgr_temp[df_mgr_temp["관리지사"] == sel_branch_all]

    mgr_opts = ["전체"] + sorted(
        df_mgr_temp["구역담당자_통합"].dropna().astype(str).unique().tolist()
    )
    sel_mgr_all = col2.radio("담당자 선택", mgr_opts, horizontal=True)

    c1, c2, c3 = st.columns(3)
    q_cn = c1.text_input("계약번호 검색")
    q_name = c2.text_input("상호 검색")
    q_addr = c3.text_input("주소 검색")

    temp = voc_filtered_global.copy()
    if sel_branch_all != "전체":
        temp = temp[temp["관리지사"] == sel_branch_all]
    if sel_mgr_all != "전체":
        temp = temp[temp["구역담당자_통합"].astype(str) == sel_mgr_all]

    if q_cn:
        temp = temp[temp["계약번호_정제"].astype(str).str.contains(q_cn)]
    if q_name:
        if "상호" in temp.columns:
            temp = temp[temp["상호"].astype(str).str.contains(q_name)]
    if q_addr:
        cond = None
        for col in address_cols:
            if col in temp.columns:
                _c = temp[col].astype(str).str.contains(q_addr)
                cond = _c if cond is None else (cond | _c)
        if cond is not None:
            temp = temp[cond]

    if temp.empty:
        st.info("조건에 맞는 VOC 데이터가 없습니다.")
    else:
        temp2 = temp.sort_values("접수일시", ascending=False)
        grp = temp2.groupby("계약번호_정제")
        idx_latest = grp["접수일시"].idxmax()
        df_sum = temp2.loc[idx_latest].copy()
        df_sum["접수건수"] = grp.size().reindex(df_sum["계약번호_정제"]).values

        show_cols = [
            "계약번호_정제", "상호", "관리지사", "구역담당자_통합",
            "리스크등급", "경과일수",
            "매칭여부", "접수건수",
            "VOC유형", "VOC유형소",
            "등록내용", "설치주소_표시",
            fee_raw_col if fee_raw_col else None,
            "계약상태(중)", "서비스(소)",
        ]
        show_cols = [c for c in show_cols if c in df_sum.columns]

        st.dataframe(style_risk(df_sum[show_cols]), use_container_width=True, height=480)

# ============================================================
# TAB UNMATCHED — 비매칭 계약
# ============================================================
with tab_unmatched:
    st.subheader("🧯 해지방어 활동시설 (비매칭 계약 기준)")

    if unmatched_global.empty:
        st.info("비매칭 데이터가 없습니다.")
    else:
        col1, col2 = st.columns([2, 3])

        branch_opts = ["전체"] + sort_branch(unmatched_global["관리지사"].dropna().unique())
        sel_branch_u = col1.radio("지사 선택", branch_opts, horizontal=True)

        df_u_temp = unmatched_global.copy()
        if sel_branch_u != "전체":
            df_u_temp = df_u_temp[df_u_temp["관리지사"] == sel_branch_u]

        mgr_opts = ["전체"] + sorted(
            df_u_temp["구역담당자_통합"].dropna().astype(str).unique().tolist()
        )
        sel_mgr_u = col2.radio("담당자", mgr_opts, horizontal=True)

        s1, s2 = st.columns(2)
        q_cn2 = s1.text_input("계약번호 검색")
        q_name2 = s2.text_input("상호 검색")

        temp_u = df_u_temp.copy()
        if sel_mgr_u != "전체":
            temp_u = temp_u[temp_u["구역담당자_통합"].astype(str) == sel_mgr_u]
        if q_cn2:
            temp_u = temp_u[temp_u["계약번호_정제"].astype(str).str.contains(q_cn2)]
        if q_name2:
            temp_u = temp_u[temp_u["상호"].astype(str).str.contains(q_name2)]

        if temp_u.empty:
            st.info("조건에 맞는 비매칭 계약 없음.")
        else:
            temp_sorted = temp_u.sort_values("접수일시", ascending=False)
            grp_u = temp_sorted.groupby("계약번호_정제")
            idx_latest = grp_u["접수일시"].idxmax()

            df_u_sum = temp_sorted.loc[idx_latest].copy()
            df_u_sum["접수건수"] = grp_u.size().reindex(df_u_sum["계약번호_정제"]).values

            cols_u = [
                "계약번호_정제", "상호", "관리지사", "구역담당자_통합",
                "리스크등급", "경과일수", "접수건수",
                "VOC유형", "VOC유형소", "등록내용", "설치주소_표시",
            ]
            cols_u = [c for c in cols_u if c in df_u_sum.columns]

            st.dataframe(style_risk(df_u_sum[cols_u]), use_container_width=True, height=420)

            st.markdown("### 📂 계약 상세 보기")

            cn_list = df_u_sum["계약번호_정제"].astype(str).tolist()
            sel_u_cn = st.selectbox("계약 선택", ["(선택)"] + cn_list)

            if sel_u_cn != "(선택)":
                detail = temp_u[temp_u["계약번호_정제"].astype(str)==sel_u_cn]
                detail = detail.sort_values("접수일시", ascending=False)
                st.dataframe(style_risk(detail[display_cols]), use_container_width=True)

# ============================================================
# TAB DRILL — 드릴다운
# ============================================================
with tab_drill:
    st.subheader("🔍 계약별 VOC 드릴다운")

    base = voc_filtered_global.copy()

    match_sel = st.radio("매칭여부", ["전체","매칭(O)","비매칭(X)"], horizontal=True)
    if match_sel != "전체":
        base = base[base["매칭여부"] == match_sel]

    c1, c2 = st.columns([2, 3])

    branch_opts = ["전체"] + sort_branch(base["관리지사"].dropna().unique())
    sel_br_d = c1.radio("지사 선택", branch_opts, horizontal=True)

    df_temp_d = base.copy()
    if sel_br_d != "전체":
        df_temp_d = df_temp_d[df_temp_d["관리지사"]==sel_br_d]

    mgr_opts = ["전체"] + sorted(
        df_temp_d["구역담당자_통합"].dropna().astype(str).unique().tolist()
    )
    sel_mgr_d = c2.radio("담당자 선택", mgr_opts, horizontal=True)

    f1, f2 = st.columns(2)
    q_cnd = f1.text_input("계약번호 검색")
    q_named = f2.text_input("상호 검색")

    drill = base.copy()
    if sel_br_d!="전체": drill = drill[drill["관리지사"]==sel_br_d]
    if sel_mgr_d!="전체": drill = drill[drill["구역담당자_통합"].astype(str)==sel_mgr_d]

    if q_cnd: drill = drill[drill["계약번호_정제"].astype(str).str.contains(q_cnd)]
    if q_named and "상호" in drill: drill = drill[drill["상호"].astype(str).str.contains(q_named)]

    if drill.empty:
        st.info("조건에 맞는 계약 없음.")
        sel_cn = None
    else:
        d2 = drill.sort_values("접수일시", ascending=False)
        grp = d2.groupby("계약번호_정제")
        idx = grp["접수일시"].idxmax()
        df_d_sum = d2.loc[idx].copy()
        df_d_sum["접수건수"] = grp.size().reindex(df_d_sum["계약번호_정제"]).values

        st.dataframe(style_risk(df_d_sum), use_container_width=True, height=260)

        cn_list = df_d_sum["계약번호_정제"].astype(str).tolist()

        sel_cn = st.selectbox("VOC 상세 보기", cn_list)

        if sel_cn:
            voc_hist = df_voc[df_voc["계약번호_정제"].astype(str)==sel_cn]
            voc_hist = voc_hist.sort_values("접수일시",ascending=False)

            other_hist = df_other[df_other["계약번호_정제"].astype(str)==sel_cn]

            st.markdown("### VOC 이력")
            st.dataframe(style_risk(voc_hist[display_cols]), use_container_width=True)

            st.markdown("### 기타 출처 이력")
            st.dataframe(other_hist, use_container_width=True)

# ============================================================
# TAB FILTER — 정밀 필터 (안내용)
# ============================================================
with tab_filter:
    st.subheader("🎯 정밀 필터 (VOC유형소 기준)")
    st.info("향후 확장 가능… 현재는 안내용입니다.")

# ============================================================
# TAB ALERT — 담당자 알림
# ============================================================
with tab_alert:
    st.subheader("📨 담당자 알림 발송")

    if contact_df.empty:
        st.warning("contact_map.xlsx 파일이 없어 자동 매핑 불가.")
    else:
        st.success(f"담당자 매핑 {len(contact_df)}명 불러옴")

        df_alert = unmatched_global.groupby("구역담당자_통합")["계약번호_정제"].nunique().reset_index()
        df_alert.columns = ["담당자","비매칭 계약수"]
        df_alert["이메일"] = df_alert["담당자"].apply(lambda x: manager_contacts.get(x,{}).get("email",""))

        st.dataframe(df_alert, use_container_width=True, height=300)

        sel_alert_mgr = st.selectbox("담당자 선택", ["(선택)"] + df_alert["담당자"].tolist())

        if sel_alert_mgr != "(선택)":
            mgr_email = manager_contacts.get(sel_alert_mgr,{}).get("email","")
            custom_email = st.text_input("이메일 주소", value=mgr_email)

            df_mgr_rows = unmatched_global[
                unmatched_global["구역담당자_통합"].astype(str)==sel_alert_mgr
            ]

            df_sorted = df_mgr_rows.sort_values("접수일시", ascending=False)
            grp = df_sorted.groupby("계약번호_정제")
            idx = grp["접수일시"].idxmax()
            df_mgr_latest = df_sorted.loc[idx].copy()

            st.dataframe(df_mgr_latest, use_container_width=True)

            subject = f"[해지VOC] {sel_alert_mgr} 담당자 비매칭 안내"
            body = f"{sel_alert_mgr} 담당자님,\n비매칭 계약이 {len(df_mgr_latest)}건 존재합니다."

            if st.button("📤 이메일 발송"):
                if not custom_email:
                    st.error("이메일 입력 필요")
                else:
                    try:
                        msg = EmailMessage()
                        msg["Subject"] = subject
                        msg["From"] = SMTP_USER
                        msg["To"] = custom_email
                        msg.set_content(body)

                        csv_data = df_mgr_latest.to_csv(index=False).encode("utf-8-sig")
                        msg.add_attachment(csv_data, maintype="application", subtype="octet-stream",
                                           filename=f"비매칭_{sel_alert_mgr}.csv")

                        with smtplib.SMTP(SMTP_HOST,SMTP_PORT) as smtp:
                            smtp.starttls()
                            if SMTP_USER and SMTP_PASSWORD:
                                smtp.login(SMTP_USER,SMTP_PASSWORD)
                            smtp.send_message(msg)

                        st.success("발송 완료.")
                    except Exception as e:
                        st.error(f"발송 실패: {e}")

# ============================================================
# 14. 글로벌 피드백 입력/이력 관리 (선택 계약 sel_cn 기준)
# ============================================================

st.markdown(
    '<div class="section-card"><div class="section-title">📝 해지상담대상 활동등록 (고객대응 / 현장 처리내역)</div>',
    unsafe_allow_html=True,
)

# sel_cn 은 TAB DRILL 에서 선택됨
if "sel_cn" not in locals() or sel_cn is None:
    st.info("🔍 위의 '해지상담대상 활동등록' 탭에서 먼저 계약을 선택해주세요.")
else:

    st.caption(f"📌 현재 선택된 계약번호: **{sel_cn}** 에 대한 처리내역 관리")

    fb_all = st.session_state["feedback_df"]
    fb_sel = fb_all[fb_all["계약번호_정제"].astype(str) == str(sel_cn)].copy()
    fb_sel = fb_sel.sort_values("등록일자", ascending=False)

    # 관리자 삭제 권한
    ADMIN_CODE = "C3A"
    admin_pw = st.text_input("관리자 비밀번호 (삭제 시 필요)", type="password")
    is_admin = admin_pw == ADMIN_CODE

    # ---------------------------
    # 기존 처리내역 리스트
    # ---------------------------
    st.markdown("### 📄 등록된 처리내역")

    if fb_sel.empty:
        st.info("등록된 처리내역이 없습니다.")
    else:
        for idx, row in fb_sel.iterrows():
            with st.container():
                st.markdown('<div class="feedback-item">', unsafe_allow_html=True)

                c1, c2 = st.columns([6, 1])

                with c1:
                    st.write(f"**내용:** {row['고객대응내용']}")
                    st.markdown(
                        f"<div class='feedback-meta'>등록자: {row['등록자']} | 등록일: {row['등록일자']}</div>",
                        unsafe_allow_html=True
                    )
                    if row.get("비고"):
                        st.markdown(
                            f"<div class='feedback-note'>비고: {safe_str(row['비고'])}</div>",
                            unsafe_allow_html=True
                        )

                with c2:
                    if is_admin:
                        if st.button("🗑 삭제", key=f"fb_del_{idx}"):
                            fb_all = fb_all.drop(index=idx)
                            st.session_state["feedback_df"] = fb_all
                            save_feedback(FEEDBACK_PATH, fb_all)
                            st.success("삭제되었습니다!")
                            st.rerun()

                st.markdown("</div>", unsafe_allow_html=True)

    # ---------------------------
    # 새로운 처리내역 입력
    # ---------------------------
    st.markdown("### ➕ 새 처리내용 등록")

    new_content = st.text_area("고객대응 / 현장 처리내용", key="new_fb_content")
    new_writer = st.text_input("등록자", key="new_fb_writer")
    new_note = st.text_input("비고", key="new_fb_note")

    if st.button("등록하기", key="btn_add_feedback"):
        if not new_content.strip():
            st.warning("내용을 입력해주세요.")
        elif not new_writer.strip():
            st.warning("등록자를 입력해주세요.")
        else:
            new_row = {
                "계약번호_정제": sel_cn,
                "고객대응내용": new_content,
                "등록자": new_writer,
                "등록일자": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                "비고": new_note,
            }
            fb_all = pd.concat([fb_all, pd.DataFrame([new_row])], ignore_index=True)

            st.session_state["feedback_df"] = fb_all
            save_feedback(FEEDBACK_PATH, fb_all)

            st.success("등록 완료!")
            st.rerun()

st.markdown("</div>", unsafe_allow_html=True)

# ============================================================
# 15. 담당자 알림 발송(베타)
# ============================================================
with tab_alert:
    st.subheader("📨 담당자 알림 발송 (베타)")

    st.markdown(
        """
        비매칭(X) 계약을 **담당자별로 분리하여 이메일로 발송**할 수 있습니다.  
        담당자 매핑 파일(**contact_map.xlsx**)을 통해 이메일 자동 매칭됩니다.
        """
    )

    # 1) 담당자 파일 체크
    if contact_df.empty:
        st.warning("⚠ contact_map.xlsx 파일이 없어 이메일 자동 매핑이 비활성화되었습니다.")
        st.info("이메일 주소를 직접 입력하여 발송할 수 있습니다.")
        manager_contacts = {}  # fallback
    else:
        st.success(f"담당자 매핑 파일 로드 완료 — 총 {len(contact_df)}명")

    # 현재 비매칭 데이터
    unmatched_alert = unmatched_global.copy()
    grouped = unmatched_alert.groupby("구역담당자_통합")

    st.markdown("### 📧 담당자별 비매칭 계약 수 요약")

    alert_list = []
    for mgr, g in grouped:
        mgr = safe_str(mgr)
        if not mgr:
            continue
        count = g["계약번호_정제"].nunique()
        email = manager_contacts.get(mgr, {}).get("email", "")
        alert_list.append([mgr, email, count])

    alert_df = pd.DataFrame(alert_list, columns=["담당자", "이메일", "비매칭 계약수"])
    st.dataframe(alert_df, use_container_width=True, height=260)

    st.markdown("---")

    # 2) 개별 발송 UI
    st.markdown("### ✉ 개별 발송 (계약번호 중복 제거 + VOC 핵심정보 포함)")

    sel_mgr_alert = st.selectbox(
        "담당자 선택",
        options=["(선택)"] + alert_df["담당자"].tolist(),
        key="alert_mgr",
    )

    if sel_mgr_alert != "(선택)":

        # 기본 이메일
        registered_email = manager_contacts.get(sel_mgr_alert, {}).get("email", "")
        st.write(f"📮 등록된 이메일: **{registered_email or '(없음)'}**")

        # 사용자 수정 가능
        custom_email = st.text_input("이메일 직접 입력 또는 수정", value=registered_email)

        df_mgr = unmatched_alert[
            unmatched_alert["구역담당자_통합"].astype(str) == sel_mgr_alert
        ].copy()

        # 계약번호 중복 제거 → 최신 VOC 1건만 정제
        if not df_mgr.empty:
            df_mgr_sorted = df_mgr.sort_values("접수일시", ascending=False)
            grp_mgr = df_mgr_sorted.groupby("계약번호_정제")
            idx_latest_mgr = grp_mgr["접수일시"].idxmax()
            df_mgr_latest = df_mgr_sorted.loc[idx_latest_mgr].copy()
        else:
            df_mgr_latest = df_mgr

        st.write(f"🔍 발송 예정 유니크 계약 수: **{len(df_mgr_latest)}건**")

        # ----------------------
        # 미리보기 테이블
        # ----------------------
        if not df_mgr_latest.empty:
            preview_cols = [
                "계약번호_정제",
                "상호",
                "관리지사",
                "구역담당자_통합",
                "VOC유형",
                "VOC유형소",
                "등록내용",
                "설치주소_표시",
                "리스크등급",
                "경과일수",
            ]
            preview_cols = [c for c in preview_cols if c in df_mgr_latest.columns]

            st.dataframe(
                df_mgr_latest[preview_cols],
                use_container_width=True,
                height=260,
            )
        else:
            st.info("해당 담당자에게 비매칭 계약이 없습니다.")

        subject = f"[해지VOC] {sel_mgr_alert} 담당자 비매칭 계약 알림"
        body = (
            f"{sel_mgr_alert} 담당자님,\n\n"
            f"비매칭 해지 VOC가 {len(df_mgr_latest)}건 접수되었습니다.\n"
            "첨부된 CSV 파일을 확인하시기 바랍니다.\n\n"
            "- 해지VOC 시스템 -"
        )

        # ----------------------
        # 이메일 전송 버튼
        # ----------------------
        if st.button("📤 이메일 발송하기", key="send_alert_email"):

            if not custom_email:
                st.error("이메일 주소를 입력해주세요.")
            elif df_mgr_latest.empty:
                st.error("전송할 데이터가 없습니다.")
            elif not SMTP_USER or not SMTP_PASSWORD:
                st.error("SMTP 설정(secrets)이 누락되었습니다.")
            else:
                try:
                    # CSV 생성
                    export_cols = [
                        "계약번호_정제",
                        "상호",
                        "관리지사",
                        "구역담당자_통합",
                        "VOC유형",
                        "VOC유형소",
                        "등록내용",
                        "설치주소_표시",
                        "리스크등급",
                        "경과일수",
                    ]
                    export_cols = [c for c in export_cols if c in df_mgr_latest.columns]

                    csv_bytes = (
                        df_mgr_latest[export_cols]
                        .sort_values("리스크등급")
                        .to_csv(index=False)
                        .encode("utf-8-sig")
                    )

                    # 이메일 구성
                    msg = EmailMessage()
                    msg["Subject"] = subject
                    msg["From"] = f"{SENDER_NAME} <{SMTP_USER}>"
                    msg["To"] = custom_email
                    msg.set_content(body)

                    # 첨부파일 추가
                    msg.add_attachment(
                        csv_bytes,
                        maintype="application",
                        subtype="octet-stream",
                        filename=f"비매칭계약_{sel_mgr_alert}.csv",
                    )

                    # SMTP 전송
                    with smtplib.SMTP(SMTP_HOST, SMTP_PORT) as smtp:
                        smtp.starttls()
                        smtp.login(SMTP_USER, SMTP_PASSWORD)
                        smtp.send_message(msg)

                    st.success(f"✅ 이메일 발송 성공 → {custom_email}")

                except smtplib.SMTPAuthenticationError:
                    st.error("❌ SMTP 계정 인증 실패 (535). Gmail 앱 비밀번호를 다시 확인해주세요.")
                except Exception as e:
                    st.error(f"❌ 이메일 전송 중 오류 발생: {e}")


