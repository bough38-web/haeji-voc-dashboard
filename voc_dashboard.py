 # ====================================================
# PART 1 — 기본 설정, 스타일, 유틸, 데이터 로딩
# ====================================================
import os
from datetime import datetime, date
import smtplib
from email.message import EmailMessage

import numpy as np
import pandas as pd
import streamlit as st

# Plotly (없으면 fallback)
try:
    import plotly.express as px
    HAS_PLOTLY = True
except:
    HAS_PLOTLY = False

# ----------------------------------------------------
# 페이지 설정 + CSS (라이트 모드 고정)
# ----------------------------------------------------
st.set_page_config(page_title="해지 VOC 종합 대시보드", layout="wide")

st.markdown("""
<style>
html, body, .stApp {
    background-color:#f5f5f7 !important;
    color:#111827 !important;
    font-family:-apple-system,BlinkMacSystemFont,"Segoe UI",sans-serif;
}
.block-container {padding-top:0.8rem !important;}
section[data-testid="stSidebar"] {background:#fafafa !important;}

.branch-grid {
    display:grid;
    grid-template-columns:repeat(4,1fr);
    gap:14px;
}
@media (max-width:1200px){.branch-grid{grid-template-columns:repeat(2,1fr);} }
@media (max-width:700px){.branch-grid{grid-template-columns:repeat(1,1fr);} }

.branch-card {
    background:#fff;
    border:1px solid #e5e7eb;
    border-radius:14px;
    padding:1rem;
    box-shadow:0 3px 8px rgba(0,0,0,0.05);
    transition:0.12s ease;
}
.branch-card:hover {
    transform:translateY(-4px);
    box-shadow:0 8px 18px rgba(0,0,0,0.08);
}
.badge-high {background:#ef4444;color:white;padding:2px 7px;border-radius:7px;}
.badge-medium{background:#f59e0b;color:white;padding:2px 7px;border-radius:7px;}
.badge-low {background:#3b82f6;color:white;padding:2px 7px;border-radius:7px;}

.js-plotly-plot .plotly {background-color:transparent !important;}
</style>
""", unsafe_allow_html=True)

# ----------------------------------------------------
# SMTP
# ----------------------------------------------------
if "SMTP_HOST" in st.secrets:
    SMTP_HOST = st.secrets["SMTP_HOST"]
    SMTP_PORT = int(st.secrets["SMTP_PORT"])
    SMTP_USER = st.secrets["SMTP_USER"]
    SMTP_PASSWORD = st.secrets["SMTP_PASSWORD"]
    SENDER_NAME = st.secrets["SENDER_NAME"]
else:
    from dotenv import load_dotenv
    load_dotenv()
    SMTP_HOST = os.getenv("SMTP_HOST", "")
    SMTP_PORT = int(os.getenv("SMTP_PORT", "587"))
    SMTP_USER = os.getenv("SMTP_USER", "")
    SMTP_PASSWORD = os.getenv("SMTP_PASSWORD", "")
    SENDER_NAME = os.getenv("SENDER_NAME", "해지VOC 관리자")

MERGED_PATH = "merged.xlsx"
CONTACT_PATH = "contact_map.xlsx"
FEEDBACK_PATH = "feedback.csv"

# ----------------------------------------------------
# 유틸
# ----------------------------------------------------
def safe_str(x):
    return "" if pd.isna(x) else str(x).strip()

def detect_column(df, keywords):
    for k in keywords:
        if k in df.columns:
            return k
    for col in df.columns:
        for k in keywords:
            if k.lower() in col.lower():
                return col
    return None

# ====================================================
# PART 2 — VOC 데이터 로딩 / 정제 / 매칭 / 리스크 계산
# ====================================================

@st.cache_data
def load_voc_data(path: str):
    if not os.path.exists(path):
        st.error(f"❌ '{path}' 파일을 찾을 수 없습니다.")
        return pd.DataFrame()
    df = pd.read_excel(path)

    # 계약번호 정제
    if "계약번호" in df.columns:
        df["계약번호_정제"] = (
            df["계약번호"].astype(str)
            .str.replace(r"[^0-9A-Za-z]", "", regex=True)
            .str.strip()
        )
    else:
        df["계약번호_정제"] = ""

    # 출처 정제
    if "출처" in df.columns:
        df["출처"] = df["출처"].replace({"고객리스트": "해지시설"})

    # 접수일시 → datetime
    if "접수일시" in df.columns:
        df["접수일시"] = pd.to_datetime(df["접수일시"], errors="coerce")

    return df


@st.cache_data
def load_contact_map(path: str):
    if not os.path.exists(path):
        st.warning(f"⚠ 담당자 매핑 파일 '{path}' 없음 → 이메일 직접 입력 기능만 사용 가능")
        return pd.DataFrame(), {}

    df = pd.read_excel(path)

    name_col = detect_column(df, ["구역담당자", "담당자", "성명", "이름"])
    email_col = detect_column(df, ["이메일", "메일", "email"])
    phone_col = detect_column(df, ["휴대폰", "전화", "연락처"])

    if not (name_col and email_col):
        st.warning("⚠ 담당자 / 이메일 컬럼을 찾지 못했습니다.")
        return df, {}

    df = df[[name_col, email_col] + ([phone_col] if phone_col else [])].copy()
    rename_map = {name_col: "구역담당자_통합", email_col: "이메일"}
    if phone_col:
        rename_map[phone_col] = "휴대폰"
    df.rename(columns=rename_map, inplace=True)

    contacts = {}
    for _, r in df.iterrows():
        n = safe_str(r["구역담당자_통합"])
        if n:
            contacts[n] = {
                "email": safe_str(r.get("이메일", "")),
                "phone": safe_str(r.get("휴대폰", "")),
            }
    return df, contacts


# ---------------- 실제 데이터 로드 ----------------
df = load_voc_data(MERGED_PATH)
if df.empty:
    st.stop()

contact_df, manager_contacts = load_contact_map(CONTACT_PATH)


# ----------------------------------------------------
# 영업구역 / 담당자 통합
# ----------------------------------------------------
def pick_manager(row):
    for c in ["구역담당자", "담당자", "처리자"]:
        if c in row and pd.notna(row[c]) and str(row[c]).strip():
            return row[c]
    return ""

df["구역담당자_통합"] = df.apply(pick_manager, axis=1)


# ----------------------------------------------------
# 출처 분리 및 매칭여부
# ----------------------------------------------------
df_voc = df[df.get("출처") == "해지VOC"].copy()
df_other = df[df.get("출처") != "해지VOC"].copy()

other_union = set(df_other["계약번호_정제"].dropna())

df_voc["매칭여부"] = df_voc["계약번호_정제"].apply(
    lambda x: "매칭(O)" if x in other_union else "비매칭(X)"
)


# ----------------------------------------------------
# 월정료 정제 + 구간 나누기
# ----------------------------------------------------
def parse_fee(val):
    if pd.isna(val): return np.nan
    s = str(val).replace(",", "")
    s2 = "".join(ch for ch in s if ch.isdigit())
    if not s2: return np.nan
    v = float(s2)
    if v >= 200000: v /= 10
    return v

fee_col = None
for c in ["시설_KTT월정료(조정)", "KTT월정료(조정)"]:
    if c in df_voc.columns:
        fee_col = c
        break

if fee_col:
    df_voc["월정료_수치"] = df_voc[fee_col].apply(parse_fee)
else:
    df_voc["월정료_수치"] = np.nan

def band(v):
    if pd.isna(v): return "미기재"
    return "10만 이상" if v >= 100000 else "10만 미만"

df_voc["월정료구간"] = df_voc["월정료_수치"].apply(band)


# ----------------------------------------------------
# 리스크 등급 / 경과일 계산
# ----------------------------------------------------
today = date.today()
def compute_risk(row):
    dt = row.get("접수일시")
    if pd.isna(dt): return np.nan, "LOW"
    if not isinstance(dt, (pd.Timestamp, datetime)):
        dt = pd.to_datetime(dt, errors="coerce")
    if pd.isna(dt): return np.nan, "LOW"

    days = (today - dt.date()).days
    if days <= 3: level = "HIGH"
    elif days <= 10: level = "MEDIUM"
    else: level = "LOW"
    return days, level

df_voc["경과일수"], df_voc["리스크등급"] = zip(
    *df_voc.apply(compute_risk, axis=1)
)

df_unmatched = df_voc[df_voc["매칭여부"] == "비매칭(X)"]

# ====================================================
# PART 3 — 글로벌 필터 / KPI / 고급 시각화
# ====================================================

st.sidebar.title("🔧 글로벌 필터")

# 날짜 필터
if "접수일시" in df_voc.columns and df_voc["접수일시"].notna().any():
    d_min = df_voc["접수일시"].min().date()
    d_max = df_voc["접수일시"].max().date()
    dr = st.sidebar.date_input(
        "📅 접수일자 범위",
        value=(d_min, d_max),
        min_value=d_min,
        max_value=d_max,
        key="flt_date"
    )
else:
    dr = None

# 지사 필터
branches_all = sort_branch(df_voc["관리지사"].dropna().unique())
sel_branches = st.sidebar.multiselect(
    "🏢 관리지사",
    options=branches_all,
    default=branches_all,
    key="flt_branch"
)

# 리스크 필터
risk_opts = ["HIGH", "MEDIUM", "LOW"]
sel_risk = st.sidebar.multiselect(
    "⚠ 리스크등급",
    options=risk_opts,
    default=risk_opts,
    key="flt_risk"
)

# 매칭 필터
sel_match = st.sidebar.multiselect(
    "🔗 매칭 여부",
    options=["매칭(O)", "비매칭(X)"],
    default=["매칭(O)", "비매칭(X)"],
    key="flt_match"
)

# 월정료 구간
fee_global = st.sidebar.radio(
    "💰 월정료 구간",
    ["전체", "10만 미만", "10만 이상"],
    index=0,
    key="flt_fee_band"
)

st.sidebar.markdown("---")
st.sidebar.caption("※ 이 필터는 모든 탭에 공통 적용됩니다.")

# ----------------------------------------------------
# 1. 글로벌 필터 적용
# ----------------------------------------------------
voc_filtered_global = df_voc.copy()

# 날짜
if dr:
    s, e = dr
    voc_filtered_global = voc_filtered_global[
        (voc_filtered_global["접수일시"] >= pd.to_datetime(s)) &
        (voc_filtered_global["접수일시"] < pd.to_datetime(e) + pd.Timedelta(days=1))
    ]

# 지사
voc_filtered_global = voc_filtered_global[
    voc_filtered_global["관리지사"].isin(sel_branches)
]

# 리스크
voc_filtered_global = voc_filtered_global[
    voc_filtered_global["리스크등급"].isin(sel_risk)
]

# 매칭
voc_filtered_global = voc_filtered_global[
    voc_filtered_global["매칭여부"].isin(sel_match)
]

# 월정료
if fee_global == "10만 이상":
    voc_filtered_global = voc_filtered_global[
        voc_filtered_global["월정료_수치"] >= 100000
    ]
elif fee_global == "10만 미만":
    voc_filtered_global = voc_filtered_global[
        (voc_filtered_global["월정료_수치"] < 100000) &
        voc_filtered_global["월정료_수치"].notna()
    ]

# 비매칭 데이터
unmatched_global = voc_filtered_global[
    voc_filtered_global["매칭여부"] == "비매칭(X)"
]

# ----------------------------------------------------
# 2. KPI 카드
# ----------------------------------------------------
st.markdown("## 📊 해지 VOC 종합 대시보드")

k1, k2, k3, k4 = st.columns(4)

k1.metric("총 VOC 건수(행)", f"{len(voc_filtered_global):,}")
k2.metric("VOC 계약 수", f"{voc_filtered_global['계약번호_정제'].nunique():,}")
k3.metric("비매칭 계약 수", f"{unmatched_global['계약번호_정제'].nunique():,}")
k4.metric(
    "매칭 계약 수",
    f"{voc_filtered_global[voc_filtered_global['매칭여부']=='매칭(O)']['계약번호_정제'].nunique():,}"
)

st.markdown("---")

# ----------------------------------------------------
# 3. 탭 구성
# ----------------------------------------------------
tab_viz, tab_all, tab_unmatched, tab_drill, tab_alert = st.tabs(
    [
        "📊 지사/담당자 시각화",
        "📘 VOC 전체",
        "🧯 해지방어 활동시설",
        "🔍 계약별 상세조회",
        "📨 담당자 알림"
    ]
)

# ====================================================
# TAB VIZ — 지사 / 담당자 시각화 (완성본)
# ====================================================
with tab_viz:
    st.subheader("📊 지사 / 담당자별 비매칭 리스크 현황")

    # 비매칭 데이터 없을 때
    if unmatched_global.empty:
        st.info("현재 조건에서 비매칭(X) 데이터가 없습니다.")
        st.stop()

    clean_df = unmatched_global.dropna(subset=["관리지사"])

    # =========================================================
    # 0) CSS (지사 카드 스타일)
    # =========================================================
    st.markdown("""
        <style>
        .branch-grid {
            display: grid;
            grid-template-columns: repeat(4, 1fr);
            gap: 14px;
        }
        @media (max-width:1200px){
            .branch-grid { grid-template-columns: repeat(2, 1fr); }
        }
        @media (max-width:700px){
            .branch-grid { grid-template-columns: repeat(1, 1fr); }
        }

        .branch-card {
            background: #ffffff;
            border-radius: 14px;
            padding: 1rem;
            border: 1px solid #e5e7eb;
            box-shadow: 0 3px 8px rgba(0,0,0,0.05);
            transition: 0.12s ease;
        }
        .branch-card:hover {
            transform: translateY(-4px);
            box-shadow: 0 8px 18px rgba(0,0,0,0.08);
        }
        .branch-card-header {
            font-size: 1.05rem;
            font-weight: 600;
            margin-bottom: 4px;
            color: #111827;
        }
        .branch-card-sub {
            font-size: 0.82rem;
            margin-top: 4px;
            color: #374151;
        }

        .badge-high {
            color: white;
            background: #ef4444;
            padding: 2px 7px;
            border-radius: 7px;
            font-size: 0.75rem;
        }
        .badge-medium {
            color: white;
            background: #f59e0b;
            padding: 2px 7px;
            border-radius: 7px;
            font-size: 0.75rem;
        }
        .badge-low {
            color: white;
            background: #3b82f6;
            padding: 2px 7px;
            border-radius: 7px;
            font-size: 0.75rem;
        }
        </style>
    """, unsafe_allow_html=True)

    # =========================================================
    # 1) 지사 요약 카드
    # =========================================================
    st.markdown("### 🏢 지사별 비매칭 요약")

    branch_stats = (
        clean_df.groupby("관리지사")
        .agg(
            계약수=("계약번호_정제", "nunique"),
            HIGH=("리스크등급", lambda s: (s == "HIGH").sum()),
            MEDIUM=("리스크등급", lambda s: (s == "MEDIUM").sum()),
            LOW=("리스크등급", lambda s: (s == "LOW").sum()),
        )
    )
    branch_stats = branch_stats.reindex(BRANCH_ORDER).dropna(how="all")

    html = '<div class="branch-grid">'
    for branch, row in branch_stats.iterrows():
        html += f"""
        <div class="branch-card">
            <div class="branch-card-header">{branch}</div>
            <div class="branch-card-sub">계약 {int(row['계약수'])}건</div>

            <span class="badge-high">HIGH {int(row['HIGH'])}</span>&nbsp;
            <span class="badge-medium">MED {int(row['MEDIUM'])}</span>&nbsp;
            <span class="badge-low">LOW {int(row['LOW'])}</span>
        </div>
        """
    html += "</div>"
    st.markdown(html, unsafe_allow_html=True)

    st.markdown("---")

    # =========================================================
    # 2) 필터 (지사 / 담당자)
    # =========================================================
    f1, f2, f3 = st.columns([1.2, 1.2, 1])

    branch_opts = ["전체"] + sort_branch(clean_df["관리지사"].unique())
    sel_branch = f1.selectbox("지사 선택", branch_opts)

    df_mgr_scope = (
        clean_df if sel_branch == "전체"
        else clean_df[clean_df["관리지사"] == sel_branch]
    )

    mgr_list = (
        df_mgr_scope["구역담당자_통합"]
        .dropna()
        .astype(str)
        .unique()
        .tolist()
    )
    mgr_list = sorted([m for m in mgr_list if m.strip() != ""])

    sel_mgr = f2.selectbox("담당자 선택", ["(전체)"] + mgr_list)

    scope_df = (
        df_mgr_scope if sel_mgr == "(전체)"
        else df_mgr_scope[df_mgr_scope["구역담당자_통합"].astype(str) == sel_mgr]
    )
    f3.metric("선택 범위 계약 수", f"{scope_df['계약번호_정제'].nunique():,}")

    st.caption(
        f"HIGH {(scope_df['리스크등급']=='HIGH').sum()}건 / "
        f"MEDIUM {(scope_df['리스크등급']=='MEDIUM').sum()}건 / "
        f"LOW {(scope_df['리스크등급']=='LOW').sum()}건"
    )

    st.markdown("---")

    # =========================================================
    # 3) 지사별 리스크 STACK BAR
    # =========================================================
    st.markdown("### 🧱 지사별 리스크 분포")

    risk_by_branch = (
        clean_df.groupby(["관리지사", "리스크등급"])["계약번호_정제"]
        .nunique()
        .reset_index()
    )

    risk_by_branch["관리지사"] = pd.Categorical(
        risk_by_branch["관리지사"], categories=BRANCH_ORDER, ordered=True
    )
    risk_by_branch["리스크등급"] = pd.Categorical(
        risk_by_branch["리스크등급"], categories=["HIGH", "MEDIUM", "LOW"], ordered=True
    )

    risk_by_branch = risk_by_branch.sort_values(["관리지사", "리스크등급"])

    fig_stack = px.bar(
        risk_by_branch,
        x="관리지사",
        y="계약번호_정제",
        color="리스크등급",
        barmode="stack",
        color_discrete_map={
            "HIGH": "#ef4444",
            "MEDIUM": "#f59e0b",
            "LOW": "#3b82f6",
        }
    )
    fig_stack.update_layout(
        height=360,
        margin=dict(l=10, r=10, t=40, b=40),
        xaxis_title="지사",
        yaxis_title="계약 수"
    )
    st.plotly_chart(fig_stack, use_container_width=True)

    st.markdown("---")

    # =========================================================
    # 4) 담당자 TOP15 + 전체 리스크 도넛
    # =========================================================
    g1, g2 = st.columns(2)

    # 담당자 TOP 15
    with g1:
        st.markdown("#### 👤 담당자별 비매칭 TOP 15")
        scope_df2 = (
            clean_df if sel_branch == "전체"
            else clean_df[clean_df["관리지사"] == sel_branch]
        )

        top15 = (
            scope_df2.groupby("구역담당자_통합")["계약번호_정제"]
            .nunique()
            .sort_values(ascending=False)
            .head(15)
            .reset_index()
        )

        fig_top = px.bar(
            top15,
            x="구역담당자_통합",
            y="계약번호_정제",
            text="계약번호_정제",
            color="계약번호_정제",
            color_continuous_scale="Blues"
        )
        fig_top.update_traces(textposition="outside")
        fig_top.update_layout(height=330, xaxis_tickangle=-40)
        st.plotly_chart(fig_top, use_container_width=True)

    # 전체 리스크 도넛
    with g2:
        st.markdown("#### 🍩 전체 비매칭 리스크 비율")

        rc = clean_df["리스크등급"].value_counts().reindex(["HIGH", "MEDIUM", "LOW"]).fillna(0)
        rc_df = rc.reset_index()
        rc_df.columns = ["리스크등급", "건수"]

        fig_pie = px.pie(
            rc_df,
            names="리스크등급",
            values="건수",
            hole=0.45,
            color="리스크등급",
            color_discrete_map={
                "HIGH": "#ef4444",
                "MEDIUM": "#f59e0b",
                "LOW": "#3b82f6",
            }
        )
        fig_pie.update_layout(height=330)
        st.plotly_chart(fig_pie, use_container_width=True)

    st.markdown("---")

    # =========================================================
    # 5) 일자별 추이 + 담당자 리스크 차트
    # =========================================================
    t1, t2 = st.columns(2)

    # 일자별 추이
    with t1:
        st.markdown("#### 📈 일별 비매칭 추이")
        if "접수일시" in clean_df:
            df_trend = clean_df.assign(접수일=clean_df["접수일시"].dt.date)
            trend = (
                df_trend.groupby("접수일")["계약번호_정제"]
                .nunique()
                .reset_index()
            )
            fig_trend = px.line(trend, x="접수일", y="계약번호_정제")
            fig_trend.update_layout(height=260)
            st.plotly_chart(fig_trend, use_container_width=True)

    # 담당자 리스크 요약
    with t2:
        st.markdown("#### 🌐 선택 담당자 리스크 비율")
        if sel_mgr == "(전체)":
            st.info("담당자를 선택하면 표시됩니다.")
        else:
            mgr_df = clean_df[
                clean_df["구역담당자_통합"].astype(str) == sel_mgr
            ]
            rc_mgr = mgr_df["리스크등급"].value_counts().reindex(["HIGH", "MEDIUM", "LOW"]).fillna(0)

            fig_mgr = px.bar(
                rc_mgr.reset_index(),
                x="index",
                y="리스크등급",
                text="리스크등급",
                color="index",
                color_discrete_map={
                    "HIGH": "#ef4444",
                    "MEDIUM": "#f59e0b",
                    "LOW": "#3b82f6",
                }
            )
            fig_mgr.update_traces(textposition="outside")
            fig_mgr.update_layout(height=260)
            st.plotly_chart(fig_mgr, use_container_width=True)

# ====================================================
# TAB ALL — VOC 전체 (계약번호 기준 요약)
# ====================================================
with tab_all:
    st.subheader("📘 VOC 전체 (계약번호 기준 조회)")

    # ------------------------------
    # 1) 상단 필터 (지사 / 담당자)
    # ------------------------------
    col_b1, col_b2 = st.columns([2, 3])

    branches_tab = ["전체"] + sort_branch(voc_filtered_global["관리지사"].dropna().unique())

    sel_branch_tab = col_b1.radio(
        "지사 선택",
        options=branches_tab,
        horizontal=True,
        key="all_branch",
    )

    df_mgr_scope_tab = voc_filtered_global.copy()
    if sel_branch_tab != "전체":
        df_mgr_scope_tab = df_mgr_scope_tab[df_mgr_scope_tab["관리지사"] == sel_branch_tab]

    mgr_opts_tab = (
        ["전체"]
        + sorted(
            df_mgr_scope_tab["구역담당자_통합"]
            .dropna()
            .astype(str)
            .unique()
            .tolist()
        )
    )

    sel_mgr_tab = col_b2.radio(
        "담당자 선택",
        options=mgr_opts_tab,
        horizontal=True,
        key="all_mgr",
    )

    # ------------------------------
    # 2) 상세 텍스트 검색 필터
    # ------------------------------
    s1, s2, s3 = st.columns(3)

    q_cn = s1.text_input("계약번호 검색(부분)", key="all_cn")
    q_name = s2.text_input("상호 검색(부분)", key="all_name")
    q_addr = s3.text_input("주소 검색(부분)", key="all_addr")

    df_all = voc_filtered_global.copy()

    if sel_branch_tab != "전체":
        df_all = df_all[df_all["관리지사"] == sel_branch_tab]
    if sel_mgr_tab != "전체":
        df_all = df_all[df_all["구역담당자_통합"].astype(str) == sel_mgr_tab]

    if q_cn:
        df_all = df_all[df_all["계약번호_정제"].astype(str).str.contains(q_cn.strip())]

    if q_name and "상호" in df_all.columns:
        df_all = df_all[df_all["상호"].astype(str).str.contains(q_name.strip())]

    if q_addr:
        cond = None
        if "설치주소_표시" in df_all.columns:
            cond = df_all["설치주소_표시"].astype(str).str.contains(q_addr.strip())
        else:
            for col in address_cols:
                if col in df_all.columns:
                    c = df_all[col].astype(str).str.contains(q_addr.strip())
                    cond = c if cond is None else (cond | c)
        if cond is not None:
            df_all = df_all[cond]

    # ------------------------------
    # 3) 계약번호 기준 대표 1건 요약 정리
    # ------------------------------
    if df_all.empty:
        st.info("조건에 맞는 VOC 데이터가 없습니다.")
    else:
        df_all_sorted = df_all.sort_values("접수일시", ascending=False)

        grp = df_all_sorted.groupby("계약번호_정제")
        idx_latest = grp["접수일시"].idxmax()

        df_summary = df_all_sorted.loc[idx_latest].copy()
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
            "VOC유형",
            "VOC유형소",
            "등록내용",
            "설치주소_표시",
            fee_raw_col if fee_raw_col else None,
            "계약상태(중)",
            "서비스(소)",
        ]
        summary_cols = [c for c in summary_cols if c and c in df_summary.columns]
        summary_cols = filter_valid_columns(summary_cols, df_summary)

        st.markdown(f"📌 표시 계약 수: **{len(df_summary):,} 건**")
        st.dataframe(
            style_risk(df_summary[summary_cols]),
            use_container_width=True,
            height=520,
        )

# ====================================================
# TAB UNMATCHED — 해지방어 활동시설(비매칭)
# ====================================================
with tab_unmatched:
    st.subheader("🧯 해지방어 활동시설 (비매칭, 계약번호 기준)")

    st.caption("비매칭(X) = 해지 VOC 접수 후 시스템상 활동내역이 확인되지 않은 시설")

    if unmatched_global.empty:
        st.info("현재 글로벌 필터 조건에서 비매칭(X) 계약이 없습니다.")
        st.stop()

    # -----------------------------------------------
    # 1) 상단 필터 (지사 / 담당자)
    # -----------------------------------------------
    col1, col2 = st.columns([2, 3])

    branch_opts_u = ["전체"] + sort_branch(
        unmatched_global["관리지사"].dropna().unique()
    )
    sel_branch_u = col1.radio(
        "지사 선택",
        options=branch_opts_u,
        horizontal=True,
        key="un_branch",
    )

    df_mgr_u = unmatched_global.copy()
    if sel_branch_u != "전체":
        df_mgr_u = df_mgr_u[df_mgr_u["관리지사"] == sel_branch_u]

    mgr_opts_u = (
        ["전체"]
        + sorted(
            df_mgr_u["구역담당자_통합"]
            .dropna()
            .astype(str)
            .unique()
            .tolist()
        )
    )

    sel_mgr_u = col2.radio(
        "담당자 선택",
        options=mgr_opts_u,
        horizontal=True,
        key="un_mgr",
    )

    # -----------------------------------------------
    # 2) 텍스트 검색 (계약 / 상호)
    # -----------------------------------------------
    s1, s2 = st.columns(2)
    q_cn_u = s1.text_input("계약번호 검색(부분)", key="un_cn")
    q_name_u = s2.text_input("상호 검색(부분)", key="un_name")

    df_u = unmatched_global.copy()

    if sel_branch_u != "전체":
        df_u = df_u[df_u["관리지사"] == sel_branch_u]

    if sel_mgr_u != "전체":
        df_u = df_u[
            df_u["구역담당자_통합"].astype(str) == sel_mgr_u
        ]

    if q_cn_u:
        df_u = df_u[
            df_u["계약번호_정제"].astype(str).str.contains(q_cn_u.strip())
        ]
    if q_name_u and "상호" in df_u.columns:
        df_u = df_u[df_u["상호"].astype(str).str.contains(q_name_u.strip())]

    # -----------------------------------------------
    # 3) 계약번호 대표행(최신 VOC 1건) 요약
    # -----------------------------------------------
    if df_u.empty:
        st.info("조건에 맞는 해지방어 활동시설(비매칭) 계약이 없습니다.")
        st.stop()

    df_u_sorted = df_u.sort_values("접수일시", ascending=False)

    grp_u = df_u_sorted.groupby("계약번호_정제")
    idx_latest_u = grp_u["접수일시"].idxmax()

    df_u_summary = df_u_sorted.loc[idx_latest_u].copy()
    df_u_summary["접수건수"] = grp_u.size().reindex(df_u_summary["계약번호_정제"]).values

    summary_cols_u = [
        "계약번호_정제",
        "상호",
        "관리지사",
        "구역담당자_통합",
        "리스크등급",
        "경과일수",
        "접수건수",
        "VOC유형",
        "VOC유형소",
        "등록내용",
        "설치주소_표시",
        fee_raw_col if fee_raw_col else None,
        "계약상태(중)",
        "서비스(소)",
    ]
    summary_cols_u = [c for c in summary_cols_u if c in df_u_summary.columns]
    summary_cols_u = filter_valid_columns(summary_cols_u, df_u_summary)

    st.markdown(f"⚠ 비매칭 계약 수: **{len(df_u_summary):,} 건**")

    st.dataframe(
        style_risk(df_u_summary[summary_cols_u]),
        use_container_width=True,
        height=420,
    )

    # -----------------------------------------------
    # 4) 선택 계약 상세 VOC 이력
    # -----------------------------------------------
    st.markdown("### 📂 선택한 계약번호 VOC 상세 이력")

    contract_list = df_u_summary["계약번호_정제"].astype(str).tolist()

    sel_cn_u = st.selectbox(
        "상세 볼 계약 선택",
        options=["(선택)"] + contract_list,
        key="un_detail_select",
    )

    if sel_cn_u == "(선택)":
        st.stop()

    detail_voc = df_u[df_u["계약번호_정제"].astype(str) == sel_cn_u].copy()
    detail_voc = detail_voc.sort_values("접수일시", ascending=False)

    st.markdown(
        f"#### 🔍 `{sel_cn_u}` VOC 상세 이력 ({len(detail_voc)}건)"
    )
    st.dataframe(
        style_risk(detail_voc[display_cols]),
        use_container_width=True,
        height=350,
    )

    # -----------------------------------------------
    # 5) CSV 다운로드
    # -----------------------------------------------
    st.download_button(
        "📥 선택 계약 VOC 상세 다운로드 (CSV)",
        detail_voc.to_csv(index=False).encode("utf-8-sig"),
        file_name=f"비매칭상세_{sel_cn_u}.csv",
        mime="text/csv",
    )

