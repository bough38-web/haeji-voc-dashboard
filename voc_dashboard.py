import os
from datetime import datetime, date

import numpy as np
import pandas as pd
import streamlit as st

# ----------------------------------------------------
# 0. 기본 설정 (라이트톤)
# ----------------------------------------------------
st.set_page_config(page_title="해지 VOC 종합 대시보드", layout="wide")

st.markdown(
    """
    <style>
    .stApp {
        background-color: #f8fafc;
        color: #111827;
        font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif;
    }
    [data-testid="stHeader"] {
        background-color: #f8fafc;
    }
    .dataframe tbody tr:nth-child(odd) {
        background-color: #f9fafb;
    }
    .dataframe tbody tr:nth-child(even) {
        background-color: #eef2ff;
    }
    h2, h3, h4 {
        margin-top: 0.6rem;
        margin-bottom: 0.3rem;
    }
    </style>
    """,
    unsafe_allow_html=True,
)

# ----------------------------------------------------
# 1. 파일 경로
# ----------------------------------------------------
MERGED_PATH = "merged.xlsx"
FEEDBACK_PATH = "feedback.csv"

# ----------------------------------------------------
# 2. 데이터 로딩
# ----------------------------------------------------
@st.cache_data
def load_voc_data(path: str) -> pd.DataFrame:
    if not os.path.exists(path):
        st.error("❌ 'merged.xlsx' 파일이 존재하지 않습니다. 저장소 루트 위치를 확인해주세요.")
        return pd.DataFrame()

    df = pd.read_excel(path)

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


df = load_voc_data(MERGED_PATH)
if df.empty:
    st.stop()

if "feedback_df" not in st.session_state:
    st.session_state["feedback_df"] = load_feedback(FEEDBACK_PATH)

# ----------------------------------------------------
# 3. 공통 전처리 (지사, 담당자, 월정료 등)
# ----------------------------------------------------
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

BRANCH_ORDER = ["중앙", "강북", "서대문", "고양", "의정부", "남양주", "강릉", "원주"]


def sort_branch(series):
    return sorted(
        [s for s in series if s in BRANCH_ORDER],
        key=lambda x: BRANCH_ORDER.index(x),
    )


# 통합 구역 / 담당자
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

# 주소 컬럼(검색용) : 시설_설치주소 우선
address_cols = []
for col in df.columns:
    if "설치주소" in col or "주소" in col:
        address_cols.append(col)

# KTT 월정료(조정) 파싱 (시설_ 우선)
fee_base_col = None
for cand in ["시설_KTT월정료(조정)", "KTT월정료(조정)"]:
    if cand in df.columns:
        fee_base_col = cand
        break

if fee_base_col is not None:
    def parse_fee(x):
        if pd.isna(x):
            return np.nan
        s = str(x)
        s = s.replace(",", "")
        digits = "".join(ch for ch in s if ch.isdigit())
        if digits == "":
            return np.nan
        try:
            return float(digits)
        except Exception:
            return np.nan

    df["월정료_수치"] = df[fee_base_col].apply(parse_fee)

    # 천단위 콤마 표시
    def format_fee(v):
        if pd.isna(v):
            return np.nan
        try:
            return f"{int(v):,}"
        except Exception:
            return np.nan

    df["월정료_표시"] = df["월정료_수치"].apply(format_fee)

    def fee_band(v):
        if pd.isna(v):
            return "미기재"
        if v >= 100000:
            return "10만 이상"
        return "10만 미만"

    df["월정료구간"] = df["월정료_수치"].apply(fee_band)
else:
    df["월정료_수치"] = np.nan
    df["월정료_표시"] = np.nan
    df["월정료구간"] = "미기재"

# ----------------------------------------------------
# 4. 출처 분리 & 매칭여부
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
# 5. 리스크 계산 (접수일시 안전 처리)
# ----------------------------------------------------
today = date.today()


def compute_risk_from_dt(dt_value):
    if pd.isna(dt_value):
        return np.nan, "LOW"

    if not isinstance(dt_value, (pd.Timestamp, datetime)):
        try:
            dt_value = pd.to_datetime(dt_value)
        except Exception:
            return np.nan, "LOW"

    if pd.isna(dt_value):
        return np.nan, "LOW"

    days = (today - dt_value.date()).days

    if days <= 3:
        level = "HIGH"
    elif days <= 10:
        level = "MEDIUM"
    else:
        level = "LOW"
    return days, level


if "접수일시" in df_voc.columns:
    df_voc["경과일수"], df_voc["리스크등급"] = zip(
        *df_voc["접수일시"].apply(compute_risk_from_dt)
    )
else:
    df_voc["경과일수"] = np.nan
    df_voc["리스크등급"] = "LOW"

# ----------------------------------------------------
# 6. 스타일 함수 & 공통 유틸
# ----------------------------------------------------
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


def filter_nonempty_columns(df_src: pd.DataFrame, cols: list[str]) -> list[str]:
    """해당 df에서 전부 None/NaN 인 컬럼은 제외한 실제 표시용 컬럼 리스트"""
    real_cols: list[str] = []
    for c in cols:
        if c in df_src.columns and df_src[c].notna().any():
            real_cols.append(c)
    return real_cols


# 표시 후보 컬럼 (시설_ 컬럼 그대로 사용)
BASE_COLS_CANDIDATES = [
    "계약번호_정제",
    "상호",
    "관리지사",
    "구역담당자_통합",
    "리스크등급",
    "경과일수",
    "매칭여부",
    "접수건수",
    # 시설 정보
    "시설_설치주소",
    "시설_KTT월정료(조정)",
    "시설_계약상태(중)",
    "시설_서비스(소)",
    # 파생
    "월정료_표시",
    "월정료구간",
]

# ----------------------------------------------------
# 7. 사이드바 글로벌 필터 (날짜/지사/리스크/매칭/월정료)
# ----------------------------------------------------
st.sidebar.title("🔧 글로벌 필터")

# 날짜 필터
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

# 지사 필터
branches_all = sort_branch(df_voc["관리지사"].dropna().unique())
sel_branches = st.sidebar.multiselect(
    "관리지사(복수 선택)",
    options=branches_all,
    default=branches_all,
    key="global_branches",
)

# 리스크 필터
risk_all = ["HIGH", "MEDIUM", "LOW"]
sel_risk = st.sidebar.multiselect(
    "리스크등급",
    options=risk_all,
    default=risk_all,
    key="global_risk",
)

# 매칭여부 필터
match_all = ["매칭(O)", "비매칭(X)"]
sel_match = st.sidebar.multiselect(
    "매칭여부",
    options=match_all,
    default=match_all,
    key="global_match",
)

# 월정료 필터
if "월정료_수치" in df_voc.columns and df_voc["월정료_수치"].notna().any():
    fee_filter = st.sidebar.radio(
        "월정료 구간",
        options=["전체", "10만 미만", "10만 이상"],
        index=0,
        key="global_fee_band",
    )
else:
    fee_filter = "전체"

st.sidebar.markdown("---")
st.sidebar.caption(
    f"마지막 갱신: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}"
)

# ----------------------------------------------------
# 8. 글로벌 필터 적용
# ----------------------------------------------------
voc_filtered_global = df_voc.copy()

# 날짜
if dr and isinstance(dr, tuple) and len(dr) == 2:
    start_d, end_d = dr
    if isinstance(start_d, date) and isinstance(end_d, date):
        voc_filtered_global = voc_filtered_global[
            (voc_filtered_global["접수일시"] >= pd.to_datetime(start_d))
            & (
                voc_filtered_global["접수일시"]
                < pd.to_datetime(end_d) + pd.Timedelta(days=1)
            )
        ]

# 지사
if sel_branches:
    voc_filtered_global = voc_filtered_global[
        voc_filtered_global["관리지사"].isin(sel_branches)
    ]

# 리스크
if sel_risk:
    voc_filtered_global = voc_filtered_global[
        voc_filtered_global["리스크등급"].isin(sel_risk)
    ]

# 매칭
if sel_match:
    voc_filtered_global = voc_filtered_global[
        voc_filtered_global["매칭여부"].isin(sel_match)
    ]

# 월정료
if fee_filter != "전체" and "월정료_수치" in voc_filtered_global.columns:
    if fee_filter == "10만 이상":
        voc_filtered_global = voc_filtered_global[
            voc_filtered_global["월정료_수치"] >= 100000
        ]
    elif fee_filter == "10만 미만":
        voc_filtered_global = voc_filtered_global[
            (voc_filtered_global["월정료_수치"] < 100000)
            & voc_filtered_global["월정료_수치"].notna()
        ]

unmatched_global = voc_filtered_global[
    voc_filtered_global["매칭여부"] == "비매칭(X)"
].copy()

# ----------------------------------------------------
# 9. 상단 KPI
# ----------------------------------------------------
st.markdown("## 📊 해지 VOC 종합 대시보드")

total_rows = len(voc_filtered_global)
unique_cn = voc_filtered_global["계약번호_정제"].nunique()
unmatched_contracts = unmatched_global["계약번호_정제"].nunique()
matched_contracts = (
    voc_filtered_global[voc_filtered_global["매칭여부"] == "매칭(O)"]["계약번호_정제"]
    .nunique()
)

k1, k2, k3, k4 = st.columns(4)
k1.metric("VOC 접수건수", f"{total_rows:,}")
k2.metric("VOC 계약 수(유니크)", f"{unique_cn:,}")
k3.metric("비매칭(X) 계약 수", f"{unmatched_contracts:,}")
k4.metric("매칭(O) 계약 수", f"{matched_contracts:,}")

st.markdown("---")

# ----------------------------------------------------
# 10. 탭 구성
# ----------------------------------------------------
tab1, tab2, tab3, tab4, tab5 = st.tabs(
    [
        "📋 글로벌 VOC 리스트(행 단위)",
        "🚨 비매칭 계약 요약",
        "📊 지사/담당자 현황",
        "🔍 계약별 드릴다운 + 피드백",
        "🎯 비매칭 정밀 필터",
    ]
)

# ====================================================
# TAB 1 — 글로벌 VOC 리스트 (행 단위)
# ====================================================
with tab1:
    st.subheader("📋 글로벌 VOC 리스트 (행 단위)")

    # 담당자 / 지사 라디오 (빠른 선택용)
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

    tmp_mgr = voc_filtered_global.copy()
    if selected_branch_tab1 != "전체":
        tmp_mgr = tmp_mgr[tmp_mgr["관리지사"] == selected_branch_tab1]

    mgr_options_tab1 = (
        ["전체"]
        + sorted(
            tmp_mgr["구역담당자_통합"]
            .dropna()
            .astype(str)
            .unique()
            .tolist()
        )
        if "구역담당자_통합" in tmp_mgr.columns
        else ["전체"]
    )

    selected_mgr_tab1 = row1_col2.radio(
        "담당자 선택",
        options=mgr_options_tab1,
        horizontal=True,
        key="tab1_mgr_radio",
    )

    # 검색 (계약번호 / 상호 / 주소)
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
    if q_addr and address_cols:
        cond = False
        for col in address_cols:
            cond = cond | temp[col].astype(str).str.contains(q_addr.strip())
        temp = temp[cond]

    if temp.empty:
        st.info("조건에 맞는 VOC 데이터가 없습니다.")
    else:
        temp_sorted = temp.sort_values("접수일시", ascending=False)

        # 계약번호 기준 접수건수 계산
        grp = temp_sorted.groupby("계약번호_정제")
        temp_sorted["접수건수"] = grp["계약번호_정제"].transform("size")

        # 표시 컬럼 (None만 있는 컬럼은 자동 제외)
        show_cols = filter_nonempty_columns(temp_sorted, BASE_COLS_CANDIDATES)

        st.markdown(f"📌 표시 계약 수: **{temp_sorted['계약번호_정제'].nunique():,} 건**")
        st.dataframe(
            style_risk(temp_sorted[show_cols]),
            use_container_width=True,
            height=520,
        )

# ====================================================
# TAB 2 — 비매칭 계약 요약 (계약번호 기준)
# ====================================================
with tab2:
    st.subheader("🚨 비매칭(X) 계약 요약 (계약번호 기준)")

    if unmatched_global.empty:
        st.info("비매칭(X) 계약이 없습니다.")
    else:
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

        tmp_mgr_u = unmatched_global.copy()
        if selected_branch_u != "전체":
            tmp_mgr_u = tmp_mgr_u[tmp_mgr_u["관리지사"] == selected_branch_u]

        mgr_options_u = (
            ["전체"]
            + sorted(
                tmp_mgr_u["구역담당자_통합"]
                .dropna()
                .astype(str)
                .unique()
                .tolist()
            )
            if "구역담당자_통합" in tmp_mgr_u.columns
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
            st.info("조건에 맞는 비매칭(X) 계약이 없습니다.")
        else:
            temp_u_sorted = temp_u.sort_values("접수일시", ascending=False)
            grp_u = temp_u_sorted.groupby("계약번호_정제")
            # 계약별 최신 1건 + 접수건수
            idx_latest_u = grp_u["접수일시"].idxmax()
            df_u_summary = temp_u_sorted.loc[idx_latest_u].copy()
            df_u_summary["접수건수"] = grp_u.size().reindex(
                df_u_summary["계약번호_정제"]
            ).values

            show_cols_u = filter_nonempty_columns(
                df_u_summary, BASE_COLS_CANDIDATES
            )

            st.markdown(
                f"⚠ 활동대상 비매칭(X) 계약 수: **{len(df_u_summary):,} 건**"
            )

            st.dataframe(
                style_risk(df_u_summary[show_cols_u]),
                use_container_width=True,
                height=520,
            )

# ====================================================
# TAB 3 — 지사/담당자 현황
# ====================================================
with tab3:
    st.subheader("📊 지사 / 담당자별 비매칭 리스크 현황")

    if unmatched_global.empty:
        st.info("비매칭(X) 데이터가 없습니다.")
    else:
        c1, c2, c3 = st.columns(3)

        # 지사별 비매칭 계약 수
        bc = (
            unmatched_global.groupby("관리지사")["계약번호_정제"]
            .nunique()
            .rename("비매칭계약수")
        )
        bc = bc[bc.index.isin(BRANCH_ORDER)].reindex(BRANCH_ORDER).dropna()

        with c1:
            st.markdown("#### 🏢 지사별 비매칭 계약 수(계약 기준)")
            st.bar_chart(bc, use_container_width=True)

        # 담당자별 TOP 15
        mc = (
            unmatched_global.groupby("구역담당자_통합")["계약번호_정제"]
            .nunique()
            .rename("비매칭계약수")
            .sort_values(ascending=False)
        )
        mc = mc[mc.index.astype(str).str.strip() != ""].head(15)

        with c2:
            st.markdown("#### 👤 담당자별 비매칭 TOP 15")
            st.bar_chart(mc, use_container_width=True)

        # 리스크 분포
        rc = (
            unmatched_global["리스크등급"]
            .value_counts()
            .reindex(["HIGH", "MEDIUM", "LOW"])
            .fillna(0)
        )

        with c3:
            st.markdown("#### 🔥 비매칭 리스크 등급 분포(행 기준)")
            st.bar_chart(rc, use_container_width=True)

        st.markdown("---")

        # 일별 비매칭 추이
        if "접수일시" in unmatched_global.columns:
            trend = (
                unmatched_global.assign(접수일=unmatched_global["접수일시"].dt.date)
                .groupby("접수일")["계약번호_정제"]
                .nunique()
                .rename("비매칭계약수")
                .sort_index()
            )
            st.markdown("#### 📈 일별 비매칭 계약 추이")
            st.line_chart(trend, use_container_width=True)

# ====================================================
# TAB 4 — 계약별 드릴다운 + 피드백
# ====================================================
with tab4:
    st.subheader("🔍 계약번호 기준 드릴다운 + 피드백")

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
    else:
        drill_sorted = drill.sort_values("접수일시", ascending=False)
        g = drill_sorted.groupby("계약번호_정제")
        idx_latest_d = g["접수일시"].idxmax()
        df_d_summary = drill_sorted.loc[idx_latest_d].copy()
        df_d_summary["접수건수"] = g.size().reindex(
            df_d_summary["계약번호_정제"]
        ).values

        sum_cols_d = filter_nonempty_columns(
            df_d_summary, BASE_COLS_CANDIDATES
        )

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
                m2_1.metric("접수건수", f"{len(voc_hist):,}건")
                m2_2.metric("리스크등급", str(base_info.get("리스크등급", "")))
                m2_3.metric("매칭여부", str(base_info.get("매칭여부", "")))

                st.caption(
                    f"📍 시설_설치주소: {str(base_info.get('시설_설치주소', ''))}"
                )
                st.caption(
                    f"💰 시설_KTT월정료(조정): {str(base_info.get('시설_KTT월정료(조정)', ''))}"
                )

            st.markdown("---")

            c_left, c_right = st.columns(2)

            # VOC 이력
            with c_left:
                st.markdown("#### 📘 VOC 이력 (전체)")
                if voc_hist.empty:
                    st.info("VOC 이력이 없습니다.")
                else:
                    show_cols_hist = filter_nonempty_columns(
                        voc_hist, BASE_COLS_CANDIDATES
                    )
                    st.dataframe(
                        style_risk(voc_hist[show_cols_hist]),
                        use_container_width=True,
                        height=320,
                    )

            # 기타 출처
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

            # 피드백 이력 + 입력
            st.markdown("#### 📝 고객대응 / 현장 처리내역")

            fb_all = st.session_state["feedback_df"]
            fb_sel = fb_all[
                fb_all["계약번호_정제"].astype(str) == str(sel_cn)
            ].copy()
            fb_sel = fb_sel.sort_values("등록일자", ascending=False)

            if fb_sel.empty:
                st.info("등록된 처리 이력이 없습니다.")
            else:
                st.dataframe(
                    fb_sel,
                    use_container_width=True,
                    height=220,
                )

            st.markdown("##### ✏️ 새 처리내용 등록")

            fb1, fb2 = st.columns([3, 1])
            new_fb = fb1.text_area("고객대응 / 현장 처리내용", key="fb_content")
            new_user = fb2.text_input("등록자", key="fb_user")
            new_note = fb2.text_input("비고", key="fb_note")

            if st.button("💾 처리내역 저장", key="fb_save_btn"):
                if not new_fb.strip():
                    st.warning("처리내용을 입력하세요.")
                elif not new_user.strip():
                    st.warning("등록자를 입력하세요.")
                else:
                    new_row = pd.DataFrame(
                        [
                            {
                                "계약번호_정제": sel_cn,
                                "고객대응내용": new_fb.strip(),
                                "등록자": new_user.strip(),
                                "등록일자": datetime.now().strftime(
                                    "%Y-%m-%d %H:%M:%S"
                                ),
                                "비고": new_note.strip(),
                            }
                        ]
                    )
                    st.session_state["feedback_df"] = pd.concat(
                        [st.session_state["feedback_df"], new_row],
                        ignore_index=True,
                    )
                    save_feedback(FEEDBACK_PATH, st.session_state["feedback_df"])
                    st.success("처리내역이 저장되었습니다.")
                    st.experimental_rerun()

# ====================================================
# TAB 5 — 비매칭 정밀 필터
# ====================================================
with tab5:
    st.subheader("🎯 비매칭(X) 활동대상 정밀 필터")

    df_u = unmatched_global.copy()

    if df_u.empty:
        st.info("비매칭(X) 데이터가 없습니다.")
    else:
        f1, f2, f3 = st.columns([2, 2, 3])

        branches_5 = ["전체"] + sort_branch(df_u["관리지사"].dropna().unique())
        sel_branch_5 = f1.radio(
            "지사 선택",
            options=branches_5,
            horizontal=True,
            key="tab5_branch_radio",
        )

        tmp_mgr_5 = df_u.copy()
        if sel_branch_5 != "전체":
            tmp_mgr_5 = tmp_mgr_5[tmp_mgr_5["관리지사"] == sel_branch_5]

        mgr_options_5 = (
            ["전체"]
            + sorted(
                tmp_mgr_5["구역담당자_통합"]
                .dropna()
                .astype(str)
                .unique()
                .tolist()
            )
            if "구역담당자_통합" in tmp_mgr_5.columns
            else ["전체"]
        )

        sel_mgr_5 = f2.radio(
            "담당자 선택",
            options=mgr_options_5,
            horizontal=True,
            key="tab5_mgr_radio",
        )

        defense_types = ["지사방어", "센터방어"]
        filter_type = f3.radio(
            "VOC유형소 필터 방식",
            options=[
                "전체 보기",
                "지사방어만 보기",
                "센터방어만 보기",
                "지사·센터방어 제외한 실제 활동대상 보기",
            ],
            horizontal=False,
            key="tab5_filter_radio",
        )

        a1, a2 = st.columns(2)
        addr_kw = a1.text_input("시설_설치주소 검색(부분)", key="tab5_addr_kw")

        # 월정료 추가 범위 필터 (글로벌 필터에서 한 번 걸렸지만, 여기서 범위 재조정 가능)
        ktt_min, ktt_max = None, None
        if "월정료_수치" in df_u.columns and df_u["월정료_수치"].notna().any():
            valid_fee = df_u["월정료_수치"].dropna()
            min_val = int(valid_fee.min())
            max_val = int(valid_fee.max())
            ktt_min, ktt_max = a2.slider(
                "KTT월정료(조정) 범위(추가 필터)",
                min_value=min_val,
                max_value=max_val,
                value=(min_val, max_val),
                step=1000,
                key="tab5_ktt_slider",
            )

        df_filtered = df_u.copy()

        if sel_branch_5 != "전체":
            df_filtered = df_filtered[df_filtered["관리지사"] == sel_branch_5]
        if sel_mgr_5 != "전체":
            df_filtered = df_filtered[
                df_filtered["구역담당자_통합"].astype(str) == sel_mgr_5
            ]

        if filter_type == "지사방어만 보기":
            df_filtered = df_filtered[df_filtered["VOC유형소"] == "지사방어"]
        elif filter_type == "센터방어만 보기":
            df_filtered = df_filtered[df_filtered["VOC유형소"] == "센터방어"]
        elif filter_type == "지사·센터방어 제외한 실제 활동대상 보기":
            df_filtered = df_filtered[
                ~df_filtered["VOC유형소"].isin(defense_types)
            ]

        if addr_kw and "시설_설치주소" in df_filtered.columns:
            df_filtered = df_filtered[
                df_filtered["시설_설치주소"]
                .astype(str)
                .str.contains(addr_kw.strip())
            ]

        if (
            ktt_min is not None
            and ktt_max is not None
            and "월정료_수치" in df_filtered.columns
        ):
            df_filtered = df_filtered[
                df_filtered["월정료_수치"].between(ktt_min, ktt_max)
            ]

        st.markdown(
            f"📌 **필터 적용 후 비매칭 계약 수 : {df_filtered['계약번호_정제'].nunique():,} 건**"
        )

        if df_filtered.empty:
            st.warning("조건에 해당하는 데이터가 없습니다.")
        else:
            df_sorted = df_filtered.sort_values("접수일시", ascending=False)
            grp5 = df_sorted.groupby("계약번호_정제")
            idx_latest5 = grp5["접수일시"].idxmax()
            df_summary5 = df_sorted.loc[idx_latest5].copy()
            df_summary5["접수건수"] = grp5.size().reindex(
                df_summary5["계약번호_정제"]
            ).values

            sum_cols5 = filter_nonempty_columns(
                df_summary5, BASE_COLS_CANDIDATES
            )

            st.dataframe(
                style_risk(df_summary5[sum_cols5]),
                use_container_width=True,
                height=420,
            )
