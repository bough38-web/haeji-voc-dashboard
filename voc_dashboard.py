import os
import platform
import pandas as pd
import numpy as np
import streamlit as st
from datetime import datetime

# ----------------------------------------------------
# 기본 설정
# ----------------------------------------------------
st.set_page_config(page_title="해지 VOC 종합 대시보드", layout="wide")

# ----------------------------------------------------
# 파일 경로 설정
# ----------------------------------------------------
current_os = platform.system()
default_path = (
    r"C:\Users\User\Downloads\해지VOC관리시스템\merged.xlsx"
    if current_os == "Windows"
    else "/Users/heebonpark/Downloads/해지VOC관리시스템/merged.xlsx"
)

st.sidebar.header("📁 데이터 파일 경로 설정")
MERGED_PATH = st.sidebar.text_input("merged.xlsx 파일 경로", default_path)

# ----------------------------------------------------
# 데이터 로딩
# ----------------------------------------------------
@st.cache_data
def load_data(path: str) -> pd.DataFrame:
    if not os.path.exists(path):
        st.error(f"❌ 파일 없음: {path}")
        return pd.DataFrame()

    df = pd.read_excel(path)

    # 콤마 제거 처리
    for col in ["계약번호", "고객번호"]:
        if col in df.columns:
            df[col] = (
                df[col]
                .astype(str)
                .str.replace(",", "")
                .str.strip()
            )

    # 출처 정제
    df["출처"] = df["출처"].replace({"고객리스트": "해지시설"})

    # 계약번호 정제
    df["계약번호_정제"] = (
        df["계약번호"].astype(str)
        .str.replace(r"[^0-9A-Za-z]", "", regex=True)
        .str.strip()
    )

    # 날짜 처리
    if "접수일시" in df.columns:
        df["접수일시"] = pd.to_datetime(df["접수일시"], errors="coerce")

    return df


df = load_data(MERGED_PATH)
if df.empty:
    st.stop()

# ----------------------------------------------------
# 지사명 정제 (지사 → 지역명만 표시)
# ----------------------------------------------------
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

# ----------------------------------------------------
# 지사 사용자 정의 정렬
# ----------------------------------------------------
BRANCH_ORDER = [
    "중앙",
    "강북",
    "서대문",
    "고양",
    "의정부",
    "남양주",
    "강릉",
    "원주"
]

def sort_branch(series):
    return sorted(
        [s for s in series if s in BRANCH_ORDER],
        key=lambda x: BRANCH_ORDER.index(x)
    )

# ----------------------------------------------------
# 통합 구역/담당자 컬럼 생성
# ----------------------------------------------------
def make_zone(row):
    if "영업구역번호" in row and pd.notna(row["영업구역번호"]):
        return row["영업구역번호"]
    if "담당상세" in row and pd.notna(row["담당상세"]):
        return row["담당상세"]
    return ""

df["영업구역_통합"] = df.apply(make_zone, axis=1)

mgr_cols = [c for c in ["구역담당자", "담당자", "처리자"] if c in df.columns]

def pick_manager(row):
    for c in mgr_cols:
        v = row.get(c, "")
        if pd.notna(v) and str(v).strip() != "":
            return v
    return ""

df["구역담당자_통합"] = df.apply(pick_manager, axis=1)

# 주소 컬럼 자동 탐색
address_cols = [c for c in df.columns if "주소" in c]

# ----------------------------------------------------
# 출처 분리 + 매칭 계산
# ----------------------------------------------------
df_voc = df[df["출처"] == "해지VOC"].copy()
df_other = df[df["출처"] != "해지VOC"].copy()

other_sets = {
    src: set(df[df["출처"] == src]["계약번호_정제"].dropna())
    for src in ["해지시설", "해지요청", "설변", "정지", "해지파이프라인"]
    if src in df["출처"].unique()
}

other_union = set().union(*other_sets.values()) if other_sets else set()

# VOC ∧ 기타 출처 있음 → 매칭(O)
df_voc["매칭여부"] = df_voc["계약번호_정제"].apply(
    lambda x: "매칭(O)" if x in other_union else "비매칭(X)"
)

df_unmatched = df_voc[df_voc["매칭여부"] == "비매칭(X)"].copy()

# ----------------------------------------------------
# 리스크등급 + 경과일수 계산
# ----------------------------------------------------
today = datetime.today().date()

def compute_risk(row):
    dt = row.get("접수일시", pd.NaT)
    if pd.isna(dt):
        return np.nan, "MEDIUM"
    days = (today - dt.date()).days

    # 기본 리스크 룰
    if days <= 7:
        level = "HIGH"
    elif days <= 30:
        level = "MEDIUM"
    else:
        level = "LOW"

    # 해지상세 키워드 강화
    hs = str(row.get("해지상세", "") or "")
    if any(k in hs for k in ["즉시", "강성", "불만"]):
        if level == "MEDIUM":
            level = "HIGH"

    return days, level

df_voc["경과일수"], df_voc["리스크등급"] = zip(
    *df_voc.apply(lambda r: compute_risk(r), axis=1)
)

df_unmatched = df_voc[df_voc["매칭여부"] == "비매칭(X)"].copy()

# ----------------------------------------------------
# 컬럼 표시 규칙
# ----------------------------------------------------
exclude_cols = {
    "기타출처", "담당상세", "구역담당자_통합",
    "계약번호", "고객번호", "고객번호_정제",
    "고객명", "설치주소", "청구주소", "주소"
}

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
]

display_cols = [
    c for c in fixed_order
    if c in df_voc.columns and c not in exclude_cols
]

# ----------------------------------------------------
# UI 구성
# ----------------------------------------------------
st.title("📊 해지 VOC 종합 대시보드 (전문가버전)")

tab1, tab2, tab3 = st.tabs(
    ["📘 VOC 전체", "🚨 비매칭(활동대상)", "📊 지사/담당자 전문가현황"]
)

# ====================================================
# TAB 1 — VOC 전체 조회
# ====================================================
with tab1:
    st.subheader("📘 VOC 전체 조회")

    c1, c2, c3 = st.columns(3)
    key_cn = c1.text_input("계약번호 검색")
    key_name = c2.text_input("상호 검색")
    key_addr = c3.text_input("주소 검색")

    temp = df_voc.copy()

    if key_cn:
        temp = temp[temp["계약번호_정제"].str.contains(key_cn.strip())]

    if key_name and "상호" in temp.columns:
        temp = temp[temp["상호"].astype(str).str.contains(key_name.strip())]

    if key_addr and address_cols:
        cond = False
        for col in address_cols:
            cond |= temp[col].astype(str).str.contains(key_addr.strip())
        temp = temp[cond]

    temp = temp.sort_values("접수일시", ascending=False)

    st.dataframe(temp[display_cols], use_container_width=True, height=520)

# ====================================================
# TAB 2 — 비매칭(X)
# ====================================================
with tab2:
    st.subheader("🚨 비매칭(X) = 실제 활동대상")

    raw_branches = df_unmatched["관리지사"].dropna().unique().tolist()
    branches = ["전체"] + sort_branch(raw_branches)

    selected_branch = st.radio("지사 선택", branches, horizontal=True)

    temp_branch = df_unmatched.copy()
    if selected_branch != "전체":
        temp_branch = temp_branch[temp_branch["관리지사"] == selected_branch]

    mgr_list = ["전체"] + sorted(temp_branch["구역담당자_통합"].dropna().unique().tolist())

    selected_mgr = st.radio("담당자 선택", mgr_list, horizontal=True)

    temp = df_unmatched.copy()
    if selected_branch != "전체":
        temp = temp[temp["관리지사"] == selected_branch]
    if selected_mgr != "전체":
        temp = temp[temp["구역담당자_통합"] == selected_mgr]

    temp = temp.sort_values("접수일시", ascending=False)

    st.write(f"활동대상 {len(temp):,}건")
    st.dataframe(temp[display_cols], use_container_width=True, height=420)

# ====================================================
# TAB 3 — 전문가 현황
# ====================================================
with tab3:
    st.markdown("## 📊 지사·담당자 비매칭 전문가 현황")

    total_unmatched = len(df_unmatched)
    branch_count = df_unmatched["관리지사"].nunique()
    manager_count = df_unmatched["구역담당자_통합"].nunique()
    high_count = (df_unmatched["리스크등급"] == "HIGH").sum()

    k1, k2, k3, k4 = st.columns(4)
    k1.metric("🔴 비매칭", total_unmatched)
    k2.metric("🏢 지사", branch_count)
    k3.metric("👤 담당자", manager_count)
    k4.metric("🔥 HIGH 리스크", high_count)

    st.markdown("---")

    # 지사별 비매칭
    st.markdown("### 🏢 지사별 비매칭 현황")

    branch_counts = (
        df_unmatched.groupby("관리지사")["계약번호_정제"]
        .nunique()
        .reset_index()
        .rename(columns={"계약번호_정제": "비매칭건수"})
    )

    branch_counts = branch_counts[
        branch_counts["관리지사"].isin(BRANCH_ORDER)
    ]

    branch_counts = branch_counts.set_index("관리지사").reindex(BRANCH_ORDER)

    st.bar_chart(branch_counts["비매칭건수"], use_container_width=True)

    st.markdown("---")

    # 담당자별 TOP 20
    st.markdown("### 👤 담당자별 비매칭 TOP 20")

    mgr_counts = (
        df_unmatched.groupby("구역담당자_통합")["계약번호_정제"]
        .nunique()
        .reset_index()
        .rename(columns={"계약번호_정제": "비매칭건수"})
    )

    mgr_counts = mgr_counts[
        mgr_counts["구역담당자_통합"].astype(str).str.strip() != ""
    ].sort_values("비매칭건수", ascending=False).head(20)

    st.bar_chart(mgr_counts.set_index("구역담당자_통합")["비매칭건수"], use_container_width=True)

    st.markdown("---")

    # 상세 리스트 필터
    st.markdown("### 📋 상세 리스트")

    colA, colB = st.columns(2)

    branch_sel = colA.multiselect(
        "지사 선택", sort_branch(raw_branches)
    )

    temp_branch = df_unmatched.copy()
    if branch_sel:
        temp_branch = temp_branch[temp_branch["관리지사"].isin(branch_sel)]

    mgr_sel = colB.multiselect(
        "담당자 선택",
        sorted(temp_branch["구역담당자_통합"].dropna().unique().tolist())
    )

    temp = df_unmatched.copy()

    if branch_sel:
        temp = temp[temp["관리지사"].isin(branch_sel)]
    if mgr_sel:
        temp = temp[temp["구역담당자_통합"].isin(mgr_sel)]

    temp = temp.sort_values("접수일시", ascending=False)

    st.dataframe(temp[display_cols], use_container_width=True, height=420)