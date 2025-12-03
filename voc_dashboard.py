import os
import pandas as pd
import numpy as np
import streamlit as st
from datetime import datetime

# ----------------------------------------------------
# 기본 설정
# ----------------------------------------------------
st.set_page_config(page_title="해지 VOC 종합 대시보드", layout="wide")

# ----------------------------------------------------
# merged.xlsx 파일을 GitHub 저장소에서 직접 로딩
# ----------------------------------------------------
@st.cache_data
def load_data() -> pd.DataFrame:
    file_path = "merged.xlsx"   # GitHub repo root에 존재해야 함

    if not os.path.exists(file_path):
        st.error(f"❌ GitHub 저장소에서 merged.xlsx 파일을 찾을 수 없습니다.")
        return pd.DataFrame()

    df = pd.read_excel(file_path)

    # 콤마 제거
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


df = load_data()
if df.empty:
    st.stop()

# ----------------------------------------------------
# 지사명 축약 적용
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
# 지사 순서 정의
# ----------------------------------------------------
BRANCH_ORDER = [
    "중앙","강북","서대문","고양","의정부","남양주","강릉","원주"
]

def sort_branch(series):
    return sorted([s for s in series if s in BRANCH_ORDER],
                  key=lambda x: BRANCH_ORDER.index(x))

# ----------------------------------------------------
# 통합 구역/담당자 생성
# ----------------------------------------------------
def make_zone(row):
    if "영업구역번호" in row and pd.notna(row["영업구역번호"]):
        return row["영업구역번호"]
    if "담당상세" in row and pd.notna(row["담당상세"]):
        return row["담당상세"]
    return ""
df["영업구역_통합"] = df.apply(make_zone, axis=1)

mgr_priority = ["구역담당자","담당자","처리자"]

def pick_manager(row):
    for c in mgr_priority:
        if c in row and pd.notna(row[c]) and str(row[c]).strip() != "":
            return row[c]
    return ""
df["구역담당자_통합"] = df.apply(pick_manager, axis=1)

# 주소 컬럼 자동 탐색
address_cols = [c for c in df.columns if "주소" in c]

# ----------------------------------------------------
# 출처별 분리 및 매칭 계산
# ----------------------------------------------------
df_voc = df[df["출처"] == "해지VOC"].copy()
df_other = df[df["출처"] != "해지VOC"].copy()

other_sets = {
    src: set(df_other[df_other["출처"] == src]["계약번호_정제"].dropna())
    for src in ["해지시설", "해지요청", "설변", "정지", "해지파이프라인"]
}

other_union = set().union(*other_sets.values())

df_voc["매칭여부"] = df_voc["계약번호_정제"].apply(
    lambda x: "매칭(O)" if x in other_union else "비매칭(X)"
)

df_unmatched = df_voc[df_voc["매칭여부"] == "비매칭(X)"].copy()

# ----------------------------------------------------
# 리스크 등급/경과일
# ----------------------------------------------------
today = datetime.today().date()

def compute_risk(row):
    dt = row.get("접수일시")
    if pd.isna(dt):
        return np.nan, "MEDIUM"

    days = (today - dt.date()).days

    if days <= 7:
        level = "HIGH"
    elif days <= 30:
        level = "MEDIUM"
    else:
        level = "LOW"

    hs = str(row.get("해지상세",""))
    if any(k in hs for k in ["즉시","강성","불만"]):
        if level == "MEDIUM":
            level = "HIGH"

    return days, level

df_voc["경과일수"], df_voc["리스크등급"] = zip(
    *df_voc.apply(lambda r: compute_risk(r), axis=1)
)

df_unmatched = df_voc[df_voc["매칭여부"] == "비매칭(X)"]

# ----------------------------------------------------
# 컬럼 표시 규칙
# ----------------------------------------------------
exclude_cols = {
    "기타출처","담당상세","구역담당자_통합","계약번호","고객번호",
    "고객번호_정제","고객명","설치주소","청구주소","주소"
}

fixed_order = [
    "상호","계약번호_정제","매칭여부","리스크등급","경과일수",
    "출처","관리지사","영업구역번호","영업구역_통합",
    "처리자","담당유형","처리유형","처리내용",
    "접수일시","서비스개시일","계약종료일",
    "서비스중","서비스소",
    "VOC유형","VOC유형중","VOC유형소",
    "해지상세","등록내용",
]

display_cols = [c for c in fixed_order if c in df_voc.columns]

# ----------------------------------------------------
# UI Layout
# ----------------------------------------------------
st.title("📊 해지 VOC 종합 대시보드 (배포버전)")

tab1, tab2, tab3 = st.tabs(
    ["📘 VOC 전체", "🚨 비매칭(활동대상)", "📊 지사/담당자 전문가현황"]
)

# ----------------------------------------------------
# TAB 1 — 전체 VOC
# ----------------------------------------------------
with tab1:
    st.subheader("📘 VOC 전체 조회")
    c1, c2, c3 = st.columns(3)
    key_cn  = c1.text_input("계약번호 검색")
    key_nm  = c2.text_input("상호 검색")
    key_addr = c3.text_input("주소 검색")

    temp = df_voc.copy()

    if key_cn:
        temp = temp[temp["계약번호_정제"].str.contains(key_cn)]

    if key_nm:
        temp = temp[temp["상호"].astype(str).str.contains(key_nm)]

    if key_addr:
        cond = False
        for a in address_cols:
            cond |= temp[a].astype(str).str.contains(key_addr)
        temp = temp[cond]

    temp = temp.sort_values("접수일시", ascending=False)
    st.dataframe(temp[display_cols], use_container_width=True, height=520)

# ----------------------------------------------------
# TAB 2 — 비매칭
# ----------------------------------------------------
with tab2:
    st.subheader("🚨 비매칭(X) = 활동대상")

    branch_list = sort_branch(df_unmatched["관리지사"].dropna().unique())
    sel_branch = st.radio("지사 선택", ["전체"] + branch_list, horizontal=True)

    temp = df_unmatched.copy()
    if sel_branch != "전체":
        temp = temp[temp["관리지사"] == sel_branch]

    mgr_list = sorted(temp["구역담당자_통합"].dropna().unique())
    sel_mgr = st.radio("담당자 선택", ["전체"] + mgr_list, horizontal=True)

    if sel_mgr != "전체":
        temp = temp[temp["구역담당자_통합"] == sel_mgr]

    temp = temp.sort_values("접수일시", ascending=False)

    st.write(f"활동대상 {len(temp):,}건")
    st.dataframe(temp[display_cols], use_container_width=True, height=420)

# ----------------------------------------------------
# TAB 3 — 전문가 현황
# ----------------------------------------------------
with tab3:
    st.subheader("📊 지사·담당자 전문가 리스크 현황")

    total = len(df_unmatched)
    branch_cnt = df_unmatched["관리지사"].nunique()
    mgr_cnt = df_unmatched["구역담당자_통합"].nunique()
    high_cnt = (df_unmatched["리스크등급"]=="HIGH").sum()

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("🔴 비매칭", total)
    c2.metric("🏢 지사", branch_cnt)
    c3.metric("👤 담당자", mgr_cnt)
    c4.metric("🔥 HIGH", high_cnt)

    st.markdown("---")

    st.markdown("### 🏢 지사별 비매칭")
    bc = (
        df_unmatched.groupby("관리지사")["계약번호_정제"]
        .nunique().rename("비매칭건수")
    )
    bc = bc.reindex(BRANCH_ORDER)
    st.bar_chart(bc)

    st.markdown("### 👤 담당자별 비매칭 TOP 20")
    mc = (
        df_unmatched.groupby("구역담당자_통합")["계약번호_정제"]
        .nunique().rename("비매칭건수")
        .sort_values(ascending=False).head(20)
    )
    st.bar_chart(mc)
