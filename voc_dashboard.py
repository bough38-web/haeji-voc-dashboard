import os
from datetime import datetime, date
import numpy as np
import pandas as pd
import streamlit as st

# ----------------------------------------------------
# 0. 기본 설정
# ----------------------------------------------------
st.set_page_config(page_title="해지 VOC 종합 대시보드", layout="wide")

# ----------------------------------------------------
# 1. 파일 경로
# ----------------------------------------------------
MERGED_PATH = "merged.xlsx"
FEEDBACK_PATH = "feedback.csv"

# ----------------------------------------------------
# 2. 데이터 로딩
# ----------------------------------------------------
@st.cache_data
def load_data(path: str):
    if not os.path.exists(path):
        st.error("❌ merged.xlsx 파일을 찾지 못했습니다.")
        return pd.DataFrame()

    df = pd.read_excel(path)

    # 시설_ 컬럼 자동 인식하여 원본 컬럼로 매핑
    rename_map = {
        "시설_설치주소": "설치주소",
        "시설_KTT월정료(조정)": "KTT월정료(조정)",
        "시설_사업자번호": "사업자번호",
        "시설_영업구역정보": "영업구역정보",
        "시설_실적채널": "실적채널",
        "시설_계약상태(중)": "계약상태(중)"
    }
    df.rename(columns={k: v for k, v in rename_map.items() if k in df.columns}, inplace=True)

    # None / NaN 컬럼 제거
    df = df.dropna(axis=1, how="all")

    # 계약번호 정제
    if "계약번호" in df.columns:
        df["계약번호_정제"] = (
            df["계약번호"].astype(str)
            .str.replace(r"[^0-9A-Za-z]", "", regex=True)
            .str.strip()
        )
    else:
        df["계약번호_정제"] = ""

    # 접수일시 변환
    if "접수일시" in df.columns:
        df["접수일시"] = pd.to_datetime(df["접수일시"], errors="coerce")

    # 월정료 처리
    if "KTT월정료(조정)" in df.columns:
        df["월정료_수치"] = (
            df["KTT월정료(조정)"].astype(str)
            .str.replace(",", "")
            .str.extract(r"(\d+)", expand=False)
            .astype(float)
        )

        df["월정료구간"] = df["월정료_수치"].apply(
            lambda v: "10만 이상" if pd.notna(v) and v >= 100000 else
            ("10만 미만" if pd.notna(v) else "미기재")
        )
    else:
        df["월정료_수치"] = np.nan
        df["월정료구간"] = "미기재"

    return df


# 피드백 로드
def load_feedback(path):
    if os.path.exists(path):
        return pd.read_csv(path, encoding="utf-8-sig")
    return pd.DataFrame(columns=["계약번호_정제", "고객대응내용", "등록자", "등록일자", "비고"])


def save_feedback(path, fb):
    fb.to_csv(path, index=False, encoding="utf-8-sig")


# ---------------- 데이터 불러오기 ----------------
df = load_data(MERGED_PATH)
if df.empty:
    st.stop()

if "feedback_df" not in st.session_state:
    st.session_state["feedback_df"] = load_feedback(FEEDBACK_PATH)

# ----------------------------------------------------
# 3. 출처 분리 및 매칭 계산
# ----------------------------------------------------
df_voc = df[df.get("출처", "") == "해지VOC"].copy()
df_other = df[df.get("출처", "") != "해지VOC"].copy()

other_union = set(df_other["계약번호_정제"].dropna())
df_voc["매칭여부"] = df_voc["계약번호_정제"].apply(
    lambda x: "매칭(O)" if x in other_union else "비매칭(X)"
)

# ----------------------------------------------------
# 4. 리스크 계산
# ----------------------------------------------------
today = date.today()

def risk(row):
    dt = row["접수일시"]
    if pd.isna(dt):
        return np.nan, "LOW"
    days = (today - dt.date()).days
    if days <= 3:
        return days, "HIGH"
    elif days <= 10:
        return days, "MEDIUM"
    return days, "LOW"

df_voc["경과일수"], df_voc["리스크등급"] = zip(*df_voc.apply(risk, axis=1))

unmatched_global = df_voc[df_voc["매칭여부"] == "비매칭(X)"]

# ----------------------------------------------------
# 5. 사이드바 글로벌 필터
# ----------------------------------------------------
st.sidebar.title("🔧 글로벌 필터")

# 날짜 범위
if df_voc["접수일시"].notna().any():
    min_d = df_voc["접수일시"].min().date()
    max_d = df_voc["접수일시"].max().date()
    dr = st.sidebar.date_input(
        "접수일자 범위", (min_d, max_d),
        min_value=min_d, max_value=max_d
    )
else:
    dr = None

# 월정료
fee_filter = st.sidebar.radio(
    "월정료 구간", ["전체", "10만 미만", "10만 이상"], index=0
)

# 매칭여부
match_filter = st.sidebar.multiselect(
    "매칭여부", ["매칭(O)", "비매칭(X)"], default=["매칭(O)", "비매칭(X)"]
)

voc_filtered = df_voc.copy()

# 날짜 필터
if dr:
    start_d, end_d = dr
    voc_filtered = voc_filtered[
        (voc_filtered["접수일시"] >= pd.to_datetime(start_d))
        & (voc_filtered["접수일시"] < pd.to_datetime(end_d) + pd.Timedelta(days=1))
    ]

# 매칭 필터
voc_filtered = voc_filtered[voc_filtered["매칭여부"].isin(match_filter)]

# 월정료 필터
if fee_filter == "10만 이상":
    voc_filtered = voc_filtered[voc_filtered["월정료_수치"] >= 100000]
elif fee_filter == "10만 미만":
    voc_filtered = voc_filtered[
        (voc_filtered["월정료_수치"] < 100000) & voc_filtered["월정료_수치"].notna()
    ]

unmatched_global = voc_filtered[voc_filtered["매칭여부"] == "비매칭(X)"]

# ----------------------------------------------------
# 6. KPI
# ----------------------------------------------------
st.markdown("## 📊 해지 VOC 종합 대시보드")

c1, c2, c3, c4 = st.columns(4)
c1.metric("VOC 접수건수", f"{len(voc_filtered):,}")
c2.metric("VOC 계약수", f"{voc_filtered['계약번호_정제'].nunique():,}")
c3.metric("비매칭(X)", f"{unmatched_global['계약번호_정제'].nunique():,}")
c4.metric("매칭(O)", f"{voc_filtered[voc_filtered['매칭여부']=='매칭(O)']['계약번호_정제'].nunique():,}")

st.markdown("---")

# ----------------------------------------------------
# 7. 탭 구성 (1~5)
# ----------------------------------------------------
tab1, tab2, tab3 = st.tabs(["📘 VOC 전체", "🚨 비매칭", "🔍 계약별 상세"])

# ----------------------------------------------------
# TAB 1
# ----------------------------------------------------
with tab1:
    st.subheader("📘 VOC 전체 요약")

    df_latest = (
        voc_filtered.sort_values("접수일시", ascending=False)
        .groupby("계약번호_정제")
        .head(1)
    )

    st.dataframe(
        df_latest[[
            "계약번호_정제", "상호", "설치주소",
            "KTT월정료(조정)", "월정료구간",
            "리스크등급", "경과일수", "매칭여부"
        ]],
        use_container_width=True,
        height=500
    )

# ----------------------------------------------------
# TAB 2
# ----------------------------------------------------
with tab2:
    st.subheader("🚨 비매칭(X) 활동대상")

    df_un = (
        unmatched_global.sort_values("접수일시", ascending=False)
        .groupby("계약번호_정제")
        .head(1)
    )

    st.dataframe(
        df_un[[
            "계약번호_정제", "상호",
            "설치주소", "KTT월정료(조정)", "월정료구간",
            "리스크등급", "경과일수"
        ]],
        use_container_width=True,
        height=500
    )

# ----------------------------------------------------
# TAB 3 — 계약별 상세
# ----------------------------------------------------
with tab3:
    st.subheader("🔍 계약별 상세")

    cn_list = voc_filtered["계약번호_정제"].unique().tolist()
    sel_cn = st.selectbox("계약 선택", ["(선택)"] + cn_list)

    if sel_cn != "(선택)":
        detail = voc_filtered[voc_filtered["계약번호_정제"] == sel_cn]

        st.dataframe(
            detail.sort_values("접수일시", ascending=False),
            use_container_width=True,
            height=500
        )
