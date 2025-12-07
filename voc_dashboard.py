# ====================================================
# 해지 VOC 종합 대시보드 (통합 고도화 실행본)
# ====================================================
# 1. 안정화 유틸
# 2. 데이터 정규화
# 3. 운영 가이드 반영
# 4. 성능/확장 고려
# ====================================================

import os
from datetime import datetime, date
import smtplib
from email.message import EmailMessage

import numpy as np
import pandas as pd
import streamlit as st

# -----------------------------
# Plotly (fallback 포함)
# -----------------------------
try:
    import plotly.express as px
    HAS_PLOTLY = True
except Exception:
    HAS_PLOTLY = False

# ====================================================
# Streamlit 기본 설정
# ====================================================
st.set_page_config(page_title="해지 VOC 종합 대시보드", layout="wide")

# ====================================================
# ✅ [공통 유틸 – 1번]
# ====================================================
def safe_str(x):
    if pd.isna(x):
        return ""
    return str(x).strip()

def safe_eq(series, value):
    if value in ["", None, "(전체)", "전체"]:
        return pd.Series([True] * len(series), index=series.index)
    return (
        series.astype(str)
        .str.strip()
        .replace({"nan": "", "None": ""})
        == str(value).strip()
    )

def safe_unique(series):
    return sorted(
        series.astype(str)
        .str.strip()
        .replace({"nan": "", "None": ""})
        .loc[lambda s: s != ""]
        .unique()
        .tolist()
    )

def latest_by_contract(df, date_col="접수일시"):
    df = df.copy()
    df[date_col] = pd.to_datetime(df[date_col], errors="coerce")
    idx = df.groupby("계약번호_정제")[date_col].idxmax()
    return df.loc[idx]

def parse_fee(x):
    if pd.isna(x):
        return np.nan
    s = str(x)
    if not any(ch.isdigit() for ch in s):
        return np.nan
    s = s.replace(",", "")
    digits = "".join(ch for ch in s if ch.isdigit() or ch == ".")
    try:
        v = float(digits)
    except Exception:
        return np.nan
    if 200000 <= v <= 2000000:  # 일부 데이터만 10배 보정
        v = v / 10
    return v

def compute_risk(row):
    dt = pd.to_datetime(row.get("접수일시"), errors="coerce")
    if pd.isna(dt):
        return np.nan, "LOW"
    days = (date.today() - dt.date()).days
    if days < 0:
        return days, "LOW"
    if days <= 3:
        return days, "HIGH"
    elif days <= 10:
        return days, "MEDIUM"
    return days, "LOW"

# ====================================================
# ✅ SMTP / 관리자 코드 – 3번
# ====================================================
SMTP_HOST = st.secrets.get("SMTP_HOST", "")
SMTP_PORT = int(st.secrets.get("SMTP_PORT", 587))
SMTP_USER = st.secrets.get("SMTP_USER", "")
SMTP_PASSWORD = st.secrets.get("SMTP_PASSWORD", "")
SENDER_NAME = st.secrets.get("SENDER_NAME", "해지VOC 관리자")
ADMIN_CODE = st.secrets.get("ADMIN_CODE", "C3A")

# ====================================================
# 파일 경로
# ====================================================
MERGED_PATH = "merged.xlsx"
FEEDBACK_PATH = "feedback.csv"
CONTACT_PATH = "contact_map.xlsx"

# ====================================================
# ✅ 데이터 로딩 – 2번
# ====================================================
@st.cache_data
def load_voc_data(path):
    if not os.path.exists(path):
        st.error("merged.xlsx 파일이 없습니다.")
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

    if "접수일시" in df.columns:
        df["접수일시"] = pd.to_datetime(df["접수일시"], errors="coerce")

    return df

df = load_voc_data(MERGED_PATH)
if df.empty:
    st.stop()

# ✅ 전역 문자열 정규화
for c in ["관리지사", "구역담당자", "담당자", "처리자", "상호", "출처"]:
    if c in df.columns:
        df[c] = df[c].astype(str).str.strip().replace({"nan": "", "None": ""})

df["계약번호_정제"] = df["계약번호_정제"].astype(str).str.strip()

# ====================================================
# 담당자 통합
# ====================================================
mgr_priority = ["구역담당자", "담당자", "처리자"]

def pick_manager(row):
    for c in mgr_priority:
        if c in row and safe_str(row[c]):
            return row[c]
    return ""

df["구역담당자_통합"] = df.apply(pick_manager, axis=1)

# ====================================================
# 출처 분리 & 매칭
# ====================================================
df_voc = df[df["출처"] == "해지VOC"].copy()
df_other = df[df["출처"] != "해지VOC"].copy()

other_union = set(df_other["계약번호_정제"].dropna())
df_voc["매칭여부"] = df_voc["계약번호_정제"].apply(
    lambda x: "매칭(O)" if x in other_union else "비매칭(X)"
)

# ====================================================
# 리스크 / 월정료
# ====================================================
df_voc["경과일수"], df_voc["리스크등급"] = zip(
    *df_voc.apply(lambda r: compute_risk(r), axis=1)
)

if "시설_KTT월정료(조정)" in df_voc.columns:
    df_voc["월정료_수치"] = df_voc["시설_KTT월정료(조정)"].apply(parse_fee)
else:
    df_voc["월정료_수치"] = np.nan

# ====================================================
# 글로벌 필터
# ====================================================
st.sidebar.title("🔧 글로벌 필터")

branches = safe_unique(df_voc["관리지사"])
sel_branches = st.sidebar.multiselect(
    "관리지사", branches, default=branches
)

risk_levels = ["HIGH", "MEDIUM", "LOW"]
sel_risk = st.sidebar.multiselect(
    "리스크등급", risk_levels, default=risk_levels
)

base = df_voc.copy()
if sel_branches:
    base = base[base["관리지사"].isin(sel_branches)]
if sel_risk:
    base = base[base["리스크등급"].isin(sel_risk)]

unmatched_global = base[base["매칭여부"] == "비매칭(X)"]

# ====================================================
# 상단 KPI
# ====================================================
st.markdown("## 📊 해지 VOC 종합 대시보드")

c1, c2, c3 = st.columns(3)
c1.metric("VOC 행 수", len(base))
c2.metric("유니크 계약 수", base["계약번호_정제"].nunique())
c3.metric("비매칭 계약 수", unmatched_global["계약번호_정제"].nunique())

st.markdown("---")

# ====================================================
# 탭 구성
# ====================================================
tab_viz, tab_all, tab_alert = st.tabs(
    ["📊 지사/담당자 현황", "📘 계약 기준 요약", "📨 담당자 알림"]
)

# ====================================================
# TAB 1
# ====================================================
with tab_viz:
    st.subheader("지사별 비매칭 계약 현황")

    if unmatched_global.empty:
        st.info("비매칭 데이터 없음")
    else:
        branch_stats = (
            unmatched_global.groupby("관리지사")["계약번호_정제"]
            .nunique()
            .sort_values(ascending=False)
        )

        if HAS_PLOTLY:
            fig = px.bar(
                branch_stats.reset_index(),
                x="관리지사",
                y="계약번호_정제",
                text="계약번호_정제"
            )
            fig.update_traces(textposition="auto")
            st.plotly_chart(fig, use_container_width=True)
        else:
            st.bar_chart(branch_stats)

# ====================================================
# TAB 2
# ====================================================
with tab_all:
    st.subheader("계약 기준 최신 VOC 요약")
    summary = latest_by_contract(base)
    st.dataframe(summary, use_container_width=True, height=500)

# ====================================================
# TAB 3
# ====================================================
with tab_alert:
    st.subheader("담당자 비매칭 계약 이메일 발송")

    mgr_list = safe_unique(unmatched_global["구역담당자_통합"])
    sel_mgr = st.selectbox("담당자 선택", ["(선택)"] + mgr_list)

    if sel_mgr != "(선택)":
        mgr_df = unmatched_global[safe_eq(unmatched_global["구역담당자_통합"], sel_mgr)]
        mgr_latest = latest_by_contract(mgr_df)

        st.write(f"유니크 계약 수: {len(mgr_latest)}")
        st.dataframe(mgr_latest, use_container_width=True, height=300)

        email = st.text_input("이메일 주소")
        if st.button("📤 이메일 발송"):
            if not email:
                st.error("이메일 입력 필요")
            else:
                try:
                    msg = EmailMessage()
                    msg["Subject"] = f"[해지VOC] {sel_mgr} 담당 비매칭 계약 안내"
                    msg["From"] = f"{SENDER_NAME} <{SMTP_USER}>"
                    msg["To"] = email
                    msg.set_content(
                        f"{sel_mgr} 담당자님,\n\n"
                        f"비매칭 해지 VOC 계약 {len(mgr_latest)}건을 공유드립니다.\n"
                        f"첨부파일을 확인해주세요.\n\n- 해지VOC 관리자 -"
                    )

                    csv_bytes = mgr_latest.to_csv(index=False).encode("utf-8-sig")
                    msg.add_attachment(
                        csv_bytes,
                        maintype="application",
                        subtype="octet-stream",
                        filename=f"unmatched_{sel_mgr}.csv",
                    )

                    with smtplib.SMTP(SMTP_HOST, SMTP_PORT, timeout=10) as smtp:
                        smtp.starttls()
                        if SMTP_USER:
                            smtp.login(SMTP_USER, SMTP_PASSWORD)
                        smtp.send_message(msg)

                    st.success("✅ 이메일 발송 완료")
                except Exception as e:
                    st.error(f"❌ 발송 실패: {e}")
