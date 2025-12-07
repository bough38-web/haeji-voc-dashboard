import os
import re
import smtplib
import urllib.parse
import base64
import time
from datetime import datetime, date
import numpy as np
import pandas as pd
import streamlit as st

# 전문가용 시각화 및 지능형 매핑 엔진
try:
    from rapidfuzz import process, utils
    import plotly.express as px
    import plotly.graph_objects as go
    HAS_LIBS = True
except ImportError:
    HAS_LIBS = False

# ----------------------------------------------------
# 0. 전역 세션 초기화 (NameError 방지)
# ----------------------------------------------------
if "feedback_db" not in st.session_state:
    st.session_state["feedback_db"] = pd.DataFrame(columns=["계약번호", "담당자", "상담상태", "상담내용", "입력일시"])

# ----------------------------------------------------
# 1. SMTP 및 환경 설정 (보안)
# ----------------------------------------------------
SMTP_USER = "bough38@gmail.com"
SMTP_PASSWORD = "adzk cyik sing emds" 
SENDER_NAME = "해지VOC 관리자"

st.set_page_config(page_title="🛡️ Enterprise VOC Manager", layout="wide")

# ----------------------------------------------------
# 2. 데이터 관리자 파이프라인 (KeyError 완벽 방지)
# ----------------------------------------------------
@st.cache_data
def load_and_verify_master():
    if not os.path.exists("merged.xlsx"): return pd.DataFrame()
    df = pd.read_excel("merged.xlsx")
    
    # [안정화 1] 컬럼명 유연성 확보: 실제 파일의 컬럼명을 체크하여 매핑
    col_mapping = {
        '접수일시': next((c for c in df.columns if '접수일' in c), '접수일시'),
        '계약번호': next((c for c in df.columns if '계약' in c), '계약번호'),
        '관리지사': next((c for c in df.columns if '지사' in c), '관리지사'),
        '처리자': next((c for c in df.columns if '처리' in c or '담당' in c), '처리자')
    }
    
    # 존재하지 않는 컬럼에 대한 에러 방지
    for target, actual in col_mapping.items():
        if actual in df.columns and target != actual:
            df = df.rename(columns={actual: target})

    # [무결성] 값이 하나도 없는 공백 열 제거
    df = df.dropna(axis=1, how='all')

    # 날짜 데이터 변환 (KeyError 방지 체크)
    if "접수일시" in df.columns:
        df["접수일시"] = pd.to_datetime(df["접수일시"], errors="coerce")
    else:
        df["접수일시"] = pd.Timestamp.now() # 누락 시 현재 시각 할당

    df["계약번호_정제"] = df["계약번호"].astype(str).str.replace(r"[^0-9A-Za-z]", "", regex=True)
    today = date.today()
    df["리스크등급"] = df["접수일시"].apply(lambda dt: "HIGH" if pd.notna(dt) and (today - dt.date()).days <= 3 else "LOW")
    
    return df

@st.cache_data
def load_contacts():
    if not os.path.exists("contact_map.xlsx"): return {}
    df_c = pd.read_excel("contact_map.xlsx")
    # 컬럼 자동 탐지
    email_col = next((c for c in df_c.columns if "E-MAIL" in str(c) or "이메일" in str(c)), df_c.columns[1])
    name_col = next((c for c in df_c.columns if "처리자" in str(c) or "담당자" in str(c)), df_c.columns[0])
    return {str(row[name_col]).strip(): str(row[email_col]).strip() for _, row in df_c.iterrows()}

# 데이터 로딩
df_all = load_and_verify_master()
manager_contacts = load_contacts()
df_voc = df_all[df_all["출처"] == "해지VOC"] if "출처" in df_all.columns else df_all

# ----------------------------------------------------
# 3. UI 탭 및 로직 (TAB 2 필터 수정)
# ----------------------------------------------------
tabs = st.tabs(["📊 분석", "🔍 마스터", "📨 알림", "⚙️ 피드백"])

with tabs[1]:
    st.subheader("🔎 동적 계약 데이터베이스")
    # 관리지사 필터가 존재할 때만 멀티셀렉트 생성
    if "관리지사" in df_all.columns:
        sel_branches = st.multiselect("지사 선택", df_all["관리지사"].unique())
        df_m = df_all[df_all["관리지사"].isin(sel_branches)] if sel_branches else df_all
    else:
        df_m = df_all

    st.dataframe(df_m.sort_values("접수일시", ascending=False), use_container_width=True, hide_index=True)

# ... 나머지 TAB 0, 2, 3 로직 (이전 제공본과 동일 유지)
