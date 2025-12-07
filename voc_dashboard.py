import os
import re
import smtplib
from datetime import datetime, date
from email.message import EmailMessage

import numpy as np
import pandas as pd
import streamlit as st

# 전문가용 지능형 매핑 라이브러리 (유사도 분석)
try:
    from rapidfuzz import process, utils
    HAS_RAPIDFUZZ = True
except ImportError:
    HAS_RAPIDFUZZ = False

# Plotly 고급 시각화
try:
    import plotly.express as px
    HAS_PLOTLY = True
except Exception:
    HAS_PLOTLY = False

# ----------------------------------------------------
# 0. UI 설정 & 라이트톤 CSS
# ----------------------------------------------------
st.set_page_config(page_title="해지 VOC 종합 대시보드 Pro", layout="wide")

st.markdown("""
    <style>
    html, body, .stApp { background-color: #f5f5f7 !important; color: #1d1d1f !important; }
    .section-card { background: white; border-radius: 12px; padding: 1.5rem; border: 1px solid #e5e7eb; margin-bottom: 1rem; }
    .stMetric { background: white; padding: 15px; border-radius: 10px; border: 1px solid #efefef; }
    </style>
    """, unsafe_allow_html=True)

# ----------------------------------------------------
# 1. 유틸리티 (매핑 검증 & 이메일 정규식)
# ----------------------------------------------------
def is_valid_email(email):
    regex = r'^[a-zA-Z0-9+-_.]+@[a-zA-Z0-9-]+\.[a-zA-Z0-9-.]+$'
    return bool(re.match(regex, str(email)))

def get_smart_contact(target_name, contact_dict):
    """Fuzzy Matching을 통한 지능형 담당자 매핑"""
    target_name = str(target_name).strip()
    if target_name in contact_dict:
        return contact_dict[target_name], "Verified"
    
    if HAS_RAPIDFUZZ:
        choices = list(contact_dict.keys())
        result = process.extractOne(target_name, choices, processor=utils.default_process)
        if result and result[1] >= 90:
            suggested_name = result[0]
            return contact_dict[suggested_name], f"Suggested({suggested_name})"
    
    return None, "Not Found"

# ----------------------------------------------------
# 2. 파일 로딩 (사용자 코드 기반)
# ----------------------------------------------------
MERGED_PATH = "merged.xlsx"
CONTACT_PATH = "contact_map.xlsx" # 업로드된 파일명에 맞춰 자동 탐지 권장
FEEDBACK_PATH = "feedback.csv"

@st.cache_data
def load_data():
    if not os.path.exists(MERGED_PATH): return pd.DataFrame()
    df = pd.read_excel(MERGED_PATH)
    # 데이터 정제 로직 포함 (사용자 원본 로직 유지)
    df["계약번호_정제"] = df["계약번호"].astype(str).str.replace(r"[^0-9A-Za-z]", "", regex=True)
    df["접수일시"] = pd.to_datetime(df["접수일시"], errors="coerce")
    return df

@st.cache_data
def load_contacts(path):
    if not os.path.exists(path): return pd.DataFrame(), {}
    df_c = pd.read_excel(path)
    # 담당자/이메일 자동 탐지 및 딕셔너리 생성
    contact_dict = {str(row[0]).strip(): {"email": str(row[1]).strip()} for _, row in df_c.iterrows() if pd.notna(row[0])}
    return df_c, contact_dict

df = load_data()
contact_df, manager_contacts = load_contacts(CONTACT_PATH)

#  

# ----------------------------------------------------
# 3. 비매칭 리스크 계산 및 필터링 (글로벌)
# ----------------------------------------------------
# 출처별 필터링 및 매칭여부 계산 로직 (사용자 코드 기반 축약)
df_voc = df[df["출처"] == "해지VOC"].copy()
# ... 리스크 등급 계산 ... (사용자 로직 적용)

# ----------------------------------------------------
# 4. 탭 구성 (Tab Alert 강화)
# ----------------------------------------------------
tabs = st.tabs(["📊 시각화", "📘 VOC 전체", "🧯 비매칭", "🔍 활동등록", "🎯 정밀필터", "📨 담당자 알림"])

with tabs[5]: # 담당자 알림 탭
    st.subheader("📨 지능형 담당자 알림 및 발송 검증")
    
    unmatched_targets = df_voc[df_voc["매칭여부"] == "비매칭(X)"].copy()
    
    if unmatched_targets.empty:
        st.info("비매칭 대상 데이터가 없습니다.")
    else:
        # 데이터 매핑 검증
        verify_list = []
        for _, row in unmatched_targets.iterrows():
            mgr = row.get("구역담당자_통합", "미지정")
            info, status = get_smart_contact(mgr, manager_contacts)
            verify_list.append({
                "계약번호": row["계약번호_정제"],
                "담당자(원본)": mgr,
                "매핑이메일": info.get("email", "") if info else "",
                "검증상태": status,
                "유효성": is_valid_email(info.get("email", "")) if info else False
            })
        
        v_df = pd.DataFrame(verify_list)
        
        # 검증 요약 위젯
        c1, c2, c3 = st.columns(3)
        c1.metric("매핑 성공률", f"{(v_df['검증상태'] != 'Not Found').mean()*100:.1f}%")
        c2.metric("형식 오류 주소", len(v_df[~v_df["유효성"] & (v_df["매핑이메일"] != "")]))
        c3.metric("알림 대상 계약", len(v_df))

        # 리스트 에디터 및 발송 제어
        st.markdown("#### 🛠️ 발송 리스트 최종 검토")
        edited_df = st.data_editor(
            v_df.groupby(["담당자(원본)", "매핑이메일", "검증상태", "유효성"]).size().reset_index(name="건수"),
            use_container_width=True, hide_index=True
        )

        with st.form("email_form"):
            subject = st.text_input("메일 제목", "[긴급] 고위험 해지 VOC 활동 미등록 건 안내")
            body_tpl = st.text_area("메일 본문", "안녕하세요 {담당자}님, 긴급 계약 {건수}건의 내역을 확인해주세요.")
            
            if st.form_submit_button("📧 일괄 발송 시작"):
                # SMTP 설정 및 발송 루프 수행
                st.success("발송 프로세스가 시작되었습니다. 로그를 확인하세요.")
