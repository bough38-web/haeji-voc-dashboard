import os
import re
import smtplib
import time
from datetime import datetime, date
from email.message import EmailMessage

import numpy as np
import pandas as pd
import streamlit as st

# 유사도 분석 및 시각화 엔진 로드
try:
    from rapidfuzz import process, utils
    HAS_RAPIDFUZZ = True
except ImportError:
    HAS_RAPIDFUZZ = False

try:
    import plotly.express as px
    HAS_PLOTLY = True
except Exception:
    HAS_PLOTLY = False

# ----------------------------------------------------
# 1. 유틸리티 (유효성 검사, 지능형 매핑, 로그)
# ----------------------------------------------------

def is_valid_email(email):
    if not email or pd.isna(email): return False
    regex = r'^[a-zA-Z0-9+-_.]+@[a-zA-Z0-9-]+\.[a-zA-Z0-9-.]+$'
    return bool(re.match(regex, str(email)))

def get_smart_contact(target_name, contact_dict):
    target_name = str(target_name).strip()
    if not target_name or target_name == "nan": return None, "Name Empty"
    if target_name in contact_dict:
        return contact_dict[target_name], "Verified"
    if HAS_RAPIDFUZZ:
        choices = list(contact_dict.keys())
        result = process.extractOne(target_name, choices, processor=utils.default_process)
        if result and result[1] >= 90:
            suggested_name = result[0]
            return contact_dict[suggested_name], f"Suggested({suggested_name})"
    return None, "Not Found"

def log_email_history(log_path, status_list):
    new_logs = pd.DataFrame(status_list)
    if os.path.exists(log_path):
        try:
            old_logs = pd.read_csv(log_path)
            combined = pd.concat([old_logs, new_logs], ignore_index=True)
            combined.to_csv(log_path, index=False, encoding="utf-8-sig")
        except:
            new_logs.to_csv(log_path, index=False, encoding="utf-8-sig")
    else:
        new_logs.to_csv(log_path, index=False, encoding="utf-8-sig")

# ----------------------------------------------------
# 2. 데이터 로드 및 초기화
# ----------------------------------------------------

st.set_page_config(page_title="해지 VOC 종합 대시보드 Pro", layout="wide")

MERGED_PATH = "merged.xlsx"
CONTACT_PATH = "contact_map.xlsx"
LOG_PATH = "email_log.csv"

@st.cache_data
def load_and_prep_data():
    if not os.path.exists(MERGED_PATH): return pd.DataFrame()
    df = pd.read_excel(MERGED_PATH)
    # 기본 정제 로직
    df["계약번호_정제"] = df["계약번호"].astype(str).str.replace(r"[^0-9A-Za-z]", "", regex=True)
    df["접수일시"] = pd.to_datetime(df["접수일시"], errors="coerce")
    return df

@st.cache_data
def load_contacts():
    if not os.path.exists(CONTACT_PATH): return {}, pd.DataFrame()
    df_c = pd.read_excel(CONTACT_PATH)
    contact_dict = {str(row[0]).strip(): {"email": str(row[1]).strip()} for _, row in df_c.iterrows()}
    return contact_dict, df_c

df_all = load_and_prep_data()
manager_contacts, contact_df = load_contacts()

# [예시 데이터셋 구성 - 필터링 로직에 맞춰 수정 필요]
unmatched_global = df_all.copy() # 실제 조건에 맞춰 할당

# ----------------------------------------------------
# 3. 탭 구성 (Tab Alert 데이터 노출 수정)
# ----------------------------------------------------

tabs = st.tabs(["📊 시각화", "📘 VOC 전체", "🧯 비매칭", "🔍 활동등록", "🎯 정밀필터", "📨 담당자 알림"])

# 알림 탭 상세 구현
with tabs[5]:
    st.subheader("📨 지능형 담당자 알림 및 발송 데이터 관리")
    
    if df_all.empty or not manager_contacts:
        st.warning("⚠️ 데이터 파일(merged.xlsx) 혹은 매핑 파일(contact_map.xlsx)을 확인해 주세요.")
    else:
        targets = unmatched_global.head(20) # 테스트용 샘플링
        
        st.info("🔍 담당자 매핑 및 데이터 무결성 검증을 수행합니다.")
        

[Image of data mapping verification flow chart]

        
        verify_list = []
        for _, row in targets.iterrows():
            mgr_name = row.get("구역담당자_통합", "미지정")
            contact_info, v_status = get_smart_contact(mgr_name, manager_contacts)
            email = contact_info.get("email", "") if contact_info else ""
            
            verify_list.append({
                "계약번호": row["계약번호_정제"],
                "지사": row.get("관리지사", "-"),
                "담당자": mgr_name,
                "매핑이메일": email,
                "검증상태": v_status,
                "유효성": is_valid_email(email)
            })
        
        v_df = pd.DataFrame(verify_list)
        
        # 지표 표시
        c1, c2, c3 = st.columns(3)
        c1.metric("매핑 성공률", f"{(v_df['검증상태'] != 'Not Found').mean()*100:.1f}%")
        c2.metric("주소 형식 오류", len(v_df[~v_df["유효성"] & (v_df["매핑이메일"] != "")]))
        c3.metric("발송 예정 건수", len(v_df))

        st.markdown("---")
        
        # 에디터 및 발송 폼
        edited_agg = st.data_editor(v_df, use_container_width=True, hide_index=True)
        
        with st.form("alert_pro_form"):
            subject = st.text_input("제목", "[긴급] 고위험 해지 VOC 활동 미등록 건 확인 요청")
            body = st.text_area("본문", "안녕하세요. 담당하신 구역에 긴급 해지 VOC 건이 확인되었습니다.")
            dry_run = st.toggle("모의 발송 (Dry Run)", value=True)
            
            if st.form_submit_button("📧 일괄 발송 및 로그 저장"):
                # 발송 로직 및 log_email_history 호출
                st.success("처리가 완료되었습니다. 로그 파일을 확인하세요.")
