import os
import re
import smtplib
from datetime import datetime, date
from email.message import EmailMessage

import numpy as np
import pandas as pd
import streamlit as st

# 

[Image of data mapping verification flow chart]


# 1. 라이브러리 체크 및 자동 폴백(Fallback) 설정
try:
    from rapidfuzz import process, utils
    HAS_RAPIDFUZZ = True
except ImportError:
    HAS_RAPIDFUZZ = False

# ----------------------------------------------------
# 2. 지능형 매핑 및 유효성 검사 유틸리티
# ----------------------------------------------------

def is_valid_email(email):
    """이메일 정규식 유효성 검사"""
    if not email: return False
    regex = r'^[a-zA-Z0-9+-_.]+@[a-zA-Z0-9-]+\.[a-zA-Z0-9-.]+$'
    return bool(re.match(regex, str(email)))

def get_smart_contact(target_name, contact_dict):
    """라이브러리가 있으면 Fuzzy Matching, 없으면 정확한 일치만 수행"""
    target_name = str(target_name).strip()
    
    # 정확 일치 확인
    if target_name in contact_dict:
        return contact_dict[target_name], "Verified"
    
    # 유사도 분석 (라이브러리 설치된 경우에만 작동)
    if HAS_RAPIDFUZZ:
        choices = list(contact_dict.keys())
        result = process.extractOne(target_name, choices, processor=utils.default_process)
        if result and result[1] >= 90:
            suggested_name = result[0]
            return contact_dict[suggested_name], f"Suggested({suggested_name})"
    
    return None, "Not Found"

# ----------------------------------------------------
# 3. 발송 히스토리 로깅 (CSV 기록)
# ----------------------------------------------------
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
# 4. 담당자 알림 탭 상세 구현
# ----------------------------------------------------
# [주의: tab_alert는 st.tabs() 객체 중 하나여야 함]
def render_alert_tab(tab_alert, unmatched_df, manager_contacts):
    with tab_alert:
        st.subheader("📨 담당자 알림 및 발송 데이터 검증")
        
        # 비매칭 고위험 대상 추출
        targets = unmatched_df[unmatched_df["리스크등급"] == "HIGH"].copy()
        
        if targets.empty:
            st.success("🎉 현재 발송 대상(비매칭 고위험)이 없습니다.")
            return

        # 무결성 검증 수행
        verify_list = []
        for _, row in targets.iterrows():
            mgr_name = row["구역담당자_통합"]
            info, status = get_smart_contact(mgr_name, manager_contacts)
            email = info.get("email", "") if info else ""
            
            verify_list.append({
                "계약번호": row["계약번호_정제"],
                "지사": row["관리지사"],
                "담당자": mgr_name,
                "매핑이메일": email,
                "상태": status,
                "유효성": is_valid_email(email)
            })
        
        v_df = pd.DataFrame(verify_list)
        
        # 지표 시각화
        c1, c2, c3 = st.columns(3)
        c1.metric("담당자 매핑률", f"{(v_df['상태'] != 'Not Found').mean()*100:.1f}%")
        c2.metric("형식 오류 주소", len(v_df[~v_df["유효성"] & (v_df["매핑이메일"] != "")]))
        c3.metric("알림 대상 계약", len(v_df))

        # 데이터 에디터 및 발송 제어
        edited_df = st.data_editor(v_df, use_container_width=True, hide_index=True)
        
        # (이후 발송 로직 및 로그 호출 생략)
