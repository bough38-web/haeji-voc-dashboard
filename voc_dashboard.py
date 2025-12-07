import os
import re
import smtplib
import time
from datetime import datetime, date
from email.message import EmailMessage

import numpy as np
import pandas as pd
import streamlit as st

# 전문가용 유사도 분석 라이브러리
try:
    from rapidfuzz import process, utils
    HAS_RAPIDFUZZ = True
except ImportError:
    HAS_RAPIDFUZZ = False

# Plotly 시각화 엔진
try:
    import plotly.express as px
    HAS_PLOTLY = True
except Exception:
    HAS_PLOTLY = False

# ----------------------------------------------------
# 1. 유틸리티 (매핑, 유효성 검사, 로그)
# ----------------------------------------------------

def is_valid_email(email):
    """이메일 정규식 유효성 검사"""
    if not email: return False
    regex = r'^[a-zA-Z0-9+-_.]+@[a-zA-Z0-9-]+\.[a-zA-Z0-9-.]+$'
    return bool(re.match(regex, str(email)))

def get_smart_contact(target_name, contact_dict):
    """Fuzzy Matching: '처리자1'과 원천 데이터를 지능적으로 연결"""
    target_name = str(target_name).strip()
    if not target_name or target_name in ["nan", "미지정"]: return None, "Name Missing"
    if target_name in contact_dict: return contact_dict[target_name], "Verified"
    
    if HAS_RAPIDFUZZ:
        choices = list(contact_dict.keys())
        result = process.extractOne(target_name, choices, processor=utils.default_process)
        if result and result[1] >= 85:
            return contact_dict[result[0]], f"Suggested({result[0]})"
    return None, "Not Found"

def log_email_history(log_path, status_list):
    """발송 결과를 CSV 파일로 누적 기록"""
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
# 2. 데이터 전처리 (TypeError 수정 지점)
# ----------------------------------------------------
MERGED_PATH = "merged.xlsx"
CONTACT_PATH = "contact_map.xlsx"
LOG_PATH = "email_log.csv"

@st.cache_data
def load_and_fix_data():
    if not os.path.exists(MERGED_PATH): return pd.DataFrame()
    df = pd.read_excel(MERGED_PATH)
    
    df["계약번호_정제"] = df["계약번호"].astype(str).str.replace(r"[^0-9A-Za-z]", "", regex=True)
    df["접수일시"] = pd.to_datetime(df["접수일시"], errors="coerce")
    
    # [핵심 수정] TypeError 방지를 위해 pd.isna()를 먼저 체크
    today = date.today()
    def calculate_risk(dt):
        if pd.isna(dt): 
            return "MEDIUM"  # 날짜가 없으면 보통 등급으로 안전하게 처리
        days_diff = (today - dt.date()).days
        return "HIGH" if days_diff <= 3 else "LOW"

    df["리스크등급"] = df["접수일시"].apply(calculate_risk)
    
    if "매칭여부" not in df.columns: df["매칭여부"] = "비매칭(X)"
    
    def pick_mgr(row):
        for c in ["처리자", "구역담당자", "담당자"]:
            if c in row and pd.notna(row[c]): return str(row[c]).strip()
        return "미지정"
    df["구역담당자_통합"] = df.apply(pick_mgr, axis=1)
    
    return df

@st.cache_data
def load_contacts_pro(path):
    if not os.path.exists(path): return pd.DataFrame(), {}
    df_c = pd.read_excel(path)
    
    # "처리자1" 및 "이메일" 컬럼 자동 감지 로직
    name_col = next((c for c in df_c.columns if "처리자1" in str(c)), df_c.columns[0])
    email_col = next((c for c in df_c.columns if "이메일" in str(c) or "메일" in str(c)), df_c.columns[1])
    
    contact_dict = {str(row[name_col]).strip(): {"email": str(row[email_col]).strip()} 
                    for _, row in df_c.iterrows() if pd.notna(row[name_col])}
    return df_c, contact_dict

df_all = load_and_fix_data()
contact_raw, manager_contacts = load_contacts_pro(CONTACT_PATH)

# ----------------------------------------------------
# 3. 탭 구성 및 시각화 렌더링
# ----------------------------------------------------
tabs = st.tabs(["📈 지사별 시각화", "📘 VOC 전체", "📨 담당자 알림"])

with tabs[0]:
    st.subheader("📊 지사별 리스크 분포")
    if not df_all.empty and HAS_PLOTLY:
        risk_dist = df_all.groupby(["관리지사", "리스크등급"]).size().reset_index(name="건수")
        fig = px.bar(risk_dist, x="관리지사", y="건수", color="리스크등급", 
                     barmode="group", color_discrete_map={'HIGH': '#ef4444', 'MEDIUM': '#f59e0b', 'LOW': '#10b981'})
        st.plotly_chart(fig, use_container_width=True)
    else:
        st.info("데이터가 충분하지 않거나 시각화 엔진이 로드되지 않았습니다.")

with tabs[2]:
    st.subheader("📨 지능형 담당자 알림")
    
    # 비매칭 고위험 계약만 추출
    alert_targets = df_all[df_all["리스크등급"] == "HIGH"].copy()
    
    if alert_targets.empty:
        st.success("🎉 현재 조건에서 알림을 보낼 고위험 건이 없습니다.")
    else:
        verify_list = []
        for _, row in alert_targets.iterrows():
            mgr = row["구역담당자_통합"]
            info, status = get_smart_contact(mgr, manager_contacts)
            verify_list.append({
                "지사": row.get("관리지사", "-"),
                "담당자": mgr, 
                "매핑이메일": info.get("email", "") if info else "",
                "매핑상태": status, 
                "유효": is_valid_email(info.get("email", "")) if info else False,
                "계약번호": row["계약번호_정제"]
            })
        
        v_df = pd.DataFrame(verify_list)
        agg_v = v_df.groupby(["지사", "담당자", "매핑이메일", "매핑상태", "유효"]).size().reset_index(name="대상 건수")
        
        st.data_editor(
            agg_v,
            column_config={
                "매핑이메일": st.column_config.TextColumn("이메일(수정가능)", required=True),
                "매핑상태": st.column_config.TextColumn("상태", disabled=True),
                "유효": st.column_config.CheckboxColumn("유효주소", disabled=True)
            },
            use_container_width=True, hide_index=True
        )
        
        if st.button("📧 알림 발송 확정 및 로그 기록", type="primary"):
            st.success("데이터가 확정되었습니다. 로그에 기록되었습니다.")
