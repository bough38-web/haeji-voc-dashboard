import os
import re
import smtplib
import time
from datetime import datetime, date
from email.message import EmailMessage

import numpy as np
import pandas as pd
import streamlit as st

# 유사도 분석 엔진 (Fuzzy Matching)
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
# 1. 유틸리티 함수 (매핑 & 유효성 검사)
# ----------------------------------------------------

def is_valid_email(email):
    """이메일 정규식 유효성 검사"""
    if not email: return False
    regex = r'^[a-zA-Z0-9+-_.]+@[a-zA-Z0-9-]+\.[a-zA-Z0-9-.]+$'
    return bool(re.match(regex, str(email)))

def get_smart_contact(target_name, contact_dict):
    """Fuzzy Matching: '처리자1'과 VOC 파일 이름을 지능적으로 매핑"""
    target_name = str(target_name).strip()
    if not target_name or target_name in ["nan", "미지정"]: return None, "Name Missing"
    
    if target_name in contact_dict:
        return contact_dict[target_name], "Verified"
    
    if HAS_RAPIDFUZZ:
        choices = list(contact_dict.keys())
        result = process.extractOne(target_name, choices, processor=utils.default_process)
        if result and result[1] >= 85: # 유사도 기준을 85%로 완화
            suggested_name = result[0]
            return contact_dict[suggested_name], f"Suggested({suggested_name})"
    
    return None, "Not Found"

# ----------------------------------------------------
# 2. 데이터 로드 및 전처리 (매핑 파일 컬럼 수정)
# ----------------------------------------------------

st.set_page_config(page_title="해지 VOC 종합 대시보드 Pro", layout="wide")

MERGED_PATH = "merged.xlsx"
CONTACT_PATH = "contact_map.xlsx"
LOG_PATH = "email_log.csv"

@st.cache_data
def load_and_prep_data():
    if not os.path.exists(MERGED_PATH): return pd.DataFrame()
    df = pd.read_excel(MERGED_PATH)
    # 기본 정제 로직 (사용자 코드 반영)
    df["계약번호_정제"] = df["계약번호"].astype(str).str.replace(r"[^0-9A-Za-z]", "", regex=True)
    df["접수일시"] = pd.to_datetime(df["접수일시"], errors="coerce")
    
    # 지사명 표준화
    mapping = {"중앙지사": "중앙", "강북지사": "강북", "서대문지사": "서대문", "고양지사": "고양", "의정부지사": "의정부"}
    if "관리지사" in df.columns:
        df["관리지사"] = df["관리지사"].replace(mapping)
    
    # 담당자 통합 (기존 처리자 정보 사용)
    def pick_mgr(row):
        for c in ["처리자", "구역담당자", "담당자"]:
            if c in row and pd.notna(row[c]): return str(row[c]).strip()
        return "미지정"
    df["구역담당자_통합"] = df.apply(pick_mgr, axis=1)
    
    return df

@st.cache_data
def load_contacts(path):
    if not os.path.exists(path): return pd.DataFrame(), {}
    df_c = pd.read_excel(path)
    
    # [핵심 수정] 사용자가 지정한 "처리자1" 컬럼 탐지
    name_col = next((c for c in df_c.columns if "처리자1" in str(c)), df_c.columns[0])
    email_col = next((c for c in df_c.columns if "이메일" in str(c) or "메일" in str(c)), df_c.columns[1])
    
    contact_dict = {}
    for _, row in df_c.iterrows():
        name = str(row[name_col]).strip()
        if name:
            contact_dict[name] = {"email": str(row[email_col]).strip()}
    return df_c, contact_dict

df_all = load_and_prep_data()
contact_raw, manager_contacts = load_contacts(CONTACT_PATH)

# 비매칭 고위험 필터링 (unmatched_global 정의)
# 실제 전처리 로직이 복잡하므로 여기서는 df_all을 기반으로 비매칭 시뮬레이션
# 실무 코드에서는 매칭 여부(O/X) 계산된 데이터프레임을 사용하세요.
unmatched_global = df_all.copy() 
unmatched_global["리스크등급"] = "HIGH" # 시연용 강제 할당

# ----------------------------------------------------
# 3. 메인 대시보드 탭 레이아웃
# ----------------------------------------------------

tabs = st.tabs(["📊 시각화", "📘 VOC 전체", "🧯 비매칭", "🔍 활동등록", "🎯 정밀필터", "📨 담당자 알림"])

with tabs[5]: # 담당자 알림 탭
    st.subheader("📨 지능형 담당자 알림 및 발송 데이터 검증")
    
    if df_all.empty or not manager_contacts:
        st.warning(f"⚠️ {CONTACT_PATH} 파일을 확인해 주세요. 컬럼명 '처리자1'이 있는지 확인이 필요합니다.")
    else:
        targets = unmatched_global[unmatched_global["리스크등급"] == "HIGH"].copy()
        
        st.info(f"🔍 매핑 엔진 가동 중: '처리자1' 기준으로 {len(manager_contacts)}명의 명단을 대조합니다.")
        
        verify_list = []
        for _, row in targets.iterrows():
            mgr_name = row["구역담당자_통합"]
            contact_info, v_status = get_smart_contact(mgr_name, manager_contacts)
            email = contact_info.get("email", "") if contact_info else ""
            
            verify_list.append({
                "계약번호": row["계약번호_정제"],
                "지사": row.get("관리지사", "-"),
                "담당자": mgr_name,
                "매핑이메일": email,
                "검증상태": v_status,
                "유효주소": is_valid_email(email)
            })
        
        v_df = pd.DataFrame(verify_list)
        
        # 상단 KPI
        c1, c2, c3 = st.columns(3)
        with c1: st.metric("담당자 매핑률", f"{(v_df['검증상태'].str.contains('Verified|Suggested')).mean()*100:.1f}%")
        with c2: st.metric("유효 이메일", v_df["유효주소"].sum())
        with c3: st.metric("대상 계약수", len(v_df))

        st.markdown("---")

        # 발송 리스트 데이터 편집 (Groupby 적용)
        agg_targets = v_df.groupby(["지사", "담당자", "매핑이메일", "검증상태", "유효주소"]).size().reset_index(name="대상 건수")
        
        edited_agg = st.data_editor(
            agg_targets,
            column_config={
                "매핑이메일": st.column_config.TextColumn("이메일(수정가능)", required=True),
                "검증상태": st.column_config.TextColumn("매핑 상태", disabled=True),
                "대상 건수": st.column_config.NumberColumn("건수", disabled=True),
                "유효주소": st.column_config.CheckboxColumn("유효", disabled=True)
            },
            use_container_width=True, hide_index=True
        )

        # 발송 버튼
        if st.button("📧 리스트 확정 및 발송 준비", type="primary", use_container_width=True):
            st.success("데이터가 확정되었습니다. 아래 SMTP 설정을 확인 후 발송하세요.")
