import os
import re
import smtplib
import time
from datetime import datetime, date
from email.message import EmailMessage

import numpy as np
import pandas as pd
import streamlit as st

# 전문가용 유사도 분석 라이브러리 (pip install rapidfuzz 필수)
try:
    from rapidfuzz import process, utils
    HAS_RAPIDFUZZ = True
except ImportError:
    HAS_RAPIDFUZZ = False

# Plotly 시각화 라이브러리
try:
    import plotly.express as px
    HAS_PLOTLY = True
except ImportError:
    HAS_PLOTLY = False

# ----------------------------------------------------
# 1. 유틸리티 함수 (매핑 검증 및 이메일 유효성)
# ----------------------------------------------------
def is_valid_email(email):
    """이메일 정규식 유효성 검사"""
    if not email or pd.isna(email): return False
    regex = r'^[a-zA-Z0-9+-_.]+@[a-zA-Z0-9-]+\.[a-zA-Z0-9-.]+$'
    return bool(re.match(regex, str(email)))

def get_smart_contact(target_name, contact_dict):
    """Fuzzy Matching: 오타나 직급 차이를 지능적으로 매핑"""
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

# ----------------------------------------------------
# 2. 데이터 로딩 및 전처리 (KeyError 방지)
# ----------------------------------------------------
st.set_page_config(page_title="해지 VOC 종합 대시보드 Pro", layout="wide")

# 파일 경로 정의
MERGED_PATH = "merged.xlsx"
CONTACT_PATH = "contact_map.xlsx"
LOG_PATH = "email_log.csv"

@st.cache_data
def load_all_data():
    if not os.path.exists(MERGED_PATH): return pd.DataFrame(), pd.DataFrame()
    df = pd.read_excel(MERGED_PATH)
    
    # 1. 컬럼 정제
    if "계약번호" in df.columns:
        df["계약번호_정제"] = df["계약번호"].astype(str).str.replace(r"[^0-9A-Za-z]", "", regex=True).str.strip()
    
    # 2. 매칭여부 컬럼 생성 (KeyError 해결 핵심)
    df_voc = df[df.get("출처") == "해지VOC"].copy()
    df_other = df[df.get("출처") != "해지VOC"].copy()
    other_contract_set = set(df_other["계약번호_정제"].dropna().unique())
    df_voc["매칭여부"] = df_voc["계약번호_정제"].apply(
        lambda x: "매칭(O)" if x in other_contract_set else "비매칭(X)"
    )
    
    # 3. 리스크 등급 계산
    today = date.today()
    df_voc["접수일시"] = pd.to_datetime(df_voc.get("접수일시"), errors="coerce")
    df_voc["리스크등급"] = df_voc["접수일시"].apply(
        lambda dt: "HIGH" if pd.notna(dt) and (today - dt.date()).days <= 3 else "LOW"
    )
    
    # 4. 담당자 통합
    def pick_manager(row):
        for c in ["구역담당자", "담당자", "처리자"]:
            if c in row and pd.notna(row[c]) and str(row[c]).strip():
                return str(row[c]).strip()
        return "미지정"
    df_voc["구역담당자_통합"] = df_voc.apply(pick_manager, axis=1)
    
    return df_voc, df

@st.cache_data
def load_manager_map(path):
    if not os.path.exists(path): return {}
    df_c = pd.read_excel(path)
    contact_dict = {}
    for _, row in df_c.iterrows():
        name = str(row.iloc[0]).strip()
        if name: contact_dict[name] = {"email": str(row.iloc[1]).strip()}
    return contact_dict

# 데이터 로드 실행
df_voc, df_raw = load_all_data()
manager_contacts = load_manager_map(CONTACT_PATH)
unmatched_global = df_voc[df_voc["매칭여부"] == "비매칭(X)"].copy()

# ----------------------------------------------------
# 3. 메인 탭 구성 (NameError 해결)
# ----------------------------------------------------
tabs = st.tabs(["📊 시각화", "📘 VOC 전체", "🧯 비매칭", "🔍 활동등록", "🎯 정밀필터", "📨 담당자 알림"])

with tabs[5]:
    st.subheader("📨 지능형 담당자 알림 및 발송 데이터 검증")
    
    if not manager_contacts:
        st.warning("⚠️ 담당자 매핑 파일(contact_map.xlsx)이 필요합니다.")
    else:
        # 고위험 비매칭 대상 필터
        alert_targets = unmatched_global[unmatched_global["리스크등급"] == "HIGH"].copy()
        
        if alert_targets.empty:
            st.success("🎉 발송 대상(비매칭 고위험) 계약이 없습니다.")
        else:
            st.info("🔍 데이터 매핑 및 이메일 무결성 검증 프로세스를 실행합니다.")
            
            # [수정 완료] SyntaxError 유발 구문을 코드 밖으로 처리
            # 담당자별 데이터 매핑 프로세스 차트 개념 적용
            
            verify_list = []
            for _, row in alert_targets.iterrows():
                mgr_name = row.get("구역담당자_통합", "미지정")
                contact_info, v_status = get_smart_contact(mgr_name, manager_contacts)
                email = contact_info.get("email", "") if contact_info else ""
                
                verify_list.append({
                    "계약번호": row.get("계약번호_정제", "-"),
                    "지사": row.get("관리지사", "-"),
                    "담당자": mgr_name,
                    "매핑이메일": email,
                    "검증상태": v_status,
                    "유효성": is_valid_email(email)
                })
            
            v_df = pd.DataFrame(verify_list)
            
            # 무결성 요약 지표
            col_v1, col_v2, col_v3 = st.columns(3)
            with col_v1:
                st.metric("담당자 매핑률", f"{(v_df['검증상태'] != 'Not Found').mean()*100:.1f}%")
            with col_v2:
                invalid_cnt = v_df[~v_df["유효성"] & (v_df["매핑이메일"] != "")].shape[0]
                st.metric("형식 오류 주소", f"{invalid_cnt}건", delta_color="inverse")
            with col_v3:
                st.metric("발송 예정 총 계약", len(v_df))

            st.markdown("---")

            # 일괄 발송 리스트 데이터 편집기
            st.markdown("#### 🛠️ 발송 리스트 데이터 편집 및 확정")
            agg_targets = v_df.groupby(["지사", "담당자", "매핑이메일", "검증상태", "유효성"]).size().reset_index(name="건수")
            
            edited_agg = st.data_editor(
                agg_targets,
                column_config={
                    "매핑이메일": st.column_config.TextColumn("이메일(수정가능)", required=True),
                    "건수": st.column_config.NumberColumn("대상 건수", disabled=True),
                    "유효성": st.column_config.CheckboxColumn("유효 주소", disabled=True)
                },
                use_container_width=True,
                key="alert_batch_editor_verified",
                hide_index=True
            )
            
            # (이후 발송 로직 및 로그 호출은 생략 - 필요 시 추가 구현 가능)
