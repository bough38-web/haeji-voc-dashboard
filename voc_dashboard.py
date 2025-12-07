import os
import re
import smtplib
from datetime import datetime, date
from email.message import EmailMessage
import time

import numpy as np
import pandas as pd
import streamlit as st

# Plotly 및 유사도 분석 엔진 로드 (Fallback 설정)
try:
    import plotly.express as px
    HAS_PLOTLY = True
except ImportError:
    HAS_PLOTLY = False

# rapidfuzz는 requirements.txt에 추가가 필요할 수 있습니다.
try:
    from rapidfuzz import process, utils
    HAS_RAPIDFUZZ = True
except ImportError:
    HAS_RAPIDFUZZ = False

# ----------------------------------------------------
# 1. 유틸리티 함수 (이메일 검증 및 매핑)
# ----------------------------------------------------
def is_valid_email(email):
    """이메일 정규식 유효성 검사"""
    if not email or pd.isna(email): return False
    regex = r'^[a-zA-Z0-9+-_.]+@[a-zA-Z0-9-]+\.[a-zA-Z0-9-.]+$'
    return bool(re.match(regex, str(email)))

def get_smart_contact(target_name, contact_dict):
    """Fuzzy Matching: 담당자 이름 오타나 직급 차이를 지능적으로 매핑"""
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
# 2. 데이터 로딩 및 초기 설정 (KeyError 방지)
# ----------------------------------------------------
st.set_page_config(page_title="해지 VOC 종합 대시보드 Pro", layout="wide")

MERGED_PATH = "merged.xlsx"
CONTACT_PATH = "contact_map.xlsx"
LOG_PATH = "email_log.csv"

# [데이터 로딩 로직 - 기존 load_voc_data 등 함수 유지 필수]
# unmatched_global 데이터가 전처리 단계에서 정의되었다고 가정합니다.

# ----------------------------------------------------
# 3. 메인 탭 구성 (NameError 해결 지점)
# ----------------------------------------------------
# 탭을 변수에 할당하여 정의합니다.
tabs = st.tabs(["📊 시각화", "📘 VOC 전체", "🧯 비매칭", "🔍 활동등록", "🎯 정밀필터", "📨 담당자 알림"])

# 알림 탭 (마지막 탭) 로직 구현
with tabs[5]:
    st.subheader("📨 지능형 담당자 알림 및 발송 데이터 관리")
    
    # manager_contacts는 load_contact_map 함수로 생성된 딕셔너리여야 함
    if 'manager_contacts' not in locals() or not manager_contacts:
        st.warning("⚠️ 담당자 매핑 파일(contact_map.xlsx)을 로드하지 못했습니다.")
    else:
        # 비매칭 고위험 계약 데이터 추출
        # 'unmatched_global' 변수가 코드 상단에서 정의되어 있는지 확인하세요.
        try:
            alert_targets = unmatched_global[unmatched_global["리스크등급"] == "HIGH"].copy()
        except NameError:
            st.error("unmatched_global 데이터가 정의되지 않았습니다. 상단 필터 로직을 확인하세요.")
            alert_targets = pd.DataFrame()

        if alert_targets.empty:
            st.success("🎉 현재 알림 발송 대상(비매칭 고위험)이 없습니다.")
        else:
            st.info("🔍 담당자 매핑 및 이메일 무결성 검증을 수행합니다.")
            
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
                    "매핑상태": v_status,
                    "유효성": is_valid_email(email)
                })
            
            v_df = pd.DataFrame(verify_list)
            
            # 지표 표시 (무결성 통계)
            col_v1, col_v2, col_v3 = st.columns(3)
            with col_v1:
                st.metric("담당자 매핑률", f"{(v_df['매핑상태'] != 'Not Found').mean()*100:.1f}%")
            with col_v2:
                invalid_cnt = v_df[~v_df["유효성"] & (v_df["매핑이메일"] != "")].shape[0]
                st.metric("형식 오류 주소", f"{invalid_cnt}건", delta_color="inverse")
            with col_v3:
                st.metric("발송 예정 건수", f"{len(v_df)}건")

            st.markdown("---")

            # 리스트 에디터 및 발송 제어
            st.markdown("#### 🛠️ 발송 데이터 최종 확인")
            edited_agg = st.data_editor(
                v_df.groupby(["지사", "담당자", "매핑이메일", "매핑상태", "유효성"]).size().reset_index(name="건수"),
                use_container_width=True, hide_index=True, key="alert_batch_editor"
            )

            with st.form("alert_send_form"):
                subject = st.text_input("제목", "[긴급] 해지방어 활동 미등록 건 확인 요청")
                body_tpl = st.text_area("본문", "안녕하세요, {담당자}님. 긴급 고위험 계약 {건수}건을 확인하세요.")
                dry_run = st.toggle("모의 발송 (로그만 기록)", value=True)
                
                if st.form_submit_button("📧 일괄 발송 시작"):
                    st.info("메일 발송 기능을 활성화하려면 SMTP 설정이 필요합니다.")
                    # 발송 이력 로그를 저장하는 log_email_history 함수 호출 등의 로직 추가 가능
