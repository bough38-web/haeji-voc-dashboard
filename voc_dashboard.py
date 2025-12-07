import os
import re
import smtplib
import time
from datetime import datetime, date
from email.message import EmailMessage

import numpy as np
import pandas as pd
import streamlit as st

# ----------------------------------------------------
# 1. 고급 라이브러리 로드 (Fuzzy Matching & Visualization)
# ----------------------------------------------------
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
# 2. 공통 유틸리티 함수 (매핑, 검증, 로그)
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
# 3. 데이터 로딩 및 초기 설정
# ----------------------------------------------------
st.set_page_config(page_title="해지 VOC 종합 대시보드 Pro", layout="wide")

MERGED_PATH = "merged.xlsx"
CONTACT_PATH = "contact_map.xlsx"
FEEDBACK_PATH = "feedback.csv"
LOG_PATH = "email_log.csv"

# SMTP 설정 (Secrets 로드)
SMTP_HOST = st.secrets.get("SMTP_HOST", "")
SMTP_PORT = int(st.secrets.get("SMTP_PORT", 587))
SMTP_USER = st.secrets.get("SMTP_USER", "")
SMTP_PASSWORD = st.secrets.get("SMTP_PASSWORD", "")
SENDER_NAME = st.secrets.get("SENDER_NAME", "해지VOC 관리자")

@st.cache_data
def load_voc_data(path):
    if not os.path.exists(path): return pd.DataFrame()
    df = pd.read_excel(path)
    if "계약번호" in df.columns:
        df["계약번호_정제"] = df["계약번호"].astype(str).str.replace(r"[^0-9A-Za-z]", "", regex=True).str.strip()
    if "접수일시" in df.columns:
        df["접수일시"] = pd.to_datetime(df["접수일시"], errors="coerce")
    if "관리지사" in df.columns:
        mapping = {"중앙지사":"중앙", "강북지사":"강북", "서대문지사":"서대문", "고양지사":"고양", "의정부지사":"의정부", "남양주지사":"남양주"}
        df["관리지사"] = df["관리지사"].replace(mapping)
    return df

@st.cache_data
def load_contact_map(path):
    if not os.path.exists(path): return pd.DataFrame(), {}
    df_c = pd.read_excel(path)
    contact_dict = {}
    # 첫번째 컬럼: 담당자, 두번째 컬럼: 이메일로 가정 (detect_column 로직 대체)
    for _, row in df_c.iterrows():
        name = str(row[0]).strip()
        if name: contact_dict[name] = {"email": str(row[1]).strip()}
    return df_c, contact_dict

# 데이터 실제 로드
df = load_voc_data(MERGED_PATH)
contact_df, manager_contacts = load_contact_map(CONTACT_PATH)

# 필터링 로직 (비매칭 리스트 생성)
df_voc = df[df["출처"] == "해지VOC"].copy()
# (간략화된 글로벌 필터 적용)
unmatched_global = df_voc[df_voc["매칭여부"].isna()].copy() # 실제 조건에 맞게 수정 필요

# ----------------------------------------------------
# 4. 메인 탭 구성
# ----------------------------------------------------
tab_viz, tab_all, tab_unmatched, tab_drill, tab_filter, tab_alert = st.tabs(
    ["📊 시각화", "📘 VOC 전체", "🧯 비매칭", "🔍 활동등록", "🎯 정밀필터", "📨 담당자 알림"]
)

with tab_alert:
    st.subheader("📨 지능형 담당자 알림 시스템")
    
    if not manager_contacts:
        st.warning("⚠️ 담당자 매핑 파일(contact_map.xlsx)이 필요합니다.")
    else:
        # 고위험 비매칭 리스트 필터 (임시 조건)
        targets = unmatched_global.head(20) # 실제 리스크 필터 적용 필요
        
        st.info("🔍 담당자 매핑 및 이메일 무결성 검증 프로세스 실행")
        
        # 검증 리스트 생성
        verify_list = []
        for _, row in targets.iterrows():
            mgr_name = row.get("구역담당자_통합", "미지정")
            info, status = get_smart_contact(mgr_name, manager_contacts)
            email = info.get("email", "") if info else ""
            
            verify_list.append({
                "계약번호": row.get("계약번호_정제", "-"),
                "지사": row.get("관리지사", "-"),
                "담당자": mgr_name,
                "이메일": email,
                "상태": status,
                "유효성": is_valid_email(email)
            })
        
        v_df = pd.DataFrame(verify_list)
        
        # 지표 대시보드
        c1, c2, c3 = st.columns(3)
        with c1: st.metric("매핑 성공률", f"{(v_df['상태'] != 'Not Found').mean()*100:.1f}%")
        with c2: st.metric("주소 형식 오류", len(v_df[~v_df["유효성"] & (v_df["이메일"] != "")]))
        with c3: st.metric("발송 예정 건수", len(v_df))

        st.markdown("---")

        # 발송 리스트 검토
        edited_agg = st.data_editor(
            v_df.groupby(["지사", "담당자", "이메일", "상태", "유효성"]).size().reset_index(name="건수"),
            use_container_width=True, key="email_batch_editor", hide_index=True
        )

        # 발송 폼
        with st.form("alert_send_form"):
            subject = st.text_input("제목", "[긴급] 해지방어 활동 미등록 건 확인 요청")
            body_tpl = st.text_area("본문", "안녕하세요, {담당자}님. 긴급 고위험 계약 {건수}건을 확인하세요.")
            dry_run = st.toggle("모의 발송 (로그만 기록)", value=True)
            
            if st.form_submit_button("📧 일괄 발송 시작"):
                progress = st.progress(0)
                status_log = []
                for i, row in edited_agg.iterrows():
                    mgr, dest, cnt = row["담당자"], row["이메일"], row["건수"]
                    log_entry = {"time": datetime.now(), "target": mgr, "email": dest, "cnt": cnt, "mode": "Dry" if dry_run else "Actual"}
                    
                    if not dest or not row["유효성"]:
                        log_entry["result"] = "FAIL(Address)"
                    else:
                        try:
                            if not dry_run:
                                # Email 전송 로직
                                msg = EmailMessage()
                                msg["To"] = dest
                                msg["Subject"] = subject
                                msg.set_content(body_tpl.format(담당자=mgr, 건수=cnt))
                                with smtplib.SMTP(SMTP_HOST, SMTP_PORT) as s:
                                    s.starttls()
                                    s.login(SMTP_USER, SMTP_PASSWORD)
                                    s.send_message(msg)
                            log_entry["result"] = "SUCCESS"
                        except Exception as e:
                            log_entry["result"] = f"ERROR({str(e)})"
                    
                    status_log.append(log_entry)
                    progress.progress((i+1)/len(edited_agg))
                
                log_email_history(LOG_PATH, status_log)
                st.success(f"처리 완료! 로그가 {LOG_PATH}에 저장되었습니다.")

        if os.path.exists(LOG_PATH):
            st.markdown("#### 📊 발송 히스토리")
            st.dataframe(pd.read_csv(LOG_PATH).tail(10), use_container_width=True)
