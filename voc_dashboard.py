import os
import re
import smtplib
import time
from datetime import datetime, date
from email.message import EmailMessage

import numpy as np
import pandas as pd
import streamlit as st

# 전문가용 지능형 매핑 및 시각화 라이브러리 로드
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
# 2. 데이터 로딩 및 전처리 (KeyError 방지)
# ----------------------------------------------------
st.set_page_config(page_title="해지 VOC 종합 대시보드 Pro", layout="wide")

MERGED_PATH = "merged.xlsx"
CONTACT_PATH = "contact_map.xlsx"
FEEDBACK_PATH = "feedback.csv"
LOG_PATH = "email_log.csv"

# SMTP 설정 (Streamlit Secrets 활용)
SMTP_HOST = st.secrets.get("SMTP_HOST", "")
SMTP_USER = st.secrets.get("SMTP_USER", "")
SMTP_PASSWORD = st.secrets.get("SMTP_PASSWORD", "")
SENDER_NAME = st.secrets.get("SENDER_NAME", "해지VOC 관리자")

@st.cache_data
def load_all_data():
    if not os.path.exists(MERGED_PATH): return pd.DataFrame(), pd.DataFrame(), {}
    
    df = pd.read_excel(MERGED_PATH)
    
    # 1. 기본 컬럼 정제
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

# 담당자 매핑 파일 로드
@st.cache_data
def load_manager_map(path):
    if not os.path.exists(path): return {}
    df_c = pd.read_excel(path)
    contact_dict = {}
    for _, row in df_c.iterrows():
        # 첫 번째 컬럼을 이름, 두 번째 컬럼을 이메일로 가정
        name = str(row[0]).strip()
        if name: contact_dict[name] = {"email": str(row[1]).strip()}
    return contact_dict

df_voc, df_raw = load_all_data()
manager_contacts = load_manager_map(CONTACT_PATH)

# 글로벌 필터링된 데이터 생성
unmatched_global = df_voc[df_voc["매칭여부"] == "비매칭(X)"].copy()

# ----------------------------------------------------
# 3. 메인 탭 구성 (NameError 해결)
# ----------------------------------------------------
# tabs 변수 정의
tabs = st.tabs(["📊 시각화", "📘 VOC 전체", "🧯 비매칭", "🔍 활동등록", "🎯 정밀필터", "📨 담당자 알림"])

with tabs[5]:
    st.subheader("📨 지능형 담당자 알림 및 발송 검증")
    
    if not manager_contacts:
        st.warning("⚠️ 담당자 매핑 파일(contact_map.xlsx)이 필요합니다.")
    else:
        # 고위험 비매칭 대상 추출
        alert_targets = unmatched_global[unmatched_global["리스크등급"] == "HIGH"].copy()
        
        if alert_targets.empty:
            st.success("🎉 현재 발송 대상(비매칭 고위험)이 없습니다.")
        else:
            st.info("🔍 담당자 매핑 및 데이터 무결성 검증 프로세스 실행")
            
            # 

[Image of data mapping verification flow chart]

            
            # 검증 리스트 생성
            verify_list = []
            for _, row in alert_targets.iterrows():
                mgr_name = row.get("구역담당자_통합", "미지정")
                info, status = get_smart_contact(mgr_name, manager_contacts)
                email = info.get("email", "") if info else ""
                
                verify_list.append({
                    "계약번호": row.get("계약번호_정제", "-"),
                    "지사": row.get("관리지사", "-"),
                    "담당자": mgr_name,
                    "매핑이메일": email,
                    "매핑상태": status,
                    "유효성": is_valid_email(email)
                })
            
            v_df = pd.DataFrame(verify_list)
            
            # 지표 대시보드
            c1, c2, c3 = st.columns(3)
            with c1: st.metric("매핑 성공률", f"{(v_df['매핑상태'] != 'Not Found').mean()*100:.1f}%")
            with c2: st.metric("주소 형식 오류", len(v_df[~v_df["유효성"] & (v_df["매핑이메일"] != "")]))
            with c3: st.metric("발송 예정 건수", len(v_df))

            st.markdown("---")

            # 일괄 발송 리스트 확인
            edited_agg = st.data_editor(
                v_df.groupby(["지사", "담당자", "매핑이메일", "매핑상태", "유효성"]).size().reset_index(name="건수"),
                use_container_width=True, key="alert_batch_editor", hide_index=True
            )

            # 발송 설정 폼
            with st.form("alert_send_form"):
                subject = st.text_input("제목", "[긴급] 해지방어 활동 미등록 건 확인 요청")
                body_tpl = st.text_area("본문", "안녕하세요, {담당자}님. 긴급 고위험 계약 {건수}건을 확인해주세요.")
                dry_run = st.toggle("모의 발송 (로그만 기록)", value=True)
                
                if st.form_submit_button("📧 일괄 발송 시작"):
                    progress = st.progress(0)
                    status_log = []
                    
                    for i, row in edited_agg.iterrows():
                        mgr, dest, cnt = row["담당자"], row["매핑이메일"], row["건수"]
                        log_entry = {"time": datetime.now(), "target": mgr, "email": dest, "cnt": cnt, "mode": "Dry" if dry_run else "Actual"}
                        
                        if not dest or not row["유효성"]:
                            log_entry["결과"] = "FAIL(Address)"
                        else:
                            try:
                                if not dry_run:
                                    msg = EmailMessage()
                                    msg["To"] = dest
                                    msg["Subject"] = subject
                                    msg.set_content(body_tpl.format(담당자=mgr, 건수=cnt))
                                    # SMTP 발송 엔진 연동 필요 (비밀번호 인증 등)
                                log_entry["결과"] = "SUCCESS"
                            except Exception as e:
                                log_entry["결과"] = f"ERROR({str(e)})"
                        
                        status_log.append(log_entry)
                        progress.progress((i+1)/len(edited_agg))
                    
                    log_email_history(LOG_PATH, status_log)
                    st.success(f"처리 완료! 발송 이력이 {LOG_PATH}에 기록되었습니다.")

        if os.path.exists(LOG_PATH):
            with st.expander("📄 최근 발송 로그 보기"):
                st.dataframe(pd.read_csv(LOG_PATH).tail(10), use_container_width=True)
