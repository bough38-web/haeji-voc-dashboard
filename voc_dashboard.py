import os
import re
import smtplib
import time
from datetime import datetime, date
from email.message import EmailMessage

import numpy as np
import pandas as pd
import streamlit as st

# 1. 유사도 분석 및 시각화 엔진 로드
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
# 2. 핵심 유틸리티 (유효성 검사, 매핑, 로그)
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
    
    # 정확 일치 확인
    if target_name in contact_dict:
        return contact_dict[target_name], "Verified"
    
    # 유사도 분석 (임계값 90%)
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
# 3. Streamlit 앱 설정 및 데이터 로드
# ----------------------------------------------------

st.set_page_config(page_title="해지 VOC 종합 대시보드 Pro", layout="wide")

# CSS 스타일 적용 (사용자 원본 유지)
st.markdown("""<style>... 생략 ...</style>""", unsafe_allow_html=True)

# 파일 경로 정의
MERGED_PATH = "merged.xlsx"
CONTACT_PATH = "contact_map.xlsx"
FEEDBACK_PATH = "feedback.csv"
LOG_PATH = "email_log.csv"

# [데이터 로딩 함수들: load_voc_data, load_feedback, load_contact_map 정의 생략 - 이전과 동일하게 유지]
# ... (생략된 부분: 기존 사용자의 데이터 로드 함수 코드) ...

# 데이터 실제 로드 (예시 변수명)
# df = load_voc_data(MERGED_PATH)
# contact_df, manager_contacts = load_contact_map(CONTACT_PATH)
# unmatched_global = ... (필터링된 비매칭 데이터) ...

# ----------------------------------------------------
# 4. 담당자 알림 및 발송 이력 관리 (TAB ALERT)
# ----------------------------------------------------

tabs = st.tabs(["📊 시각화", "📘 VOC 전체", "🧯 비매칭", "🔍 활동등록", "🎯 정밀필터", "📨 담당자 알림"])
tab_alert = tabs[5]

with tab_alert:
    st.subheader("📨 지능형 담당자 알림 및 발송 관리 (Pro)")
    
    if 'manager_contacts' not in locals() or not manager_contacts:
        st.warning("⚠️ 담당자 매핑 파일(contact_map.xlsx)이 필요합니다.")
    else:
        # 비매칭 고위험 계약 추출
        targets = unmatched_global[unmatched_global["리스크등급"] == "HIGH"].copy()
        
        if targets.empty:
            st.success("🎉 현재 발송 대상(비매칭 고위험)이 없습니다.")
        else:
            st.info("🔍 담당자 매핑 및 이메일 유효성 검증을 수행합니다.")
            
            # 매핑 검증
            verify_list = []
            for _, row in targets.iterrows():
                mgr_name = row["구역담당자_통합"]
                contact_info, status = get_smart_contact(mgr_name, manager_contacts)
                email = contact_info.get("email", "") if contact_info else ""
                
                verify_list.append({
                    "계약번호": row["계약번호_정제"],
                    "지사": row["관리지사"],
                    "담당자": mgr_name,
                    "매핑이메일": email,
                    "상태": status,
                    "유효성": is_valid_email(email)
                })
            
            v_df = pd.DataFrame(verify_list)
            
            # 지표 대시보드
            col_v1, col_v2, col_v3 = st.columns(3)
            with col_v1:
                match_rate = (v_df["상태"] != "Not Found").mean() * 100
                st.metric("담당자 매핑률", f"{match_rate:.1f}%")
            with col_v2:
                bad_cnt = v_df[~v_df["유효성"] & (v_df["매핑이메일"] != "")].shape[0]
                st.metric("형식 오류 주소", f"{bad_cnt}건", delta_color="inverse")
            with col_v3:
                st.metric("발송 예정 건수", f"{len(v_df)}건")

            st.markdown("---")

            # 리스트 편집기
            st.markdown("#### 🛠️ 발송 데이터 최종 검토")
            agg_targets = v_df.groupby(["지사", "담당자", "매핑이메일", "상태", "유효성"]).size().reset_index(name="건수")
            
            edited_agg = st.data_editor(
                agg_targets,
                column_config={
                    "매핑이메일": st.column_config.TextColumn("이메일(수정 가능)", required=True),
                    "건수": st.column_config.NumberColumn("계약 건수", disabled=True),
                    "유효성": st.column_config.CheckboxColumn("유효 주소", disabled=True)
                },
                use_container_width=True,
                key="alert_editor_pro",
                hide_index=True
            )

            # 발송 엔진 및 로그 설정
            st.markdown("#### 🚀 발송 실행 제어")
            c_m1, c_m2 = st.columns([2, 1])
            with c_m1:
                subject = st.text_input("제목", f"[긴급] 해지방어 활동 미등록 건 확인 요청")
                body_tpl = st.text_area("본문", "안녕하세요, {담당자}님. 긴급 고위험 계약 {건수}건의 활동 내역을 등록해주세요.")
            
            with c_m2:
                dry_run = st.toggle("모의 발송 (Dry Run)", value=True)
                if st.button("📧 일괄 발송 및 로그 저장", type="primary", use_container_width=True):
                    progress = st.progress(0)
                    status_log = []
                    
                    for i, row in edited_agg.iterrows():
                        mgr, dest, cnt, branch = row["담당자"], row["매핑이메일"], row["건수"], row["지사"]
                        
                        log_entry = {
                            "timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                            "지사": branch, "담당자": mgr, "이메일": dest, "건수": cnt, "모드": "Dry Run" if dry_run else "Actual"
                        }
                        
                        if not dest or not row["유효성"]:
                            log_entry["결과"] = "FAIL(Address Error)"
                            status_log.append(log_entry)
                            continue
                        
                        try:
                            if not dry_run:
                                # SMTP 전송 (st.secrets의 설정값 사용)
                                em = EmailMessage()
                                em["Subject"] = subject
                                em["From"] = f"{SENDER_NAME} <{SMTP_USER}>"
                                em["To"] = dest
                                em.set_content(body_tpl.format(담당자=mgr, 건수=cnt))
                                
                                with smtplib.SMTP(SMTP_HOST, SMTP_PORT) as server:
                                    server.starttls()
                                    server.login(SMTP_USER, SMTP_PASSWORD)
                                    server.send_message(em)
                            log_entry["결과"] = "SUCCESS"
                        except Exception as e:
                            log_entry["결과"] = f"ERROR({str(e)})"
                        
                        status_log.append(log_entry)
                        progress.progress((i + 1) / len(edited_agg))
                    
                    # 히스토리 로그 저장
                    log_email_history(LOG_PATH, status_log)
                    st.success("발송이 완료되었습니다. 이력이 로그 파일에 저장되었습니다.")

            # 지사별 발송 리포트 시각화
            if os.path.exists(LOG_PATH):
                st.markdown("---")
                st.markdown("#### 📊 지사별 발송 성공 통계")
                log_data = pd.read_csv(LOG_PATH)
                if not log_data.empty and HAS_PLOTLY:
                    success_data = log_data[(log_data["결과"] == "SUCCESS") & (log_data["모드"] == "Actual")]
                    if not success_data.empty:
                        fig = px.bar(success_data.groupby("지사").size().reset_index(name="건수"), 
                                     x="지사", y="건수", text="건수", title="지사별 누적 발송 성공 리포트")
                        st.plotly_chart(fig, use_container_width=True)
                    else:
                        st.info("실제 발송 성공 데이터가 아직 없습니다.")
