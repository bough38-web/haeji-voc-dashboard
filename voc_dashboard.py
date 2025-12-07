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
# 1. 지능형 매핑 및 유효성 검사 유틸리티
# ----------------------------------------------------

def is_valid_email(email):
    """이메일 정규식 유효성 검사"""
    if not email: return False
    regex = r'^[a-zA-Z0-9+-_.]+@[a-zA-Z0-9-]+\.[a-zA-Z0-9-.]+$'
    return bool(re.match(regex, str(email)))

def get_smart_contact(target_name, contact_dict):
    """Fuzzy Matching: 오타나 직급이 섞인 담당자명을 지능적으로 매핑"""
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
# 2. 데이터 로드 및 환경 설정 (Tab Alert 연동용)
# ----------------------------------------------------

st.set_page_config(page_title="해지 VOC 종합 대시보드 Pro", layout="wide")

MERGED_PATH = "merged.xlsx"
CONTACT_PATH = "contact_map.xlsx"
LOG_PATH = "email_log.csv"

# [가정: 이전 코드들에서 정의된 load_voc_data, load_contact_map 함수가 정의되어 있음]
# 데이터 전처리 후 unmatched_global, manager_contacts가 생성된 상태에서 아래 탭이 실행됩니다.

# ----------------------------------------------------
# 3. 담당자 알림 및 발송 히스토리 관리 (Tab Alert)
# ----------------------------------------------------

# tab_alert 객체는 st.tabs() 중 하나
with tab_alert:
    st.subheader("📨 전문가용 담당자 알림 및 발송 데이터 관리")
    
    if 'manager_contacts' not in locals() or not manager_contacts:
        st.warning("⚠️ 담당자 매핑 파일(contact_map.xlsx) 로드가 필요합니다.")
    else:
        # 1. 발송 대상 추출 (비매칭 & HIGH 리스크)
        # unmatched_global은 이전 전처리 과정에서 필터링된 데이터프레임
        targets = unmatched_global[unmatched_global["리스크등급"] == "HIGH"].copy()
        
        if targets.empty:
            st.success("🎉 현재 알림 발송 대상(비매칭 + 고위험)이 없습니다.")
        else:
            st.caption("🔍 데이터 매핑 및 이메일 유효성 검증 프로세스 실행 중...")
            

[Image of data mapping verification flow chart]

            
            # 2. 검증 수행
            verify_list = []
            for _, row in targets.iterrows():
                mgr_name = row["구역담당자_통합"]
                contact_info, v_status = get_smart_contact(mgr_name, manager_contacts)
                email = contact_info.get("email", "") if contact_info else ""
                
                verify_list.append({
                    "계약번호": row["계약번호_정제"],
                    "지사": row["관리지사"],
                    "담당자": mgr_name,
                    "매핑이메일": email,
                    "매핑상태": v_status,
                    "유효성": is_valid_email(email)
                })
            
            v_df = pd.DataFrame(verify_list)
            
            # 3. 무결성 요약 지표
            col_v1, col_v2, col_v3 = st.columns(3)
            with col_v1:
                st.metric("담당자 매핑률", f"{(v_df['매핑상태'] != 'Not Found').mean()*100:.1f}%")
            with col_v2:
                invalid_cnt = v_df[~v_df["유효성"] & (v_df["매핑이메일"] != "")].shape[0]
                st.metric("형식 오류 주소", f"{invalid_cnt}건", delta_color="inverse")
            with col_v3:
                st.metric("발송 예정 총 계약", f"{len(v_df)}건")

            st.markdown("---")

            # 4. 발송 리스트 최종 확인 (Data Editor)
            st.markdown("#### 🛠️ 발송 리스트 무결성 검증 및 수정")
            agg_targets = v_df.groupby(["지사", "담당자", "매핑이메일", "매핑상태", "유효성"]).size().reset_index(name="건수")
            
            edited_agg = st.data_editor(
                agg_targets,
                column_config={
                    "매핑이메일": st.column_config.TextColumn("이메일(수정가능)", required=True),
                    "건수": st.column_config.NumberColumn("대상 건수", disabled=True),
                    "매핑상태": st.column_config.TextColumn("상태", disabled=True),
                    "유효성": st.column_config.CheckboxColumn("주소 유효함", disabled=True)
                },
                use_container_width=True,
                key="alert_batch_editor_verified",
                hide_index=True
            )

            # 5. 발송 엔진 및 로그 설정
            st.markdown("#### 🚀 메일 엔진 및 히스토리 로그 제어")
            c_m1, c_m2 = st.columns([2, 1])
            with c_m1:
                subject = st.text_input("메일 제목", f"[긴급] 해지방어 활동 미등록 건 확인 요청")
                body_tpl = st.text_area("메일 본문", 
                    "안녕하세요, {담당자} 담당자님.\n\n해지 VOC 접수 후 아직 피드백이 등록되지 않은 고위험 계약이 {건수}건 확인되었습니다.\n신속히 결과를 시스템에 등록해 주시기 바랍니다.")
            
            with c_m2:
                dry_run = st.toggle("모의 발송 (Dry Run)", value=True, help="실제 발송 없이 로그만 생성합니다.")
                if st.button("📧 일괄 발송 및 로그 저장", type="primary", use_container_width=True):
                    progress = st.progress(0)
                    status_log = []
                    
                    for i, row in edited_agg.iterrows():
                        mgr = row["담당자"]
                        dest = row["매핑이메일"]
                        cnt = row["건수"]
                        
                        log_entry = {
                            "timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                            "지사": row["지사"],
                            "담당자": mgr,
                            "이메일": dest,
                            "건수": cnt,
                            "모드": "Dry Run" if dry_run else "Actual"
                        }
                        
                        if not dest or not row["유효성"]:
                            log_entry["result"] = "FAIL(Address Error)"
                            status_log.append(log_entry)
                            continue
                        
                        try:
                            if not dry_run:
                                em = EmailMessage()
                                em["Subject"] = subject
                                em["From"] = f"{SENDER_NAME} <{SMTP_USER}>"
                                em["To"] = dest
                                em.set_content(body_tpl.format(담당자=mgr, 건수=cnt))
                                
                                with smtplib.SMTP(SMTP_HOST, SMTP_PORT) as server:
                                    server.starttls()
                                    server.login(SMTP_USER, SMTP_PASSWORD)
                                    server.send_message(em)
                            log_entry["result"] = "SUCCESS"
                        except Exception as e:
                            log_entry["result"] = f"ERROR({str(e)})"
                        
                        status_log.append(log_entry)
                        progress.progress((i + 1) / len(edited_agg))
                    
                    log_email_history(LOG_PATH, status_log)
                    st.success(f"처리 완료! 발송 이력이 '{LOG_PATH}'에 저장되었습니다.")

            # 6. 발송 현황 시각화 리포트
            if os.path.exists(LOG_PATH):
                st.markdown("---")
                st.markdown("#### 📊 지사별 발송 성공 현황 리포트")
                log_data = pd.read_csv(LOG_PATH)
                if not log_data.empty and HAS_PLOTLY:
                    # 성공한 실제 발송 건만 집계
                    success_report = log_data[(log_data["result"] == "SUCCESS") & (log_data["모드"] == "Actual")]
                    if not success_report.empty:
                        fig = px.bar(success_report.groupby("지사").size().reset_index(name="건수"), 
                                     x="지사", y="건수", text="건수", title="지사별 누적 발송 성공 건수")
                        st.plotly_chart(fig, use_container_width=True)
                    else:
                        st.info("실제 발송 성공 데이터가 아직 없습니다.")
                
                with st.expander("📄 최근 발송 로그 보기"):
                    st.dataframe(log_data.tail(15), use_container_width=True)
