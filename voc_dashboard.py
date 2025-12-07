import os
import re
import smtplib
import time
from datetime import datetime, date
from email.message import EmailMessage

import numpy as np
import pandas as pd
import streamlit as st

# 유사도 분석 엔진 (설치 필요: pip install rapidfuzz)
try:
    from rapidfuzz import process, utils
    HAS_RAPIDFUZZ = True
except ImportError:
    HAS_RAPIDFUZZ = False

# Plotly 고급 시각화 설정
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
    
    # 1. 100% 일치 확인
    if target_name in contact_dict:
        return contact_dict[target_name], "Verified"
    
    # 2. 유사도 기반 추천 (유사도 90% 임계값)
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
# [주의: 기존의 load_voc_data, load_feedback, load_contact_map 함수 및 
# 데이터 전처리 로직이 이 자리에 위치해야 합니다. (이전 소스코드 복사 권장)]
# ----------------------------------------------------

# ----------------------------------------------------
# 탭 4: 전문가용 지사 필터 기반 알림 및 히스토리 관리
# ----------------------------------------------------

LOG_PATH = "email_log.csv"

with tab_alert:
    st.subheader("📨 전문가용 담당자 알림 및 발송 이력 관리")
    
    # 1. 담당자 파일 로드 체크
    if 'contact_df' not in locals() or contact_df.empty:
        st.warning("⚠️ 담당자 매핑 파일(contact_map.xlsx)이 필요합니다.")
    else:
        # 2. 지사 필터링 추가 (Global 필터와 별개로 알림 대상만 별도 필터)
        alert_branches = st.multiselect(
            "발송 지사 선택", 
            options=BRANCH_ORDER, 
            default=BRANCH_ORDER,
            key="alert_branch_filter"
        )
        
        # 3. 발송 대상 추출 (비매칭 & HIGH 리스크 & 선택 지사)
        targets = unmatched_global[
            (unmatched_global["리스크등급"] == "HIGH") & 
            (unmatched_global["관리지사"].isin(alert_branches))
        ].copy()
        
        if targets.empty:
            st.success("🎉 현재 조건에서 알림 발송 대상(비매칭 + 고위험)이 없습니다.")
        else:
            # 4. 데이터 무결성 검증 수행
            verify_list = []
            for _, row in targets.iterrows():
                mgr_name = row["구역담당자_통합"]
                contact_info, v_status = get_smart_contact(mgr_name, manager_contacts)
                email = contact_info.get("email", "") if contact_info else ""
                
                verify_list.append({
                    "계약번호": row["계약번호_정제"],
                    "관리지사": row["관리지사"],
                    "담당자(원본)": mgr_name,
                    "매핑이메일": email,
                    "매핑상태": v_status,
                    "유효성": is_valid_email(email)
                })
            
            v_df = pd.DataFrame(verify_list)
            
            # 5. 검증 요약 및 시각화
            

[Image of data mapping verification flow chart]

            col_v1, col_v2, col_v3 = st.columns(3)
            with col_v1:
                match_rate = (v_df["매핑상태"] != "Not Found").mean() * 100
                st.metric("담당자 매핑률", f"{match_rate:.1f}%")
            with col_v2:
                invalid_cnt = v_df[~v_df["유효성"] & (v_df["매핑이메일"] != "")].shape[0]
                st.metric("형식 오류 주소", f"{invalid_cnt}건", delta_color="inverse")
            with col_v3:
                st.metric("발송 예정 건수", f"{len(v_df)}건")

            st.markdown("---")

            # 6. 최종 편집 UI (Data Editor)
            st.markdown("#### 🛠️ 일괄 발송 리스트 무결성 검증")
            agg_targets = v_df.groupby(["관리지사", "담당자(원본)", "매핑이메일", "매핑상태", "유효성"]).size().reset_index(name="건수")
            
            edited_agg = st.data_editor(
                agg_targets,
                column_config={
                    "매핑이메일": st.column_config.TextColumn("이메일(수정가능)", required=True),
                    "건수": st.column_config.NumberColumn("대상 건수", disabled=True),
                    "관리지사": st.column_config.TextColumn("지사", disabled=True),
                    "유효성": st.column_config.CheckboxColumn("주소 유효함", disabled=True)
                },
                use_container_width=True,
                key="alert_editor_advanced",
                hide_index=True
            )

            # 7. 발송 제어부
            st.markdown("#### 🚀 메일 엔진 및 로그 설정")
            c_m1, c_m2 = st.columns([2, 1])
            with c_m1:
                subject = st.text_input("메일 제목", f"[긴급] 지사별 미처리 고위험 해지 VOC 안내")
                body_tpl = st.text_area("템플릿 본문", 
                    "안녕하세요, {담당자}님.\n해지 VOC 접수 후 아직 현장 조치 활동이 없는 긴급 건이 {건수}건 있습니다.\n신속한 대응 및 결과 등록을 요청드립니다.")
            
            with c_m2:
                dry_run = st.toggle("모의 발송 (Dry Run)", value=True)
                
                if st.button("📧 선택 지사 일괄 발송", type="primary", use_container_width=True):
                    progress = st.progress(0)
                    msg_txt = st.empty()
                    status_log = []
                    
                    for i, row in edited_agg.iterrows():
                        mgr = row["담당자(원본)"]
                        dest = row["매핑이메일"]
                        cnt = row["건수"]
                        branch = row["관리지사"]
                        
                        log_entry = {
                            "timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                            "관리지사": branch,
                            "담당자": mgr,
                            "이메일": dest,
                            "대상건수": cnt,
                            "모드": "Dry Run" if dry_run else "Actual"
                        }
                        
                        if not dest or not row["유효성"]:
                            log_entry["결과"] = "FAIL(Address Invalid)"
                            status_log.append(log_entry)
                            continue
                        
                        try:
                            if not dry_run:
                                # SMTP 발송 (st.secrets 로드 정보 활용)
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
                    
                    # 8. 이력 누적 저장
                    log_email_history(LOG_PATH, status_log)
                    st.success(f"처리 완료! 발송 히스토리가 '{LOG_PATH}'에 저장되었습니다.")

            # 9. 로그 뷰어
            if os.path.exists(LOG_PATH):
                with st.expander("📄 발송 히스토리(Log) 탐색", expanded=False):
                    try:
                        st.dataframe(pd.read_csv(LOG_PATH).tail(20), use_container_width=True)
                    except:
                        st.error("로그 데이터 로딩 중 문제가 발생했습니다.")
