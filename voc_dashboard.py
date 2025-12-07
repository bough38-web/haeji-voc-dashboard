import os
import re
import smtplib
import time
from datetime import datetime, date
from email.message import EmailMessage

import numpy as np
import pandas as pd
import streamlit as st

# 유사도 분석 엔진 (설치 필수: pip install rapidfuzz)
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
    
    # 2. 유사도 기반 추천 (유사도 90% 이상일 때만 제안)
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
# 2. 데이터 로드 및 초기 설정
# ----------------------------------------------------

# (중략: 기존 코드의 load_voc_data, load_feedback, load_contact_map 정의 유지)
# unmatched_global 등 필수 변수가 이전 탭 전처리 과정에서 생성된 상태여야 합니다.

# ----------------------------------------------------
# 탭 구성: 담당자 알림 (SyntaxError 수정 버전)
# ----------------------------------------------------

# 탭 객체 중 담당자 알림 탭 선택
with tab_alert:
    st.subheader("📨 전문가용 담당자 알림 및 발송 데이터 관리")
    
    # 1. 담당자 매핑 데이터 체크
    if 'manager_contacts' not in locals() or not manager_contacts:
        st.warning("⚠️ 담당자 매핑 파일(contact_map.xlsx)을 업로드하거나 확인해주세요.")
    else:
        # 2. 발송 대상 추출 (비매칭 & HIGH 리스크)
        targets = unmatched_global[unmatched_global["리스크등급"] == "HIGH"].copy()
        
        if targets.empty:
            st.success("🎉 현재 알림 발송 대상(비매칭 고위험)이 없습니다.")
        else:
            st.info("🔍 데이터 매핑 및 이메일 유효성 검증을 수행합니다.")
            
            # [기존 오류 지점 수정] 이미지 설명을 텍스트와 지표로 대체
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
            
            # 3. 무결성 요약 대시보드
            col_v1, col_v2, col_v3 = st.columns(3)
            with col_v1:
                st.metric("담당자 매핑률", f"{(v_df['매핑상태'] != 'Not Found').mean()*100:.1f}%")
            with col_v2:
                invalid_cnt = v_df[~v_df["유효성"] & (v_df["매핑이메일"] != "")].shape[0]
                st.metric("형식 오류 주소", f"{invalid_cnt}건", delta_color="inverse")
            with col_v3:
                st.metric("발송 예정 계약", f"{len(v_df)}건")

            st.markdown("---")

            # 4. 발송 데이터 에디터 (최종 확인)
            st.markdown("#### 🛠️ 발송 리스트 데이터 편집 및 확정")
            agg_targets = v_df.groupby(["지사", "담당자", "매핑이메일", "매핑상태", "유효성"]).size().reset_index(name="건수")
            
            edited_agg = st.data_editor(
                agg_targets,
                column_config={
                    "매핑이메일": st.column_config.TextColumn("이메일(수정가능)", required=True),
                    "건수": st.column_config.NumberColumn("대상 건수", disabled=True),
                    "매핑상태": st.column_config.TextColumn("상태", disabled=True),
                    "유효성": st.column_config.CheckboxColumn("유효 주소", disabled=True)
                },
                use_container_width=True,
                hide_index=True,
                key="alert_pro_editor"
            )

            # 5. 발송 엔진 설정
            c_mail1, c_mail2 = st.columns([2, 1])
            with c_mail1:
                subject = st.text_input("메일 제목", f"[긴급] 미처리 해지 VOC {len(targets)}건 안내")
                body_tpl = st.text_area("메일 본문", 
                    "안녕하세요, {담당자}님.\n해지 VOC 접수 후 처리 내역이 없는 긴급 건 {건수}건의 확인을 요청드립니다.")
            
            with c_mail2:
                dry_run = st.toggle("모의 발송 (Dry Run)", value=True, help="실제 발송 없이 로그만 생성")
                if st.button("📧 일괄 발송 시작", type="primary", use_container_width=True):
                    progress = st.progress(0)
                    msg_txt = st.empty()
                    status_log = []
                    
                    for i, row in edited_agg.iterrows():
                        mgr, dest, cnt, branch = row["담당자"], row["매핑이메일"], row["건수"], row["지사"]
                        
                        log_entry = {
                            "timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                            "지사": branch, "담당자": mgr, "이메일": dest, "건수": cnt, "모드": "Dry Run" if dry_run else "Actual"
                        }
                        
                        if not dest or not row["유효성"]:
                            log_entry["result"] = "FAIL(Bad Address)"
                            status_log.append(log_entry)
                            continue
                        
                        try:
                            if not dry_run:
                                # SMTP 발송 로직 (st.secrets 환경변수 사용)
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
                    st.success(f"처리 완료! 로그가 '{LOG_PATH}'에 기록되었습니다.")

            # 6. 지사별 발송 현황 시각화
            if os.path.exists(LOG_PATH):
                st.markdown("---")
                st.markdown("#### 📊 지사별 누적 발송 현황 리포트")
                log_data = pd.read_csv(LOG_PATH)
                if not log_data.empty and HAS_PLOTLY:
                    success_data = log_data[(log_data["result"] == "SUCCESS") & (log_data["모드"] == "Actual")]
                    if not success_data.empty:
                        fig = px.bar(success_data.groupby("지사").size().reset_index(name="건수"), 
                                     x="지사", y="건수", text="건수", title="지사별 실제 발송 성공 건수")
                        st.plotly_chart(fig, use_container_width=True)
                
                with st.expander("📄 최근 발송 로그 상세 보기"):
                    st.dataframe(log_data.tail(10), use_container_width=True)
