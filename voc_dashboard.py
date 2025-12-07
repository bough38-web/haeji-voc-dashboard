import os
import re
import time
import smtplib
from datetime import datetime, date
from email.message import EmailMessage

import numpy as np
import pandas as pd
import streamlit as st

# 전문가용 지능형 매핑 라이브러리 (유사도 분석)
try:
    from rapidfuzz import process, utils
    HAS_RAPIDFUZZ = True
except ImportError:
    HAS_RAPIDFUZZ = False

# Plotly
try:
    import plotly.express as px
    HAS_PLOTLY = True
except ImportError:
    HAS_PLOTLY = False

# [주의] st.set_page_config 등 기본 설정은 기존 코드 유지

# ----------------------------------------------------
# 1. 고도화된 유틸리티 (매핑 및 유효성 검사)
# ----------------------------------------------------
def is_valid_email(email):
    """이메일 정규식 검증"""
    regex = r'^[a-zA-Z0-9+-_.]+@[a-zA-Z0-9-]+\.[a-zA-Z0-9-.]+$'
    return bool(re.match(regex, str(email or "")))

def get_smart_contact(target_name, contact_dict):
    """Fuzzy Matching: '홍길동 주임'과 '홍길동'을 매핑"""
    target_name = str(target_name).strip()
    if target_name in contact_dict:
        return contact_dict[target_name], "Verified"
    
    if HAS_RAPIDFUZZ and target_name:
        choices = list(contact_dict.keys())
        result = process.extractOne(target_name, choices, processor=utils.default_process)
        if result and result[1] >= 90:
            suggested_name = result[0]
            return contact_dict[suggested_name], f"Suggested({suggested_name})"
    
    return None, "Not Found"

# ----------------------------------------------------
# 2. 데이터 로드 및 전처리 (KeyError 방지 보완)
# ----------------------------------------------------
@st.cache_data(ttl=600)
def load_and_process_data():
    if not os.path.exists(MERGED_PATH):
        return pd.DataFrame(), pd.DataFrame(), {}, "FILE_NOT_FOUND"
    
    df = pd.read_excel(MERGED_PATH)
    
    # 컬럼 존재 여부 확인 및 생성
    if "계약번호" in df.columns:
        df["계약번호_정제"] = df["계약번호"].astype(str).str.replace(",", "").str.strip()
    else:
        df["계약번호_정제"] = "UNKNOWN"

    df["접수일시"] = pd.to_datetime(df.get("접수일시", datetime.now()), errors="coerce")
    
    # [핵심 수정보완] 매칭여부 컬럼 생성 로직 강화
    df_voc = df[df.get("출처") == "해지VOC"].copy()
    df_other = df[df.get("출처") != "해지VOC"].copy()
    
    other_contract_set = set(df_other["계약번호_정제"].dropna().unique())
    
    # df_voc에 명확하게 매칭여부 컬럼 할당
    df_voc["매칭여부"] = df_voc["계약번호_정제"].apply(
        lambda x: "매칭(O)" if x in other_contract_set else "비매칭(X)"
    )
    
    # 리스크 등급 등 기타 로직 (기존 유지)
    today = date.today()
    df_voc["리스크등급"] = df_voc["접수일시"].apply(lambda dt: "HIGH" if (today - dt.date()).days <= 3 else "LOW")
    
    return df_voc, df, {}, "SUCCESS"

# [기존 데이터 로딩 로직 실행 코드 생략]

# ----------------------------------------------------
# 3. 고도화된 알림 발송 탭 (TAB 4)
# ----------------------------------------------------
with tab4:
    st.subheader("📨 지능형 담당자 알림 및 발송 검증")
    
    # 발송 대상 추출 (비매칭 & HIGH 리스크)
    targets = filtered_voc[
        (filtered_voc["매칭여부"] == "비매칭(X)") & 
        (filtered_voc["리스크등급"] == "HIGH")
    ].copy()
    
    if targets.empty:
        st.success("🎉 현재 조건에서 발송 대상(비매칭 고위험)이 없습니다.")
    else:
        # 데이터 매핑 및 유효성 실시간 검증
        verify_list = []
        for idx, row in targets.iterrows():
            mgr_name = row["구역담당자_통합"]
            email, status = get_smart_contact(mgr_name, contact_map)
            
            verify_list.append({
                "계약번호": row["계약번호_정제"],
                "상호": row["상호"],
                "담당자": mgr_name,
                "매핑이메일": email or "",
                "검증상태": status,
                "유효성": is_valid_email(email)
            })
        
        v_df = pd.DataFrame(verify_list)
        
        # 📊 발송 전 무결성 지표
        m1, m2, m3 = st.columns(3)
        m1.metric("매핑 성공률", f"{(v_df['검증상태'] != 'Not Found').mean()*100:.1f}%")
        m2.metric("형식 오류 주소", v_df[~v_df["유효성"] & (v_df["매핑이메일"] != "")].shape[0], delta_color="inverse")
        m3.metric("알림 대상 계약", len(v_df))

        # 리스트 에디터 (데이터 확인 및 수동 편집)
        st.markdown("#### 🛠️ 발송 데이터 리스트")
        agg_targets = v_df.groupby(["담당자", "매핑이메일", "검증상태", "유효성"]).size().reset_index(name="건수")
        
        # 

[Image of data mapping verification flow chart]

        
        edited_agg = st.data_editor(
            agg_targets,
            column_config={
                "매핑이메일": st.column_config.TextColumn("이메일(수정가능)", required=True),
                "검증상태": st.column_config.TextColumn("상태", disabled=True),
                "건수": st.column_config.NumberColumn("건수", disabled=True)
            },
            use_container_width=True,
            hide_index=True
        )

        # 발송 엔진
        st.markdown("---")
        with st.form("alert_send_form"):
            subject_input = st.text_input("메일 제목", "[긴급] 해지 VOC 미조치/고위험 건 확인 요청")
            body_input = st.text_area("메일 본문", "안녕하세요 {담당자}님. 긴급 고위험 계약 {건수}건을 확인하세요.")
            
            c_btn1, c_btn2 = st.columns([1, 1])
            dry_run = c_btn1.toggle("모의 발송 (로그만 확인)", value=True)
            submit = st.form_submit_button("📧 일괄 발송 시작", type="primary", use_container_width=True)
            
            if submit:
                # SMTP 연동 발송 로직
                progress = st.progress(0)
                status_txt = st.empty()
                success_cnt, fail_cnt = 0, 0
                
                for i, row in edited_agg.iterrows():
                    mgr, email, cnt = row["담당자"], row["매핑이메일"], row["건수"]
                    
                    if not row["유효성"] or not email:
                        fail_cnt += 1
                        continue
                    
                    status_txt.text(f"발송 중... ({i+1}/{len(edited_agg)}) - {mgr}")
                    
                    try:
                        if not dry_run:
                            # 실제 이메일 전송 (EmailMessage 라이브러리 활용)
                            msg = EmailMessage()
                            msg["Subject"] = subject_input
                            msg["From"] = SENDER_NAME
                            msg["To"] = email
                            msg.set_content(body_input.format(담당자=mgr, 건수=cnt))
                            
                            with smtplib.SMTP(SMTP_HOST, SMTP_PORT) as s:
                                s.starttls()
                                s.login(SMTP_USER, SMTP_PASSWORD)
                                s.send_message(msg)
                        success_cnt += 1
                    except Exception as e:
                        st.error(f"{mgr} 발송 실패: {str(e)}")
                        fail_cnt += 1
                    
                    progress.progress((i+1)/len(edited_agg))
                
                st.success(f"완료: 성공 {success_cnt}건 / 실패 {fail_cnt}건 (모드: {'모의' if dry_run else '실제'})")
