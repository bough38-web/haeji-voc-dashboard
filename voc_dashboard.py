import os
import re
import smtplib
from datetime import datetime
from email.message import EmailMessage

import numpy as np
import pandas as pd
import streamlit as st

# ====================================================
# 1. 설정 및 유틸리티 함수
# ====================================================
st.set_page_config(page_title="해지 VOC 통합 대시보드", layout="wide")

# SMTP 설정 (보안상 실제 운영 시에는 st.secrets 사용 권장)
SMTP_HOST = "smtp.gmail.com"
SMTP_PORT = 587
SMTP_USER = "bough38@gmail.com"
SMTP_PASSWORD = "adzk cyik sing emds"
SENDER_NAME = "해지VOC 관리자"

# 파일 경로 설정
MERGED_PATH = "merged.xlsx"
CONTACT_PATH = "contact_map.xlsx" # 또는 영업구역담당자_251204.xlsx
FEEDBACK_PATH = "feedback.csv"

def safe_str(val):
    """NaN이나 None을 빈 문자열로 변환"""
    if pd.isna(val) or val is None:
        return ""
    return str(val).strip()

def is_valid_email(email):
    """이메일 유효성 검사"""
    email = safe_str(email)
    if not email:
        return False
    regex = r'^[a-zA-Z0-9+-_.]+@[a-zA-Z0-9-]+\.[a-zA-Z0-9-.]+$'
    return bool(re.match(regex, email))

def sort_branch(branches):
    """지사 이름 정렬 (가나다순)"""
    return sorted([safe_str(b) for b in branches if safe_str(b)])

def style_risk(df):
    """리스크 등급에 따른 스타일링 함수 (표시용 - 원본 반환)"""
    return df

def filter_valid_columns(cols, df):
    """존재하는 컬럼만 필터링"""
    return [c for c in cols if c in df.columns]

def save_feedback(path, df):
    """피드백 데이터 저장"""
    try:
        df.to_csv(path, index=False, encoding="utf-8-sig")
    except Exception as e:
        st.error(f"피드백 저장 실패: {e}")

# ====================================================
# 2. 데이터 로드 및 전처리
# ====================================================
@st.cache_data
def load_data():
    """메인 데이터 및 담당자 매핑 데이터 로드"""
    # 1. VOC 데이터 로드
    if not os.path.exists(MERGED_PATH):
        st.error(f"데이터 파일({MERGED_PATH})이 없습니다.")
        return pd.DataFrame(), pd.DataFrame(), pd.DataFrame(), {}, None

    df = pd.read_excel(MERGED_PATH)
    
    # 필수 전처리
    if "계약번호" in df.columns:
        df["계약번호_정제"] = df["계약번호"].astype(str).str.replace(r"[^0-9A-Za-z]", "", regex=True)
    
    # 날짜 변환
    if "접수일시" in df.columns:
        df["접수일시"] = pd.to_datetime(df["접수일시"], errors="coerce")
    
    # 컬럼 존재 여부 확인 및 기본값 설정
    required_cols = ["관리지사", "구역담당자_통합", "상호", "리스크등급", "매칭여부", "설치주소_표시"]
    for col in required_cols:
        if col not in df.columns:
            df[col] = "" # 없는 컬럼은 빈 문자열로 채움 (에러 방지)

    # 2. 담당자 매핑 데이터 로드
    contact_df = pd.DataFrame()
    manager_contacts = {}
    
    # 파일 이름 유연하게 처리
    c_path = CONTACT_PATH
    if not os.path.exists(c_path) and os.path.exists("영업구역담당자_251204.xlsx"):
        c_path = "영업구역담당자_251204.xlsx"

    if os.path.exists(c_path):
        contact_df = pd.read_excel(c_path)
        # 담당자 이름 / 이메일 컬럼 찾기 (가정: 첫 번째가 이름, 두 번째가 이메일)
        if len(contact_df.columns) >= 2:
            name_col = contact_df.columns[0]
            email_col = contact_df.columns[1]
            
            for _, row in contact_df.iterrows():
                mgr_name = safe_str(row[name_col])
                email = safe_str(row[email_col])
                if mgr_name:
                    manager_contacts[mgr_name] = {"email": email}
    
    # 3. 월정료 컬럼 찾기
    fee_col = None
    for c in df.columns:
        if "월정료" in str(c):
            fee_col = c
            break
            
    return df, contact_df, manager_contacts, fee_col

# 데이터 로딩 실행
voc_df, contact_df, manager_contacts, fee_raw_col = load_data()

# 전역 필터링용 데이터프레임 생성 (초기 상태)
voc_filtered_global = voc_df.copy()

# 비매칭 데이터 별도 분리
unmatched_global = voc_df[voc_df["매칭여부"] == "비매칭(X)"].copy()
df_other = voc_df[voc_df["출처"] != "해지VOC"].copy() # 기타 출처 데이터
df_voc = voc_df[voc_df["출처"] == "해지VOC"].copy() # 순수 해지VOC

# 주소 컬럼 후보군
address_cols = ["설치주소", "주소", "설치장소"] 

# 피드백 데이터 로드 (세션 상태 활용)
if "feedback_df" not in st.session_state:
    if os.path.exists(FEEDBACK_PATH):
        try:
            st.session_state["feedback_df"] = pd.read_csv(FEEDBACK_PATH)
        except:
            st.session_state["feedback_df"] = pd.DataFrame(
                columns=["계약번호_정제", "고객대응내용", "등록자", "등록일자", "비고"]
            )
    else:
        st.session_state["feedback_df"] = pd.DataFrame(
            columns=["계약번호_정제", "고객대응내용", "등록자", "등록일자", "비고"]
        )

# ====================================================
# 3. 메인 UI 구성 (탭)
# ====================================================
st.title("📊 해지 VOC 종합 관리 대시보드")

if voc_df.empty:
    st.warning("데이터가 없습니다. 엑셀 파일을 확인해주세요.")
else:
    # 탭 구성
    tab_all, tab_unmatched, tab_drill, tab_filter, tab_alert = st.tabs(
        ["📘 VOC 전체", "🧯 비매칭 관리", "🔍 활동등록(상세)", "🎯 정밀필터", "📨 담당자 알림"]
    )

    # ====================================================
    # TAB 1: VOC 전체 (계약번호 기준 요약)
    # ====================================================
    with tab_all:
        st.subheader("📘 VOC 전체 (계약번호 기준 요약)")

        row1_col1, row1_col2 = st.columns([2, 3])

        # 지사 선택
        all_branches = sort_branch(voc_filtered_global["관리지사"].unique())
        branches_for_tab1 = ["전체"] + all_branches
        
        selected_branch_tab1 = row1_col1.radio(
            "지사 선택",
            options=branches_for_tab1,
            horizontal=True,
            key="tab1_branch_radio",
        )

        # 담당자 선택 (지사 선택에 따라 필터링)
        temp_for_mgr = voc_filtered_global.copy()
        if selected_branch_tab1 != "전체":
            temp_for_mgr = temp_for_mgr[temp_for_mgr["관리지사"] == selected_branch_tab1]

        mgr_list = sorted(temp_for_mgr["구역담당자_통합"].astype(str).unique().tolist())
        mgr_options_tab1 = ["전체"] + [m for m in mgr_list if m != "nan" and m != ""]

        selected_mgr_tab1 = row1_col2.radio(
            "담당자 선택",
            options=mgr_options_tab1,
            horizontal=True,
            key="tab1_mgr_radio",
        )

        # 검색 필터
        s1, s2, s3 = st.columns(3)
        q_cn = s1.text_input("계약번호 검색(부분)", key="tab1_cn")
        q_name = s2.text_input("상호 검색(부분)", key="tab1_name")
        q_addr = s3.text_input("주소 검색(부분)", key="tab1_addr")

        # 필터링 적용
        temp = voc_filtered_global.copy()

        if selected_branch_tab1 != "전체":
            temp = temp[temp["관리지사"] == selected_branch_tab1]
        if selected_mgr_tab1 != "전체":
            temp = temp[temp["구역담당자_통합"].astype(str) == selected_mgr_tab1]

        if q_cn:
            temp = temp[temp["계약번호_정제"].astype(str).str.contains(q_cn.strip())]
        if q_name and "상호" in temp.columns:
            temp = temp[temp["상호"].astype(str).str.contains(q_name.strip())]
        if q_addr:
            # 주소 컬럼 통합 검색
            addr_mask = pd.Series(False, index=temp.index)
            if "설치주소_표시" in temp.columns:
                addr_mask |= temp["설치주소_표시"].astype(str).str.contains(q_addr.strip())
            for col in address_cols:
                if col in temp.columns:
                    addr_mask |= temp[col].astype(str).str.contains(q_addr.strip())
            temp = temp[addr_mask]

        if temp.empty:
            st.info("조건에 맞는 VOC 데이터가 없습니다.")
        else:
            # 최신 접수일시 기준 요약 (계약번호별 1행)
            temp_sorted = temp.sort_values("접수일시", ascending=False)
            df_summary = temp_sorted.drop_duplicates("계약번호_정제").copy()
            
            # 접수건수 계산
            counts = temp["계약번호_정제"].value_counts()
            df_summary["접수건수"] = df_summary["계약번호_정제"].map(counts)

            summary_cols = [
                "계약번호_정제", "상호", "관리지사", "구역담당자_통합", "리스크등급",
                "경과일수", "매칭여부", "접수건수", "설치주소_표시", fee_raw_col,
                "계약상태(중)", "서비스(소)"
            ]
            # 존재하는 컬럼만 선택
            display_cols = filter_valid_columns(summary_cols, df_summary)

            st.markdown(f"📌 표시 계약 수: **{len(df_summary):,} 건**")
            st.dataframe(
                df_summary[display_cols],
                use_container_width=True,
                height=480,
                hide_index=True
            )

    # ====================================================
    # TAB 2: 비매칭 관리
    # ====================================================
    with tab_unmatched:
        st.subheader("🧯 해지방어 활동시설 (비매칭)")
        st.caption("비매칭(X) = 해지 VOC 접수 후 시스템상 활동내역이 확인되지 않은 시설")

        if unmatched_global.empty:
            st.info("현재 비매칭(X) 계약이 없습니다.")
        else:
            u_col1, u_col2 = st.columns([2, 3])
            
            # 지사 필터
            u_branches = ["전체"] + sort_branch(unmatched_global["관리지사"].unique())
            sel_branch_u = u_col1.radio("지사 선택", u_branches, horizontal=True, key="tab2_br")
            
            # 담당자 필터
            temp_u = unmatched_global.copy()
            if sel_branch_u != "전체":
                temp_u = temp_u[temp_u["관리지사"] == sel_branch_u]
            
            u_mgrs = sorted(temp_u["구역담당자_통합"].astype(str).unique().tolist())
            u_mgr_opts = ["전체"] + [m for m in u_mgrs if m != "nan" and m != ""]
            sel_mgr_u = u_col2.radio("담당자 선택", u_mgr_opts, horizontal=True, key="tab2_mgr")

            # 검색 필터
            us1, us2 = st.columns(2)
            uq_cn = us1.text_input("계약번호 검색", key="tab2_cn")
            uq_name = us2.text_input("상호 검색", key="tab2_nm")

            # 필터링 적용
            if sel_mgr_u != "전체":
                temp_u = temp_u[temp_u["구역담당자_통합"].astype(str) == sel_mgr_u]
            
            if uq_cn:
                temp_u = temp_u[temp_u["계약번호_정제"].str.contains(uq_cn.strip())]
            if uq_name:
                temp_u = temp_u[temp_u["상호"].astype(str).str.contains(uq_name.strip())]

            if temp_u.empty:
                st.info("조건에 맞는 데이터가 없습니다.")
            else:
                # 요약표 생성
                u_summary = temp_u.sort_values("접수일시", ascending=False).drop_duplicates("계약번호_정제").copy()
                u_counts = temp_u["계약번호_정제"].value_counts()
                u_summary["접수건수"] = u_summary["계약번호_정제"].map(u_counts)

                u_cols = ["계약번호_정제", "상호", "관리지사", "구역담당자_통합", "리스크등급", "접수건수", "설치주소_표시"]
                st.dataframe(
                    u_summary[filter_valid_columns(u_cols, u_summary)],
                    use_container_width=True,
                    hide_index=True
                )

                # 상세 보기 및 피드백 입력
                st.markdown("### 📂 상세 VOC 이력 및 조치 등록")
                
                # 계약번호 선택 박스 생성 (계약번호 | 상호 | 지사)
                u_summary["display"] = u_summary.apply(
                    lambda x: f"{x['계약번호_정제']} | {x.get('상호','')} | {x.get('관리지사','')}", axis=1
                )
                contract_opts = ["(선택)"] + u_summary["display"].tolist()
                
                sel_display = st.selectbox("계약 선택", contract_opts)
                
                if sel_display != "(선택)":
                    sel_cn_u = sel_display.split(" | ")[0]
                    
                    # 상세 이력 표시
                    st.markdown(f"#### 🔍 계약번호: `{sel_cn_u}` 상세 이력")
                    detail_df = temp_u[temp_u["계약번호_정제"] == sel_cn_u]
                    st.dataframe(detail_df, use_container_width=True)
                    
                    # 피드백 입력 폼
                    with st.form("feedback_form_u"):
                        st.write("📝 **조치 내역 등록**")
                        fb_content = st.text_area("내용 입력")
                        fb_writer = st.text_input("등록자")
                        
                        if st.form_submit_button("등록"):
                            if fb_content and fb_writer:
                                new_fb = {
                                    "계약번호_정제": sel_cn_u,
                                    "고객대응내용": fb_content,
                                    "등록자": fb_writer,
                                    "등록일자": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                                    "비고": "비매칭 탭에서 등록"
                                }
                                st.session_state["feedback_df"] = pd.concat(
                                    [st.session_state["feedback_df"], pd.DataFrame([new_fb])], 
                                    ignore_index=True
                                )
                                save_feedback(FEEDBACK_PATH, st.session_state["feedback_df"])
                                st.success("등록되었습니다.")
                            else:
                                st.warning("내용과 등록자를 입력해주세요.")

    # ====================================================
    # TAB 3: 상세 활동등록 (Drill-down)
    # ====================================================
    with tab_drill:
        st.subheader("🔍 해지상담대상 활동등록 (계약별 상세)")
        
        # 필터링 UI (탭 1, 2와 유사하지만 독립적으로 동작)
        d_col1, d_col2 = st.columns(2)
        
        # 전체 데이터 기준 필터
        drill_df = voc_filtered_global.copy()
        
        d_branches = ["전체"] + sort_branch(drill_df["관리지사"].unique())
        sel_br_d = d_col1.selectbox("지사", d_branches, key="drill_br")
        
        if sel_br_d != "전체":
            drill_df = drill_df[drill_df["관리지사"] == sel_br_d]
            
        d_mgrs = ["전체"] + sorted([x for x in drill_df["구역담당자_통합"].astype(str).unique() if x != "nan"])
        sel_mgr_d = d_col2.selectbox("담당자", d_mgrs, key="drill_mgr")
        
        if sel_mgr_d != "전체":
            drill_df = drill_df[drill_df["구역담당자_통합"].astype(str) == sel_mgr_d]
            
        # 계약 선택
        if drill_df.empty:
            st.info("데이터가 없습니다.")
        else:
            drill_summary = drill_df.drop_duplicates("계약번호_정제")
            drill_summary["display"] = drill_summary.apply(
                lambda x: f"{x['계약번호_정제']} ({x.get('상호','Unknown')})", axis=1
            )
            
            sel_drill_cn = st.selectbox(
                "계약번호 선택", 
                ["(선택)"] + drill_summary["display"].tolist(),
                key="drill_cn_sel"
            )
            
            if sel_drill_cn != "(선택)":
                real_cn = sel_drill_cn.split(" (")[0]
                
                # 1. 기본 정보 표시
                info_row = drill_summary[drill_summary["계약번호_정제"] == real_cn].iloc[0]
                
                m1, m2, m3 = st.columns(3)
                m1.metric("상호", info_row.get("상호", "-"))
                m2.metric("관리지사", info_row.get("관리지사", "-"))
                m3.metric("담당자", info_row.get("구역담당자_통합", "-"))
                
                st.info(f"주소: {info_row.get('설치주소_표시', '-')}")
                
                # 2. VOC 이력 / 기타 이력 분리 표시
                col_v1, col_v2 = st.columns(2)
                
                with col_v1:
                    st.markdown("##### 📘 해지 VOC 이력")
                    v_hist = df_voc[df_voc["계약번호_정제"] == real_cn]
                    st.dataframe(v_hist, use_container_width=True)
                    
                with col_v2:
                    st.markdown("##### 📂 기타 이력 (요청/설변 등)")
                    o_hist = df_other[df_other["계약번호_정제"] == real_cn]
                    if o_hist.empty:
                        st.caption("기타 이력이 없습니다.")
                    else:
                        st.dataframe(o_hist, use_container_width=True)
                
                # 3. 통합 피드백 관리
                st.markdown("---")
                st.markdown("#### 📝 활동 내역 관리")
                
                # 기존 이력 표시
                curr_fb = st.session_state["feedback_df"]
                my_fb = curr_fb[curr_fb["계약번호_정제"].astype(str) == str(real_cn)]
                
                if not my_fb.empty:
                    for i, r in my_fb.iterrows():
                        st.text_area(
                            f"{r['등록일자']} - {r['등록자']}",
                            value=r['고객대응내용'],
                            disabled=True,
                            key=f"read_fb_{i}"
                        )
                else:
                    st.caption("등록된 활동 내역이 없습니다.")
                    
                # 신규 등록
                with st.form("drill_fb_form"):
                    txt = st.text_area("신규 활동 내용 입력")
                    writer = st.text_input("작성자")
                    if st.form_submit_button("저장"):
                        if txt and writer:
                            add_row = {
                                "계약번호_정제": real_cn,
                                "고객대응내용": txt,
                                "등록자": writer,
                                "등록일자": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                                "비고": "상세탭 등록"
                            }
                            st.session_state["feedback_df"] = pd.concat(
                                [curr_fb, pd.DataFrame([add_row])], ignore_index=True
                            )
                            save_feedback(FEEDBACK_PATH, st.session_state["feedback_df"])
                            st.success("저장되었습니다.")
                            st.rerun()
                        else:
                            st.error("내용과 작성자를 입력하세요.")

    # ====================================================
    # TAB 4: 정밀 필터 (Filter)
    # ====================================================
    with tab_filter:
        st.subheader("🎯 데이터 정밀 필터링")
        st.write("필요한 조건을 조합하여 데이터를 추출하세요.")
        
        with st.form("adv_filter"):
            c1, c2 = st.columns(2)
            f_risks = c1.multiselect("리스크 등급", ["HIGH", "MEDIUM", "LOW", "UNKNOWN"])
            f_match = c2.radio("매칭 여부", ["전체", "매칭(O)", "비매칭(X)"], horizontal=True)
            
            submitted = st.form_submit_button("필터 적용")
            
            if submitted:
                res = voc_filtered_global.copy()
                if f_risks:
                    res = res[res["리스크등급"].isin(f_risks)]
                if f_match != "전체":
                    res = res[res["매칭여부"] == f_match]
                
                st.write(f"검색 결과: {len(res)} 건")
                st.dataframe(res, use_container_width=True)

    # ====================================================
    # TAB 5: 담당자 알림 (Alert)
    # ====================================================
    with tab_alert:
        st.subheader("📨 담당자 알림 발송")
        
        if contact_df.empty:
            st.warning("담당자 매핑 파일이 없어 이메일 자동 매칭이 불가능합니다. 직접 입력하여 발송하세요.")
        
        # 알림 대상: 비매칭 데이터 기준
        targets = unmatched_global.copy()
        
        # 담당자별 집계
        mgr_counts = targets["구역담당자_통합"].value_counts().reset_index()
        mgr_counts.columns = ["담당자", "비매칭건수"]
        
        # 이메일 매핑
        mgr_counts["이메일"] = mgr_counts["담당자"].apply(
            lambda x: manager_contacts.get(x, {}).get("email", "")
        )
        
        st.markdown("#### 📮 발송 대상 목록")
        st.dataframe(mgr_counts, use_container_width=True)
        
        st.markdown("---")
        st.markdown("#### 📧 이메일 발송")
        
        sel_target_mgr = st.selectbox("발송할 담당자 선택", ["(선택)"] + mgr_counts["담당자"].tolist())
        
        if sel_target_mgr != "(선택)":
            target_info = mgr_counts[mgr_counts["담당자"] == sel_target_mgr].iloc[0]
            default_email = target_info["이메일"]
            cnt = target_info["비매칭건수"]
            
            with st.form("email_form"):
                rcpt_email = st.text_input("받는 사람 이메일", value=default_email)
                subject = st.text_input("제목", value=f"[알림] {sel_target_mgr}님, 해지방어 활동 미등록 건 확인 요청")
                msg_body = st.text_area(
                    "본문",
                    value=f"안녕하세요 {sel_target_mgr}님,\n\n현재 귀하의 담당 구역에 해지방어 활동 내역이 없는 계약이 {cnt}건 확인되었습니다.\n확인 후 조치 부탁드립니다.\n\n감사합니다.",
                    height=150
                )
                
                # 첨부파일 생성 (해당 담당자의 비매칭 목록)
                mgr_data = targets[targets["구역담당자_통합"] == sel_target_mgr]
                csv_data = mgr_data.to_csv(index=False).encode("utf-8-sig")
                
                submit_email = st.form_submit_button("전송하기")
                
                if submit_email:
                    if not is_valid_email(rcpt_email):
                        st.error("유효하지 않은 이메일 주소입니다.")
                    else:
                        try:
                            msg = EmailMessage()
                            msg["Subject"] = subject
                            msg["From"] = f"{SENDER_NAME} <{SMTP_USER}>"
                            msg["To"] = rcpt_email
                            msg.set_content(msg_body)
                            
                            # CSV 첨부
                            msg.add_attachment(
                                csv_data,
                                maintype="text",
                                subtype="csv",
                                filename=f"비매칭리스트_{sel_target_mgr}.csv"
                            )
                            
                            # SMTP 발송
                            with smtplib.SMTP(SMTP_HOST, SMTP_PORT) as server:
                                server.starttls()
                                server.login(SMTP_USER, SMTP_PASSWORD)
                                server.send_message(msg)
                            
                            st.success(f"✅ {rcpt_email}로 메일이 발송되었습니다!")
                            
                        except Exception as e:
                            st.error(f"❌ 전송 실패: {e}")
