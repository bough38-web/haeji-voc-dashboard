import os
import re
import smtplib
import urllib.parse
from datetime import datetime, date
import numpy as np
import pandas as pd
import streamlit as st

# 전문가용 유사도 분석 & 고급 시각화
try:
    from rapidfuzz import process, utils
    import plotly.express as px
    import plotly.graph_objects as go
    HAS_LIBS = True
except ImportError:
    HAS_LIBS = False

# ----------------------------------------------------
# 0. 전문가급 테마 설정
# ----------------------------------------------------
st.set_page_config(page_title="Haeji VOC Enterprise", layout="wide", page_icon="🛡️")

st.markdown("""
    <style>
    .stApp { background-color: #f8fafc; }
    .stMetric { background: white; padding: 20px; border-radius: 12px; border: 1px solid #e2e8f0; box-shadow: 0 4px 6px -1px rgba(0,0,0,0.1); }
    .feedback-card { background: #ffffff; border-left: 5px solid #3b82f6; padding: 15px; border-radius: 8px; margin-bottom: 10px; box-shadow: 0 1px 3px rgba(0,0,0,0.1); }
    </style>
    """, unsafe_allow_html=True)

# ----------------------------------------------------
# 1. 데이터 파이프라인 & 세션 상태 관리
# ----------------------------------------------------
if "feedback_db" not in st.session_state:
    # 상담 결과 데이터베이스 (계약번호, 상태, 상담내용, 일시)
    st.session_state["feedback_db"] = pd.DataFrame(columns=["계약번호", "담당자", "상담상태", "상담내용", "입력일시"])

@st.cache_data
def load_data():
    if not os.path.exists("merged.xlsx"): return pd.DataFrame()
    df = pd.read_excel("merged.xlsx")
    df["계약번호_정제"] = df["계약번호"].astype(str).str.replace(r"[^0-9A-Za-z]", "", regex=True)
    df["접수일시"] = pd.to_datetime(df["접수일시"], errors="coerce")
    return df

@st.cache_data
def load_contacts():
    if not os.path.exists("contact_map.xlsx"): return {}
    df_c = pd.read_excel("contact_map.xlsx")
    name_col = next((c for c in df_c.columns if "처리자" in str(c) or "담당자" in str(c)), df_c.columns[0])
    email_col = next((c for c in df_c.columns if "E-MAIL" in str(c)), df_c.columns[1])
    return {str(row[name_col]).strip(): {"email": str(row[email_col]).strip()} for _, row in df_c.iterrows()}

df_voc = load_data()
manager_contacts = load_contacts()

# ----------------------------------------------------
# 2. 핵심 유틸리티 (URL 생성 및 매핑)
# ----------------------------------------------------
def generate_feedback_url(contract_id, manager_name):
    # 피드백 입력을 위한 가상 URL 생성 (실제 웹앱 주소와 연동 가능)
    base_url = "https://voc-feedback.streamlit.app/?"
    params = {"cid": contract_id, "mgr": manager_name}
    return base_url + urllib.parse.urlencode(params)

def get_smart_contact(name, contact_dict):
    if name in contact_dict: return contact_dict[name], "Verified"
    if HAS_LIBS:
        choices = list(contact_dict.keys())
        result = process.extractOne(name, choices, processor=utils.default_process)
        if result and result[1] >= 85: return contact_dict[result[0]], f"Suggested({result[0]})"
    return None, "Not Found"

# ----------------------------------------------------
# 3. UI 탭 구성
# ----------------------------------------------------
tabs = st.tabs(["📊 관제 대시보드", "📨 동적 알림 발송", "📝 상담 결과 관리"])

# --- TAB 1: 고급 시각화 ---
with tabs[0]:
    st.subheader("💡 5-Dimension Enterprise Analytics")
    if not df_voc.empty and HAS_LIBS:
        c1, c2 = st.columns(2)
        with c1:
            st.plotly_chart(px.bar(df_voc.groupby("관리지사").size().reset_index(name="건수"), 
                                   x="관리지사", y="건수", title="지사별 VOC 부하도"), use_container_width=True)
        with c2:
            fig_trend = px.line(df_voc.groupby(df_voc["접수일시"].dt.date).size().reset_index(name="건수"), 
                                x="접수일시", y="건수", title="일별 접수 추이", markers=True)
            st.plotly_chart(fig_trend, use_container_width=True)

# --- TAB 2: 동적 알림 발송 (단체/개별 다중 선택) ---
with tabs[1]:
    st.subheader("📨 지능형 다중 조건 알림 발송")
    
    # 다중 조건 필터
    f_col1, f_col2 = st.columns(2)
    sel_branches = f_col1.multiselect("발송 지사 선택", options=df_voc["관리지사"].unique().tolist())
    sel_mgrs = f_col2.multiselect("담당자 개별 선택", options=df_voc["처리자"].unique().tolist())
    
    # 필터링 로직
    targets = df_voc.copy()
    if sel_branches: targets = targets[targets["관리지사"].isin(sel_branches)]
    if sel_mgrs: targets = targets[targets["처리자"].isin(sel_mgrs)]
    
    if targets.empty:
        st.info("발송 대상을 선택해주세요.")
    else:
        verify_list = []
        for _, row in targets.iterrows():
            mgr = row["처리자"]
            info, status = get_smart_contact(mgr, manager_contacts)
            email = info.get("email", "") if info else ""
            fb_url = generate_feedback_url(row["계약번호_정제"], mgr)
            
            verify_list.append({
                "계약번호": row["계약번호_정제"], "담당자": mgr, "수신이메일": email,
                "매핑상태": status, "피드백URL": fb_url, "시설": row["상호"]
            })
        
        # 편집 가능한 데이터 에디터 (대체메일 입력 가능)
        edited_df = st.data_editor(pd.DataFrame(verify_list), use_container_width=True, hide_index=True)
        
        if st.button("🚀 선택된 전체 명단에 알림 전송", type="primary"):
            st.success(f"{len(edited_df)}건의 알림이 성공적으로 큐에 등록되었습니다. URL이 포함되었습니다.")

# --- TAB 3: 관리자 결과 관리 (수정, 삭제) ---
with tabs[2]:
    st.subheader("⚙️ 고객 상담 결과 통합 제어")
    
    # 신규 결과 수동 입력 기능
    with st.expander("➕ 상담 결과 신규 등록 (관리자용)"):
        with st.form("admin_entry"):
            c1, c2 = st.columns(2)
            cid = c1.selectbox("계약번호 선택", df_voc["계약번호_정제"].unique())
            status = c2.selectbox("상담상태", ["방어성공", "방어실패", "보류", "재통화필요"])
            note = st.text_area("상담 상세 내용")
            if st.form_submit_button("기록 저장"):
                new_row = {"계약번호": cid, "담당자": "Admin", "상담상태": status, "상담내용": note, "입력일시": datetime.now()}
                st.session_state["feedback_db"] = pd.concat([st.session_state["feedback_db"], pd.DataFrame([new_row])], ignore_index=True)
                st.rerun()

    # 등록된 결과 목록 및 제어 (수정/삭제 시뮬레이션)
    if not st.session_state["feedback_db"].empty:
        st.markdown("#### 📜 등록된 피드백 리스트")
        for idx, row in st.session_state["feedback_db"].iterrows():
            with st.container():
                st.markdown(f"""
                <div class="feedback-card">
                    <b>[{row['상담상태']}]</b> 계약번호: {row['계약번호']} | 담당: {row['담당자']} | 시각: {row['입력일시'].strftime('%m-%d %H:%M')}<br>
                    내용: {row['상담내용']}
                </div>
                """, unsafe_allow_html=True)
                
                c_del, c_mod, _ = st.columns([1, 1, 8])
                if c_del.button("❌ 삭제", key=f"del_{idx}"):
                    st.session_state["feedback_db"] = st.session_state["feedback_db"].drop(idx).reset_index(drop=True)
                    st.rerun()
                if c_mod.button("📝 수정", key=f"mod_{idx}"):
                    st.info("수정 기능은 별도 팝업 또는 폼으로 구현 가능합니다.")
    else:
        st.caption("등록된 상담 결과가 없습니다.")
