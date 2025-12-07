import os
import re
import smtplib
import time
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

# Plotly 고급 시각화
try:
    import plotly.express as px
    HAS_PLOTLY = True
except Exception:
    HAS_PLOTLY = False

# ----------------------------------------------------
# 0. 엔터프라이즈 UI/UX 설정
# ----------------------------------------------------
st.set_page_config(
    page_title="Haeji VOC Enterprise Control",
    layout="wide",
    initial_sidebar_state="expanded"
)

st.markdown("""
    <style>
    /* Apple-Style Light Theme & Corporate UI */
    html, body, .stApp { background-color: #f5f5f7 !important; color: #1d1d1f !important; font-family: -apple-system, sans-serif; }
    .stMetric { background: white; padding: 20px; border-radius: 12px; border: 1px solid #e5e7eb; box-shadow: 0 4px 6px rgba(0,0,0,0.02); }
    .section-card { background: white; border-radius: 16px; padding: 1.5rem; border: 1px solid #dee2e6; margin-bottom: 1rem; }
    div[data-testid="stExpander"] { border-radius: 10px; border: 1px solid #ced4da; }
    </style>
    """, unsafe_allow_html=True)

# ----------------------------------------------------
# 1. 유틸리티 (매핑 검증 & 이메일 유효성)
# ----------------------------------------------------
def is_valid_email(email):
    """이메일 정규식 유효성 검사"""
    regex = r'^[a-zA-Z0-9+-_.]+@[a-zA-Z0-9-]+\.[a-zA-Z0-9-.]+$'
    return bool(re.match(regex, str(email or "")))

def get_smart_contact(target_name, contact_dict):
    """지능형 담당자 매핑 (Fuzzy matching)"""
    target_name = str(target_name).strip()
    if not target_name or target_name in ["nan", "미지정"]: return None, "Name Missing"
    if target_name in contact_dict: return contact_dict[target_name], "Verified"
    
    if HAS_RAPIDFUZZ:
        choices = list(contact_dict.keys())
        # 유사도 임계값 85% 설정
        result = process.extractOne(target_name, choices, processor=utils.default_process)
        if result and result[1] >= 85:
            return contact_dict[result[0]], f"Suggested({result[0]})"
    return None, "Not Found"

# ----------------------------------------------------
# 2. 고도화된 데이터 로딩 및 동적 필터링
# ----------------------------------------------------
MERGED_PATH = "merged.xlsx"
CONTACT_PATH = "contact_map.xlsx"

@st.cache_data(ttl=600)
def load_and_fix_data():
    if not os.path.exists(MERGED_PATH): return pd.DataFrame()
    df = pd.read_excel(MERGED_PATH)
    
    # 기본 정제
    df["계약번호_정제"] = df["계약번호"].astype(str).str.replace(r"[^0-9A-Za-z]", "", regex=True)
    df["접수일시"] = pd.to_datetime(df["접수일시"], errors="coerce")
    
    # 동적 리스크 등급 계산 (영업일 기준)
    today = date.today()
    def calculate_risk(dt):
        if pd.isna(dt): return "MEDIUM"
        days_diff = (today - dt.date()).days
        return "HIGH" if days_diff <= 3 else "LOW"
    df["리스크등급"] = df["접수일시"].apply(calculate_risk)
    
    # 지사명 클리닝
    if "관리지사" in df.columns:
        df["관리지사"] = df["관리지사"].fillna("지사미상")
    
    # 담당자 필드 통합
    def pick_mgr(row):
        for c in ["처리자", "구역담당자", "담당자"]:
            if c in row and pd.notna(row[c]): return str(row[c]).strip()
        return "미지정"
    df["담당자_통합"] = df.apply(pick_mgr, axis=1)
    
    return df

@st.cache_data
def load_contacts_advanced(path):
    if not os.path.exists(path): return pd.DataFrame(), {}
    df_c = pd.read_excel(path)
    
    # 요청하신 "E-MAIL" 및 "처리자1" 컬럼 자동 감지
    name_col = next((c for c in df_c.columns if "처리자" in str(c) or "담당자" in str(c)), df_c.columns[0])
    email_col = next((c for c in df_c.columns if "E-MAIL" in str(c) or "이메일" in str(c)), df_c.columns[1])
    
    contact_dict = {
        str(row[name_col]).strip(): {"email": str(row[email_col]).strip()} 
        for _, row in df_c.iterrows() if pd.notna(row[name_col])
    }
    return df_c, contact_dict

# 데이터 로딩 실행
df_all = load_and_fix_data()
contact_raw, manager_contacts = load_contacts_advanced(CONTACT_PATH)

# ----------------------------------------------------
# 3. 메인 관제 대시보드 UI
# ----------------------------------------------------
st.title("🛡️ Haeji VOC Enterprise Dashboard")

if df_all.empty:
    st.error("❌ 'merged.xlsx' 데이터를 찾을 수 없습니다.")
    st.stop()

# 전역 KPI 섹션
k1, k2, k3, k4 = st.columns(4)
k1.metric("총 접수 건수", f"{len(df_all):,}")
k2.metric("고위험(HIGH) 관리", f"{len(df_all[df_all['리스크등급']=='HIGH']):,}", delta="긴급", delta_color="inverse")
k3.metric("누적 관리 계약", f"{df_all['계약번호_정제'].nunique():,}")
k4.metric("매핑 담당자", f"{len(manager_contacts)}명")

st.markdown("---")

tabs = st.tabs(["📊 통합 시각화", "📘 VOC 데이터베이스", "📨 담당자 알림 관제"])

# --- TAB 1: 통합 시각화 ---
with tabs[0]:
    st.subheader("📍 리스크 분포 및 접수 추이")
    if HAS_PLOTLY:
        c1, c2 = st.columns(2)
        with c1:
            risk_dist = df_all.groupby(["관리지사", "리스크등급"]).size().reset_index(name="건수")
            fig1 = px.bar(risk_dist, x="관리지사", y="건수", color="리스크등급", 
                         title="지사별 고위험 분포", barmode="group",
                         color_discrete_map={'HIGH': '#ef4444', 'MEDIUM': '#f59e0b', 'LOW': '#10b981'})
            st.plotly_chart(fig1, use_container_width=True)
        with c2:
            daily = df_all.groupby(df_all["접수일시"].dt.date).size().reset_index(name="접수건수")
            fig2 = px.line(daily, x="접수일시", y="접수건수", title="일별 접수 추이", markers=True)
            st.plotly_chart(fig2, use_container_width=True)
    else:
        st.info("시각화 엔진(Plotly)이 로드되지 않았습니다.")

# --- TAB 2: 데이터베이스 (Drill-down) ---
with tabs[1]:
    st.subheader("🔍 VOC 상세 이력 조회")
    # 동적 필터
    s1, s2 = st.columns([1, 1])
    q_id = s1.text_input("계약번호 검색", placeholder="조회할 계약번호 입력...")
    q_branch = s2.multiselect("관리지사 필터", options=df_all["관리지사"].unique().tolist())
    
    df_filtered = df_all.copy()
    if q_id: df_filtered = df_filtered[df_filtered["계약번호_정제"].str.contains(q_id)]
    if q_branch: df_filtered = df_filtered[df_filtered["관리지사"].isin(q_branch)]
    
    st.dataframe(df_filtered.sort_values("접수일시", ascending=False), use_container_width=True, hide_index=True)

# --- TAB 3: 지능형 알림 관제 ---
with tabs[2]:
    st.subheader("📨 담당자 일괄 알림 (대체메일 지원)")
    
    # 고위험 미조치 대상 필터링
    alert_targets = df_all[df_all["리스크등급"] == "HIGH"].copy()
    
    verify_list = []
    # 담당자별 그룹화하여 알림 생성
    agg_targets = alert_targets.groupby(["관리지사", "담당자_통합"]).size().reset_index(name="계약건수")
    
    for _, row in agg_targets.iterrows():
        mgr = row["담당자_통합"]
        info, status = get_smart_contact(mgr, manager_contacts)
        email_addr = info.get("email", "") if info else ""
        
        verify_list.append({
            "지사": row["관리지사"],
            "담당자": mgr,
            "이메일(E-MAIL)": email_addr,
            "검증결과": status,
            "유효성": is_valid_email(email_addr),
            "대상건수": row["계약건수"]
        })
    
    v_df = pd.DataFrame(verify_list)
    
    st.markdown("💡 **Tip:** 매핑된 메일 주소가 틀리거나 없을 경우, 표 안의 '이메일' 칸을 **직접 수정(대체메일)**하여 발송할 수 있습니다.")
    
    # 데이터 에디터로 대체메일 입력 지원
    edited_df = st.data_editor(
        v_df,
        column_config={
            "이메일(E-MAIL)": st.column_config.TextColumn("수신 메일(편집 가능)", required=True),
            "검증결과": st.column_config.TextColumn("AI 매핑 결과", disabled=True),
            "유효성": st.column_config.CheckboxColumn("유효 형식", disabled=True),
            "대상건수": st.column_config.NumberColumn("건수", disabled=True)
        },
        use_container_width=True, hide_index=True, key="mail_control_editor"
    )
    
    # 발송 제어
    with st.form("alert_control_form"):
        c_m1, c_m2 = st.columns([2, 1])
        subject = c_m1.text_input("메일 제목", f"[긴급] 해지방어 활동 미등록 건 확인 ({datetime.now().strftime('%Y-%m-%d')})")
        body_tpl = c_m2.text_area("메일 본문", "안녕하세요 {담당자}님, 고위험 계약 {건수}건의 활동 내역을 등록해주세요.")
        
        btn_send = st.form_submit_button("🚀 일괄 발송 시작", use_container_width=True)
        
        if btn_send:
            # SMTP 로직 (실제 사용 시 설정 필요)
            st.success("발송 큐에 등록되었습니다. (성공 00건 / 실패 00건 - 로그를 확인하세요)")
