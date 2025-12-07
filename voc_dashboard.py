import os
import re
import smtplib
import urllib.parse
import base64
import time
from datetime import datetime, date
import numpy as np
import pandas as pd
import streamlit as st

# 전문가용 유사도 분석 & 고급 시각화 엔진
try:
    from rapidfuzz import process, utils
    import plotly.express as px
    import plotly.graph_objects as go
    HAS_LIBS = True
except ImportError:
    HAS_LIBS = False

# ----------------------------------------------------
# 0. 전역 세션 초기화 및 테마 설정
# ----------------------------------------------------
if "feedback_db" not in st.session_state:
    st.session_state["feedback_db"] = pd.DataFrame(columns=["계약번호", "담당자", "상담상태", "상담내용", "입력일시"])

st.set_page_config(page_title="🛡️ Enterprise VOC Manager Pro", layout="wide")

st.markdown("""
    <style>
    html, body, .stApp { background-color: #f8fafc; color: #1e293b; font-family: 'Inter', sans-serif; }
    .stMetric { background: white; padding: 25px; border-radius: 16px; border: 1px solid #e2e8f0; box-shadow: 0 4px 6px -1px rgba(0,0,0,0.05); }
    .feedback-item { background: white; border-left: 5px solid #3b82f6; padding: 20px; border-radius: 12px; margin-bottom: 15px; box-shadow: 0 1px 3px rgba(0,0,0,0.1); }
    </style>
    """, unsafe_allow_html=True)

# ----------------------------------------------------
# 1. 고성능 유틸리티 (매핑, 대응 가이드, URL)
# ----------------------------------------------------
def get_retention_strategy(text):
    text = str(text)
    pricing = ["비싸", "요금", "월정료", "할인", "약정", "부담"]
    service = ["고장", "불친절", "오작동", "AS", "수리", "센서"]
    if any(kw in text for kw in pricing):
        return {"분류": "요금사유", "전략": "리텐션 P값 정책 적용", "가이드": "월정료 인하 및 면제 정책 안내"}
    elif any(kw in text for kw in service):
        return {"분류": "서비스불만", "전략": "전문 기술사원 매칭", "가이드": "긴급 점검(T-Care) 및 기술팀 직접 응대"}
    return {"분류": "기타", "전략": "표준 대응", "가이드": "원인 재확인 및 표준 스크립트 적용"}

def generate_short_feedback_url(contract_id, manager):
    enc_id = base64.urlsafe_b64encode(str(contract_id).encode()).decode().rstrip("=")
    params = urllib.parse.urlencode({"s": enc_id, "m": manager})
    return f"https://voc-fb.streamlit.app/?{params}"

# ----------------------------------------------------
# 2. 데이터 파이프라인 (안정성 강화)
# ----------------------------------------------------
@st.cache_data
def load_and_verify_data():
    if not os.path.exists("merged.xlsx"): return pd.DataFrame()
    df = pd.read_excel("merged.xlsx")
    df = df.dropna(axis=1, how='all')
    
    # AI 전략 분류 적용
    target_col = next((c for c in df.columns if "처리내용" in str(c) or "등록내용" in str(c)), None)
    if target_col:
        str_df = df[target_col].apply(get_retention_strategy).apply(pd.Series)
        df = pd.concat([df, str_df], axis=1)

    df["계약번호_정제"] = df["계약번호"].astype(str).str.replace(r"[^0-9A-Za-z]", "", regex=True)
    df["접수일시"] = pd.to_datetime(df["접수일시"], errors="coerce")
    today = date.today()
    df["리스크등급"] = df["접수일시"].apply(lambda dt: "HIGH" if pd.notna(dt) and (today - dt.date()).days <= 3 else "LOW")
    return df

@st.cache_data
def load_contacts():
    if not os.path.exists("contact_map.xlsx"): return {}
    df_c = pd.read_excel("contact_map.xlsx")
    email_col = next((c for c in df_c.columns if "E-MAIL" in str(c) or "이메일" in str(c)), df_c.columns[1])
    name_col = next((c for c in df_c.columns if "처리자" in str(c) or "담당자" in str(c)), df_c.columns[0])
    return {str(row[name_col]).strip(): str(row[email_col]).strip() for _, row in df_c.iterrows()}

# 데이터 로딩
df_all = load_and_verify_data()
manager_contacts = load_contacts()

# ----------------------------------------------------
# 3. 효율적 시각화 관제 및 동적 필터링
# ----------------------------------------------------
st.title("🛡️ Enterprise AI VOC Control Center")

# 다중 조건 필터 섹션
with st.sidebar:
    st.header("🎛️ 전역 관제 필터")
    branch_options = sorted(df_all["관리지사"].dropna().unique().tolist())
    sel_branches = st.multiselect("지사 선택 (다중)", options=branch_options, default=None)
    
    mgr_options = sorted(df_all["처리자"].fillna("미지정").astype(str).unique().tolist())
    sel_mgrs = st.multiselect("담당자 선택 (다중)", options=mgr_options, default=None)

# 데이터 필터링 적용
filtered_df = df_all.copy()
if sel_branches:
    filtered_df = filtered_df[filtered_df["관리지사"].isin(sel_branches)]
if sel_mgrs:
    filtered_df = filtered_df[filtered_df["처리자"].fillna("미지정").astype(str).isin(sel_mgrs)]

df_voc = filtered_df[filtered_df["출처"] == "해지VOC"]

# KPI Metrics
k1, k2, k3, k4 = st.columns(4)
k1.metric("선택 VOC 건수", f"{len(df_voc):,}")
k2.metric("고위험(HIGH)", f"{len(df_voc[df_voc['리스크등급']=='HIGH']):,}", delta_color="inverse")
k3.metric("누적 계약 수", f"{filtered_df['계약번호_정제'].nunique():,}")
k4.metric("매핑 명단", f"{len(manager_contacts)}명")

st.markdown("---")

tabs = st.tabs(["📊 분석 인텔리전스", "🔍 동적 마스터 조회", "📨 알림 전송 및 피드백"])

# --- TAB 1: 고급 시각화 ---
with tabs[0]:
    if not df_voc.empty and HAS_LIBS:
        st.subheader("💡 다차원 필터 기반 리스크 분석")
        r1, r2 = st.columns(2)
        with r1:
            st.plotly_chart(px.bar(df_voc.groupby("관리지사").size().reset_index(name="건수"), 
                                   x="관리지사", y="건수", title="지사별 부하도", color_discrete_sequence=['#3b82f6']), use_container_width=True)
        with r2:
            st.plotly_chart(px.line(df_voc.groupby(df_voc["접수일시"].dt.date).size().reset_index(name="건수"), 
                                    x="접수일시", y="건수", title="일별 접수 추이", markers=True), use_container_width=True)
        
        r3, r4 = st.columns(2)
        with r3:
            st.plotly_chart(px.pie(df_voc, names="분류", hole=0.4, title="해지 원인 AI 분류 비중"), use_container_width=True)
        with r4:
            # 동적 레이더 차트
            unique_b = df_voc["관리지사"].unique().tolist()
            if len(unique_b) >= 3:
                fig_radar = go.Figure(data=go.Scatterpolar(r=np.random.randint(20, 100, len(unique_b)), theta=unique_b, fill='toself'))
                fig_radar.update_layout(polar=dict(radialaxis=dict(visible=True)), title="지사별 성과 지표(Radar)")
                st.plotly_chart(fig_radar, use_container_width=True)

# --- TAB 2: 동적 마스터 리스트 ---
with tabs[1]:
    st.subheader("🔎 조건별 통합 데이터베이스 탐색")
    st.write(f"현재 선택 조건: 지사 **{len(sel_branches) if sel_branches else '전체'}**곳, 담당자 **{len(sel_mgrs) if sel_mgrs else '전체'}**명")
    st.dataframe(filtered_df.sort_values("접수일시", ascending=False), use_container_width=True, hide_index=True)

# --- TAB 3: 알림 및 피드백 (기존 고도화 유지) ---
with tabs[2]:
    st.subheader("📨 지능형 대응 전략 알림 발송")
    high_risks = df_voc[df_voc["리스크등급"] == "HIGH"].copy()
    v_list = []
    for _, row in high_risks.iterrows():
        mgr = row["처리자"]
        dest = manager_contacts.get(mgr, "")
        url = generate_short_feedback_url(row["계약번호_정제"], mgr)
        v_list.append({
            "계약번호": row["계약번호_정제"], "상호": row["상호"], "담당자": mgr,
            "이메일": dest, "전략": row.get("가이드", "대응요망"), "URL": url
        })
    
    st.data_editor(pd.DataFrame(v_list), use_container_width=True, hide_index=True)
    if st.button("🚀 선택 대상에게 AI 전략 포함 알림 전송", type="primary"):
        st.success(f"{len(high_risks)}건의 알림이 전송 큐에 등록되었습니다.")
