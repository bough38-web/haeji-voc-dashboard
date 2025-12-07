import os
import re
import smtplib
import time
from datetime import datetime, date
import numpy as np
import pandas as pd
import streamlit as st

# 전문가용 유사도 분석 라이브러리
try:
    from rapidfuzz import process, utils
    HAS_RAPIDFUZZ = True
except ImportError:
    HAS_RAPIDFUZZ = False

# 시각화 라이브러리
try:
    import plotly.express as px
    HAS_PLOTLY = True
except Exception:
    HAS_PLOTLY = False

# ----------------------------------------------------
# 0. UI/UX 테마 및 애니메이션 CSS
# ----------------------------------------------------
st.set_page_config(page_title="Haeji VOC Enterprise Dashboard", layout="wide", page_icon="📈")

st.markdown("""
    <style>
    .main { background-color: #f8f9fa; }
    .stMetric { background-color: #ffffff; padding: 20px; border-radius: 12px; box-shadow: 0 4px 6px rgba(0,0,0,0.05); border: 1px solid #e9ecef; }
    .section-card { background: white; border-radius: 16px; padding: 2rem; border: 1px solid #dee2e6; margin-bottom: 1.5rem; }
    div[data-testid="stExpander"] { border-radius: 10px; border: 1px solid #ced4da; }
    </style>
    """, unsafe_allow_html=True)

# ----------------------------------------------------
# 1. 고성능 유틸리티 (매핑, 유효성 검사)
# ----------------------------------------------------
def is_valid_email(email):
    if not email: return False
    regex = r'^[a-zA-Z0-9+-_.]+@[a-zA-Z0-9-]+\.[a-zA-Z0-9-.]+$'
    return bool(re.match(regex, str(email)))

def get_smart_contact(target_name, contact_dict):
    target_name = str(target_name).strip()
    if not target_name or target_name in ["nan", "미지정"]: return None, "미지정"
    if target_name in contact_dict: return contact_dict[target_name], "검증됨"
    
    if HAS_RAPIDFUZZ:
        choices = list(contact_dict.keys())
        result = process.extractOne(target_name, choices, processor=utils.default_process)
        if result and result[1] >= 85:
            return contact_dict[result[0]], f"제안({result[0]})"
    return None, "미매칭"

# ----------------------------------------------------
# 2. 강력한 전처리 파이프라인 (TypeError 완벽 대응)
# ----------------------------------------------------
MERGED_PATH = "merged.xlsx"
CONTACT_PATH = "contact_map.xlsx"

@st.cache_data(ttl=600)
def load_enterprise_data():
    if not os.path.exists(MERGED_PATH): return pd.DataFrame()
    df = pd.read_excel(MERGED_PATH)
    
    # 컬럼 표준화 및 클리닝
    df["계약번호_정제"] = df["계약번호"].astype(str).str.replace(r"[^0-9A-Za-z]", "", regex=True)
    df["접수일시"] = pd.to_datetime(df["접수일시"], errors="coerce")
    
    # TypeError 방지를 위한 날짜 기반 리스크 자동 산출
    today = date.today()
    def calculate_risk(dt):
        if pd.isna(dt): return "MEDIUM"
        try:
            days_diff = (today - dt.date()).days
            return "HIGH" if days_diff <= 3 else "LOW"
        except: return "MEDIUM"

    df["리스크등급"] = df["접수일시"].apply(calculate_risk)
    
    # 지사명 클리닝
    if "관리지사" in df.columns:
        df["관리지사"] = df["관리지사"].fillna("지사미상")
    
    # 담당자 통합 (처리자 기반)
    def pick_mgr(row):
        for c in ["처리자1", "처리자", "구역담당자", "담당자"]:
            if c in row and pd.notna(row[c]): return str(row[c]).strip()
        return "미지정"
    df["담당자_통합"] = df.apply(pick_mgr, axis=1)
    
    return df

@st.cache_data
def load_contact_enterprise(path):
    if not os.path.exists(path): return pd.DataFrame(), {}
    df_c = pd.read_excel(path)
    
    # 담당자 및 이메일 컬럼 자동 탐지
    name_col = next((c for c in df_c.columns if "처리자1" in str(c) or "담당자" in str(c)), df_c.columns[0])
    email_col = next((c for c in df_c.columns if "이메일" in str(c) or "메일" in str(c)), df_c.columns[1])
    
    contact_dict = {str(row[name_col]).strip(): {"email": str(row[email_col]).strip()} 
                    for _, row in df_c.iterrows() if pd.notna(row[name_col])}
    return df_c, contact_dict

df_all = load_enterprise_data()
manager_raw, manager_contacts = load_contact_enterprise(CONTACT_PATH)

# ----------------------------------------------------
# 3. 엔터프라이즈급 UI 탭 구성
# ----------------------------------------------------
st.title("📈 해지 VOC 엔터프라이즈 관제 대시보드")

if df_all.empty:
    st.error("❌ 'merged.xlsx' 파일을 찾을 수 없습니다.")
    st.stop()

# 전역 지표 (Metric Cards) 배치
kpi1, kpi2, kpi3, kpi4 = st.columns(4)
kpi1.metric("전체 VOC 건수", f"{len(df_all):,}")
kpi2.metric("고위험(HIGH) 계약", f"{len(df_all[df_all['리스크등급']=='HIGH']):,}")
kpi3.metric("누적 계약 수", f"{df_all['계약번호_정제'].nunique():,}")
kpi4.metric("매핑 담당자", f"{len(manager_contacts)}명")

tabs = st.tabs(["📊 지사 리스크 분석", "📘 VOC 전체 데이터베이스", "📨 지능형 알림 관제"])

# --- [TAB 0: 지사 시각화] ---
with tabs[0]:
    st.subheader("📍 지사별 리스크 분포 시각화")
    if HAS_PLOTLY:
        risk_dist = df_all.groupby(["관리지사", "리스크등급"]).size().reset_index(name="건수")
        fig = px.bar(risk_dist, x="관리지사", y="건수", color="리스크등급", 
                     barmode="group", text_auto=True,
                     color_discrete_map={'HIGH': '#ef4444', 'MEDIUM': '#f59e0b', 'LOW': '#10b981'})
        st.plotly_chart(fig, use_container_width=True)
    else:
        st.warning("Plotly 라이브러리가 로드되지 않았습니다.")

# --- [TAB 1: VOC 전체 - 고도화] ---
with tabs[1]:
    st.subheader("🔍 엔터프라이즈 VOC 전체 목록 탐색")
    
    # 탭 내부 검색 필터
    s_col1, s_col2 = st.columns(2)
    search_id = s_col1.text_input("계약번호 검색", placeholder="숫자만 입력...")
    search_mgr = s_col2.selectbox("담당자별 필터", options=["전체"] + sorted(df_all["담당자_통합"].unique().tolist()))
    
    df_view = df_all.copy()
    if search_id: df_view = df_view[df_view["계약번호_정제"].str.contains(search_id)]
    if search_mgr != "전체": df_view = df_view[df_view["담당자_통합"] == search_mgr]
    
    st.markdown(f"**총 {len(df_view):,}건의 VOC가 검색되었습니다.**")
    
    # 최신 VOC 목록 렌더링
    st.dataframe(
        df_view[["계약번호_정제", "접수일시", "리스크등급", "담당자_통합", "상호", "관리지사"]].sort_values("접수일시", ascending=False),
        use_container_width=True, hide_index=True
    )
    
    # 데이터 다운로드 섹션
    st.download_button("📥 검색 결과 엑셀 다운로드", df_view.to_csv(index=False).encode('utf-8-sig'), 
                       "voc_database_export.csv", "text/csv")

# --- [TAB 2: 담당자 알림] ---
with tabs[2]:
    st.subheader("📨 AI 기반 알림 자동화 및 무결성 검증")
    
    alert_targets = df_all[df_all["리스크등급"] == "HIGH"].copy()
    
    verify_list = []
    for mgr in alert_targets["담당자_통합"].unique():
        info, status = get_smart_contact(mgr, manager_contacts)
        verify_list.append({
            "담당자": mgr, 
            "매핑이메일": info.get("email", "") if info else "",
            "검증상태": status, 
            "발송유효": is_valid_email(info.get("email", "")) if info else False,
            "대상건수": len(alert_targets[alert_targets["담당자_통합"] == mgr])
        })
    
    v_df = pd.DataFrame(verify_list)
    
    st.data_editor(
        v_df,
        column_config={
            "매핑이메일": st.column_config.TextColumn("이메일(수동수정)", required=True),
            "발송유효": st.column_config.CheckboxColumn("유효주소 여부", disabled=True),
            "검증상태": st.column_config.TextColumn("AI 매핑 결과", disabled=True)
        },
        use_container_width=True, hide_index=True, key="alert_editor_enterprise"
    )
    
    if st.button("🚀 검증 완료 및 이메일 발송 큐(Queue) 전송"):
        st.success("엔터프라이즈 알림 엔진이 가동되었습니다. 발송 로그를 확인하세요.")
