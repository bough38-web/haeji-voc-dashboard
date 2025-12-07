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

# 전문가용 유사도 분석 & 고급 시각화 엔진 로드
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
# 1. 고성능 유틸리티 (인코딩 & 매핑 엔진)
# ----------------------------------------------------
def encode_id(text):
    try: return base64.urlsafe_b64encode(str(text).encode()).decode().rstrip("=")
    except: return text

def generate_short_feedback_url(contract_id, manager):
    params = urllib.parse.urlencode({"s": encode_id(contract_id), "m": manager})
    return f"https://voc-fb.streamlit.app/?{params}"

def get_retention_strategy(text):
    text = str(text)
    pricing = ["비싸", "요금", "월정료", "할인", "약정", "부담"]
    if any(kw in text for kw in pricing):
        return {"분류": "요금사유", "전략": "리텐션 P값 적용", "가이드": "할인 및 면제 정책"}
    return {"분류": "일반", "전략": "표준 대응", "가이드": "원인 재확인"}

# ----------------------------------------------------
# 2. 데이터 관리 파이프라인 (본부 필터링 포함)
# ----------------------------------------------------
@st.cache_data
def load_and_verify_data():
    if not os.path.exists("merged.xlsx"): return pd.DataFrame()
    df = pd.read_excel("merged.xlsx")
    df = df.dropna(axis=1, how='all')
    
    # [핵심] 관리본부 필터 적용: 강북/강원본부, 강원본부만 포함
    if "관리본부명" in df.columns:
        valid_hq = ["강북/강원본부", "강원본부"]
        df = df[df["관리본부명"].isin(valid_hq)]
    
    df["계약번호_정제"] = df["계약번호"].astype(str).str.replace(r"[^0-9A-Za-z]", "", regex=True)
    df["접수일시"] = pd.to_datetime(df["접수일시"], errors="coerce")
    today = date.today()
    df["리스크등급"] = df["접수일시"].apply(lambda dt: "HIGH" if pd.notna(dt) and (today - dt.date()).days <= 3 else "LOW")
    
    # 전략 적용
    target_col = next((c for c in df.columns if "처리내용" in str(c) or "등록내용" in str(c)), None)
    if target_col:
        str_df = df[target_col].apply(get_retention_strategy).apply(pd.Series)
        df = pd.concat([df, str_df], axis=1)
        
    return df

@st.cache_data
def load_contacts():
    if not os.path.exists("contact_map.xlsx"): return {}
    df_c = pd.read_excel("contact_map.xlsx")
    email_col = next((c for c in df_c.columns if "E-MAIL" in str(c) or "이메일" in str(c)), df_c.columns[1])
    name_col = next((c for c in df_c.columns if "처리자" in str(c) or "담당자" in str(c)), df_c.columns[0])
    return {str(row[name_col]).strip(): str(row[email_col]).strip() for _, row in df_c.iterrows()}

df_all = load_and_verify_data()
manager_contacts = load_contacts()
df_voc = df_all[df_all["출처"] == "해지VOC"]

# ----------------------------------------------------
# 3. 동적 관제 UI 및 탭 구성
# ----------------------------------------------------
st.title("🛡️ Enterprise AI Retention Control Center")

tabs = st.tabs(["📊 분석 인텔리전스", "🔍 동적 계약 마스터", "📨 AI 알림 및 피드백"])

# --- TAB 1: 고급 시각화 ---
with tabs[0]:
    if not df_voc.empty:
        st.subheader("💡 다차원 통합 분석 리포트")
        r1, r2 = st.columns(2)
        with r1: st.plotly_chart(px.bar(df_voc.groupby("관리지사").size().reset_index(name="건수"), x="관리지사", y="건수", title="지사별 부하도"), use_container_width=True)
        with r2: st.plotly_chart(px.pie(df_voc, names="분류", title="AI 이슈 분류"), use_container_width=True)

# --- TAB 2: 동적 계약 마스터 (지사-담당자 연동 필터) ---
with tabs[1]:
    st.subheader("🔎 조건별 통합 데이터베이스 탐색")
    
    c1, c2 = st.columns(2)
    
    # 1. 지사 다중 선택
    unique_branches = sorted(df_all["관리지사"].unique().tolist())
    sel_branches = c1.multiselect("관리지사 선택", options=unique_branches)
    
    # 2. 지사 선택에 따른 담당자 목록 동적 갱신
    if sel_branches:
        subset_df = df_all[df_all["관리지사"].isin(sel_branches)]
    else:
        subset_df = df_all
        
    unique_mgrs = sorted(subset_df["처리자"].fillna("미지정").unique().tolist())
    sel_mgrs = c2.multiselect("담당자 선택 (지사별 동적 목록)", options=unique_mgrs)
    
    # 최종 필터링 데이터 구성
    df_m = subset_df.copy()
    if sel_mgrs:
        df_m = df_m[df_m["처리자"].fillna("미지정").isin(sel_mgrs)]

    # 정렬 필수 컬럼 확보 및 공백열 자동 제외
    available_cols = ["계약번호_정제", "상호", "리스크등급", "관리지사", "처리자", "시설_설치주소", "시설_KTT월정료(조정)", "접수일시"]
    final_cols = df_m[[c for c in available_cols if c in df_m.columns]].dropna(axis=1, how='all').columns.tolist()
    
    if "접수일시" in final_cols:
        display_df = df_m[final_cols].sort_values("접수일시", ascending=False)
    else:
        display_df = df_m[final_cols]
        
    st.write(f"**총 {len(df_m)}건 검색됨**")
    st.dataframe(display_df, use_container_width=True, hide_index=True)

# --- TAB 3: 알림 전송 (동적 연동 반영) ---
with tabs[2]:
    st.subheader("📨 전략 기반 자동 알림 및 가이드")
    # 탭 2의 필터링 결과를 그대로 알림 대상으로 활용 가능하도록 연결
    high_targets = df_m[df_m["리스크등급"] == "HIGH"].copy()
    
    v_list = []
    for _, row in high_targets.iterrows():
        m = str(row["처리자"])
        email = manager_contacts.get(m, "")
        url = generate_short_feedback_url(row["계약번호_정제"], m)
        v_list.append({
            "계약번호": row["계약번호_정제"], "상호": row["상호"], "담당자": m,
            "이메일": email, "전략가이드": row.get("가이드", "확인요망"), "URL": url
        })
    
    edited_agg = st.data_editor(pd.DataFrame(v_list), use_container_width=True, hide_index=True)
    if st.button("🚀 선택 대상 스마트 알림 발송", type="primary"):
        st.success(f"{len(edited_agg)}건 전송 큐 등록 완료")
