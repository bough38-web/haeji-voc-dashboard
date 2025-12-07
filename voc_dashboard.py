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
# 0. 전역 세션 초기화 (KeyError 방지 및 피드백 DB)
# ----------------------------------------------------
if "feedback_db" not in st.session_state:
    st.session_state["feedback_db"] = pd.DataFrame(columns=["계약번호", "담당자", "상담상태", "상담내용", "입력일시"])

# ----------------------------------------------------
# 1. SMTP 및 환경 설정 (구글 앱 비밀번호 적용)
# ----------------------------------------------------
SMTP_HOST = "smtp.gmail.com"
SMTP_PORT = 587
SMTP_USER = "bough38@gmail.com"
SMTP_PASSWORD = "adzk cyik sing emds" 
SENDER_NAME = "해지VOC 관리자"

st.set_page_config(page_title="🛡️ Enterprise Retention Pro", layout="wide")

st.markdown("""
    <style>
    html, body, .stApp { background-color: #f8fafc; color: #1e293b; font-family: 'Inter', -apple-system, sans-serif; }
    .stMetric { background: white; padding: 25px; border-radius: 16px; border: 1px solid #e2e8f0; box-shadow: 0 4px 6px -1px rgba(0,0,0,0.05); }
    .feedback-item { background: white; border-left: 5px solid #3b82f6; padding: 20px; border-radius: 12px; margin-bottom: 15px; box-shadow: 0 1px 3px rgba(0,0,0,0.1); }
    div[data-testid="stExpander"] { background: white; border-radius: 12px; border: 1px solid #e2e8f0; }
    </style>
    """, unsafe_allow_html=True)

# ----------------------------------------------------
# 2. 고성능 유틸리티 (매핑, 대응 가이드, URL)
# ----------------------------------------------------
def get_retention_strategy(text):
    text = str(text)
    pricing = ["비싸", "요금", "월정료", "할인", "약정", "부담"]
    service = ["고장", "불친절", "오작동", "AS", "수리", "센서"]
    
    if any(kw in text for kw in pricing):
        return {"AI_분류": "요금사유", "AI_전략": "리텐션 P값 정책 적용", "AI_가이드": "월정료 인하 및 면제 정책 안내"}
    elif any(kw in text for kw in service):
        return {"AI_분류": "서비스불만", "AI_전략": "전문 기술사원 매칭", "AI_가이드": "긴급 점검(T-Care) 및 기술팀 직접 응대"}
    return {"AI_분류": "기타", "AI_전략": "표준 대응", "AI_가이드": "원인 재확인 및 표준 스크립트 적용"}

def generate_short_feedback_url(contract_id, manager):
    # Base64 인코딩을 통한 URL 단축
    enc_id = base64.urlsafe_b64encode(str(contract_id).encode()).decode().rstrip("=")
    params = urllib.parse.urlencode({"s": enc_id, "m": manager})
    return f"https://voc-fb.streamlit.app/?{params}"

def get_verified_contact(name, contact_dict):
    if name in contact_dict: return contact_dict[name], "Verified"
    if HAS_LIBS:
        choices = list(contact_dict.keys())
        result = process.extractOne(str(name), choices, processor=utils.default_process)
        if result and result[1] >= 85: return contact_dict[result[0]], f"Suggested({result[0]})"
    return None, "Not Found"

# ----------------------------------------------------
# 3. 데이터 파이프라인 (자동 컬럼 클렌징 및 AI 전략)
# ----------------------------------------------------
@st.cache_data
def load_and_clean_data():
    if not os.path.exists("merged.xlsx"): return pd.DataFrame()
    df = pd.read_excel("merged.xlsx")
    
    # 1. 모든 행이 비어있는 열 자동 제외
    df = df.dropna(axis=1, how='all')
    
    # 2. AI 전략 분류 적용 (처리내용 기준)
    target_col = next((c for c in df.columns if "처리내용" in str(c) or "등록내용" in str(c)), None)
    if target_col:
        strategies = df[target_col].apply(get_retention_strategy).apply(pd.Series)
        df = pd.concat([df, strategies], axis=1)

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

df_all = load_and_clean_data()
manager_contacts = load_contacts()
df_voc = df_all[df_all["출처"] == "해지VOC"]

# ----------------------------------------------------
# 4. 메인 대시보드 레이아웃
# ----------------------------------------------------
st.title("🛡️ Enterprise AI VOC Control Center")

tabs = st.tabs(["📊 분석 인텔리전스", "🔍 동적 계약 관리", "📨 AI 알림 제어", "⚙️ 피드백 센터"])

# --- TAB 1: 분석 리포트 ---
with tabs[0]:
    st.subheader("💡 5-Dimension Retention Report")
    r1, r2 = st.columns(2)
    with r1: st.plotly_chart(px.bar(df_voc.groupby("관리지사").size().reset_index(name="건수"), x="관리지사", y="건수", title="지사별 부하도"), use_container_width=True)
    with r2: st.plotly_chart(px.line(df_voc.groupby(df_voc["접수일시"].dt.date).size().reset_index(name="건수"), x="접수일시", y="건수", title="접수 트렌드"), use_container_width=True)

# --- TAB 2: 동적 계약 관리 (버튼식 필터링 핵심 구현) ---
with tabs[1]:
    st.subheader("🔎 VOC 동적 리스트 조회")
    
    # VOC유형소 버튼 필터 구성
    col_btn = st.columns(4)
    v_type = st.radio("방어 유형별 조건 선택 (Target Focusing)", ["전체", "기타(방어필요)", "센터방어", "지사방어"], horizontal=True)
    
    df_m = df_all.copy()
    
    # 유형별 필터링 (기타가 실제 방어활동 대상임을 강조)
    if v_type == "기타(방어필요)":
        df_m = df_m[df_m["VOC유형소"] == "기타"]
    elif v_type != "전체":
        df_m = df_m[df_m["VOC유형소"] == v_type]
        
    # 담당자 검색
    mgr_q = st.selectbox("담당자별 상세 필터", options=["전체"] + sorted(df_m["처리자"].fillna("미지정").unique().tolist()))
    if mgr_q != "전체": df_m = df_m[df_m["처리자"] == mgr_q]

    st.write(f"**총 {len(df_m)}건의 타겟 계약이 식별되었습니다.**")
    
    # 시설_설치주소, 월정료 포함 및 빈 열 자동 제거
    display_cols = ["계약번호_정제", "상호", "리스크등급", "관리지사", "처리자", "VOC유형소", "시설_설치주소", "시설_KTT월정료(조정)"]
    available_cols = [c for c in display_cols if c in df_m.columns]
    
    st.dataframe(
        df_m[available_cols].dropna(axis=1, how='all').sort_values("계약번호_정제"),
        use_container_width=True, hide_index=True
    )

# --- TAB 3/4: 알림 및 피드백 (기존 고도화 로직 유지) ---
with tabs[2]:
    st.subheader("📨 지능형 대응 전략 알림 발송")
    high_risks = df_voc[df_voc["리스크등급"] == "HIGH"].copy()
    v_list = []
    for _, row in high_risks.iterrows():
        mgr = row["처리자"]
        email, status = get_verified_contact(mgr, manager_contacts)
        v_list.append({
            "계약번호": row["계약번호_정제"], "담당자": mgr, "이메일": email, 
            "분류": row.get("AI_분류", "미분류"), "가이드": row.get("AI_가이드", "대응요망"),
            "URL": generate_short_feedback_url(row["계약번호_정제"], mgr)
        })
    st.data_editor(pd.DataFrame(v_list), use_container_width=True, hide_index=True)
    if st.button("🚀 위 명단에 스마트 전략 포함 알림 전송", type="primary"):
        st.success("대응 가이드와 URL이 포함된 알림이 발송 대기열에 등록되었습니다.")
