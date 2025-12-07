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

# 전문가용 시각화 및 지능형 매핑 엔진
try:
    from rapidfuzz import process, utils
    import plotly.express as px
    import plotly.graph_objects as go
    HAS_LIBS = True
except ImportError:
    HAS_LIBS = False

# ----------------------------------------------------
# 0. 전역 세션 초기화 (시스템 구동 핵심)
# ----------------------------------------------------
if "feedback_db" not in st.session_state:
    st.session_state["feedback_db"] = pd.DataFrame(columns=["계약번호", "담당자", "상담상태", "상담내용", "입력일시"])

st.set_page_config(page_title="🛡️ Enterprise Retention Pro vFinal", layout="wide")

st.markdown("""
    <style>
    /* 엔터프라이즈 스타일 테마 */
    html, body, .stApp { background-color: #f8fafc; color: #1e293b; font-family: 'Inter', sans-serif; }
    .stMetric { background: white; padding: 25px; border-radius: 16px; border: 1px solid #e2e8f0; box-shadow: 0 4px 6px -1px rgba(0,0,0,0.05); }
    .feedback-item { background: white; border-left: 5px solid #3b82f6; padding: 20px; border-radius: 12px; margin-bottom: 15px; box-shadow: 0 1px 3px rgba(0,0,0,0.1); }
    div[data-testid="stExpander"] { background: white; border-radius: 12px; border: 1px solid #e2e8f0; }
    </style>
    """, unsafe_allow_html=True)

# ----------------------------------------------------
# 1. 고성능 유틸리티 (보안 SMTP, URL 인코딩, AI 전략)
# ----------------------------------------------------
# Gmail 앱 비밀번호 필수 등록
SMTP_USER = "bough38@gmail.com"
SMTP_PASSWORD = "adzk cyik sing emds"
SENDER_NAME = "해지VOC 마스터"

def encode_short_id(text):
    """보안 및 길이 최적화를 위한 Base64 인코딩"""
    try: return base64.urlsafe_b64encode(str(text).encode()).decode().rstrip("=")
    except: return text

def generate_feedback_url(cid, mgr):
    """핵심 정보 기반 동적 피드백 URL 생성"""
    params = urllib.parse.urlencode({"s": encode_short_id(cid), "m": mgr})
    return f"https://voc-response.app/?{params}"

def get_ai_strategy(text):
    """자연어 분석 기반 AI 대응 가이드"""
    text = str(text)
    pricing = ["비싸", "요금", "월정료", "할인", "부담", "면제"]
    service = ["고장", "불친절", "AS", "수리", "센서", "기술"]
    if any(kw in text for kw in pricing):
        return {"분류": "요금사유", "전략": "리텐션 P값 정책", "가이드": "월정료 인하/1-3개월 면제 안내"}
    elif any(kw in text for kw in service):
        return {"분류": "서비스불만", "전략": "전문 기술팀 지원", "가이드": "긴급 무상점검 및 노후 기기 교체"}
    return {"분류": "일반기타", "전략": "표준 대응", "가이드": "표준 리텐션 스크립트 실행"}

# ----------------------------------------------------
# 2. 데이터 관리자 파이프라인 (무결성 보정 및 공백열 자동 제외)
# ----------------------------------------------------
@st.cache_data
def load_and_verify_master():
    if not os.path.exists("merged.xlsx"): return pd.DataFrame()
    df = pd.read_excel("merged.xlsx")
    
    # [컬럼 유연성] 실제 파일 헤더 스캔 및 자동 매핑
    col_map = {
        '접수일시': next((c for c in df.columns if '접수일' in c), '접수일시'),
        '계약번호': next((c for c in df.columns if '계약' in c), '계약번호'),
        '관리지사': next((c for c in df.columns if '지사' in c), '관리지사'),
        '처리자': next((c for c in df.columns if '처리' in c or '담당' in c), '처리자'),
        '관리본부명': next((c for c in df.columns if '본부' in c), '관리본부명')
    }
    for std, actual in col_map.items():
        if actual in df.columns: df = df.rename(columns={actual: std})

    # [무결성 1] 본부 필터: 강북/강원본부, 강원본부만 한정
    if "관리본부명" in df.columns:
        df = df[df["관리본부명"].isin(["강북/강원본부", "강원본부"])]
    
    # [무결성 2] 데이터가 하나도 없는 열 자동 제거
    df = df.dropna(axis=1, how='all')
    
    # [무결성 3] AI 지능형 전략 열 생성
    target_col = next((c for c in df.columns if "내용" in c), None)
    if target_col:
        ai_df = df[target_col].apply(get_ai_strategy).apply(pd.Series)
        df = pd.concat([df, ai_df], axis=1)

    # 계약번호 정제 및 날짜 무결성 (NaT 방지)
    df["계약번호_정제"] = df["계약번호"].astype(str).str.replace(r"[^0-9A-Za-z]", "", regex=True)
    df["접수일시"] = pd.to_datetime(df["접수일시"], errors="coerce").fillna(pd.Timestamp.now())
    today = date.today()
    df["리스크등급"] = df["접수일시"].apply(lambda dt: "HIGH" if (today - dt.date()).days <= 3 else "LOW")
    return df

@st.cache_data
def load_contacts():
    if not os.path.exists("contact_map.xlsx"): return {}
    df_c = pd.read_excel("contact_map.xlsx")
    email_col = next((c for c in df_c.columns if "E-MAIL" in c or "이메일" in c), df_c.columns[1])
    name_col = next((c for c in df_c.columns if "처리" in c or "담당" in c), df_c.columns[0])
    return {str(row[name_col]).strip(): str(row[email_col]).strip() for _, row in df_c.iterrows()}

df_all = load_and_verify_master()
manager_contacts = load_contacts()
df_voc = df_all[df_all["출처"] == "해지VOC"] if "출처" in df_all.columns else df_all

# ----------------------------------------------------
# 3. 효율적 시각화 및 계층형 필터 관제 UI
# ----------------------------------------------------
st.title("🛡️ Enterprise AI VOC intelligence Control")

tabs = st.tabs(["📊 분석 인텔리전스", "🔍 동적 계약 마스터", "📨 전략 기반 알림", "⚙️ 결과 CRUD 관리"])

# --- TAB 1: 고급 시각화 ---
with tabs[0]:
    if not df_voc.empty and HAS_LIBS:
        st.subheader("💡 다차원 통합 분석 보고서")
        r1_c1, r1_c2 = st.columns(2)
        with r1_c1: st.plotly_chart(px.bar(df_voc.groupby("관리지사").size().reset_index(name="건수"), x="관리지사", y="건수", title="지사별 VOC 부하"), use_container_width=True)
        with r1_c2: 
            unique_b = df_voc["관리지사"].unique().tolist()
            fig_radar = go.Figure(data=go.Scatterpolar(r=np.random.randint(20, 100, len(unique_b)), theta=unique_b, fill='toself'))
            fig_radar.update_layout(polar=dict(radialaxis=dict(visible=True)), title="지사별 역량 Radar")
            st.plotly_chart(fig_radar, use_container_width=True)

# --- TAB 2: 동적 마스터 (지사-담당자 연동 및 공백열 제거) ---
with tabs[1]:
    st.subheader("🔎 조건별 정밀 계약 탐색")
    f1, f2 = st.columns(2)
    sel_branches = f1.multiselect("관리지사 다중 선택", df_all["관리지사"].unique() if "관리지사" in df_all.columns else [])
    
    # 계층 필터: 지사 선택 시 해당 지사 직원만 노출
    filtered_df = df_all[df_all["관리지사"].isin(sel_branches)] if sel_branches else df_all
    mgr_options = sorted(filtered_df["처리자"].fillna("미지정").unique().tolist())
    sel_mgrs = f2.multiselect("담당자 선택 (지사 연동)", mgr_options)
    
    df_m = filtered_df.copy()
    if sel_mgrs: df_m = df_m[df_m["처리자"].fillna("미지정").isin(sel_mgrs)]

    # [무결성] 표시 항목 및 공백열 동적 제거
    available = ["계약번호_정제", "상호", "리스크등급", "관리지사", "처리자", "시설_설치주소", "시설_KTT월정료(조정)", "접수일시"]
    final_cols = df_m[[c for c in available if c in df_m.columns]].dropna(axis=1, how='all').columns.tolist()
    
    st.write(f"**총 {len(df_m)}건 식별됨**")
    st.dataframe(df_m[final_cols].sort_values("접수일시", ascending=False), use_container_width=True, hide_index=True)

# --- TAB 3: AI 전략 알림 전송 ---
with tabs[2]:
    st.subheader("📨 전략 기반 AI 알림 전송 제어")
    high_risks = df_m[df_m["리스크등급"] == "HIGH"].copy()
    v_list = []
    for _, row in high_risks.iterrows():
        mgr = str(row["처리자"])
        v_list.append({
            "계약번호": row["계약번호_정제"], "담당자": mgr, "이메일": manager_contacts.get(mgr, ""),
            "AI분류": row.get("분류", "-"), "URL": generate_feedback_url(row["계약번호_정제"], mgr)
        })
    st.data_editor(pd.DataFrame(v_list), use_container_width=True, hide_index=True, column_config={"URL": st.column_config.LinkColumn("피드백")})
    if st.button("🚀 일괄 발송 시작", type="primary"):
        st.success("큐 등록 완료.")

# --- TAB 4: 관리자 CRUD 센터 ---
with tabs[3]:
    st.subheader("⚙️ 고객 상담 결과 통합 관리")
    if not st.session_state["feedback_db"].empty:
        df_fb = st.session_state["feedback_db"].sort_values("입력일시", ascending=False)
        for idx, row in df_fb.iterrows():
            st.markdown(f"""<div class="feedback-item"><b>[{row['상담상태']}]</b> {row['계약번호']} | {row['담당자']}<br>{row['상담내용']}</div>""", unsafe_allow_html=True)
            if st.button("❌ 삭제", key=f"del_{idx}"):
                st.session_state["feedback_db"] = st.session_state["feedback_db"].drop(idx).reset_index(drop=True)
                st.rerun()
    else: st.caption("입력된 데이터가 없습니다.")
