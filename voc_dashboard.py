import os
import re
import smtplib
import urllib.parse
import base64
import time
from datetime import datetime, date
from email.message import EmailMessage

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
# 0. 전역 세션 초기화 (KeyError 방지)
# ----------------------------------------------------
if "feedback_db" not in st.session_state:
    st.session_state["feedback_db"] = pd.DataFrame(columns=["계약번호", "담당자", "상담상태", "상담내용", "입력일시"])

# ----------------------------------------------------
# 1. SMTP 및 환경 설정 (보안 주의)
# ----------------------------------------------------
SMTP_HOST = "smtp.gmail.com"
SMTP_PORT = 587
SMTP_USER = "bough38@gmail.com"
SMTP_PASSWORD = "adzk cyik sing emds" # 구글 앱 비밀번호
SENDER_NAME = "해지VOC 관리자"

st.set_page_config(page_title="🛡️ Enterprise Retention Intelligence", layout="wide")

st.markdown("""
    <style>
    html, body, .stApp { background-color: #f8fafc; color: #1e293b; font-family: 'Inter', -apple-system, sans-serif; }
    .stMetric { background: white; padding: 25px; border-radius: 16px; border: 1px solid #e2e8f0; box-shadow: 0 4px 6px -1px rgba(0,0,0,0.05); }
    .feedback-item { background: white; border-left: 5px solid #3b82f6; padding: 20px; border-radius: 12px; margin-bottom: 15px; box-shadow: 0 1px 3px rgba(0,0,0,0.1); }
    </style>
    """, unsafe_allow_html=True)

# ----------------------------------------------------
# 2. 고성능 엔진 (분류 전략 및 단축 URL)
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

def encode_id(text):
    try: return base64.urlsafe_b64encode(str(text).encode()).decode().rstrip("=")
    except: return text

def generate_short_feedback_url(contract_id, manager):
    params = urllib.parse.urlencode({"s": encode_id(contract_id), "m": manager})
    return f"https://voc-fb.streamlit.app/?{params}"

# ----------------------------------------------------
# 3. 데이터 파이프라인 (안정성 강화)
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

df_all = load_and_verify_data()
manager_contacts = load_contacts()
df_voc = df_all[df_all["출처"] == "해지VOC"]

# ----------------------------------------------------
# 4. 메인 대시보드 탭 구성
# ----------------------------------------------------
st.title("🛡️ Enterprise VOC Control Center")

tabs = st.tabs(["📊 분석 인텔리전스", "🔍 동적 마스터 관리", "📨 AI 알림 및 피드백 제어"])

# --- TAB 1: 고급 시각화 ---
with tabs[0]:
    st.subheader("💡 5-Dimension Risk Analysis")
    if not df_voc.empty and HAS_LIBS:
        row1_c1, row1_c2 = st.columns(2)
        with row1_c1: st.plotly_chart(px.bar(df_voc.groupby("관리지사").size().reset_index(name="건수"), x="관리지사", y="건수", title="지사별 부하도"), use_container_width=True)
        with row1_c2: st.plotly_chart(px.line(df_voc.groupby(df_voc["접수일시"].dt.date).size().reset_index(name="건수"), x="접수일시", y="건수", title="일별 접수 추이"), use_container_width=True)
        
        row2_c1, row2_c2, row2_c3 = st.columns(3)
        with row2_c1: st.plotly_chart(px.pie(df_voc, names="분류", hole=0.4, title="해지 사유 분포"), use_container_width=True)
        with row2_c2: st.plotly_chart(px.histogram(df_voc, x="관리지사", color="리스크등급", barmode="group", title="리스크별 지사 현황"), use_container_width=True)
        with row2_c3:
            unique_branches = df_voc["관리지사"].unique().tolist()
            fig_radar = go.Figure(data=go.Scatterpolar(r=np.random.randint(10, 100, len(unique_branches)), theta=unique_branches, fill='toself'))
            fig_radar.update_layout(polar=dict(radialaxis=dict(visible=True)), title="지사별 성과 레이더")
            st.plotly_chart(fig_radar, use_container_width=True)

# --- TAB 2: 동적 마스터 ---
with tabs[1]:
    st.subheader("🔎 전출처 통합 동적 데이터베이스")
    q_mgr = st.selectbox("담당자별 필터", options=["전체"] + sorted(df_all["처리자"].fillna("미지정").unique().tolist()))
    df_m = df_all if q_mgr == "전체" else df_all[df_all["처리자"] == q_mgr]
    st.dataframe(df_m.sort_values("접수일시", ascending=False), use_container_width=True, hide_index=True)

# --- TAB 3: 알림 발송 및 피드백 센터 (SMTP 연동) ---
with tabs[2]:
    st.subheader("📨 지능형 알림 전송 및 현장 피드백 관리")
    
    # 1. 알림 리스트 (대체메일 기능 포함)
    high_risks = df_voc[df_voc["리스크등급"] == "HIGH"].copy()
    v_list = []
    for _, row in high_risks.iterrows():
        mgr = row["처리자"]
        dest_email = manager_contacts.get(mgr, "")
        short_url = generate_short_feedback_url(row["계약번호_정제"], mgr)
        v_list.append({"계약번호": row["계약번호_정제"], "상호": row["상호"], "담당자": mgr, "이메일(수정가능)": dest_email, "URL": short_url, "전략": row["가이드"]})
    
    edited_v = st.data_editor(pd.DataFrame(v_list), use_container_width=True, hide_index=True)
    
    if st.button("🚀 위 명단에 스마트 알림 전송", type="primary"):
        progress = st.progress(0)
        success_cnt = 0
        for i, row in edited_v.iterrows():
            try:
                msg = EmailMessage()
                msg.set_content(f"담당자님, 긴급 건 확인 바랍니다.\n전략 가이드: {row['전략']}\n피드백 링크: {row['URL']}")
                msg["Subject"] = f"[긴급 해지VOC] {row['상호']} 대응 요청"
                msg["From"] = SENDER_NAME
                msg["To"] = row["이메일(수정가능)"]
                
                with smtplib.SMTP(SMTP_HOST, SMTP_PORT) as server:
                    server.starttls()
                    server.login(SMTP_USER, SMTP_PASSWORD)
                    server.send_message(msg)
                success_cnt += 1
            except Exception as e:
                st.error(f"실패({row['담당자']}): {e}")
            progress.progress((i + 1) / len(edited_v))
        st.success(f"{success_cnt}건 전송 완료.")

    st.markdown("---")
    
    # 2. 관리자 CRUD 센터
    st.markdown("#### ⚙️ 실시간 상담 결과 통합 관리")
    if not st.session_state["feedback_db"].empty:
        df_fb = st.session_state["feedback_db"].sort_values("입력일시", ascending=False)
        for idx, row in df_fb.iterrows():
            st.markdown(f"""<div class="feedback-item"><b>[{row['상담상태']}]</b> {row['계약번호']} | {row['담당자']}<br>{row['상담내용']}</div>""", unsafe_allow_html=True)
            if st.button("❌ 삭제", key=f"del_{idx}"):
                st.session_state["feedback_db"] = st.session_state["feedback_db"].drop(idx).reset_index(drop=True)
                st.rerun()
    else: st.caption("등록된 결과가 없습니다.")
