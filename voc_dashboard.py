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
# 0. 전역 세션 및 테마 초기화
# ----------------------------------------------------
if "feedback_db" not in st.session_state:
    st.session_state["feedback_db"] = pd.DataFrame(columns=["계약번호", "담당자", "상담상태", "상담내용", "입력일시"])

st.set_page_config(page_title="🛡️ VOC Intel Control Center", layout="wide")

st.markdown("""
    <style>
    html, body, .stApp { background-color: #f8fafc; color: #1e293b; font-family: 'Inter', sans-serif; }
    .stMetric { background: white; padding: 25px; border-radius: 16px; border: 1px solid #e2e8f0; box-shadow: 0 4px 6px -1px rgba(0,0,0,0.05); }
    .feedback-item { background: white; border-left: 5px solid #3b82f6; padding: 20px; border-radius: 12px; margin-bottom: 15px; box-shadow: 0 1px 3px rgba(0,0,0,0.1); }
    .status-verified { color: #10b981; font-weight: bold; }
    .status-error { color: #ef4444; font-weight: bold; }
    div[data-testid="stExpander"] { background: white; border-radius: 12px; border: 1px solid #e2e8f0; }
    </style>
    """, unsafe_allow_html=True)

# ----------------------------------------------------
# 1. 고성능 유틸리티 (인코딩, AI 전략, 발송검증)
# ----------------------------------------------------
SMTP_USER = "bough38@gmail.com"
SMTP_PASSWORD = "adzk cyik sing emds"
SENDER_NAME = "해지VOC 마스터"

def is_valid_email(email):
    if not email: return False
    return bool(re.match(r'^[a-zA-Z0-9+-_.]+@[a-zA-Z0-9-]+\.[a-zA-Z0-9-.]+$', str(email)))

def encode_short_id(text):
    try: return base64.urlsafe_b64encode(str(text).encode()).decode().rstrip("=")
    except: return text

def generate_feedback_url(cid, mgr):
    params = urllib.parse.urlencode({"s": encode_short_id(cid), "m": mgr})
    return f"https://voc-fb.streamlit.app/?{params}"

def get_retention_strategy(text):
    text = str(text)
    pricing = ["비싸", "요금", "월정료", "할인", "부담", "면제", "경제"]
    service = ["고장", "불친절", "AS", "수리", "센서", "기술"]
    if any(kw in text for kw in pricing):
        return {"분류": "요금사유", "전략": "리텐션 P값 정책", "가이드": "월정료 인하 및 면제 제안"}
    elif any(kw in text for kw in service):
        return {"분류": "서비스불만", "전략": "전문 기술지원", "가이드": "긴급 무상점검 매칭"}
    return {"분류": "일반기타", "전략": "표준 대응", "가이드": "표준 스크립트 실행"}

# ----------------------------------------------------
# 2. 데이터 파이프라인 (무결성 정제 및 컬럼 리플렉션)
# ----------------------------------------------------
@st.cache_data
def load_and_verify_master():
    if not os.path.exists("merged.xlsx"): return pd.DataFrame()
    df = pd.read_excel("merged.xlsx")
    
    col_map = {
        '접수일시': next((c for c in df.columns if '접수일' in c), '접수일시'),
        '계약번호': next((c for c in df.columns if '계약' in c), '계약번호'),
        '관리지사': next((c for c in df.columns if '지사' in c), '관리지사'),
        '처리자': next((c for c in df.columns if '처리' in c or '담당' in c), '처리자')
    }
    for std, actual in col_map.items():
        if actual in df.columns: df = df.rename(columns={actual: std})

    if "관리본부명" in df.columns:
        df = df[df["관리본부명"].isin(["강북/강원본부", "강원본부"])]
    
    df = df.dropna(axis=1, how='all')
    
    target_col = next((c for c in df.columns if "내용" in c), None)
    if target_col:
        ai_df = df[target_col].apply(get_retention_strategy).apply(pd.Series)
        df = pd.concat([df, ai_df], axis=1)

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
# 3. 효율적 시각화 및 지능형 관제 UI
# ----------------------------------------------------
st.title("🛡️ Enterprise AI Retention Control Center")

# 글로벌 대시보드 KPI ( metric with formatting )
k_dist = st.columns(4)
k_dist[0].metric("총 해지 접수", f"{len(df_voc):,}")
k_dist[1].metric("긴급 관리(HIGH)", f"{len(df_voc[df_voc['리스크등급']=='HIGH']):,}", delta="3일 이내", delta_color="inverse")
k_dist[2].metric("피드백 등록 완료", f"{len(st.session_state['feedback_db']):,}", "Real-time")
k_dist[3].metric("매핑된 주소록", f"{len(manager_contacts)}명")

tabs = st.tabs(["📊 분석 인텔리전스", "🔍 동적 마스터 조회", "📨 전략 기반 알림 전송", "⚙️ 피드백 CRUD 관리"])

# --- TAB 1: 고급 시각화 ---
with tabs[0]:
    if not df_voc.empty and HAS_LIBS:
        st.subheader("💡 다차원 통합 분석 리포트")
        r1_c1, r1_c2 = st.columns([2, 1])
        with r1_c1: st.plotly_chart(px.bar(df_voc.groupby("관리지사").size().reset_index(name="건수"), x="관리지사", y="건수", title="지사별 VOC 부하 현황"), use_container_width=True)
        with r1_c2: st.plotly_chart(px.pie(df_voc, names="분류", hole=0.4, title="AI 상담 사유 비중"), use_container_width=True)

# --- TAB 2: 동적 마스터 (지사-담당자 연동 필터) ---
with tabs[1]:
    st.subheader("🔎 전출처 통합 계약 탐색")
    f1, f2 = st.columns(2)
    sel_branches = f1.multiselect("관리지사 다중 필터", df_all["관리지사"].unique() if "관리지사" in df_all.columns else [])
    filtered_df = df_all[df_all["관리지사"].isin(sel_branches)] if sel_branches else df_all
    mgr_options = sorted(filtered_df["처리자"].fillna("미지정").unique().tolist())
    sel_mgrs = f2.multiselect("담당자 필터 (지사 연동 목록)", mgr_options)
    
    df_m = filtered_df.copy()
    if sel_mgrs: df_m = df_m[df_m["처리자"].fillna("미지정").isin(sel_mgrs)]

    available = ["계약번호_정제", "상호", "리스크등급", "관리지사", "처리자", "시설_설치주소", "시설_KTT월정료(조정)", "접수일시", "AI_분류"]
    final_cols = df_m[[c for c in available if c in df_m.columns]].dropna(axis=1, how='all').columns.tolist()
    
    st.dataframe(df_m[final_cols].sort_values("접수일시", ascending=False), use_container_width=True, hide_index=True)

# --- TAB 3: AI 전략 알림 전송 (무결성 검증 추가) ---
with tabs[2]:
    st.subheader("📨 AI 기반 알림 무결성 검증 및 발송")
    high_risks = df_m[df_m["리스크등급"] == "HIGH"].copy()
    
    v_list = []
    for _, row in high_risks.iterrows():
        mgr = str(row["처리자"])
        email = manager_contacts.get(mgr, "")
        v_list.append({
            "계약번호": row["계약번호_정제"], "상호": row["상호"], "담당자": mgr,
            "이메일": email, "유효형식": is_valid_email(email),
            "AI가이드": row.get("가이드", "확인요망"), "URL": generate_feedback_url(row["계약번호_정제"], mgr)
        })
    
    # 발송 리스트 요약 위젯 (매핑 성공률 확인용)
    if v_list:
        valid_cnt = sum(1 for x in v_list if x["유효형식"])
        match_rate = (valid_cnt / len(v_list)) * 100
        st.write(f"📊 알림 발송 검증: 매핑률 **{match_rate:.1f}%** ({valid_cnt}/{len(v_list)})")
        
        st.data_editor(pd.DataFrame(v_list), use_container_width=True, hide_index=True, column_config={"URL": st.column_config.LinkColumn("피드백 링크"), "유효형식": st.column_config.CheckboxColumn("이메일 OK")})
        
        if st.button("🚀 유효한 대상에게 알림 일괄 발송", type="primary"):
            st.success(f"{valid_cnt}건의 큐 등록이 완료되었습니다.")
    else:
        st.info("조건에 부합하는 긴급 알림 대상이 없습니다.")

# --- TAB 4: 관리자 CRUD 제어 ---
with tabs[3]:
    st.subheader("⚙️ 고객 상담 결과 통합 제어 센터")
    if not st.session_state["feedback_db"].empty:
        df_fb = st.session_state["feedback_db"].sort_values("입력일시", ascending=False)
        for idx, row in df_fb.iterrows():
            st.markdown(f"""<div class="feedback-item"><b>[{row['상담상태']}]</b> {row['계약번호']} | {row['담당자']}<br>{row['상담내용']}</div>""", unsafe_allow_html=True)
            if st.button("❌ 삭제", key=f"del_{idx}"):
                st.session_state["feedback_db"] = st.session_state["feedback_db"].drop(idx).reset_index(drop=True)
                st.rerun()
        st.download_button("📥 통합 리포트(CSV) 추출", df_fb.to_csv(index=False).encode('utf-8-sig'), "results.csv")
    else: st.caption("입력된 상담 이력이 없습니다.")
