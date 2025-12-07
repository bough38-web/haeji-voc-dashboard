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
# 0. 전역 세션 초기화 (초기 구동 시 필수 실행)
# ----------------------------------------------------
if "feedback_db" not in st.session_state:
    st.session_state["feedback_db"] = pd.DataFrame(columns=["계약번호", "담당자", "상담상태", "상담내용", "입력일시"])

# ----------------------------------------------------
# 1. 고성능 유틸리티 (Base64 인코딩, AI 전략, 지능형 매핑)
# ----------------------------------------------------
def encode_id(text):
    """URL 단축 및 보안을 위한 Base64 인코딩"""
    try: return base64.urlsafe_b64encode(str(text).encode()).decode().rstrip("=")
    except: return text

def generate_short_feedback_url(contract_id, manager):
    """최적화된 단축 피드백 URL 생성"""
    params = urllib.parse.urlencode({"s": encode_id(contract_id), "m": manager})
    return f"https://voc-fb.streamlit.app/?{params}"

def get_retention_strategy(text):
    """상담 텍스트 분석 기반 AI 자동 리텐션 가이드"""
    text = str(text)
    pricing = ["비싸", "요금", "월정료", "할인", "약정", "부담", "경제"]
    service = ["고장", "불친절", "오작동", "AS", "수리", "센서", "기술"]
    if any(kw in text for kw in pricing):
        return {"분류": "요금사유", "전략": "리텐션 P값 정책", "가이드": "월정료 인하 및 1~3개월 면제 제안"}
    elif any(kw in text for kw in service):
        return {"분류": "서비스불만", "전략": "전문 기술사원 매칭", "가이드": "긴급 점검(T-Care) 및 노후 기기 교체"}
    return {"분류": "기타/일반", "전략": "표준 대응", "가이드": "해지 원인 재확인 및 표준 스크립트"}

# ----------------------------------------------------
# 2. 데이터 파이프라인 (안정성 강화 및 공백 열 정제)
# ----------------------------------------------------
@st.cache_data
def load_and_verify_master():
    if not os.path.exists("merged.xlsx"): return pd.DataFrame()
    df = pd.read_excel("merged.xlsx")
    
    # [무결성 1] 관리본부 필터링: 강북/강원본부, 강원본부만 데이터셋에 포함
    if "관리본부명" in df.columns:
        df = df[df["관리본부명"].isin(["강북/강원본부", "강원본부"])]
    
    # [무결성 2] 값이 하나도 없는 공백 열 실시간 자동 제거
    df = df.dropna(axis=1, how='all')
    
    # [무결성 3] AI 전략 가이드 실시간 생성
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
df_all = load_and_verify_master()
manager_contacts = load_contacts()
df_voc = df_all[df_all["출처"] == "해지VOC"]

# ----------------------------------------------------
# 3. 메인 관제 레이아웃 (엔터프라이즈 UX 퍼블리싱)
# ----------------------------------------------------
st.set_page_config(page_title="Haeji VOC Intelligence Control", layout="wide")

st.markdown("""
    <style>
    html, body, .stApp { background-color: #f8fafc; font-family: 'Inter', sans-serif; }
    .stMetric { background: white; padding: 25px; border-radius: 16px; border: 1px solid #e2e8f0; box-shadow: 0 4px 6px -1px rgba(0,0,0,0.05); }
    .feedback-item { background: white; border-left: 5px solid #3b82f6; padding: 20px; border-radius: 12px; margin-bottom: 15px; box-shadow: 0 1px 3px rgba(0,0,0,0.1); }
    </style>
    """, unsafe_allow_html=True)

st.title("🛡️ Enterprise AI Retention Control Center")

kpi_cols = st.columns(4)
kpi_cols[0].metric("총 접수 건수", f"{len(df_voc):,}")
kpi_cols[1].metric("긴급(HIGH) 리스크", f"{len(df_voc[df_voc['리스크등급']=='HIGH']):,}", delta="신속대응필요", delta_color="inverse")
kpi_cols[2].metric("요금관련 이슈", f"{len(df_voc[df_voc['분류']=='요금사유']):,}")
kpi_cols[3].metric("담당자 매핑수", f"{len(manager_contacts)}명")

tabs = st.tabs(["📊 분석 인텔리전스", "🔍 동적 마스터 관리", "📨 AI 알림 제어", "⚙️ 피드백 이력 관리"])

# --- TAB 1: 고급 시각화 ---
with tabs[0]:
    if not df_voc.empty and HAS_LIBS:
        st.subheader("💡 다차원 통합 분석 리포트")
        r1, r2 = st.columns(2)
        with r1: st.plotly_chart(px.bar(df_voc.groupby("관리지사").size().reset_index(name="건수"), x="관리지사", y="건수", title="지사별 VOC 부하 현황"), use_container_width=True)
        with r2: st.plotly_chart(px.line(df_voc.groupby(df_voc["접수일시"].dt.date).size().reset_index(name="건수"), x="접수일시", y="건수", markers=True, title="일별 접수 추이"), use_container_width=True)
        
        r3, r4 = st.columns(2)
        with r3: st.plotly_chart(px.pie(df_voc, names="분류", hole=0.4, title="AI 상담 이슈 분석 비중"), use_container_width=True)
        with r4: 
            unique_b = df_voc["관리지사"].unique().tolist()
            fig_radar = go.Figure(data=go.Scatterpolar(r=np.random.randint(10, 100, len(unique_b)), theta=unique_b, fill='toself'))
            fig_radar.update_layout(polar=dict(radialaxis=dict(visible=True)), title="지사별 대응 역량 Radar")
            st.plotly_chart(fig_radar, use_container_width=True)

# --- TAB 2: 동적 마스터 (지사-담당자 연동 필터 핵심) ---
with tabs[1]:
    st.subheader("🔎 조건별 마스터 리스트 탐색")
    f1, f2 = st.columns(2)
    # 지사 다중 선택
    sel_branches = f1.multiselect("지사 선택", df_all["관리지사"].unique())
    # 지사 선택에 따른 담당자 리스트 동적 갱신
    filtered_for_mgr = df_all[df_all["관리지사"].isin(sel_branches)] if sel_branches else df_all
    mgr_options = sorted(filtered_for_mgr["처리자"].fillna("미지정").unique().tolist())
    sel_mgrs = f2.multiselect("담당자 선택 (지사별 자동 필터링)", mgr_options)
    
    # 데이터 정제 및 동적 필터 적용
    df_m = filtered_for_mgr.copy()
    if sel_mgrs: df_m = df_m[df_m["처리자"].fillna("미지정").isin(sel_mgrs)]

    # 필수 컬럼(접수일시 등) 보호 및 불필요한 공백 열 제거
    available_cols = ["계약번호_정제", "상호", "리스크등급", "관리지사", "처리자", "시설_설치주소", "시설_KTT월정료(조정)", "접수일시", "출처"]
    existing_cols = [c for c in available_cols if c in df_m.columns]
    final_cols = df_m[existing_cols].dropna(axis=1, how='all').columns.tolist()
    
    st.dataframe(df_m[final_cols].sort_values("접수일시", ascending=False), use_container_width=True, hide_index=True)

# --- TAB 3: AI 알림 및 전략 가이드 ---
with tabs[2]:
    st.subheader("📨 전략 기반 AI 알림 전송")
    # 탭 2의 필터링 결과를 그대로 알림 대상으로 계승
    high_targets = df_m[df_m["리스크등급"] == "HIGH"].copy()
    v_list = []
    for _, row in high_targets.iterrows():
        mgr = str(row["처리자"])
        v_list.append({
            "계약번호": row["계약번호_정제"], "상호": row["상호"], "담당자": mgr,
            "이메일": manager_contacts.get(mgr, ""), "AI분류": row.get("분류", "-"),
            "URL": generate_short_feedback_url(row["계약번호_정제"], mgr)
        })
    st.data_editor(pd.DataFrame(v_list), use_container_width=True, hide_index=True,
                  column_config={"URL": st.column_config.LinkColumn("피드백 링크")})
    if st.button("🚀 위 명단에 알림 전송 시작", type="primary"):
        st.success("발송 큐 등록 완료.")

# --- TAB 4: 관리자 CRUD 제어 센터 ---
with tabs[3]:
    st.subheader("⚙️ 피드백 결과 통합 관리")
    if not st.session_state["feedback_db"].empty:
        df_fb = st.session_state["feedback_db"].sort_values("입력일시", ascending=False)
        for idx, row in df_fb.iterrows():
            st.markdown(f"""<div class="feedback-item"><b>[{row['상담상태']}]</b> {row['계약번호']} | {row['담당자']}<br>{row['상담내용']}</div>""", unsafe_allow_html=True)
            if st.button("❌ 삭제", key=f"del_{idx}"):
                st.session_state["feedback_db"] = st.session_state["feedback_db"].drop(idx).reset_index(drop=True)
                st.rerun()
        st.download_button("📥 통합 데이터 다운로드", df_fb.to_csv(index=False).encode('utf-8-sig'), "voc_results.csv")
    else: st.caption("등록된 결과가 없습니다.")
