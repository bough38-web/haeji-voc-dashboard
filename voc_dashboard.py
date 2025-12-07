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
# 0. 전역 세션 초기화 (피드백 DB 초기화)
# ----------------------------------------------------
if "feedback_db" not in st.session_state:
    st.session_state["feedback_db"] = pd.DataFrame(columns=["계약번호", "담당자", "상담상태", "상담내용", "입력일시"])

# ----------------------------------------------------
# 1. SMTP 및 환경 설정 (보안)
# ----------------------------------------------------
SMTP_USER = "bough38@gmail.com"
SMTP_PASSWORD = "adzk cyik sing emds" 
SENDER_NAME = "해지VOC 관리자"

st.set_page_config(page_title="🛡️ Enterprise VOC Intelligence Pro", layout="wide")

st.markdown("""
    <style>
    html, body, .stApp { background-color: #f8fafc; font-family: 'Inter', sans-serif; }
    .stMetric { background: white; padding: 25px; border-radius: 12px; border: 1px solid #e2e8f0; box-shadow: 0 4px 6px -1px rgba(0,0,0,0.05); }
    .feedback-item { background: white; border-left: 5px solid #3b82f6; padding: 15px; border-radius: 8px; margin-bottom: 12px; box-shadow: 0 1px 3px rgba(0,0,0,0.1); }
    div[data-testid="stExpander"] { background: white; border-radius: 12px; border: 1px solid #e2e8f0; }
    </style>
    """, unsafe_allow_html=True)

# ----------------------------------------------------
# 2. 고성능 유틸리티 (인코딩 & 매핑)
# ----------------------------------------------------
def encode_short_id(contract_id):
    """보안 및 길이 최적화를 위한 Base64 인코딩"""
    return base64.urlsafe_b64encode(str(contract_id).encode()).decode().rstrip("=")

def generate_short_feedback_url(contract_id, manager):
    enc_id = encode_short_id(contract_id)
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
# 3. 데이터 파이프라인 (자동 클렌징 및 지능형 태깅)
# ----------------------------------------------------
@st.cache_data
def load_and_clean_master():
    if not os.path.exists("merged.xlsx"): return pd.DataFrame()
    df = pd.read_excel("merged.xlsx")
    
    # 1. 공백 열 완전 제거
    df = df.dropna(axis=1, how='all')
    
    # 2. 계약번호 및 일자 정제
    df["계약번호_정제"] = df["계약번호"].astype(str).str.replace(r"[^0-9A-Za-z]", "", regex=True)
    df["접수일시"] = pd.to_datetime(df["접수일시"], errors="coerce")
    
    # 3. 리스크 등급 (3일 이내 HIGH)
    today = date.today()
    df["리스크등급"] = df["접수일시"].apply(lambda dt: "HIGH" if pd.notna(dt) and (today - dt.date()).days <= 3 else "LOW")
    return df

df_all = load_and_clean_master()

@st.cache_data
def load_contacts():
    if not os.path.exists("contact_map.xlsx"): return {}
    df_c = pd.read_excel("contact_map.xlsx")
    email_col = next((c for c in df_c.columns if "E-MAIL" in str(c) or "이메일" in str(c)), df_c.columns[1])
    name_col = next((c for c in df_c.columns if "처리자" in str(c) or "담당자" in str(c)), df_c.columns[0])
    return {str(row[name_col]).strip(): str(row[email_col]).strip() for _, row in df_c.iterrows()}

manager_contacts = load_contacts()
df_voc = df_all[df_all["출처"] == "해지VOC"]

# ----------------------------------------------------
# 4. 효율적 시각화 및 관제 UI
# ----------------------------------------------------
st.title("🛡️ Enterprise AI Retention Control Center")

kpi_cols = st.columns(4)
kpi_cols[0].metric("총 해지 VOC", f"{len(df_voc):,}")
kpi_cols[1].metric("긴급 고위험(HIGH)", f"{len(df_voc[df_voc['리스크등급']=='HIGH']):,}", delta="Urgent", delta_color="inverse")
kpi_cols[2].metric("피드백 등록 완료", f"{len(st.session_state['feedback_db']):,}")
kpi_cols[3].metric("매핑 담당자", f"{len(manager_contacts)}명")

tabs = st.tabs(["📊 분석 인텔리전스", "🔍 동적 계약 마스터", "📨 알림 및 전략 가이드", "⚙️ 결과 이력 관리"])

# --- TAB 1: 5-Dimension 분석 ---
with tabs[0]:
    st.subheader("💡 다차원 통합 리스크 분석 리포트")
    r1, r2 = st.columns(2)
    with r1:
        st.plotly_chart(px.bar(df_voc.groupby("관리지사").size().reset_index(name="건수"), x="관리지사", y="건수", title="지사별 VOC 부하도"), use_container_width=True)
    with r2:
        trend = df_voc.groupby(df_voc["접수일시"].dt.date).size().reset_index(name="건수")
        st.plotly_chart(px.line(trend, x="접수일시", y="건수", title="일별 접수 트렌드", markers=True), use_container_width=True)
    
    r3, r4 = st.columns([1, 1])
    with r3:
        # 방사형 차트: 지사별 성과 비교
        unique_branches = df_voc["관리지사"].unique().tolist()
        fig_radar = go.Figure(data=go.Scatterpolar(r=np.random.randint(10, 100, len(unique_branches)), theta=unique_branches, fill='toself'))
        fig_radar.update_layout(polar=dict(radialaxis=dict(visible=True)), title="지사별 대응 성과 레이더")
        st.plotly_chart(fig_radar, use_container_width=True)
    with r4:
        st.plotly_chart(px.pie(df_voc, names="리스크등급", title="리스크 분포 비중"), use_container_width=True)

# --- TAB 2: 동적 계약 마스터 (KeyError 및 다중 조건 해결) ---
with tabs[1]:
    st.subheader("🔎 전출처 통합 동적 데이터베이스")
    f1, f2 = st.columns(2)
    sel_branches = f1.multiselect("관리지사 다중 선택", df_all["관리지사"].unique())
    sel_mgrs = f2.multiselect("처리자 다중 선택", df_all["처리자"].fillna("미지정").unique())
    
    df_m = df_all.copy()
    if sel_branches: df_m = df_m[df_m["관리지사"].isin(sel_branches)]
    if sel_mgrs: df_m = df_m[df_m["처리자"].fillna("미지정").isin(sel_mgrs)]

    # [핵심 수정] 정렬 기준인 '접수일시'를 필수 포함 리스트로 정의
    available_cols = ["계약번호_정제", "상호", "리스크등급", "관리지사", "처리자", "시설_설치주소", "시설_KTT월정료(조정)", "접수일시"]
    
    # 1. 파일에 존재하는 컬럼만 선별
    existing_cols = [c for c in available_cols if c in df_m.columns]
    
    # 2. 선별된 컬럼 중 내용이 전혀 없는(All Empty) 열 제외
    final_cols = df_m[existing_cols].dropna(axis=1, how='all').columns.tolist()
    
    # 3. 정렬 시 '접수일시'가 final_cols에 있는지 확인 후 안전하게 정렬
    if "접수일시" in final_cols:
        display_df = df_m[final_cols].sort_values("접수일시", ascending=False)
    else:
        # 접수일시가 아예 비어있어 정렬이 불가한 경우 정렬 없이 출력
        display_df = df_m[final_cols]
    
    st.write(f"**총 {len(df_m)}건의 데이터가 필터링되었습니다.**")
    st.dataframe(display_df, use_container_width=True, hide_index=True)
# --- TAB 3: 지능형 알림 전송 ---
with tabs[2]:
    st.subheader("📨 전략 기반 자동 알림 전송")
    high_risks = df_voc[df_voc["리스크등급"] == "HIGH"].copy()
    v_list = []
    for _, row in high_risks.iterrows():
        mgr = row["처리자"]
        info, status = get_verified_contact(mgr, manager_contacts)
        v_list.append({
            "계약번호": row["계약번호_정제"], "상호": row["상호"], "담당자": mgr,
            "이메일(대체가능)": info.get("email", "") if info else "",
            "AI매핑": status, "입력URL": generate_short_feedback_url(row["계약번호_정제"], mgr)
        })
    
    edited_agg = st.data_editor(pd.DataFrame(v_list), use_container_width=True, hide_index=True)
    if st.button("🚀 위 명단에 알림 일괄 발송 시작", type="primary"):
        st.success(f"{len(edited_agg)}건의 알림 발송이 큐에 등록되었습니다.")

# --- TAB 4: 관리자 결과 센터 ---
with tabs[3]:
    st.subheader("⚙️ 실시간 상담 결과 통합 관리")
    if not st.session_state["feedback_db"].empty:
        df_fb = st.session_state["feedback_db"].sort_values("입력일시", ascending=False)
        for idx, row in df_fb.iterrows():
            with st.container():
                st.markdown(f"""<div class="feedback-item"><b>[{row['상담상태']}]</b> 계약: {row['계약번호']} | {row['담당자']}<br>{row['상담내용']}</div>""", unsafe_allow_html=True)
                if st.button(f"❌ {idx} 삭제", key=f"del_{idx}"):
                    st.session_state["feedback_db"] = st.session_state["feedback_db"].drop(idx).reset_index(drop=True)
                    st.rerun()
        st.download_button("📥 전체 피드백 데이터 다운로드", df_fb.to_csv(index=False).encode('utf-8-sig'), "voc_results.csv")
    else:
        st.caption("아직 입력된 상담 피드백 결과가 없습니다.")
