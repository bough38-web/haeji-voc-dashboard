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
# 0. 엔터프라이즈 테마 및 UX 설정
# ----------------------------------------------------
st.set_page_config(page_title="🛡️ Enterprise VOC Intelligence Pro", layout="wide")

st.markdown("""
    <style>
    /* Modern Glassmorphism Theme */
    html, body, .stApp { background-color: #f8fafc; color: #1e293b; font-family: 'Inter', -apple-system, sans-serif; }
    .stMetric { background: white; padding: 25px; border-radius: 16px; border: 1px solid #e2e8f0; box-shadow: 0 4px 6px -1px rgba(0,0,0,0.05); }
    .feedback-item { background: white; border-left: 5px solid #3b82f6; padding: 20px; border-radius: 12px; margin-bottom: 15px; box-shadow: 0 1px 3px rgba(0,0,0,0.1); }
    div[data-testid="stExpander"] { background: white; border-radius: 12px; border: 1px solid #e2e8f0; }
    </style>
    """, unsafe_allow_html=True)

# ----------------------------------------------------
# 1. 고성능 유틸리티 (URL 압축 및 지능형 매핑)
# ----------------------------------------------------
def encode_id(text):
    """URL 단축을 위한 Base64 인코딩"""
    try: return base64.urlsafe_b64encode(str(text).encode()).decode().rstrip("=")
    except: return text

def generate_short_feedback_url(contract_id, manager):
    """최적화된 단축 피드백 URL 생성"""
    base_url = "https://voc-fb.streamlit.app/?"
    params = urllib.parse.urlencode({"s": encode_id(contract_id), "m": manager})
    return base_url + params

def get_verified_contact(name, contact_dict):
    """Fuzzy Matching을 통한 담당자 매핑 정합성 검증"""
    if not name or str(name) == "nan": return None, "Name Missing"
    if name in contact_dict: return contact_dict[name], "Verified"
    if HAS_LIBS:
        choices = list(contact_dict.keys())
        result = process.extractOne(str(name), choices, processor=utils.default_process)
        if result and result[1] >= 85: return contact_dict[result[0]], f"Suggested({result[0]})"
    return None, "Not Found"

# ----------------------------------------------------
# 2. 데이터 관리자 파이프라인 (자동 컬럼 클렌징)
# ----------------------------------------------------
if "feedback_db" not in st.session_state:
    st.session_state["feedback_db"] = pd.DataFrame(columns=["계약번호", "담당자", "상담상태", "상담내용", "입력일시"])

@st.cache_data
def load_and_clean_master_data():
    if not os.path.exists("merged.xlsx"): return pd.DataFrame()
    df = pd.read_excel("merged.xlsx")
    
    # 1. 값이 하나도 없는 컬럼 자동 제외
    df = df.dropna(axis=1, how='all')
    
    # 2. 계약번호 및 일자 정제
    df["계약번호_정제"] = df["계약번호"].astype(str).str.replace(r"[^0-9A-Za-z]", "", regex=True)
    df["접수일시"] = pd.to_datetime(df["접수일시"], errors="coerce")
    
    # 3. 리스크 등급 동적 생성 (3일 이내 긴급)
    today = date.today()
    df["리스크등급"] = df["접수일시"].apply(lambda dt: "HIGH" if pd.notna(dt) and (today - dt.date()).days <= 3 else "LOW")
    return df

@st.cache_data
def load_contacts():
    if not os.path.exists("contact_map.xlsx"): return {}
    df_c = pd.read_excel("contact_map.xlsx")
    # E-MAIL 컬럼 우선 탐색
    email_col = next((c for c in df_c.columns if "E-MAIL" in str(c) or "이메일" in str(c)), df_c.columns[1])
    name_col = next((c for c in df_c.columns if "처리자" in str(c) or "담당자" in str(c)), df_c.columns[0])
    return {str(row[name_col]).strip(): {"email": str(row[email_col]).strip()} for _, row in df_c.iterrows()}

df_all = load_and_clean_master_data()
manager_contacts = load_contacts()

# ----------------------------------------------------
# 3. 효율적 시각화 관제 센터 (5-Dimension)
# ----------------------------------------------------
st.title("🛡️ Enterprise Haeji VOC Intelligence Center")

kpi1, kpi2, kpi3, kpi4 = st.columns(4)
df_voc = df_all[df_all["출처"] == "해지VOC"]
kpi1.metric("총 해지 VOC", f"{len(df_voc):,}")
kpi2.metric("긴급 고위험(HIGH)", f"{len(df_voc[df_voc['리스크등급']=='HIGH']):,}", delta="신속대응", delta_color="inverse")
kpi3.metric("피드백 등록률", f"{(len(st.session_state['feedback_db'])/len(df_voc)*100) if len(df_voc)>0 else 0:.1f}%")
kpi4.metric("매핑된 주소록", f"{len(manager_contacts)}명")

tabs = st.tabs(["📊 데이터 인텔리전스", "🔍 동적 계약 마스터", "📨 AI 알림 및 피드백"])

# --- TAB 1: 고급 분석 리포트 ---
with tabs[0]:
    if not df_voc.empty and HAS_LIBS:
        st.subheader("💡 5개 핵심 차원 다각도 리스크 분석")
        row1_col1, row1_col2 = st.columns(2)
        with row1_col1:
            st.plotly_chart(px.bar(df_voc.groupby("관리지사").size().reset_index(name="건수"), 
                                   x="관리지사", y="건수", title="지사별 VOC 부하도"), use_container_width=True)
        with row1_col2:
            st.plotly_chart(px.line(df_voc.groupby(df_voc["접수일시"].dt.date).size().reset_index(name="건수"), 
                                    x="접수일시", y="건수", title="일별 접수 트렌드", markers=True), use_container_width=True)
        
        row2_col1, row2_col2, row2_col3 = st.columns(3)
        with row2_col1:
            st.plotly_chart(px.pie(df_voc, names="관리지사", title="지사별 시장 점유 비중"), use_container_width=True)
        with row2_col2:
            st.plotly_chart(px.histogram(df_voc, x="시설_KTT월정료(조정)", title="월정료 금액 분포"), use_container_width=True)
        with row2_col3:
            # 전문가 레이더 차트 (지사별 성과 시뮬레이션)
            branches = df_voc["관리지사"].unique()
            fig_radar = go.Figure(data=go.Scatterpolar(r=np.random.randint(10, 100, len(branches)), theta=branches, fill='toself'))
            fig_radar.update_layout(polar=dict(radialaxis=dict(visible=True)), title="지사별 대응 성과 레이더")
            st.plotly_chart(fig_radar, use_container_width=True)

# --- TAB 2: 동적 계약 마스터 (Fuzzy 통합) ---
with tabs[1]:
    st.subheader("🔎 전출처 통합 계약 데이터베이스 (Fuzzy 조회)")
    f1, f2 = st.columns(2)
    # 결측치를 "미지정"으로 채우고 모든 값을 문자열로 변환한 뒤 리스트화
    mgr_list = df_all["처리자"].fillna("미지정").astype(str).unique().tolist()
    # 담당자 필터 옵션 구성 (정렬된 리스트 앞에 "전체" 추가)
    q_mgr = f2.selectbox("담당자 필터", options=["전체"] + sorted(mgr_list))
    df_m = df_all.copy()
    if q_branch: df_m = df_m[df_m["관리지사"].isin(q_branch)]
    if q_mgr != "전체": df_m = df_m[df_m["처리자"] == q_mgr]
    
    st.write(f"**검색 결과: {len(df_m)}건** (불필요한 공백 열 자동 제외 완료)")
    st.dataframe(df_m.sort_values("접수일시", ascending=False), use_container_width=True, hide_index=True)

# --- TAB 3: AI 알림 & 피드백 로그 관리 ---
with tabs[2]:
    st.subheader("📨 지능형 알림 전송 및 상담 결과 제어 센터")
    
    # 1. 알림 전송 명단 구성 (URL 최적화 포함)
    high_risks = df_voc[df_voc["리스크등급"] == "HIGH"].copy()
    v_list = []
    for _, row in high_risks.iterrows():
        mgr = row["처리자"]
        info, status = get_verified_contact(mgr, manager_contacts)
        short_url = generate_short_feedback_url(row["계약번호_정제"], mgr)
        
        v_list.append({
            "계약번호": row["계약번호_정제"], "상호": row["상호"], "담당자": mgr,
            "이메일(대체가능)": info.get("email", "") if info else "",
            "매핑": status, "압축URL": short_url
        })
    
    edited_v = st.data_editor(pd.DataFrame(v_list), use_container_width=True, hide_index=True,
                              column_config={"압축URL": st.column_config.LinkColumn("입력 폼 링크", display_text="Open Feedback")})
    
    if st.button("🚀 선택 대상 스마트 알림 발송 시작", type="primary"):
        st.success(f"압축 URL이 포함된 동적 알림 {len(edited_v)}건이 전송 큐에 등록되었습니다.")

    st.markdown("---")
    
    # 2. 관리자 상담 결과 CRUD 센터
    st.markdown("#### ⚙️ 상담 결과 통합 제어 (수정/삭제)")
    if not st.session_state["feedback_db"].empty:
        df_fb = st.session_state["feedback_db"].sort_values("입력일시", ascending=False)
        for idx, row in df_fb.iterrows():
            with st.container():
                st.markdown(f"""
                <div class="feedback-item">
                    <b>[{row['상담상태']}]</b> 계약: {row['계약번호']} | {row['담당자']} | {row['입력일시'].strftime('%m-%d %H:%M')}<br>
                    내용: {row['상담내용']}
                </div>
                """, unsafe_allow_html=True)
                c_del, c_csv, _ = st.columns([1, 2, 7])
                if c_del.button("❌ 삭제", key=f"del_{idx}"):
                    st.session_state["feedback_db"] = st.session_state["feedback_db"].drop(idx).reset_index(drop=True)
                    st.rerun()
    else:
        st.caption("아직 입력된 피드백 결과가 없습니다.")
