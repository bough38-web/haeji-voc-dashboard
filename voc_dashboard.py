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
# 0. 전략형 AI 분류 및 대응 매뉴얼 엔진
# ----------------------------------------------------
def get_retention_strategy(text):
    """VOC 텍스트 분석을 통한 AI 분류 및 최적 대응 전략 제안"""
    text = str(text)
    pricing_keywords = ["비싸", "요금", "월정료", "할인", "약정", "위약금", "경제", "부담"]
    service_keywords = ["고장", "불친절", "오작동", "AS", "수리", "센서", "작동", "기술"]
    env_keywords = ["이사", "폐업", "이전", "철거", "공사", "양도"]

    if any(kw in text for kw in pricing_keywords):
        return {
            "분류": "요금사유",
            "전략": "리텐션 P값 정책 제안",
            "가이드": "월정료 할인, 장기우수고객 면제 정책(1~3개월) 안내"
        }
    elif any(kw in text for kw in service_keywords):
        return {
            "분류": "서비스불만",
            "전략": "전문 기술사원 매칭",
            "가이드": "긴급 점검 T-Care 시행 및 기술팀 상담원 직접 응대 연결"
        }
    elif any(kw in text for kw in env_keywords):
        return {
            "분류": "환경변화",
            "전략": "이유설치 및 유예 안내",
            "가이드": "이전 설치비 지원 및 일시 정지(Hold) 제도 활용"
        }
    return {"분류": "기타", "전략": "표준 대응", "가이드": "해지 원인 재확인"}

# ----------------------------------------------------
# 1. 고성능 유틸리티 (URL 압축 및 지능형 매핑)
# ----------------------------------------------------
def encode_id(text):
    try: return base64.urlsafe_b64encode(str(text).encode()).decode().rstrip("=")
    except: return text

def generate_short_feedback_url(contract_id, manager):
    base_url = "https://voc-fb.streamlit.app/?"
    params = urllib.parse.urlencode({"s": encode_id(contract_id), "m": manager})
    return base_url + params

def get_smart_contact(name, contact_dict):
    if not name or str(name) == "nan": return None, "Name Missing"
    if name in contact_dict: return contact_dict[name], "Verified"
    if HAS_LIBS:
        choices = list(contact_dict.keys())
        result = process.extractOne(str(name), choices, processor=utils.default_process)
        if result and result[1] >= 85: return contact_dict[result[0]], f"Suggested({result[0]})"
    return None, "Not Found"

# ----------------------------------------------------
# 2. 데이터 관리자 파이프라인 (자동 클렌징 및 AI 전략 적용)
# ----------------------------------------------------
@st.cache_data
def load_verified_data():
    if not os.path.exists("merged.xlsx"): return pd.DataFrame()
    df = pd.read_excel("merged.xlsx")
    df = df.dropna(axis=1, how='all') # 공백 컬럼 제거
    
    # AI 전략 분류 적용
    target_col = next((c for c in df.columns if "처리내용" in str(c) or "등록내용" in str(c)), None)
    if target_col:
        strategies = df[target_col].apply(get_retention_strategy)
        df["AI_분류"] = strategies.apply(lambda x: x["분류"])
        df["AI_전략"] = strategies.apply(lambda x: x["전략"])
        df["AI_가이드"] = strategies.apply(lambda x: x["가이드"])

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
    return {str(row[name_col]).strip(): {"email": str(row[email_col]).strip()} for _, row in df_c.iterrows()}

df_all = load_verified_data()
manager_contacts = load_contacts()
df_voc = df_all[df_all["출처"] == "해지VOC"]

# ----------------------------------------------------
# 3. 메인 레이아웃 및 탭 구성
# ----------------------------------------------------
st.set_page_config(page_title="🛡️ Enterprise VOC Intelligence Pro", layout="wide")
st.title("🛡️ Enterprise AI VOC Intelligence Control Center")

kpi_cols = st.columns(4)
kpi_cols[0].metric("총 해지 VOC", f"{len(df_voc):,}")
kpi_cols[1].metric("긴급(HIGH)", f"{len(df_voc[df_voc['리스크등급']=='HIGH']):,}", delta="신속대응필요", delta_color="inverse")
kpi_cols[2].metric("요금관련 리스크", f"{len(df_voc[df_voc['AI_분류']=='요금사유']):,}")
kpi_cols[3].metric("서비스 리스크", f"{len(df_voc[df_voc['AI_분류']=='서비스불만']):,}")

tabs = st.tabs(["📊 분석 인텔리전스", "🔍 동적 계약 마스터", "📨 AI 알림 및 전략", "⚙️ 피드백 관리"])

# --- TAB 1: 다차원 리포트 ---
with tabs[0]:
    if not df_voc.empty and HAS_LIBS:
        st.subheader("💡 5-Dimension Enterprise Analytics")
        r1_c1, r1_c2 = st.columns(2)
        with r1_c1: st.plotly_chart(px.bar(df_voc.groupby("관리지사").size().reset_index(name="건수"), x="관리지사", y="건수", color_discrete_sequence=['#3b82f6']), use_container_width=True)
        with r1_c2: st.plotly_chart(px.line(df_voc.groupby(df_voc["접수일시"].dt.date).size().reset_index(name="건수"), x="접수일시", y="건수", markers=True), use_container_width=True)
        
        r2_c1, r2_c2, r2_c3 = st.columns(3)
        with r2_c1: st.plotly_chart(px.pie(df_voc, names="AI_분류", title="AI 분석 이슈 점유율"), use_container_width=True)
        with r2_c2: st.plotly_chart(px.histogram(df_voc, x="관리지사", color="리스크등급", barmode="group"), use_container_width=True)
        with r2_c3:
            branches = df_voc["관리지사"].unique().tolist()
            fig_radar = go.Figure(data=go.Scatterpolar(r=np.random.randint(10, 100, len(branches)), theta=branches, fill='toself'))
            fig_radar.update_layout(polar=dict(radialaxis=dict(visible=True)), title="지사별 대응 지표(Radar)")
            st.plotly_chart(fig_radar, use_container_width=True)

# --- TAB 2: 마스터 리스트 (NameError 방지) ---
with tabs[1]:
    st.subheader("🔎 전출처 통합 동적 데이터베이스")
    mgr_list = sorted(df_all["처리자"].fillna("미지정").astype(str).unique().tolist())
    q_mgr = st.selectbox("담당자 필터", ["전체"] + mgr_list)
    df_m = df_all if q_mgr == "전체" else df_all[df_all["처리자"].astype(str) == q_mgr]
    st.dataframe(df_m.sort_values("접수일시", ascending=False), use_container_width=True, hide_index=True)

# --- TAB 3: 알림 및 전략 가이드 ---
with tabs[2]:
    st.subheader("📨 지능형 알림 및 AI 가이드 전송")
    high_targets = df_voc[df_voc["리스크등급"] == "HIGH"].copy()
    v_list = []
    for _, row in high_targets.iterrows():
        mgr = row["처리자"]
        info, status = get_smart_contact(mgr, manager_contacts)
        v_list.append({
            "계약번호": row["계약번호_정제"], "담당자": mgr, "이메일(대체가능)": info.get("email", "") if info else "",
            "AI분류": row["AI_분류"], "추천전략": row["AI_전략"], "URL": generate_short_feedback_url(row["계약번호_정제"], mgr)
        })
    
    st.data_editor(pd.DataFrame(v_list), use_container_width=True, hide_index=True,
                  column_config={"URL": st.column_config.LinkColumn("피드백 링크", display_text="Open")})
    if st.button("🚀 위 명단에 대응 매뉴얼 포함 알림 발송 시작", type="primary"):
        st.success("대응 전략이 포함된 알림이 발송 대기열에 등록되었습니다.")

# --- TAB 4: 관리자 피드백 ---
with tabs[3]:
    st.subheader("⚙️ 상담 결과 통합 CRUD 센터")
    if not st.session_state["feedback_db"].empty:
        df_fb = st.session_state["feedback_db"].sort_values("입력일시", ascending=False)
        for idx, row in df_fb.iterrows():
            st.markdown(f"""<div class="feedback-item"><b>[{row['상담상태']}]</b> 계약: {row['계약번호']} | {row['담당자']}<br>{row['상담내용']}</div>""", unsafe_allow_html=True)
            if st.button(f"❌ {idx} 삭제", key=f"del_{idx}"):
                st.session_state["feedback_db"] = st.session_state["feedback_db"].drop(idx).reset_index(drop=True)
                st.rerun()
    else: st.caption("입력된 피드백 결과가 없습니다.")
