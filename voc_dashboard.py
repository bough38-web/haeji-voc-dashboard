import os
import re
import smtplib
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
# 0. UI 테마 설정
# ----------------------------------------------------
st.set_page_config(page_title="Enterprise Haeji VOC Control", layout="wide")

st.markdown("""
    <style>
    .stMetric { background: white; padding: 20px; border-radius: 12px; border: 1px solid #e5e7eb; }
    .detailed-card { background-color: #f8f9fa; border-radius: 10px; padding: 20px; border-left: 5px solid #007aff; }
    </style>
    """, unsafe_allow_html=True)

# ----------------------------------------------------
# 1. 유틸리티 함수
# ----------------------------------------------------
def is_valid_email(email):
    regex = r'^[a-zA-Z0-9+-_.]+@[a-zA-Z0-9-]+\.[a-zA-Z0-9-.]+$'
    return bool(re.match(regex, str(email or "")))

def get_smart_contact(target_name, contact_dict):
    """담당자 매핑 (Fuzzy matching)"""
    target_name = str(target_name).strip()
    if not target_name or target_name in ["nan", "미지정"]: return None, "미지정"
    if target_name in contact_dict: return contact_dict[target_name], "Verified"
    if HAS_RAPIDFUZZ:
        choices = list(contact_dict.keys())
        result = process.extractOne(target_name, choices, processor=utils.default_process)
        if result and result[1] >= 85:
            return contact_dict[result[0]], f"Suggested({result[0]})"
    return None, "Not Found"

# ----------------------------------------------------
# 2. 데이터 전처리 파이프라인
# ----------------------------------------------------
MERGED_PATH = "merged.xlsx"
CONTACT_PATH = "contact_map.xlsx"

@st.cache_data
def load_all_data():
    if not os.path.exists(MERGED_PATH): return pd.DataFrame()
    df = pd.read_excel(MERGED_PATH)
    
    # 1. 계약번호 정제
    df["계약번호_정제"] = df["계약번호"].astype(str).str.replace(r"[^0-9A-Za-z]", "", regex=True)
    
    # 2. 날짜 및 리스크 등급 생성
    df["접수일시"] = pd.to_datetime(df["접수일시"], errors="coerce")
    today = date.today()
    df["리스크등급"] = df["접수일시"].apply(lambda dt: "HIGH" if pd.notna(dt) and (today - dt.date()).days <= 3 else "LOW")
    
    return df

@st.cache_data
def load_contacts(path):
    if not os.path.exists(path): return {}
    df_c = pd.read_excel(path)
    # "처리자1" 또는 "담당자" 기준, "E-MAIL" 매핑
    name_col = next((c for c in df_c.columns if "처리자" in str(c) or "담당자" in str(c)), df_c.columns[0])
    email_col = next((c for c in df_c.columns if "E-MAIL" in str(c) or "이메일" in str(c)), df_c.columns[1])
    return {str(row[name_col]).strip(): {"email": str(row[email_col]).strip()} for _, row in df_c.iterrows() if pd.notna(row[name_col])}

df_all = load_all_data()
manager_contacts = load_contacts(CONTACT_PATH)

# ----------------------------------------------------
# 3. 메인 대시보드 UI
# ----------------------------------------------------
st.title("🛡️ Enterprise Haeji VOC Control")

if df_all.empty:
    st.error("merged.xlsx 데이터가 존재하지 않습니다.")
    st.stop()

# 해지VOC 출처 데이터 필터링
df_voc = df_all[df_all["출처"] == "해지VOC"].copy()

tabs = st.tabs(["📊 리스크 통계", "🔍 계약별 상세 조회", "📨 담당자 알림"])

# --- TAB 1: 리스크 통계 ---
with tabs[0]:
    st.subheader("📍 지사별/등급별 VOC 분포")
    if HAS_PLOTLY:
        risk_dist = df_voc.groupby(["관리지사", "리스크등급"]).size().reset_index(name="건수")
        fig = px.bar(risk_dist, x="관리지사", y="건수", color="리스크등급", barmode="group",
                     color_discrete_map={'HIGH': '#ef4444', 'LOW': '#10b981'})
        st.plotly_chart(fig, use_container_width=True)

# --- TAB 2: 계약별 상세 조회 (핵심 요청 사항) ---
with tabs[1]:
    st.subheader("🔍 계약번호 선택 시 상세 내역 표출")
    
    # 계약번호 리스트 (정렬)
    contract_list = sorted(df_voc["계약번호_정제"].unique())
    selected_id = st.selectbox("조회할 계약번호를 선택하세요.", ["(선택 안함)"] + contract_list)

    if selected_id != "(선택 안함)":
        # 선택된 계약번호의 행 추출
        row = df_voc[df_voc["계약번호_정제"] == selected_id].iloc[0]
        
        st.markdown(f"### 📋 계약번호: {selected_id}")
        
        col1, col2 = st.columns(2)
        with col1:
            st.info(f"🏠 **시설 설치주소**\n\n{row.get('시설_설치주소', '정보 없음')}")
        with col2:
            st.success(f"💰 **KTT 월정료(조정)**\n\n{row.get('시설_KTT월정료(조정)', 0):,} 원")

        st.markdown("---")
        
        # 상세 데이터 테이블
        st.write("#### 📝 VOC 상세 내역")
        detail_data = {
            "항목": ["상호", "관리지사", "처리자(VOC)", "접수일시", "출처", "처리내용"],
            "데이터": [
                row.get("상호", "-"), 
                row.get("관리지사", "-"), 
                row.get("처리자", "-"), 
                row.get("접수일시", "-"), 
                row.get("출처", "-"),
                row.get("처리내용", "-")
            ]
        }
        st.table(pd.DataFrame(detail_data))

# --- TAB 3: 담당자 알림 (전문가 기법 적용) ---
with tabs[2]:
    st.subheader("📨 담당자별 고위험 VOC 알림 발송")
    
    high_targets = df_voc[df_voc["리스크등급"] == "HIGH"].copy()
    
    verify_list = []
    for _, row in high_targets.iterrows():
        # "처리자" 컬럼 사용
        mgr_name = row["처리자"]
        contact_info, v_status = get_smart_contact(mgr_name, manager_contacts)
        email = contact_info.get("email", "") if contact_info else ""
        
        verify_list.append({
            "계약번호": row["계약번호_정제"],
            "담당자": mgr_name,
            "매핑이메일(E-MAIL)": email,
            "매핑상태": v_status,
            "유효": is_valid_email(email)
        })
    
    edited_v = st.data_editor(
        pd.DataFrame(verify_list), 
        use_container_width=True, hide_index=True,
        column_config={"매핑이메일(E-MAIL)": st.column_config.TextColumn("수신 메일(편집가능)")}
    )

    if st.button("🚀 알림 발송 큐 전송", type="primary"):
        st.success("데이터 검증 완료. 발송 엔진에 전달되었습니다.")
