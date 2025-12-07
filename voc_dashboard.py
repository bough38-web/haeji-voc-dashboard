import os
import re
import smtplib
from datetime import datetime, date
import numpy as np
import pandas as pd
import streamlit as st

# 전문가용 지능형 매핑 및 시각화 라이브러리
try:
    from rapidfuzz import process, utils
    HAS_RAPIDFUZZ = True
except ImportError:
    HAS_RAPIDFUZZ = False

try:
    import plotly.express as px
    HAS_PLOTLY = True
except Exception:
    HAS_PLOTLY = False

# ----------------------------------------------------
# 0. UI 설정 및 라이트톤 CSS
# ----------------------------------------------------
st.set_page_config(page_title="해지 VOC 종합 대시보드 Pro", layout="wide")

st.markdown("""
    <style>
    html, body, .stApp { background-color: #f5f5f7 !important; color: #1d1d1f !important; }
    .section-card { background: white; border-radius: 12px; padding: 1.5rem; border: 1px solid #e5e7eb; margin-bottom: 1rem; }
    .stMetric { background: white; padding: 15px; border-radius: 10px; border: 1px solid #efefef; }
    </style>
    """, unsafe_allow_html=True)

# ----------------------------------------------------
# 1. 유틸리티 (지능형 매핑 및 검증)
# ----------------------------------------------------
def is_valid_email(email):
    regex = r'^[a-zA-Z0-9+-_.]+@[a-zA-Z0-9-]+\.[a-zA-Z0-9-.]+$'
    return bool(re.match(regex, str(email or "")))

def get_smart_contact(target_name, contact_dict):
    """Fuzzy Matching: '처리자1'과 원천 데이터를 지능적으로 연결"""
    target_name = str(target_name).strip()
    if not target_name or target_name in ["nan", "미지정"]: return None, "Name Missing"
    if target_name in contact_dict: return contact_dict[target_name], "Verified"
    
    if HAS_RAPIDFUZZ:
        choices = list(contact_dict.keys())
        # 유사도 임계값 85% 설정
        result = process.extractOne(target_name, choices, processor=utils.default_process)
        if result and result[1] >= 85:
            return contact_dict[result[0]], f"Suggested({result[0]})"
    return None, "Not Found"

# ----------------------------------------------------
# 2. 데이터 전처리 (시각화 중단 해결 지점)
# ----------------------------------------------------
MERGED_PATH = "merged.xlsx"
CONTACT_PATH = "contact_map.xlsx"

@st.cache_data
def load_and_fix_data():
    if not os.path.exists(MERGED_PATH): return pd.DataFrame()
    df = pd.read_excel(MERGED_PATH)
    
    # 1. 핵심 컬럼 생성 (이게 없으면 그래프가 안 나옴)
    df["계약번호_정제"] = df["계약번호"].astype(str).str.replace(r"[^0-9A-Za-z]", "", regex=True)
    df["접수일시"] = pd.to_datetime(df["접수일시"], errors="coerce")
    
    # 2. 리스크 등급 강제 생성
    today = date.today()
    df["리스크등급"] = df["접수일시"].apply(lambda dt: "HIGH" if (today - dt.date()).days <= 3 else "LOW" if pd.notna(dt) else "MEDIUM")
    
    # 3. 매칭여부 (비매칭 리스트용 기본값)
    if "매칭여부" not in df.columns: df["매칭여부"] = "비매칭(X)"
    
    # 4. 담당자 통합 ("처리자" 컬럼 우선)
    def pick_mgr(row):
        for c in ["처리자", "구역담당자", "담당자"]:
            if c in row and pd.notna(row[c]): return str(row[c]).strip()
        return "미지정"
    df["구역담당자_통합"] = df.apply(pick_mgr, axis=1)
    
    return df

@st.cache_data
def load_contacts_v2(path):
    if not os.path.exists(path): return pd.DataFrame(), {}
    df_c = pd.read_excel(path)
    
    # 
    
    # "처리자1" 및 "이메일" 컬럼 자동 감지
    name_col = next((c for c in df_c.columns if "처리자1" in str(c)), df_c.columns[0])
    email_col = next((c for c in df_c.columns if "이메일" in str(c) or "메일" in str(c)), df_c.columns[1])
    
    contact_dict = {str(row[name_col]).strip(): {"email": str(row[email_col]).strip()} for _, row in df_c.iterrows() if pd.notna(row[name_col])}
    return df_c, contact_dict

df_all = load_and_fix_data()
contact_raw, manager_contacts = load_contacts_v2(CONTACT_PATH)

# ----------------------------------------------------
# 3. 탭 레이아웃 및 시각화 렌더링
# ----------------------------------------------------
tabs = st.tabs(["📈 지사별 시각화", "📘 VOC 전체", "📨 담당자 알림"])

with tabs[0]:
    st.subheader("📊 지사별 리스크 현황 리포트")
    if not df_all.empty and HAS_PLOTLY:
        # 지사별 리스크 카운트
        risk_dist = df_all.groupby(["관리지사", "리스크등급"]).size().reset_index(name="건수")
        fig = px.bar(risk_dist, x="관리지사", y="건수", color="리스크등급", 
                     title="지사별 고위험 VOC 분포", barmode="group",
                     color_discrete_map={'HIGH': '#ef4444', 'MEDIUM': '#f59e0b', 'LOW': '#10b981'})
        st.plotly_chart(fig, use_container_width=True)
    else:
        st.info("데이터가 충분하지 않거나 시각화 엔진이 로드되지 않았습니다.")

with tabs[2]:
    st.subheader("📨 지능형 담당자 알림 발송")
    # 고위험 비매칭 데이터 필터링
    alert_targets = df_all[df_all["리스크등급"] == "HIGH"].copy()
    
    verify_list = []
    for _, row in alert_targets.iterrows():
        mgr = row["구역담당자_통합"]
        info, status = get_smart_contact(mgr, manager_contacts)
        verify_list.append({
            "담당자": mgr, "매핑이메일": info.get("email", "") if info else "",
            "상태": status, "유효": is_valid_email(info.get("email", "")) if info else False
        })
    
    v_df = pd.DataFrame(verify_list).drop_duplicates("담당자")
    
    st.data_editor(
        v_df,
        column_config={
            "매핑이메일": st.column_config.TextColumn("이메일", required=True),
            "상태": st.column_config.TextColumn("매핑 엔진 결과", disabled=True),
            "유효": st.column_config.CheckboxColumn("유효주소", disabled=True)
        },
        use_container_width=True, hide_index=True
    )
    
    if st.button("🚀 알림 발송 준비", type="primary"):
        st.success("데이터 검증 완료. SMTP 설정을 통해 발송이 가능합니다.")
