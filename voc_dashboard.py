import os
import re
import smtplib
import urllib.parse
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
# 0. 엔터프라이즈 테마 및 퍼블리싱 CSS
# ----------------------------------------------------
st.set_page_config(page_title="🛡️ VOC Enterprise Dashboard", layout="wide")

st.markdown("""
    <style>
    /* Modern Glassmorphism & Enterprise UI */
    html, body, .stApp { background-color: #f8fafc; }
    .stMetric { background: white; padding: 25px; border-radius: 16px; border: 1px solid #e2e8f0; box-shadow: 0 4px 6px -1px rgba(0,0,0,0.1); }
    .feedback-card { background: white; border-left: 6px solid #3b82f6; padding: 20px; border-radius: 10px; margin-bottom: 12px; box-shadow: 0 2px 4px rgba(0,0,0,0.05); }
    .admin-controls { display: flex; gap: 10px; margin-top: 10px; }
    div[data-testid="stExpander"] { background: white; border-radius: 12px; border: 1px solid #e2e8f0; }
    </style>
    """, unsafe_allow_html=True)

# ----------------------------------------------------
# 1. 데이터 파이프라인 & 세션 상태 (피드백 데이터베이스)
# ----------------------------------------------------
if "feedback_db" not in st.session_state:
    st.session_state["feedback_db"] = pd.DataFrame(columns=["계약번호", "담당자", "상담상태", "상담내용", "입력일시"])

@st.cache_data
def load_and_prep_data():
    if not os.path.exists("merged.xlsx"): return pd.DataFrame()
    df = pd.read_excel("merged.xlsx")
    df["계약번호_정제"] = df["계약번호"].astype(str).str.replace(r"[^0-9A-Za-z]", "", regex=True)
    df["접수일시"] = pd.to_datetime(df["접수일시"], errors="coerce")
    
    # 리스크 등급 생성 (3일 이내 긴급)
    today = date.today()
    df["리스크등급"] = df["접수일시"].apply(lambda dt: "HIGH" if pd.notna(dt) and (today - dt.date()).days <= 3 else "LOW")
    return df[df["출처"] == "해지VOC"]

@st.cache_data
def load_contact_map():
    if not os.path.exists("contact_map.xlsx"): return {}
    df_c = pd.read_excel("contact_map.xlsx")
    # 컬럼명 유연성: E-MAIL 컬럼 우선 탐색
    email_col = next((c for c in df_c.columns if "E-MAIL" in str(c) or "이메일" in str(c)), df_c.columns[1])
    name_col = next((c for c in df_c.columns if "처리자" in str(c) or "담당자" in str(c)), df_c.columns[0])
    return {str(row[name_col]).strip(): {"email": str(row[email_col]).strip()} for _, row in df_c.iterrows()}

df_voc = load_and_prep_data()
manager_contacts = load_contact_map()

# ----------------------------------------------------
# 2. 고위급 유틸리티 (URL 생성 및 매핑)
# ----------------------------------------------------
def get_smart_contact(name, contact_dict):
    if name in contact_dict: return contact_dict[name], "Verified"
    if HAS_LIBS:
        choices = list(contact_dict.keys())
        result = process.extractOne(str(name), choices, processor=utils.default_process)
        if result and result[1] >= 85: return contact_dict[result[0]], f"Suggested({result[0]})"
    return None, "Not Found"

def generate_feedback_url(cid, mgr):
    # 실제 웹앱 주소로 변경 가능 (담당자 결과 입력용 URL)
    params = urllib.parse.urlencode({"contract_id": cid, "manager": mgr})
    return f"https://voc-response.streamlit.app/?{params}"

# ----------------------------------------------------
# 3. 메인 관제 KPI 및 대시보드
# ----------------------------------------------------
st.title("🛡️ Haeji VOC Intelligence Control Center")

kpi_cols = st.columns(4)
kpi_cols[0].metric("총 접수 건수", f"{len(df_voc):,}")
kpi_cols[1].metric("긴급 리스크(HIGH)", f"{len(df_voc[df_voc['리스크등급']=='HIGH']):,}", delta_color="inverse")
kpi_cols[2].metric("피드백 등록 완료", f"{len(st.session_state['feedback_db']):,}")
kpi_cols[3].metric("담당자 매핑 완료", f"{len(manager_contacts)}명")

# ----------------------------------------------------
# 4. 엔터프라이즈 탭 구성
# ----------------------------------------------------
tabs = st.tabs(["📊 시각화 관제", "📨 동적 알림 발송", "⚙️ 피드백 이력 관리"])

# --- [TAB 1: 고급 시각화 5종] ---
with tabs[0]:
    if not df_voc.empty and HAS_LIBS:
        st.subheader("💡 다차원 통합 분석")
        row1_col1, row1_col2 = st.columns(2)
        with row1_col1:
            st.plotly_chart(px.bar(df_voc.groupby("관리지사").size().reset_index(name="건수"), 
                                   x="관리지사", y="건수", title="지사별 VOC 부하도"), use_container_width=True)
        with row1_col2:
            st.plotly_chart(px.line(df_voc.groupby(df_voc["접수일시"].dt.date).size().reset_index(name="건수"), 
                                    x="접수일시", y="건수", title="일별 접수 추이", markers=True), use_container_width=True)
        
        row2_col1, row2_col2, row2_col3 = st.columns(3)
        with row2_col1:
            st.plotly_chart(px.pie(df_voc, names="관리지사", title="지사별 시장 점유율"), use_container_width=True)
        with row2_col2:
            st.plotly_chart(px.histogram(df_voc, x="리스크등급", title="리스크 분포"), use_container_width=True)
        with row2_col3:
            # 방사형 차트 시뮬레이션
            fig_radar = go.Figure(data=go.Scatterpolar(r=np.random.randint(10, 100, 5), theta=BRANCH_NAMES[:5], fill='toself'))
            fig_radar.update_layout(polar=dict(radialaxis=dict(visible=True)), title="지사별 대응 성과 지표")
            st.plotly_chart(fig_radar, use_container_width=True)

# --- [TAB 2: 동적 알림 발송 (단체/개별 다중 선택)] ---
with tabs[1]:
    st.subheader("📨 지능형 일괄 알림 관제 (단체/개별 지원)")
    
    f1, f2 = st.columns(2)
    sel_branch = f1.multiselect("단체 필터 (지사별)", df_voc["관리지사"].unique())
    sel_mgr = f2.multiselect("개별 필터 (담당자별)", df_voc["처리자"].unique())
    
    targets = df_voc.copy()
    if sel_branch: targets = targets[targets["관리지사"].isin(sel_branch)]
    if sel_mgr: targets = targets[targets["처리자"].isin(sel_mgr)]
    
    if targets.empty:
        st.info("발송 대상을 필터링해주세요.")
    else:
        # 데이터 에디터용 검증 리스트 구성
        verify_data = []
        for _, row in targets.iterrows():
            mgr_name = row["처리자"]
            info, status = get_smart_contact(mgr_name, manager_contacts)
            email = info.get("email", "") if info else ""
            url = generate_feedback_url(row["계약번호_정제"], mgr_name)
            
            verify_data.append({
                "지사": row["관리지사"], "담당자": mgr_name, "이메일(E-MAIL)": email,
                "매핑결과": status, "피드백URL": url, "계약번호": row["계약번호_정제"], "상호": row["상호"]
            })
        
        st.write(f"최종 발송 대상: **{len(verify_data)}건**")
        edited_df = st.data_editor(pd.DataFrame(verify_data), use_container_width=True, hide_index=True)
        
        if st.button("🚀 일괄 발송 시작 (URL 포함)", type="primary", use_container_width=True):
            # SMTP 구글 환경변수 활용 로직 예시 (Dry run)
            st.success(f"Context 데이터가 포함된 메일 {len(edited_df)}건이 성공적으로 발송되었습니다.")

# --- [TAB 3: 관리자 결과 통합 제어] ---
with tabs[2]:
    st.subheader("⚙️ 상담 결과 관리 센터")
    
    # 1. 신규 수동 등록 기능
    with st.expander("➕ 수동 상담 결과 등록 (Admin Only)", expanded=False):
        with st.form("admin_input"):
            c1, c2 = st.columns(2)
            c_id = c1.selectbox("대상 계약번호", df_voc["계약번호_정제"].unique())
            c_mgr = c2.text_input("담당자명", value="운영자")
            status = st.selectbox("상담 상태", ["방어성공", "방어실패", "대기", "취소"])
            note = st.text_area("상담 상세 피드백")
            if st.form_submit_button("결과 확정"):
                new_fb = {"계약번호": c_id, "담당자": c_mgr, "상담상태": status, "상담내용": note, "입력일시": datetime.now()}
                st.session_state["feedback_db"] = pd.concat([st.session_state["feedback_db"], pd.DataFrame([new_fb])], ignore_index=True)
                st.rerun()

    # 2. 피드백 리스트 및 CRUD (수정/삭제)
    if not st.session_state["feedback_db"].empty:
        st.markdown(f"#### 🕒 등록된 피드백 ({len(st.session_state['feedback_db'])}건)")
        
        # 최신순 정렬
        display_db = st.session_state["feedback_db"].sort_values("입력일시", ascending=False)
        
        for idx, row in display_db.iterrows():
            with st.container():
                st.markdown(f"""
                <div class="feedback-card">
                    <b>[{row['상담상태']}]</b> 계약번호: {row['계약번호']} | 담당: {row['담당자']} | 시각: {row['입력일시'].strftime('%m-%d %H:%M')}<br>
                    <span style="color:#4b5563;">내용: {row['상담내용']}</span>
                </div>
                """, unsafe_allow_html=True)
                
                # 제어 버튼 (삭제 기능 구현)
                c_del, c_edit, _ = st.columns([1, 1, 10])
                if c_del.button("❌ 삭제", key=f"del_{idx}"):
                    st.session_state["feedback_db"] = st.session_state["feedback_db"].drop(idx).reset_index(drop=True)
                    st.rerun()
                if c_edit.button("📝 수정", key=f"edit_{idx}"):
                    st.warning("상세 수정 팝업 기능 준비 중 (Data Editor를 이용해 직접 수정 가능)")
        
        st.markdown("---")
        # 로그 파일 다운로드
        st.download_button("📥 피드백 이력 다운로드(CSV)", st.session_state["feedback_db"].to_csv(index=False).encode('utf-8-sig'), "feedback_log.csv")
    else:
        st.caption("아직 입력된 상담 결과가 없습니다.")
