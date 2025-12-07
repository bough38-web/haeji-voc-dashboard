import os
import time
import smtplib
from datetime import datetime, date
from email.message import EmailMessage

import numpy as np
import pandas as pd
import streamlit as st

# Plotly
try:
    import plotly.express as px
    import plotly.graph_objects as go
    HAS_PLOTLY = True
except ImportError:
    HAS_PLOTLY = False

# ----------------------------------------------------
# 0. 기본 설정 & CSS (라이트톤/반응형)
# ----------------------------------------------------
st.set_page_config(
    page_title="해지 VOC 종합 대시보드 Pro",
    page_icon="🛡️",
    layout="wide",
    initial_sidebar_state="expanded"
)

st.markdown(
    """
    <style>
    /* Global Clean Style */
    html, body, [data-testid="stAppViewContainer"] {
        background-color: #f5f5f7;
        font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, Helvetica, Arial, sans-serif;
        color: #1d1d1f;
    }
    [data-testid="stHeader"] { background-color: #f5f5f7; }
    [data-testid="stSidebar"] { background-color: #ffffff; border-right: 1px solid #e5e5e5; }
    
    /* Card Style */
    .kpi-card {
        background-color: white;
        padding: 1.5rem;
        border-radius: 12px;
        box-shadow: 0 4px 6px rgba(0,0,0,0.02);
        border: 1px solid #f0f0f0;
        text-align: center;
    }
    .kpi-title { font-size: 0.9rem; color: #86868b; margin-bottom: 0.5rem; }
    .kpi-value { font-size: 1.8rem; font-weight: 700; color: #1d1d1f; }
    
    /* Timeline / Chat Style for Feedback */
    .timeline-item {
        background-color: white;
        border-radius: 12px;
        padding: 1rem;
        margin-bottom: 0.8rem;
        border-left: 4px solid #007aff;
        box-shadow: 0 2px 4px rgba(0,0,0,0.03);
    }
    .timeline-header { display: flex; justify-content: space-between; margin-bottom: 0.4rem; font-size: 0.85rem; color: #86868b; }
    .timeline-content { font-size: 0.95rem; color: #333; white-space: pre-wrap; line-height: 1.5; }
    
    /* Custom Metric Container adjustment */
    div[data-testid="stMetric"] { background-color: white; padding: 10px; border-radius: 10px; border: 1px solid #eee; }
    </style>
    """,
    unsafe_allow_html=True
)

# ----------------------------------------------------
# 1. 환경 변수 및 설정
# ----------------------------------------------------
# secrets.toml 혹은 환경변수 로드
SMTP_HOST = st.secrets.get("SMTP_HOST", os.getenv("SMTP_HOST", ""))
SMTP_PORT = int(st.secrets.get("SMTP_PORT", os.getenv("SMTP_PORT", 587)))
SMTP_USER = st.secrets.get("SMTP_USER", os.getenv("SMTP_USER", ""))
SMTP_PASSWORD = st.secrets.get("SMTP_PASSWORD", os.getenv("SMTP_PASSWORD", ""))
SENDER_NAME = st.secrets.get("SENDER_NAME", os.getenv("SENDER_NAME", "VOC 관리자"))

MERGED_PATH = "merged.xlsx"
FEEDBACK_PATH = "feedback.csv"
CONTACT_PATH = "contact_map.xlsx"

# ----------------------------------------------------
# 2. 데이터 로드 및 전처리 함수
# ----------------------------------------------------
@st.cache_data(ttl=600)  # 10분 캐싱
def load_and_process_data():
    # 1. VOC 파일 로드
    if not os.path.exists(MERGED_PATH):
        return pd.DataFrame(), pd.DataFrame(), {}, "FILE_NOT_FOUND"
    
    df = pd.read_excel(MERGED_PATH)
    
    # 기본 정제
    for col in ["계약번호", "고객번호"]:
        if col in df.columns:
            df[col] = df[col].astype(str).str.replace(",", "").str.strip()

    # 계약번호 정제 (특수문자 제거)
    if "계약번호" in df.columns:
        df["계약번호_정제"] = df["계약번호"].str.replace(r"[^0-9A-Za-z]", "", regex=True)
    else:
        df["계약번호_정제"] = ""
        
    # 날짜 변환
    if "접수일시" in df.columns:
        df["접수일시"] = pd.to_datetime(df["접수일시"], errors="coerce")

    # 2. 지사명 표준화
    if "관리지사" in df.columns:
        mapping = {
            "중앙지사": "중앙", "강북지사": "강북", "서대문지사": "서대문",
            "고양지사": "고양", "의정부지사": "의정부", "남양주지사": "남양주",
            "강릉지사": "강릉", "원주지사": "원주"
        }
        df["관리지사"] = df["관리지사"].replace(mapping)

    # 3. 담당자 통합
    def pick_manager(row):
        for c in ["구역담당자", "담당자", "처리자"]:
            if c in row and pd.notna(row[c]) and str(row[c]).strip():
                return str(row[c]).strip()
        return "미지정"
    df["구역담당자_통합"] = df.apply(pick_manager, axis=1)

    # 4. 해지 VOC vs 기타 데이터 분리 및 매칭
    if "출처" not in df.columns: df["출처"] = "해지VOC"
    
    df_voc = df[df["출처"] == "해지VOC"].copy()
    df_other = df[df["출처"] != "해지VOC"].copy()
    
    # 타 시스템(해지방어 활동 등)에 있는 계약번호 집합
    other_contract_set = set(df_other["계약번호_정제"].dropna())
    
    df_voc["매칭여부"] = df_voc["계약번호_정제"].apply(
        lambda x: "매칭(O)" if x in other_contract_set else "비매칭(X)"
    )

    # 5. 리스크 등급 산정 (최근 접수일수록 High)
    today = date.today()
    def get_risk(dt):
        if pd.isna(dt): return "LOW"
        days = (today - dt.date()).days
        if days <= 3: return "HIGH"
        elif days <= 10: return "MEDIUM"
        return "LOW"
        
    df_voc["리스크등급"] = df_voc["접수일시"].apply(get_risk)
    
    # 6. 월정료 처리
    fee_col = "시설_KTT월정료(조정)" if "시설_KTT월정료(조정)" in df_voc.columns else "KTT월정료(조정)"
    if fee_col in df_voc.columns:
        df_voc["월정료_수치"] = pd.to_numeric(
            df_voc[fee_col].astype(str).str.replace(",", "", regex=False), errors="coerce"
        )
        # 20만 이상이면 데이터 오류 가능성으로 1/10 보정 (예시 로직)
        df_voc["월정료_수치"] = df_voc["월정료_수치"].apply(lambda x: x/10 if x >= 200000 else x)
    else:
        df_voc["월정료_수치"] = 0

    return df_voc, df, {}, "SUCCESS"

@st.cache_data
def load_feedback_data():
    if os.path.exists(FEEDBACK_PATH):
        try:
            return pd.read_csv(FEEDBACK_PATH, encoding="utf-8-sig")
        except:
            return pd.read_csv(FEEDBACK_PATH)
    return pd.DataFrame(columns=["계약번호_정제", "고객대응내용", "등록자", "등록일자", "비고"])

def save_feedback_data(df):
    df.to_csv(FEEDBACK_PATH, index=False, encoding="utf-8-sig")

@st.cache_data
def load_contacts():
    if not os.path.exists(CONTACT_PATH): return {}
    df = pd.read_excel(CONTACT_PATH)
    # 컬럼 자동 탐지 로직 생략(간소화) - 실제론 이름, 이메일 컬럼 필요
    # 예시: 이름 -> 이메일 딕셔너리 반환
    contacts = {}
    name_col = next((c for c in df.columns if "담당" in c or "이름" in c), None)
    email_col = next((c for c in df.columns if "메일" in c or "mail" in c), None)
    if name_col and email_col:
        for _, row in df.iterrows():
            contacts[str(row[name_col]).strip()] = str(row[email_col]).strip()
    return contacts

# ----------------------------------------------------
# 3. 데이터 로딩 실행
# ----------------------------------------------------
df_voc, df_raw, _, status = load_and_process_data()
if status == "FILE_NOT_FOUND":
    st.error(f"❌ '{MERGED_PATH}' 파일을 찾을 수 없습니다. 루트 경로에 파일을 위치시켜주세요.")
    st.stop()

if "feedback_df" not in st.session_state:
    st.session_state["feedback_df"] = load_feedback_data()

contact_map = load_contacts()

# ----------------------------------------------------
# 4. 사이드바 (Global Filter)
# ----------------------------------------------------
with st.sidebar:
    st.title("🎛️ 필터 패널")
    
    # 날짜 필터
    min_date = df_voc["접수일시"].min().date()
    max_date = df_voc["접수일시"].max().date()
    date_range = st.date_input("접수일자 범위", value=(min_date, max_date))
    
    # 지사 필터
    all_branches = sorted(df_voc["관리지사"].dropna().unique().tolist())
    sel_branches = st.multiselect("관리지사", all_branches, default=all_branches)
    
    # 리스크 필터
    sel_risk = st.multiselect("리스크 등급", ["HIGH", "MEDIUM", "LOW"], default=["HIGH", "MEDIUM", "LOW"])
    
    # 매칭 여부
    sel_match = st.multiselect("매칭 여부", ["매칭(O)", "비매칭(X)"], default=["비매칭(X)"]) # 기본값을 비매칭으로
    
    st.divider()
    st.caption(f"Last Updated: {datetime.now().strftime('%Y-%m-%d %H:%M')}")

# ----------------------------------------------------
# 5. 데이터 필터링 로직
# ----------------------------------------------------
mask = (
    (df_voc["접수일시"].dt.date >= date_range[0]) &
    (df_voc["접수일시"].dt.date <= date_range[1]) &
    (df_voc["관리지사"].isin(sel_branches)) &
    (df_voc["리스크등급"].isin(sel_risk)) &
    (df_voc["매칭여부"].isin(sel_match))
)
filtered_voc = df_voc[mask].copy()

# ----------------------------------------------------
# 6. 메인 대시보드 (KPI)
# ----------------------------------------------------
st.title("📊 해지방어 활동 모니터링")

col1, col2, col3, col4 = st.columns(4)
with col1:
    st.metric("대상 VOC 건수", f"{len(filtered_voc):,}", delta="건")
with col2:
    high_risk_cnt = len(filtered_voc[filtered_voc['리스크등급']=='HIGH'])
    st.metric("High 리스크(3일이내)", f"{high_risk_cnt:,}", delta="건", delta_color="inverse")
with col3:
    unmatched_cnt = len(filtered_voc[filtered_voc['매칭여부']=='비매칭(X)'])
    st.metric("미조치(비매칭) 의심", f"{unmatched_cnt:,}", delta="건", delta_color="inverse")
with col4:
    # 처리율 (매칭/전체) - 필터 영향 안받는 전체 기준 계산 필요할 수도 있으나 여기선 필터 기준
    match_rate = (1 - (unmatched_cnt / len(filtered_voc))) * 100 if len(filtered_voc) > 0 else 0
    st.metric("시스템 등록률", f"{match_rate:.1f}%")

st.markdown("---")

# ----------------------------------------------------
# 7. 탭 구성
# ----------------------------------------------------
tab1, tab2, tab3, tab4 = st.tabs([
    "📈 종합 분석 (Chart)", 
    "📋 대상 목록 (List)", 
    "✍️ 상세/활동등록 (Detail)",
    "📨 알림 발송 (Email)"
])

# === TAB 1: 시각화 ===
with tab1:
    c1, c2 = st.columns([1, 1])
    
    with c1:
        st.subheader("📍 지사별/리스크별 현황")
        if HAS_PLOTLY:
            # Sunburst Chart: 지사 -> 리스크 -> 건수
            fig_sun = px.sunburst(
                filtered_voc, 
                path=['관리지사', '리스크등급'], 
                values='월정료_수치' if '월정료_수치' in filtered_voc.columns else None,
                color='리스크등급',
                color_discrete_map={'HIGH':'#ef4444', 'MEDIUM':'#f59e0b', 'LOW':'#10b981'},
                title="지사별 리스크 분포 (크기: 월정료 합계)"
            )
            fig_sun.update_layout(height=400, margin=dict(t=30, l=0, r=0, b=0))
            st.plotly_chart(fig_sun, use_container_width=True)
        else:
            st.bar_chart(filtered_voc["관리지사"].value_counts())

    with c2:
        st.subheader("📅 일별 접수 추이")
        daily_counts = filtered_voc.groupby(filtered_voc["접수일시"].dt.date).size().reset_index(name="건수")
        if HAS_PLOTLY:
            fig_line = px.line(daily_counts, x="접수일시", y="건수", markers=True, line_shape="spline")
            fig_line.update_layout(height=400, xaxis_title="접수일", yaxis_title="건수")
            st.plotly_chart(fig_line, use_container_width=True)
        else:
            st.line_chart(daily_counts.set_index("접수일시"))

# === TAB 2: 목록 (인터랙티브) ===
with tab2:
    st.info("💡 아래 표에서 행을 클릭(선택)하면 **'상세/활동등록'** 탭에서 내용을 바로 확인할 수 있습니다.")
    
    # 보여줄 컬럼 정의
    display_cols = [
        "계약번호_정제", "상호", "관리지사", "구역담당자_통합", 
        "리스크등급", "매칭여부", "접수일시", "처리내용", "월정료_수치"
    ]
    display_cols = [c for c in display_cols if c in filtered_voc.columns]
    
    # 최신순 정렬
    df_display = filtered_voc.sort_values("접수일시", ascending=False)[display_cols].reset_index(drop=True)
    
    # 스타일링 (리스크 하이라이트)
    def highlight_risk(val):
        color = '#ffebee' if val == 'HIGH' else ('#fff8e1' if val == 'MEDIUM' else '')
        return f'background-color: {color}'

    # Selection API 사용
    event = st.dataframe(
        df_display.style.map(highlight_risk, subset=['리스크등급']).format({"월정료_수치": "{:,.0f}"}),
        use_container_width=True,
        height=500,
        selection_mode="single-row",
        on_select="rerun",  # 선택 시 리런하여 탭3 데이터 갱신
        key="voc_list_selection"
    )

    # 선택된 계약번호 추출
    selected_contract = None
    if event.selection.rows:
        idx = event.selection.rows[0]
        selected_contract = df_display.iloc[idx]["계약번호_정제"]
        st.session_state["selected_cn"] = selected_contract  # 세션에 저장
    elif "selected_cn" in st.session_state:
        # 이전에 선택된 값이 있다면 유지 (탭 이동 시 초기화 방지)
        selected_contract = st.session_state["selected_cn"]

# === TAB 3: 상세/활동등록 ===
with tab3:
    col_d1, col_d2 = st.columns([1, 2])
    
    with col_d1:
        st.subheader("🔍 조회 대상")
        # 탭2에서 선택된 값이 있으면 자동 입력, 아니면 빈칸
        default_val = selected_contract if selected_contract else ""
        input_cn = st.text_input("계약번호", value=default_val, help="목록 탭에서 선택하면 자동 입력됩니다.")
        
        target_row = None
        if input_cn:
            subset = df_voc[df_voc["계약번호_정제"] == input_cn.strip()]
            if not subset.empty:
                target_row = subset.sort_values("접수일시", ascending=False).iloc[0]
                
                # 기본 정보 카드
                st.markdown(
                    f"""
                    <div class="kpi-card" style="text-align:left; padding:1rem;">
                        <div style="font-weight:bold; font-size:1.1rem; margin-bottom:0.5rem;">{target_row.get('상호', '상호미상')}</div>
                        <div style="color:#555; font-size:0.9rem;">
                        📍 {target_row.get('관리지사', '-')} / {target_row.get('구역담당자_통합', '-')}<br>
                        📅 접수: {target_row.get('접수일시', '-')}<br>
                        💰 월정료: {target_row.get('월정료_수치', 0):,.0f}원
                        </div>
                    </div>
                    """, 
                    unsafe_allow_html=True
                )
                
                st.markdown("#### 📜 VOC 원문")
                st.info(target_row.get("처리내용", "내용 없음"))
            else:
                st.warning("해당 계약번호 데이터를 찾을 수 없습니다.")

    with col_d2:
        if target_row is not None:
            st.subheader("💬 대응 이력 (Timeline)")
            
            # 피드백 로드
            fb_df = st.session_state["feedback_df"]
            curr_fb = fb_df[fb_df["계약번호_정제"] == input_cn.strip()].sort_values("등록일자", ascending=False)
            
            # 타임라인 UI
            with st.container(height=300):
                if not curr_fb.empty:
                    for _, row in curr_fb.iterrows():
                        st.markdown(
                            f"""
                            <div class="timeline-item">
                                <div class="timeline-header">
                                    <span>👤 {row['등록자']}</span>
                                    <span>{row['등록일자']}</span>
                                </div>
                                <div class="timeline-content">{row['고객대응내용']}</div>
                            </div>
                            """,
                            unsafe_allow_html=True
                        )
                else:
                    st.caption("아직 등록된 활동 내역이 없습니다.")
            
            st.divider()
            
            # 입력 폼
            st.markdown("#### ✍️ 신규 활동 등록")
            with st.form("new_feedback"):
                c_f1, c_f2 = st.columns(2)
                with c_f1:
                    writer = st.text_input("담당자명", value=target_row.get('구역담당자_통합', ''))
                with c_f2:
                    act_date = st.date_input("활동일", value=datetime.today())
                
                content = st.text_area("활동 상세 내용", placeholder="고객 통화 결과 및 방어 성공 여부 등...", height=100)
                
                if st.form_submit_button("등록 저장", use_container_width=True):
                    if not content.strip():
                        st.error("내용을 입력해주세요.")
                    else:
                        new_data = {
                            "계약번호_정제": input_cn.strip(),
                            "고객대응내용": content,
                            "등록자": writer,
                            "등록일자": str(act_date),
                            "비고": "Dashboard"
                        }
                        st.session_state["feedback_df"] = pd.concat(
                            [st.session_state["feedback_df"], pd.DataFrame([new_data])], 
                            ignore_index=True
                        )
                        save_feedback_data(st.session_state["feedback_df"])
                        st.success("저장되었습니다! (타임라인에 반영됩니다)")
                        time.sleep(1)
                        st.rerun()
        else:
            st.empty()

# === TAB 4: 알림 발송 ===
with tab4:
    st.subheader("📨 미조치 건 담당자 알림")
    
    # 발송 대상 로직: 비매칭 & High 리스크
    targets = filtered_voc[
        (filtered_voc["매칭여부"] == "비매칭(X)") & 
        (filtered_voc["리스크등급"] == "HIGH")
    ].copy()
    
    if targets.empty:
        st.success("현재 조건에서 알림 발송 대상(비매칭+High)이 없습니다.")
    else:
        # 담당자 이메일 매핑
        targets["이메일"] = targets["구역담당자_통합"].apply(lambda x: contact_map.get(x, ""))
        
        # 집계
        agg_targets = targets.groupby(["관리지사", "구역담당자_통합", "이메일"]).size().reset_index(name="대상건수")
        
        st.write(f"총 **{len(agg_targets)}명**의 담당자에게 **{agg_targets['대상건수'].sum()}건**에 대한 알림이 필요합니다.")
        
        with st.expander("📋 발송 리스트 확인 및 이메일 수정", expanded=True):
            edited_targets = st.data_editor(
                agg_targets,
                column_config={
                    "이메일": st.column_config.TextColumn("이메일", help="비어있으면 발송되지 않습니다.", required=True)
                },
                use_container_width=True,
                num_rows="dynamic"
            )
            
        c_mail1, c_mail2 = st.columns([2, 1])
        with c_mail1:
            subject_tpl = st.text_input("메일 제목", "[긴급] 해지방어 활동 미등록 건 확인 요청")
            body_tpl = st.text_area("메일 본문 템플릿", 
"""{담당자}님 안녕하세요.
현재 담당 구역 내 해지접수 후 방어활동이 등록되지 않은 긴급 건이 {건수}건 있습니다.
대시보드 접속 후 확인 부탁드립니다.

감사합니다.""", height=150)
            
        with c_mail2:
            st.markdown("#### 🚀 발송 옵션")
            is_dry_run = st.toggle("모의 발송 (Dry Run)", value=True, help="실제 메일을 보내지 않고 로그만 남깁니다.")
            
            if st.button("메일 발송 시작", type="primary", use_container_width=True):
                progress_bar = st.progress(0)
                status_txt = st.empty()
                
                success, fail = 0, 0
                total = len(edited_targets)
                
                for i, row in edited_targets.iterrows():
                    name = row["구역담당자_통합"]
                    email = row["이메일"]
                    count = row["대상건수"]
                    
                    status_txt.text(f"Sending to {name}...")
                    time.sleep(0.1) # UI 갱신용 딜레이
                    
                    if not email or "@" not in email:
                        fail += 1
                        continue
                        
                    try:
                        if not is_dry_run:
                            # 실제 SMTP 발송
                            msg = EmailMessage()
                            msg.set_content(body_tpl.format(담당자=name, 건수=count))
                            msg["Subject"] = subject_tpl
                            msg["From"] = SENDER_NAME
                            msg["To"] = email
                            
                            with smtplib.SMTP(SMTP_HOST, SMTP_PORT) as server:
                                server.starttls()
                                server.login(SMTP_USER, SMTP_PASSWORD)
                                server.send_message(msg)
                        success += 1
                    except Exception as e:
                        st.error(f"{name} 발송 실패: {e}")
                        fail += 1
                    
                    progress_bar.progress((i+1)/total)
                
                mode_txt = "[모의]" if is_dry_run else "[실제]"
                st.success(f"{mode_txt} 발송 완료! 성공: {success} / 실패(이메일없음 등): {fail}")
