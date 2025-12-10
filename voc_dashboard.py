import streamlit as st
import pandas as pd
import numpy as np
import os
from datetime import datetime, date, timedelta
import smtplib
from email.message import EmailMessage

# ==============================
# 0. 공통 설정 / 전역 상수
# ==============================

# 세션 초기화
if "login_type" not in st.session_state:
    st.session_state["login_type"] = None
if "login_user" not in st.session_state:
    st.session_state["login_user"] = None

ADMIN_CODE = "C3A"                 # 관리자 비밀번호
MERGED_PATH = "merged.xlsx"        # VOC 통합파일
FEEDBACK_PATH = "feedback.csv"     # 처리내역 CSV 저장 경로
CONTACT_PATH = "contact_map.xlsx"  # 담당자 매핑 파일

# Plotly 사용 여부 확인
try:
    import plotly.express as px
    HAS_PLOTLY = True
except ImportError:
    HAS_PLOTLY = False

# ==============================
# 1. 페이지 기본 설정 & CSS
# ==============================
st.set_page_config(page_title="해지 VOC 종합 대시보드", layout="wide", page_icon="📊")

st.markdown(
    """
    <style>
    /* 전체 배경 및 폰트 */
    .stApp {
        background-color: #f5f5f7;
        font-family: "Pretendard", -apple-system, BlinkMacSystemFont, system-ui, Roboto, sans-serif;
    }
    
    /* 사이드바 스타일 */
    section[data-testid="stSidebar"] {
        background-color: #ffffff;
        border-right: 1px solid #e0e0e0;
    }
    
    /* 카드 스타일 (컨테이너) */
    .section-card {
        background: #ffffff;
        border-radius: 12px;
        padding: 20px;
        border: 1px solid #e0e0e0;
        box-shadow: 0 2px 4px rgba(0,0,0,0.05);
        margin-bottom: 20px;
    }
    
    /* 피드백 아이템 스타일 */
    .feedback-item {
        background-color: #f8f9fa;
        border-radius: 8px;
        padding: 12px 16px;
        margin-bottom: 10px;
        border-left: 4px solid #3b82f6;
    }
    .feedback-meta {
        font-size: 0.85rem;
        color: #6c757d;
        margin-top: 4px;
    }
    
    /* 로그인 화면 스타일 */
    .login-wrapper {
        display: flex;
        justify-content: center;
        align-items: center;
        height: 100vh;
    }
    </style>
    """,
    unsafe_allow_html=True,
)

# ==============================
# 2. 더미 데이터 생성 (파일 없을 시)
# ==============================
def create_dummy_data():
    """실행을 위한 샘플 데이터 생성"""
    if not os.path.exists(MERGED_PATH):
        st.warning(f"⚠ '{MERGED_PATH}' 파일이 없어 샘플 데이터를 생성합니다.")
        data = {
            "계약번호": [f"1000{i}" for i in range(50)],
            "고객번호": [f"C00{i}" for i in range(50)],
            "상호": [f"고객사_{i}" for i in range(50)],
            "접수일시": [datetime.now() - timedelta(days=np.random.randint(0, 30)) for _ in range(50)],
            "관리지사": np.random.choice(["강북", "강남", "서대문", "고양", "의정부", "남양주", "강릉", "원주"], 50),
            "담당유형": ["SP"] * 50,
            "구역담당자": np.random.choice(["김철수", "이영희", "박민수", "정수진"], 50),
            "출처": np.random.choice(["해지VOC", "해지VOC", "기타"], 50),
            "시설_KTT월정료(조정)": np.random.randint(30000, 300000, 50),
            "해지상세": np.random.choice(["비싸요", "폐업", "이전", "타사도입", "불만"], 50),
            "VOC유형소": np.random.choice(["요금", "위약금", "품질", "서비스"], 50),
            "VOC유형": np.random.choice(["해지방어", "일반문의"], 50),
            "VOC유형중": ["전체"] * 50,
            "등록내용": ["상담 요청"] * 50,
            "처리내용": ["방어 성공", "실패", "부재중"] * 16 + ["완료", "처리중"],
            "매칭여부": np.random.choice(["매칭(O)", "비매칭(X)"], 50),
            "리텐션P": np.random.randint(0, 100, 50)
        }
        df = pd.DataFrame(data)
        df.to_excel(MERGED_PATH, index=False)
        st.success("✅ 샘플 데이터(merged.xlsx) 생성 완료!")

    if not os.path.exists(CONTACT_PATH):
        st.warning(f"⚠ '{CONTACT_PATH}' 파일이 없어 샘플 데이터를 생성합니다.")
        contact_data = {
            "구역담당자": ["김철수", "이영희", "박민수", "정수진"],
            "휴대폰": ["010-1111-1111", "010-2222-2222", "010-3333-3333", "010-4444-4444"],
            "이메일": ["user1@example.com", "user2@example.com", "user3@example.com", "user4@example.com"]
        }
        df_c = pd.DataFrame(contact_data)
        df_c.to_excel(CONTACT_PATH, index=False)
        st.success("✅ 샘플 담당자 파일(contact_map.xlsx) 생성 완료!")

# 앱 시작 시 데이터 확인 및 생성
create_dummy_data()

# ==============================
# 3. 유틸 함수
# ==============================
def safe_str(x):
    if pd.isna(x):
        return ""
    return str(x).strip()

def detect_column(df: pd.DataFrame, keywords: list[str]) -> str | None:
    for k in keywords:
        if k in df.columns:
            return k
    for col in df.columns:
        s = str(col)
        for k in keywords:
            if k.lower() in s.lower():
                return col
    return None

def force_bar_chart(df: pd.DataFrame, x: str, y: str, height: int = 280):
    if df.empty:
        st.info("표시할 데이터가 없습니다.")
        return
    if HAS_PLOTLY:
        fig = px.bar(df, x=x, y=y, text=y, color_discrete_sequence=['#3b82f6'])
        fig.update_traces(textposition="outside")
        fig.update_layout(height=height, margin=dict(l=20, r=20, t=30, b=20))
        st.plotly_chart(fig, use_container_width=True)
    else:
        st.bar_chart(df.set_index(x)[y])

def force_stacked_bar(df: pd.DataFrame, x: str, y_cols: list[str], height: int = 280):
    if df.empty:
        st.info("표시할 데이터가 없습니다.")
        return
    if HAS_PLOTLY:
        fig = px.bar(df, x=x, y=y_cols, height=height, barmode='stack', 
                     color_discrete_map={"HIGH": "#ef4444", "MEDIUM": "#f59e0b", "LOW": "#10b981"})
        st.plotly_chart(fig, use_container_width=True)
    else:
        st.bar_chart(df.set_index(x)[y_cols])

# ==============================
# 4. 데이터 로드
# ==============================
@st.cache_data
def load_voc_data(path: str) -> pd.DataFrame:
    try:
        df = pd.read_excel(path)
    except Exception as e:
        st.error(f"데이터 로드 중 오류 발생: {e}")
        return pd.DataFrame()

    # 전처리
    for col in ["계약번호", "고객번호"]:
        if col in df.columns:
            df[col] = df[col].astype(str).str.replace(",", "").str.strip()
    
    if "계약번호" in df.columns:
        df["계약번호_정제"] = df["계약번호"].str.replace(r"[^0-9A-Za-z]", "", regex=True)
    
    return df

@st.cache_data
def load_contact_map(path: str):
    if not os.path.exists(path):
        return pd.DataFrame(), {}
    
    df_c = pd.read_excel(path)
    name_col = detect_column(df_c, ["구역담당자", "담당자", "성명"])
    email_col = detect_column(df_c, ["이메일", "메일"])
    phone_col = detect_column(df_c, ["휴대폰", "전화", "연락처"])

    if not name_col:
        return df_c, {}

    df_c["휴대폰"] = df_c[phone_col].apply(lambda x: "".join(filter(str.isdigit, str(x)))) if phone_col else ""
    
    manager_contacts = {}
    for _, row in df_c.iterrows():
        name = safe_str(row.get(name_col))
        if name:
            manager_contacts[name] = {
                "email": safe_str(row.get(email_col)) if email_col else "",
                "phone": safe_str(row.get(phone_col)) if phone_col else ""
            }
    return df_c, manager_contacts

# 데이터 로딩 실행
df = load_voc_data(MERGED_PATH)
contact_df, manager_contacts = load_contact_map(CONTACT_PATH)

if "feedback_df" not in st.session_state:
    if os.path.exists(FEEDBACK_PATH):
        st.session_state["feedback_df"] = pd.read_csv(FEEDBACK_PATH, encoding="utf-8-sig")
    else:
        st.session_state["feedback_df"] = pd.DataFrame(columns=["계약번호_정제", "고객대응내용", "등록자", "등록일자", "비고"])

# ==============================
# 5. 로그인 처리
# ==============================
BRANCH_ADMIN_PW = {
    "중앙": "C001", "강북": "C002", "서대문": "C003", "고양": "C004",
    "의정부": "C005", "남양주": "C006", "강릉": "C007", "원주": "C008",
}

def login_form():
    st.markdown("<h2 style='text-align: center; color: #005BAC;'>🔐 해지 VOC 대시보드 로그인</h2>", unsafe_allow_html=True)
    
    tab1, tab2, tab3 = st.tabs(["관리자", "사용자(담당자)", "지사 관리자"])
    
    with tab1:
        pw = st.text_input("관리자 비밀번호", type="password", key="admin_pw")
        if st.button("관리자 로그인", use_container_width=True):
            if pw == ADMIN_CODE:
                st.session_state["login_type"] = "admin"
                st.session_state["login_user"] = "ADMIN"
                st.rerun()
            else:
                st.error("비밀번호 불일치")

    with tab2:
        name = st.text_input("성명", key="user_name")
        pw = st.text_input("휴대폰 뒷 4자리", type="password", key="user_pw")
        if st.button("사용자 로그인", use_container_width=True):
            user_info = manager_contacts.get(name.strip())
            if user_info:
                real_tel = user_info.get("phone", "")
                # 연락처에서 숫자만 추출 후 뒤 4자리 비교
                real_pw = "".join(filter(str.isdigit, real_tel))[-4:]
                if pw == real_pw:
                    st.session_state["login_type"] = "user"
                    st.session_state["login_user"] = name.strip()
                    st.rerun()
                else:
                    st.error("비밀번호(휴대폰 뒷 4자리)가 일치하지 않습니다.")
            else:
                st.error("등록되지 않은 사용자입니다.")

    with tab3:
        branch = st.selectbox("지사 선택", list(BRANCH_ADMIN_PW.keys()))
        pw = st.text_input("지사 비밀번호", type="password", key="branch_pw")
        if st.button("지사 로그인", use_container_width=True):
            if pw == BRANCH_ADMIN_PW.get(branch):
                st.session_state["login_type"] = "branch_admin"
                st.session_state["login_branch"] = branch
                st.session_state["login_user"] = f"{branch} 관리자"
                st.rerun()
            else:
                st.error("비밀번호 불일치")

if st.session_state["login_type"] is None:
    login_form()
    st.stop()

# ==============================
# 6. 데이터 전처리 및 필터링
# ==============================
LOGIN_TYPE = st.session_state["login_type"]
LOGIN_USER = st.session_state["login_user"]

# 지사명 통일
if "관리지사" in df.columns:
    df["관리지사"] = df["관리지사"].astype(str).replace({
        "중앙지사": "중앙", "강북지사": "강북", "서대문지사": "서대문",
        "고양지사": "고양", "의정부지사": "의정부", "남양주지사": "남양주",
        "강릉지사": "강릉", "원주지사": "원주"
    })

# 날짜 파싱
if "접수일시" in df.columns:
    df["접수일시"] = pd.to_datetime(df["접수일시"], errors="coerce")

# 리스크 등급 계산
today = datetime.now()
def calc_risk(row):
    dt = row.get("접수일시")
    if pd.isna(dt): return np.nan, "LOW"
    days = (today - dt).days
    if days <= 3: return days, "HIGH"
    elif days <= 10: return days, "MEDIUM"
    return days, "LOW"

df[["경과일수", "리스크등급"]] = df.apply(lambda x: pd.Series(calc_risk(x)), axis=1)

# 구역담당자 통합
df["구역담당자_통합"] = df["구역담당자"].fillna(df["처리자"]).fillna("미지정")

# 기본 필터링 (SP 유형만)
df_voc = df[df["출처"] == "해지VOC"].copy()
if "담당유형" in df_voc.columns:
    df_voc = df_voc[df_voc["담당유형"].astype(str) == "SP"]

# 권한별 데이터 필터링
if LOGIN_TYPE == "user":
    df_voc = df_voc[df_voc["구역담당자_통합"] == LOGIN_USER]
elif LOGIN_TYPE == "branch_admin":
    df_voc = df_voc[df_voc["관리지사"] == st.session_state["login_branch"]]

# ==============================
# 7. 사이드바 (필터)
# ==============================
with st.sidebar:
    st.header(f"👤 {LOGIN_USER} 님")
    if st.button("로그아웃"):
        st.session_state["login_type"] = None
        st.rerun()
    
    st.markdown("---")
    st.header("🔧 검색 필터")
    
    # 날짜 범위 (원본 코드에 누락되었던 부분 추가)
    min_date = df_voc["접수일시"].min() if not df_voc.empty else datetime.now()
    max_date = df_voc["접수일시"].max() if not df_voc.empty else datetime.now()
    
    if pd.isna(min_date): min_date = datetime.now()
    if pd.isna(max_date): max_date = datetime.now()

    dr = st.date_input("📅 접수일자 범위", [min_date, max_date])

    sel_branches = st.multiselect(
        "🏢 관리지사", 
        ["전체"] + sorted(df_voc["관리지사"].unique().tolist()), 
        default=["전체"]
    )
    
    sel_risk = st.multiselect(
        "⚠ 리스크 등급",
        ["HIGH", "MEDIUM", "LOW"],
        default=["HIGH", "MEDIUM", "LOW"]
    )

# 필터 적용
voc_filtered_global = df_voc.copy()

# 날짜 필터 적용
if len(dr) == 2:
    start_date, end_date = pd.to_datetime(dr[0]), pd.to_datetime(dr[1])
    voc_filtered_global = voc_filtered_global[
        (voc_filtered_global["접수일시"] >= start_date) & 
        (voc_filtered_global["접수일시"] <= end_date + timedelta(days=1))
    ]

# 지사 필터 적용
if "전체" not in sel_branches:
    voc_filtered_global = voc_filtered_global[voc_filtered_global["관리지사"].isin(sel_branches)]

# 리스크 필터 적용
voc_filtered_global = voc_filtered_global[voc_filtered_global["리스크등급"].isin(sel_risk)]

# ==============================
# 8. 메인 대시보드 UI
# ==============================
st.title("📊 해지 VOC 종합 대시보드")

# 상단 KPI
k1, k2, k3, k4 = st.columns(4)
total_cnt = len(voc_filtered_global)
high_risk_cnt = len(voc_filtered_global[voc_filtered_global["리스크등급"]=="HIGH"])
today_cnt = len(voc_filtered_global[voc_filtered_global["접수일시"].dt.date == date.today()])

k1.metric("총 접수 건수", f"{total_cnt}건")
k2.metric("HIGH 리스크", f"{high_risk_cnt}건", delta_color="inverse")
k3.metric("금일 접수", f"{today_cnt}건")
k4.metric("처리율", "준비중") # 실제 데이터 컬럼 필요

# 탭 구성
tab1, tab2, tab3, tab4 = st.tabs(["📊 현황 시각화", "📋 전체 리스트", "📝 활동 내역 등록", "📨 담당자 알림"])

# --- 탭 1: 시각화 ---
with tab1:
    col1, col2 = st.columns(2)
    
    with col1:
        st.subheader("지사별 접수 현황")
        if not voc_filtered_global.empty:
            df_branch = voc_filtered_global["관리지사"].value_counts().reset_index()
            df_branch.columns = ["관리지사", "건수"]
            force_bar_chart(df_branch, "관리지사", "건수")
        else:
            st.info("데이터가 없습니다.")

    with col2:
        st.subheader("리스크 등급 분포")
        if not voc_filtered_global.empty:
            df_risk = voc_filtered_global["리스크등급"].value_counts().reset_index()
            df_risk.columns = ["등급", "건수"]
            force_bar_chart(df_risk, "등급", "건수")
        else:
            st.info("데이터가 없습니다.")
            
    # 적층형 그래프 (Plotly 필요)
    if HAS_PLOTLY and not voc_filtered_global.empty:
        st.subheader("지사별 리스크 현황")
        df_stack = voc_filtered_global.groupby(["관리지사", "리스크등급"]).size().reset_index(name="건수")
        force_stacked_bar(df_stack, "관리지사", "건수")

# --- 탭 2: 전체 리스트 ---
with tab2:
    st.subheader("📋 VOC 접수 리스트")
    st.dataframe(voc_filtered_global, use_container_width=True, height=500)

# --- 탭 3: 활동 내역 등록 ---
with tab3:
    st.subheader("📝 해지방어 활동 내역 등록")
    
    # 선택할 계약번호 리스트
    contract_list = voc_filtered_global["계약번호_정제"].unique().tolist()
    selected_contract = st.selectbox("계약번호 선택", ["(선택)"] + contract_list)
    
    if selected_contract != "(선택)":
        target_row = voc_filtered_global[voc_filtered_global["계약번호_정제"] == selected_contract].iloc[0]
        st.info(f"선택된 고객: {target_row.get('상호', '미상')} ({target_row.get('관리지사', '')})")
        
        with st.form("feedback_form"):
            content = st.text_area("고객 대응 내용 입력")
            note = st.text_input("비고")
            submitted = st.form_submit_button("등록")
            
            if submitted:
                new_data = {
                    "계약번호_정제": selected_contract,
                    "고객대응내용": content,
                    "등록자": LOGIN_USER,
                    "등록일자": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                    "비고": note
                }
                st.session_state["feedback_df"] = pd.concat([st.session_state["feedback_df"], pd.DataFrame([new_data])], ignore_index=True)
                st.session_state["feedback_df"].to_csv(FEEDBACK_PATH, index=False, encoding="utf-8-sig")
                st.success("활동 내역이 저장되었습니다.")
                
    st.markdown("---")
    st.subheader("📜 최근 활동 내역")
    if not st.session_state["feedback_df"].empty:
        st.dataframe(st.session_state["feedback_df"].sort_values("등록일자", ascending=False), use_container_width=True)
    else:
        st.info("등록된 활동 내역이 없습니다.")

# --- 탭 4: 담당자 알림 ---
with tab4:
    st.subheader("📨 담당자 이메일 발송")
    
    # 알림 대상 (비매칭 건이 있는 담당자 등 로직 구현 가능)
    st.write("담당자별 VOC 건수 확인 및 알림 발송 기능입니다.")
    
    if not voc_filtered_global.empty:
        mgr_counts = voc_filtered_global.groupby("구역담당자_통합").size().reset_index(name="건수")
        st.dataframe(mgr_counts, use_container_width=True)
        
        selected_mgr = st.selectbox("발송 대상 담당자 선택", ["(선택)"] + mgr_counts["구역담당자_통합"].tolist())
        
        if selected_mgr != "(선택)":
            mgr_email = manager_contacts.get(selected_mgr, {}).get("email", "")
            email_input = st.text_input("이메일 주소", value=mgr_email)
            
            if st.button("이메일 발송 테스트"):
                st.info("이메일 발송 기능은 SMTP 설정이 필요합니다. (코드 내 secrets 설정 확인)")
                # 실제 발송 로직은 secrets 설정 후 활성화
    else:
        st.info("데이터가 없습니다.")
