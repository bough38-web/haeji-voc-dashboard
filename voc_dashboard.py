import os
from datetime import datetime, date

import numpy as np
import pandas as pd
import streamlit as st
import smtplib
from email.message import EmailMessage

# ====================================================
# 0. 기본 설정 & 스타일
# ====================================================
st.set_page_config(page_title="해지VOC 담당자 안내 및 알림", layout="wide")

st.markdown(
    """
    <style>
    .stApp {
        background-color: #f3f4f6;
        color: #111827;
        font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif;
    }

    .block-container {
        padding-top: 0.8rem;
        padding-bottom: 3rem;
        padding-left: 1.5rem;
        padding-right: 1.5rem;
    }

    [data-testid="stHeader"] {
        background-color: #f3f4f6;
    }

    section[data-testid="stSidebar"] {
        background-color: #f9fafb;
        border-right: 1px solid #e5e7eb;
    }
    section[data-testid="stSidebar"] .block-container {
        padding-top: 1.2rem;
    }

    h2, h3, h4 {
        margin-top: 0.4rem;
        margin-bottom: 0.4rem;
        font-weight: 600;
    }

    .dataframe tbody tr:nth-child(odd) {
        background-color: #f9fafb;
    }
    .dataframe tbody tr:nth-child(even) {
        background-color: #eef2ff;
    }

    textarea, input, select {
        border-radius: 8px !important;
    }

    .section-card {
        background: #ffffff;
        border-radius: 16px;
        padding: 1.0rem 1.2rem;
        border: 1px solid #e5e7eb;
        box-shadow: 0 4px 8px rgba(15, 23, 42, 0.04);
        margin-bottom: 1.2rem;
    }

    .section-title {
        font-size: 1.05rem;
        font-weight: 600;
        margin-bottom: 0.6rem;
        display: flex;
        align-items: center;
        gap: 0.25rem;
    }

    .help-text {
        font-size: 0.85rem;
        color: #4b5563;
        margin-top: 0.2rem;
    }

    .metric-badge {
        font-size: 0.8rem;
        padding: 0.15rem 0.4rem;
        border-radius: 999px;
        background-color: #eef2ff;
        color: #4f46e5;
        display: inline-block;
        margin-left: 0.4rem;
    }

    .email-preview {
        background-color: #0f172a;
        color: #e5e7eb;
        border-radius: 12px;
        padding: 0.75rem 0.9rem;
        font-family: ui-monospace, SFMono-Regular, Menlo, Monaco, Consolas, "Liberation Mono", "Courier New", monospace;
        font-size: 0.8rem;
        margin-top: 0.4rem;
        white-space: pre-line;
    }
    </style>
    """,
    unsafe_allow_html=True,
)

# ====================================================
# 1. 파일 경로 & SMTP 설정
# ====================================================
VOC_PATH = "merged.xlsx"
MANAGER_MAP_PATH = "영업구역담당자_251204.xlsx"

SMTP_HOST = "smtp.gmail.com"
SMTP_PORT = 587
SMTP_USER = "bough38@gmail.com"
SENDER_NAME = "해지VOC 관리자"

# 비밀번호는 반드시 st.secrets 또는 환경변수에 보관 (코드에 직접 쓰지 말 것!)
SMTP_PASSWORD = st.secrets.get("SMTP_PASSWORD") or os.getenv("SMTP_PASSWORD", "")

# ====================================================
# 2. 데이터 로딩 함수
# ====================================================
@st.cache_data
def load_voc_data(path: str) -> pd.DataFrame:
    if not os.path.exists(path):
        st.error(f"❌ '{path}' 파일이 존재하지 않습니다. 저장소 루트 위치를 다시 확인해주세요.")
        return pd.DataFrame()

    df = pd.read_excel(path)

    for col in ["계약번호", "고객번호"]:
        if col in df.columns:
            df[col] = (
                df[col]
                .astype(str)
                .str.replace(",", "", regex=False)
                .str.strip()
            )

    if "출처" in df.columns:
        df["출처"] = df["출처"].replace({"고객리스트": "해지시설"})

    if "계약번호" in df.columns:
        df["계약번호_정제"] = (
            df["계약번호"]
            .astype(str)
            .str.replace(r"[^0-9A-Za-z]", "", regex=True)
            .str.strip()
        )
    else:
        df["계약번호_정제"] = ""

    if "접수일시" in df.columns:
        df["접수일시"] = pd.to_datetime(df["접수일시"], errors="coerce")

    # 지사 축약
    if "관리지사" in df.columns:
        df["관리지사"] = df["관리지사"].replace(
            {
                "중앙지사": "중앙",
                "강북지사": "강북",
                "서대문지사": "서대문",
                "고양지사": "고양",
                "의정부지사": "의정부",
                "남양주지사": "남양주",
                "강릉지사": "강릉",
                "원주지사": "원주",
            }
        )
    else:
        df["관리지사"] = ""

    # 영업구역 / 담당자 통합
    def make_zone(row):
        if "영업구역번호" in row and pd.notna(row["영업구역번호"]):
            return row["영업구역번호"]
        if "담당상세" in row and pd.notna(row["담당상세"]):
            return row["담당상세"]
        if "영업구역정보" in row and pd.notna(row["영업구역정보"]):
            return row["영업구역정보"]
        return ""

    df["영업구역_통합"] = df.apply(make_zone, axis=1)

    mgr_priority = ["구역담당자", "담당자", "처리자"]

    def pick_manager(row):
        for c in mgr_priority:
            if c in row and pd.notna(row[c]) and str(row[c]).strip() != "":
                return row[c]
        return ""

    df["구역담당자_통합"] = df.apply(pick_manager, axis=1)

    # 설치주소 표시용
    def coalesce_cols(row, candidates):
        for c in candidates:
            if c in row.index:
                val = row[c]
                if pd.notna(val) and str(val).strip() not in ["", "None", "nan"]:
                    return val
        return np.nan

    df["설치주소_표시"] = df.apply(
        lambda r: coalesce_cols(r, ["시설_설치주소", "설치주소"]),
        axis=1,
    )

    return df


@st.cache_data
def load_manager_map(path: str) -> pd.DataFrame:
    if not os.path.exists(path):
        st.error(f"❌ 담당자 매핑 파일 '{path}' 이(가) 존재하지 않습니다.")
        return pd.DataFrame()

    df = pd.read_excel(path)

    # 실제 파일 컬럼 기준으로 정리
    # 예상 컬럼: 처리자1, 담당상세, 소속, 연략처, E-MAIL
    rename_map = {
        "처리자1": "담당자명",
        "담당상세": "영업구역번호",
        "소속": "관리지사",
        "연략처": "연락처",
        "E-MAIL": "이메일",
    }
    for old, new in rename_map.items():
        if old in df.columns:
            df.rename(columns={old: new}, inplace=True)

    # 필수 컬럼 체크
    for c in ["담당자명", "이메일"]:
        if c not in df.columns:
            st.error(f"❌ 담당자 매핑 파일에 '{c}' 컬럼이 없습니다. 엑셀 헤더를 확인해주세요.")
            return pd.DataFrame()

    # 이메일/담당자명 공백 제거
    df["담당자명"] = df["담당자명"].astype(str).str.strip()
    df["이메일"] = df["이메일"].astype(str).str.strip()

    df = df[(df["담당자명"] != "") & (df["이메일"] != "")]
    df = df.drop_duplicates(subset=["담당자명", "이메일"])

    if "관리지사" not in df.columns:
        df["관리지사"] = ""

    return df


# ====================================================
# 3. VOC 데이터 가공 (해지VOC + 비매칭)
# ====================================================
BRANCH_ORDER = ["중앙", "강북", "서대문", "고양", "의정부", "남양주", "강릉", "원주"]

def sort_branch(series):
    return sorted(
        [s for s in series if s in BRANCH_ORDER],
        key=lambda x: BRANCH_ORDER.index(x),
    )


@st.cache_data
def prepare_voc(df: pd.DataFrame):
    if df.empty:
        return pd.DataFrame(), pd.DataFrame()

    df_voc = df[df.get("출처") == "해지VOC"].copy()
    df_other = df[df.get("출처") != "해지VOC"].copy()

    other_sets = {
        src: set(df_other[df_other["출처"] == src]["계약번호_정제"].dropna())
        for src in ["해지시설", "해지요청", "설변", "정지", "해지파이프라인"]
        if "출처" in df_other.columns
    }
    other_union = set().union(*other_sets.values()) if other_sets else set()

    df_voc["매칭여부"] = df_voc["계약번호_정제"].apply(
        lambda x: "매칭(O)" if x in other_union else "비매칭(X)"
    )

    # 리스크 계산
    today = date.today()

    def compute_risk(row):
        dt = row.get("접수일시")
        if pd.isna(dt):
            return np.nan, "LOW"
        if not isinstance(dt, (pd.Timestamp, datetime)):
            try:
                dt = pd.to_datetime(dt, errors="coerce")
            except Exception:
                return np.nan, "LOW"
        if pd.isna(dt):
            return np.nan, "LOW"

        days = (today - dt.date()).days
        if days <= 3:
            level = "HIGH"
        elif days <= 10:
            level = "MEDIUM"
        else:
            level = "LOW"
        return days, level

    df_voc["경과일수"], df_voc["리스크등급"] = zip(
        *df_voc.apply(lambda r: compute_risk(r), axis=1)
    )

    unmatched = df_voc[df_voc["매칭여부"] == "비매칭(X)"].copy()

    return df_voc, unmatched


# ====================================================
# 4. 이메일 발송 함수
# ====================================================
def send_email(to_addrs, subject: str, body: str):
    if not SMTP_PASSWORD:
        st.error("❌ SMTP 비밀번호가 설정되어 있지 않습니다. st.secrets['SMTP_PASSWORD'] 또는 환경변수에 등록해주세요.")
        return False

    if isinstance(to_addrs, str):
        to_addrs = [to_addrs]

    msg = EmailMessage()
    msg["Subject"] = subject
    msg["From"] = f"{SENDER_NAME} <{SMTP_USER}>"
    msg["To"] = ", ".join([a for a in to_addrs if a])

    msg.set_content(body)

    try:
        with smtplib.SMTP(SMTP_HOST, SMTP_PORT) as server:
            server.starttls()
            server.login(SMTP_USER, SMTP_PASSWORD)
            server.send_message(msg)
        return True
    except Exception as e:
        st.error(f"❌ 메일 발송 중 오류가 발생했습니다: {e}")
        return False


# 템플릿 생성
def build_template(template_type: str,
                   branch: str,
                   base_manager: str,
                   num_targets: int,
                   recent_days: int = 7) -> tuple[str, str]:
    """템플릿 유형에 따른 제목/본문 생성"""
    base_title = f"[해지VOC] 중점 해지방어 활동시설 안내 - {branch} / {base_manager}"
    summary_line = f"- 기준 담당자: {base_manager}\n- 지사: {branch}\n- 대상 시설 수: {num_targets}건\n- 기준: 최근 {recent_days}일 해지VOC 중 비매칭(활동내역 미등록) 시설"

    if template_type == "업무용 요약":
        subject = base_title
        body = f"""안녕하세요.

해지VOC 대시보드 기준으로, {branch} 지사 {base_manager} 담당 구역의
'중점 해지방어 활동시설'(해지VOC 접수 후 활동내역이 확인되지 않은 시설)이 {num_targets}건 조회되었습니다.

{summary_line}

각 시설별 상세 내역은 대시보드 또는 공유된 엑셀 파일을 참고하시고,
현황 확인 후 활동내역을 등록해 주시기 바랍니다.

감사합니다.
해지VOC 운영담당 드림
"""
    elif template_type == "중점 방어 안내":
        subject = f"[중점 해지방어] 활동필요 시설 목록 공유 - {branch} / {base_manager}"
        body = f"""안녕하세요.

해지VOC 분석 결과, 아래 기준에 해당하는
'중점 해지방어 활동시설'이 {num_targets}건 확인되었습니다.

{summary_line}

▶ 중점 해지방어 활동시설이란?
해지VOC 접수 후 해지방어 또는 현장 방문 등
'활동내역이 등록되지 않은 시설'로
신속한 확인 및 방어활동이 필요한 대상입니다.

불필요 해지가 발생하지 않도록,
해당 시설 우선 확인 및 활동내역 등록을 요청드립니다.

감사합니다.
해지VOC 운영담당 드림
"""
    else:  # 긴급 재확인 요청
        subject = f"[긴급] 해지방어 활동 미등록 시설 재확인 요청 - {branch} / {base_manager}"
        body = f"""[긴급 안내]

최근 {recent_days}일 이내 접수된 해지VOC 중,
해지방어 활동내역이 등록되지 않은 시설이 {num_targets}건 확인되었습니다.

{summary_line}

특히, 장기간 방치 시 불필요 해지로 연결될 가능성이 있어
지사 차원의 신속한 확인 및 조치가 필요합니다.

가능하신 한 빠른 시일 내에
- 고객 통화 / 방문 여부
- 해지방어 결과
- 후속 조치 계획

등을 확인 후, 대시보드 또는 내부 시스템에
활동내역을 등록해 주시기 바랍니다.

감사합니다.
해지VOC 운영담당 드림
"""
    return subject, body


# ====================================================
# 5. 메인 로직
# ====================================================
st.title("📧 해지VOC 담당자 안내 & 알림 발송")

st.caption(
    "해지VOC 대시보드 기준으로 **중점 해지방어 활동시설** 현황을 확인하고, "
    "담당자 또는 대무자에게 이메일로 안내할 수 있는 화면입니다."
)

# ---------- 데이터 로딩 ----------
voc_raw = load_voc_data(VOC_PATH)
manager_map = load_manager_map(MANAGER_MAP_PATH)

if voc_raw.empty or manager_map.empty:
    st.stop()

df_voc, unmatched_global = prepare_voc(voc_raw)

# ====================================================
# 6. 사이드바: 기준 담당자 선택
# ====================================================
st.sidebar.markdown("### 🎯 기준 담당자 선택")

branches_available = sort_branch(manager_map["관리지사"].dropna().unique())
branch_sel = st.sidebar.selectbox(
    "관리지사",
    options=["(전체)"] + branches_available,
    index=0,
)

if branch_sel == "(전체)":
    mgr_base_df = manager_map.copy()
else:
    mgr_base_df = manager_map[manager_map["관리지사"] == branch_sel]

mgr_names = sorted(mgr_base_df["담당자명"].unique().tolist())

if not mgr_names:
    st.sidebar.info("선택한 지사에 등록된 담당자가 없습니다.")
    base_manager_sel = None
else:
    base_manager_sel = st.sidebar.selectbox(
        "기준 담당자 (데이터 기준)",
        options=mgr_names,
    )

st.sidebar.markdown("---")

st.sidebar.markdown("##### ℹ 중점 해지방어 활동시설 이란?")
st.sidebar.caption(
    "해지VOC 접수 후 활동내역이 확인되지 않은 시설로, "
    "신속한 확인과 해지방어 활동이 필요한 대상입니다."
)

# ====================================================
# 7. 기준 담당자별 대상 데이터 집계
# ====================================================
if base_manager_sel is None:
    st.info("좌측 사이드바에서 기준 지사와 담당자를 먼저 선택해주세요.")
    st.stop()

# VOC 데이터에서 담당자 기준 필터
base_voc = unmatched_global.copy()

if branch_sel != "(전체)":
    base_voc = base_voc[base_voc["관리지사"] == branch_sel]

# '구역담당자_통합' 이 기준 담당자 이름과 같은 건들
base_voc = base_voc[
    base_voc["구역담당자_통합"].astype(str).str.strip() == base_manager_sel
]

# 최근 N일 기준 (예: 30일)
RECENT_DAYS = 30
if "접수일시" in base_voc.columns and base_voc["접수일시"].notna().any():
    cutoff_date = pd.Timestamp.today().normalize() - pd.Timedelta(days=RECENT_DAYS)
    recent_voc = base_voc[base_voc["접수일시"] >= cutoff_date]
else:
    recent_voc = base_voc.copy()

num_total_targets = base_voc["계약번호_정제"].nunique()
num_recent_targets = recent_voc["계약번호_정제"].nunique()

# ====================================================
# 8. 상단 요약 (기준 담당자 기준)
# ====================================================
top_col1, top_col2, top_col3, top_col4 = st.columns(4)

top_col1.metric(
    "기준 지사",
    branch_sel if branch_sel != "(전체)" else "전체",
)
top_col2.metric(
    "기준 담당자",
    base_manager_sel,
)
top_col3.metric(
    "중점 해지방어 활동시설 (전체)",
    f"{num_total_targets:,} 건",
)
top_col4.metric(
    f"중점 해지방어 활동시설 (최근 {RECENT_DAYS}일)",
    f"{num_recent_targets:,} 건",
)

st.markdown("---")

# ====================================================
# 9. 대상 리스트 & 간단 시각화
# ====================================================
st.markdown(
    '<div class="section-card"><div class="section-title">📊 기준 담당자별 중점 해지방어 활동시설 현황</div>',
    unsafe_allow_html=True,
)

if base_voc.empty:
    st.info("현재 기준 조건에 해당하는 '중점 해지방어 활동시설'이 없습니다.")
else:
    # 계약번호 기준 요약 (최신 VOC만)
    temp_sorted = base_voc.sort_values("접수일시", ascending=False)
    grp = temp_sorted.groupby("계약번호_정제")
    idx_latest = grp["접수일시"].idxmax()
    df_summary = temp_sorted.loc[idx_latest].copy()
    df_summary["접수건수"] = grp.size().reindex(df_summary["계약번호_정제"]).values

    list_cols = [
        "계약번호_정제",
        "상호",
        "관리지사",
        "구역담당자_통합",
        "리스크등급",
        "경과일수",
        "설치주소_표시",
        "접수건수",
        "VOC유형소",
    ]
    list_cols = [c for c in list_cols if c in df_summary.columns]

    sub1, sub2 = st.columns([3, 2])

    with sub1:
        st.markdown("##### 📋 대상 시설 목록 (계약번호 기준, 최신 VOC 1건)")
        st.dataframe(
            df_summary[list_cols].sort_values("경과일수", ascending=False),
            use_container_width=True,
            height=360,
        )

    with sub2:
        st.markdown("##### ⏱ 리스크/경과일 분포")

        # 리스크 분포
        risk_counts = (
            df_summary["리스크등급"]
            .value_counts()
            .reindex(["HIGH", "MEDIUM", "LOW"])
            .fillna(0)
        )
        st.bar_chart(risk_counts)

        # 경과일 박스 요약
        if "경과일수" in df_summary.columns:
            days_series = df_summary["경과일수"].dropna()
            if not days_series.empty:
                st.caption(
                    f"경과일수(최소/중앙/최대): "
                    f"{int(days_series.min())}일 / "
                    f"{int(days_series.median())}일 / "
                    f"{int(days_series.max())}일"
                )

st.markdown("</div>", unsafe_allow_html=True)

# ====================================================
# 10. 수신자 선택 & 메일 템플릿
# ====================================================
st.markdown(
    '<div class="section-card"><div class="section-title">✉ 알림 수신자 선택 & 메일 작성</div>',
    unsafe_allow_html=True,
)

# 1) 수신자 선택
col_rcv1, col_rcv2 = st.columns(2)

with col_rcv1:
    st.markdown("##### 1) 수신자 선택")

    all_manager_names = sorted(manager_map["담당자명"].unique().tolist())
    default_receiver = base_manager_sel if base_manager_sel in all_manager_names else all_manager_names[0]

    receiver_name = st.selectbox(
        "알림을 보낼 담당자/대무자 선택",
        options=all_manager_names,
        index=all_manager_names.index(default_receiver),
        help="기본값은 기준 담당자입니다. 필요 시 대무자 등 다른 담당자를 선택할 수 있습니다.",
    )

    # 선택된 사람의 이메일
    receiver_email = ""
    row_match = manager_map[manager_map["담당자명"] == receiver_name]
    if not row_match.empty:
        receiver_email = row_match.iloc[0]["이메일"]

    if receiver_email:
        st.caption(f"📮 받는 이메일: **{receiver_email}**")
    else:
        st.error("선택한 담당자의 이메일 주소를 찾을 수 없습니다. 매핑 엑셀을 확인해주세요.")

with col_rcv2:
    st.markdown("##### 2) 알림 템플릿 선택")
    template_type = st.radio(
        "템플릿 유형",
        options=["업무용 요약", "중점 방어 안내", "긴급 재확인 요청"],
        horizontal=False,
    )
    st.caption(
        "- **업무용 요약**: 일상적인 안내용, 담백한 톤\n"
        "- **중점 방어 안내**: 방어 우선순위 강조\n"
        "- **긴급 재확인 요청**: 리스크가 높을 때 사용하는 긴급 템플릿"
    )

# 템플릿 기반 기본 제목/본문 생성
subject_default, body_default = build_template(
    template_type=template_type,
    branch=branch_sel if branch_sel != "(전체)" else "전체 지사",
    base_manager=base_manager_sel,
    num_targets=num_total_targets,
    recent_days=RECENT_DAYS,
)

# 템플릿 변경 시에만 기본값 갱신
if "tpl_type_prev" not in st.session_state:
    st.session_state["tpl_type_prev"] = template_type
if "email_subject" not in st.session_state:
    st.session_state["email_subject"] = subject_default
if "email_body" not in st.session_state:
    st.session_state["email_body"] = body_default

if st.session_state["tpl_type_prev"] != template_type:
    # 템플릿 바뀌면 기본값으로 다시 채워주기
    st.session_state["email_subject"] = subject_default
    st.session_state["email_body"] = body_default
    st.session_state["tpl_type_prev"] = template_type

st.markdown("##### 3) 메일 제목/본문 작성")

email_subject = st.text_input(
    "메일 제목",
    value=st.session_state["email_subject"],
    key="email_subject",
)

email_body = st.text_area(
    "메일 본문",
    value=st.session_state["email_body"],
    height=260,
    key="email_body",
)

# 미리보기
st.markdown("###### ✨ 발송 미리보기")
st.markdown(
    f"<div class='email-preview'>To: {receiver_email or '[이메일 미등록]'}\n"
    f"Subject: {email_subject}\n\n"
    f"{email_body}</div>",
    unsafe_allow_html=True,
)

st.markdown("</div>", unsafe_allow_html=True)

# ====================================================
# 11. 발송 버튼
# ====================================================
st.markdown(
    '<div class="section-card"><div class="section-title">🚀 메일 발송</div>',
    unsafe_allow_html=True,
)

st.caption(
    "※ 실제 운영에 사용하기 전에, **본인 메일 주소로 테스트 발송**을 반드시 진행한 후 사용해주세요."
)

col_send1, col_send2 = st.columns([1, 3])

with col_send1:
    send_confirm = st.checkbox("위 내용으로 메일 발송에 동의합니다.", value=False)

with col_send2:
    if st.button("✉ 이메일 발송", type="primary"):
        if not receiver_email:
            st.error("수신자 이메일 주소가 없습니다. 담당자 매핑 파일을 확인해주세요.")
        elif not email_subject.strip():
            st.error("메일 제목을 입력해주세요.")
        elif not email_body.strip():
            st.error("메일 본문을 입력해주세요.")
        elif not send_confirm:
            st.warning("발송 동의 체크박스를 먼저 선택해주세요.")
        else:
            ok = send_email(receiver_email, email_subject.strip(), email_body.strip())
            if ok:
                st.success(f"✅ {receiver_name} ({receiver_email}) 님에게 메일이 발송되었습니다.")

st.markdown("</div>", unsafe_allow_html=True)
