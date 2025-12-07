# ===========================================================
# PART 1 — 기본 설정 / Apple Glass UI 스타일 / 데이터 로딩 + 전처리
# ===========================================================

import os
from datetime import datetime, date
import smtplib
from email.message import EmailMessage

import numpy as np
import pandas as pd
import streamlit as st


# -----------------------------------------------------------
# 0. Plotly 로딩 (없어도 앱 실행 가능)
# -----------------------------------------------------------
try:
    import plotly.express as px
    HAS_PLOTLY = True
except:
    HAS_PLOTLY = False


# -----------------------------------------------------------
# 1. Streamlit 페이지 설정 (Apple Glass Mode)
# -----------------------------------------------------------
st.set_page_config(
    page_title="해지 VOC 종합 대시보드",
    layout="wide",
)

st.markdown("""
<style>
html, body, .stApp {
    background: #f5f7fa !important;
    font-family: -apple-system, BlinkMacSystemFont, "Inter", sans-serif;
}

/* Glass Surface */
.glass-card {
    background: rgba(255, 255, 255, 0.55);
    backdrop-filter: blur(12px);
    -webkit-backdrop-filter: blur(12px);
    border-radius: 18px;
    padding: 18px;
    border: 1px solid rgba(255,255,255,0.4);
    box-shadow: 0 4px 16px rgba(0,0,0,0.05);
    margin-bottom: 12px;
}

/* KPI */
.kpi-card {
    background: rgba(255,255,255,0.75);
    backdrop-filter: blur(10px);
    border-radius: 16px;
    padding: 15px;
    border: 1px solid rgba(255,255,255,0.45);
    box-shadow: 0 2px 12px rgba(0,0,0,0.06);
}

/* Sidebar Glass */
section[data-testid="stSidebar"] {
    background: rgba(255,255,255,0.55) !important;
    backdrop-filter: blur(12px);
    border-right: 1px solid rgba(255,255,255,0.4);
}

/* Plotly 투명 배경 */
.js-plotly-plot .plotly {
    background-color: transparent !important;
}
</style>
""", unsafe_allow_html=True)


# -----------------------------------------------------------
# 2. SMTP 환경변수 로딩
# -----------------------------------------------------------
if "SMTP_HOST" in st.secrets:
    SMTP_HOST = st.secrets["SMTP_HOST"]
    SMTP_PORT = int(st.secrets["SMTP_PORT"])
    SMTP_USER = st.secrets["SMTP_USER"]
    SMTP_PASSWORD = st.secrets["SMTP_PASSWORD"]
    SENDER_NAME = st.secrets["SENDER_NAME"]
else:
    # 로컬 개발자용 (Cloud에서는 secrets 사용)
    from dotenv import load_dotenv
    load_dotenv()

    SMTP_HOST = os.getenv("SMTP_HOST", "")
    SMTP_PORT = int(os.getenv("SMTP_PORT", "587"))
    SMTP_USER = os.getenv("SMTP_USER", "")
    SMTP_PASSWORD = os.getenv("SMTP_PASSWORD", "")
    SENDER_NAME = os.getenv("SENDER_NAME", "해지VOC 관리자")


# -----------------------------------------------------------
# 3. 주요 파일 경로 설정
# -----------------------------------------------------------
MERGED_PATH = "merged.xlsx"
FEEDBACK_PATH = "feedback.csv"
CONTACT_PATH = "영업구역담당자_251204.xlsx"


# -----------------------------------------------------------
# 4. Utility Functions
# -----------------------------------------------------------
def safe_str(x):
    if pd.isna(x):
        return ""
    return str(x).strip()


def detect_column(df: pd.DataFrame, keys):
    """담당자/이메일 컬럼 자동 탐색"""
    for k in keys:
        if k in df.columns:
            return k
    for col in df.columns:
        for k in keys:
            if k.lower() in col.lower():
                return col
    return None


# -----------------------------------------------------------
# 5. 데이터 로딩
# -----------------------------------------------------------
@st.cache_data
def load_voc_data(path):
    if not os.path.exists(path):
        st.error("❌ merged.xlsx 파일을 찾을 수 없습니다.")
        return pd.DataFrame()

    df = pd.read_excel(path)

    # 계약번호 정제
    if "계약번호" in df.columns:
        df["계약번호"] = (
            df["계약번호"].astype(str).str.replace(",", "").str.strip()
        )
        df["계약번호_정제"] = df["계약번호"].str.replace(r"[^0-9A-Za-z]", "", regex=True)
    else:
        df["계약번호_정제"] = ""

    # 출처 통일
    if "출처" in df.columns:
        df["출처"] = df["출처"].replace({"고객리스트": "해지시설"})

    # 날짜 변환
    if "접수일시" in df.columns:
        df["접수일시"] = pd.to_datetime(df["접수일시"], errors="coerce")

    return df


@st.cache_data
def load_feedback(path):
    """저장된 활동내역 CSV 불러오기"""
    if os.path.exists(path):
        try:
            return pd.read_csv(path, encoding="utf-8-sig")
        except:
            return pd.read_csv(path)
    return pd.DataFrame(columns=["계약번호_정제", "고객대응내용", "등록자", "등록일자", "비고"])


def save_feedback(path, df):
    df.to_csv(path, index=False, encoding="utf-8-sig")


@st.cache_data
def load_contact_map(path):
    """영업구역 담당자 매핑 파일 로딩"""
    if not os.path.exists(path):
        st.warning("⚠ 담당자 매핑 파일이 없습니다.")
        return pd.DataFrame(), {}

    df = pd.read_excel(path)

    name_col = detect_column(df, ["담당자", "구역담당자", "성명"])
    email_col = detect_column(df, ["이메일", "email"])

    if not name_col or not email_col:
        st.warning("⚠ 담당자/이메일 컬럼을 찾지 못했습니다.")
        return df, {}

    df = df[[name_col, email_col]].copy()
    df.columns = ["담당자", "이메일"]

    mapping = {
        safe_str(r["담당자"]): {"email": safe_str(r["이메일"])}
        for _, r in df.iterrows()
        if safe_str(r["담당자"]) != ""
    }

    return df, mapping


# -----------------------------------------------------------
# 6. 실제 데이터 로딩 실행
# -----------------------------------------------------------
df = load_voc_data(MERGED_PATH)
if df.empty:
    st.stop()

if "feedback_df" not in st.session_state:
    st.session_state["feedback_df"] = load_feedback(FEEDBACK_PATH)

contact_df, manager_contacts = load_contact_map(CONTACT_PATH)


# -----------------------------------------------------------
# 7. 전처리 (지사명 축약)
# -----------------------------------------------------------
if "관리지사" in df.columns:
    df["관리지사"] = df["관리지사"].replace({
        "중앙지사":"중앙", "강북지사":"강북", "서대문지사":"서대문", "고양지사":"고양",
        "의정부지사":"의정부", "남양주지사":"남양주", "강릉지사":"강릉", "원주지사":"원주"
    })
else:
    df["관리지사"] = ""


BRANCH_ORDER = ["중앙", "강북", "서대문", "고양", "의정부", "남양주", "강릉", "원주"]

def sort_branch(values):
    return [b for b in BRANCH_ORDER if b in values]


# -----------------------------------------------------------
# PART 1 끝 — 다음 PART 2에서 전처리/리스크/주소/월정료/매칭 로직이 이어짐
# -----------------------------------------------------------

# ===========================================================
# PART 2 — VOC 전처리 / 주소·월정료 통합 / 매칭 판정 / 리스크 등급 계산
# ===========================================================

# -----------------------------------------------------------
# 1) 영업구역 / 담당자 통합
# -----------------------------------------------------------
def pick_zone(r):
    for c in ["영업구역번호", "담당상세", "영업구역정보"]:
        if c in r and pd.notna(r[c]):
            return r[c]
    return ""

df["영업구역_통합"] = df.apply(pick_zone, axis=1)


def pick_manager(r):
    for c in ["구역담당자", "담당자", "처리자"]:
        if c in r and pd.notna(r[c]) and safe_str(r[c]) != "":
            return r[c]
    return ""

df["구역담당자_통합"] = df.apply(pick_manager, axis=1)


# -----------------------------------------------------------
# 2) 출처 분리 + 매칭 여부 계산
# -----------------------------------------------------------
df_voc = df[df.get("출처") == "해지VOC"].copy()
df_other = df[df.get("출처") != "해지VOC"].copy()

# 매칭 기준 계약번호 Set
other_contracts = set(df_other["계약번호_정제"].dropna().unique().tolist())

df_voc["매칭여부"] = df_voc["계약번호_정제"].apply(
    lambda x: "매칭(O)" if x in other_contracts else "비매칭(X)"
)


# -----------------------------------------------------------
# 3) 주소 통합 (시설_설치주소 → 설치주소 → None 제거)
# -----------------------------------------------------------
def merge_addr(r):
    for c in ["시설_설치주소", "설치주소"]:
        if c in r and pd.notna(r[c]) and safe_str(r[c]) not in ["", "None", "nan"]:
            return r[c]
    return np.nan

df_voc["설치주소_표시"] = df_voc.apply(merge_addr, axis=1)

address_cols = [c for c in df.columns if "주소" in c]  # 검색 필터용 주소 컬럼 목록


# -----------------------------------------------------------
# 4) 월정료 정제 (문자제거 → 숫자만 → 10배 오류보정 → 천단위 표시)
# -----------------------------------------------------------
fee_col = None
for c in ["시설_KTT월정료(조정)", "KTT월정료(조정)", "월정료"]:
    if c in df_voc.columns:
        fee_col = c
        break

def parse_fee(v):
    if pd.isna(v):
        return np.nan

    s = "".join(ch for ch in str(v) if ch.isdigit())
    if s == "":
        return np.nan

    f = float(s)

    # ✔ 55,000 → 55000 정상
    # ✔ 550,000 → 55,000 으로 자동 교정
    if f >= 200000:  
        f = f / 10

    return f

if fee_col:
    df_voc["월정료_수치"] = df_voc[fee_col].apply(parse_fee)
else:
    df_voc["월정료_수치"] = np.nan


# 구간(10만 미만 / 이상)
def fee_band(v):
    if pd.isna(v):
        return "미기재"
    return "10만 이상" if v >= 100000 else "10만 미만"

df_voc["월정료구간"] = df_voc["월정료_수치"].apply(fee_band)

# 천단위 표시용
if fee_col:
    df_voc["월정료_표시"] = df_voc["월정료_수치"].apply(
        lambda v: "" if pd.isna(v) else f"{int(v):,}"
    )


# -----------------------------------------------------------
# 5) 리스크 등급 계산 (경과일수 기준)
# -----------------------------------------------------------
today = date.today()

def calc_risk(r):
    dt = r.get("접수일시")
    if pd.isna(dt):
        return np.nan, "LOW"

    if isinstance(dt, datetime):
        dt = dt.date()

    diff = (today - dt).days

    # 리스크 레벨
    if diff <= 3:
        lv = "HIGH"
    elif diff <= 10:
        lv = "MEDIUM"
    else:
        lv = "LOW"

    return diff, lv

df_voc["경과일수"], df_voc["리스크등급"] = zip(*df_voc.apply(calc_risk, axis=1))

# 비매칭만 별도 저장
df_unmatched = df_voc[df_voc["매칭여부"] == "비매칭(X)"].copy()


# -----------------------------------------------------------
# 6) 전체가 사용하는 핵심 display 컬럼 구성
# -----------------------------------------------------------
display_cols = [
    "계약번호_정제", "상호", "관리지사", "구역담당자_통합",
    "VOC유형", "VOC유형소", "등록내용",
    "설치주소_표시",
    "리스크등급", "경과일수",
    "월정료_표시",
]

display_cols = [c for c in display_cols if c in df_voc.columns]


# -----------------------------------------------------------
# 7) 리스크 색상 적용 함수
# -----------------------------------------------------------
def style_risk(df_data):
    def color_row(row):
        if row.get("리스크등급") == "HIGH":
            return ["background-color: #ffe2e2"] * len(row)
        elif row.get("리스크등급") == "MEDIUM":
            return ["background-color: #fff6da"] * len(row)
        else:
            return ["background-color: #e8f6ff"] * len(row)
    return df_data.style.apply(color_row, axis=1)

# ===========================================================
# PART 3 — 전체 UI / 글로벌 필터 / KPI 카드 / 시각화
# (Apple Glass UI + 기업용 레이아웃 + 반응형)
# ===========================================================

import streamlit as st

# -----------------------------------------------------------
# 1) 고급 UI 스타일 (Apple Glass + Material Hybrid)
# -----------------------------------------------------------
st.markdown("""
<style>

html, body, .stApp {
    background:#f5f5f7 !important;
}

/* 본문 패딩 */
.block-container {
    padding-top:0.8rem !important;
}

/* KPI 카드 */
.kpi-card {
    background:#ffffffcc;
    padding:1rem 1.2rem;
    border-radius:16px;
    border:1px solid #e2e3e7;
    backdrop-filter: blur(12px) saturate(180%);
    box-shadow:0 8px 24px rgba(0,0,0,0.06);
}

/* 지사 그리드 */
.branch-grid {
    display:grid;
    grid-template-columns:repeat(4, minmax(0,1fr));
    gap:14px;
}
@media (max-width:1200px){
    .branch-grid { grid-template-columns:repeat(2, minmax(0,1fr)); }
}
@media (max-width:700px){
    .branch-grid { grid-template-columns:repeat(1, minmax(0,1fr)); }
}

.branch-item {
    background:#ffffffcc;
    padding:1rem;
    border-radius:14px;
    border:1px solid #e5e7eb;
    text-align:center;
    font-weight:600;
    backdrop-filter: blur(10px);
    transition: transform 0.12s ease;
}
.branch-item:hover {
    transform: translateY(-3px);
    box-shadow:0 6px 20px rgba(0,0,0,0.08);
}
</style>
""", unsafe_allow_html=True)

# -----------------------------------------------------------
# 2) 사이드바 — 글로벌 필터
# -----------------------------------------------------------
st.sidebar.title("🔧 글로벌 필터")

# 날짜 범위
if "접수일시" in df_voc.columns and df_voc["접수일시"].notna().any():
    d_min = df_voc["접수일시"].min().date()
    d_max = df_voc["접수일시"].max().date()
    dr = st.sidebar.date_input(
        "📅 접수일자 범위",
        value=(d_min, d_max),
        min_value=d_min,
        max_value=d_max,
        key="flt_date"
    )
else:
    dr = None

# 지사 선택
branch_all = sort_branch(df_voc["관리지사"].dropna().unique())
sel_branches = st.sidebar.multiselect(
    "🏢 관리지사",
    options=branch_all,
    default=branch_all,
    key="flt_branch"
)

# 리스크 등급
risk_opts = ["HIGH", "MEDIUM", "LOW"]
sel_risk = st.sidebar.multiselect(
    "⚠ 리스크 등급",
    risk_opts,
    default=risk_opts,
    key="flt_risk"
)

# 매칭
sel_match = st.sidebar.multiselect(
    "🔗 매칭 여부",
    ["매칭(O)", "비매칭(X)"],
    default=["매칭(O)", "비매칭(X)"],
    key="flt_match"
)

# 월정료
fee_global = st.sidebar.radio(
    "💰 월정료 구간",
    ["전체", "10만 미만", "10만 이상"],
    index=0,
    key="flt_fee"
)

st.sidebar.markdown("---")
st.sidebar.caption("※ 이 필터는 전체 탭에 공통 적용됩니다.")


# -----------------------------------------------------------
# 3) 글로벌 필터 적용
# -----------------------------------------------------------
voc_filtered = df_voc.copy()

# 날짜
if dr:
    start, end = dr
    voc_filtered = voc_filtered[
        (voc_filtered["접수일시"] >= pd.to_datetime(start)) &
        (voc_filtered["접수일시"] < pd.to_datetime(end) + pd.Timedelta(days=1))
    ]

# 지사
voc_filtered = voc_filtered[voc_filtered["관리지사"].isin(sel_branches)]

# 리스크
voc_filtered = voc_filtered[voc_filtered["리스크등급"].isin(sel_risk)]

# 매칭
voc_filtered = voc_filtered[voc_filtered["매칭여부"].isin(sel_match)]

# 월정료
if fee_global == "10만 이상":
    voc_filtered = voc_filtered[voc_filtered["월정료_수치"] >= 100000]
elif fee_global == "10만 미만":
    voc_filtered = voc_filtered[
        voc_filtered["월정료_수치"].notna() &
        (voc_filtered["월정료_수치"] < 100000)
    ]

# 비매칭 subset
unmatched_filtered = voc_filtered[voc_filtered["매칭여부"] == "비매칭(X)"]

# -----------------------------------------------------------
# 4) KPI 카드 출력 (Apple KPI UI)
# -----------------------------------------------------------
st.markdown("## 📊 해지 VOC 종합 대시보드 (Enterprise Edition)")

k1, k2, k3, k4 = st.columns(4)

with k1:
    st.markdown("<div class='kpi-card'>", unsafe_allow_html=True)
    st.metric("총 VOC 건수", f"{len(voc_filtered):,}")
    st.markdown("</div>", unsafe_allow_html=True)

with k2:
    st.markdown("<div class='kpi-card'>", unsafe_allow_html=True)
    st.metric("계약번호 수", f"{voc_filtered['계약번호_정제'].nunique():,}")
    st.markdown("</div>", unsafe_allow_html=True)

with k3:
    st.markdown("<div class='kpi-card'>", unsafe_allow_html=True)
    st.metric("비매칭 시설", f"{unmatched_filtered['계약번호_정제'].nunique():,}")
    st.markdown("</div>", unsafe_allow_html=True)

with k4:
    st.markdown("<div class='kpi-card'>", unsafe_allow_html=True)
    st.metric("매칭 시설", f"{voc_filtered[voc_filtered['매칭여부']=='매칭(O)']['계약번호_정제'].nunique():,}")
    st.markdown("</div>", unsafe_allow_html=True)


st.markdown("---")


# -----------------------------------------------------------
# 5) 탭 구성 (시각화, 전체 VOC, 비매칭, 상세, 알림)
# -----------------------------------------------------------
tab_viz, tab_all, tab_unmatched, tab_drill, tab_alert = st.tabs([
    "📊 지사/담당자 시각화",
    "📘 VOC 전체",
    "🧯 해지방어 활동시설",
    "🔍 계약별 상세",
    "📨 담당자 알림"
])

# ============================================================
# PART 4 — 계약별 상세 조회 + VOC 이력 + 활동등록(피드백)
# (KEY 충돌 Zero / 직관적 고급 UI / 데이터 정합성 강화)
# ============================================================

with tab_drill:
    st.subheader("🔍 계약별 상세 조회 + 처리내역(활동등록)")

    df_d = voc_filtered.copy()
    cn_list = sorted(df_d["계약번호_정제"].dropna().unique().tolist())

    # ---------------------------
    # ① 계약번호 선택
    # ---------------------------
    sel_cn = st.selectbox(
        "계약번호 선택",
        ["(선택)"] + cn_list,
        key="drill_sel_contract"
    )

    if sel_cn != "(선택)":

        # 해당 계약 전체 VOC 조회
        voc_hist = df_voc[df_voc["계약번호_정제"] == sel_cn].sort_values(
            "접수일시", ascending=False
        )

        other_hist = df_other[df_other["계약번호_정제"] == sel_cn]

        base = voc_hist.iloc[0]

        # ---------------------------
        # ② 상단 요약 카드
        # ---------------------------
        st.markdown("### 📌 계약 요약 정보")

        c1, c2, c3, c4 = st.columns(4)
        c1.metric("상호", base.get("상호", ""))
        c2.metric("관리지사", base.get("관리지사", ""))
        c3.metric("담당자", base.get("구역담당자_통합", ""))
        c4.metric("VOC 접수건수", len(voc_hist))

        st.caption(f"📍 설치주소: {base.get('설치주소_표시','')}")
        if fee_col:
            st.caption(f"💰 월정료: {base.get('월정료_표시','')}")

        st.markdown("---")

        # ---------------------------
        # ③ VOC 상세 이력
        # ---------------------------
        st.markdown("### 📘 VOC 상세 이력")

        view_cols = [
            "계약번호_정제",
            "상호",
            "VOC유형",
            "VOC유형소",
            "등록내용",
            "리스크등급",
            "경과일수",
            "설치주소_표시",
            "접수일시",
            "구역담당자_통합"
        ]
        view_cols = [c for c in view_cols if c in voc_hist.columns]

        st.dataframe(
            voc_hist[view_cols],
            use_container_width=True,
            height=360
        )

        st.markdown("---")

        # ---------------------------
        # ④ 기타 출처 이력
        # ---------------------------
        st.markdown("### 📦 기타 출처(Customer List·해지시설 등)")

        if other_hist.empty:
            st.info("기타 출처 이력이 없습니다.")
        else:
            st.dataframe(
                other_hist,
                use_container_width=True,
                height=260
            )

        st.markdown("---")

        # ---------------------------
        # ⑤ 활동등록(피드백) 시스템
        # ---------------------------
        st.markdown("## 📝 처리내역 등록")

        fb_all = st.session_state["feedback_df"]
        fb_sel = fb_all[fb_all["계약번호_정제"] == sel_cn].sort_values(
            "등록일자", ascending=False
        )

        # 기존 등록 이력
        st.markdown("### 📄 기존 등록 이력")

        if fb_sel.empty:
            st.info("등록된 처리내역이 없습니다.")
        else:
            st.dataframe(
                fb_sel,
                use_container_width=True,
                height=260
            )

        st.markdown("### ➕ 신규 처리내역 등록")

        new_content = st.text_area("고객대응 내용 입력", key="fb_new_content")
        new_writer = st.text_input("등록자", key="fb_new_writer")
        new_note = st.text_input("비고(선택)", key="fb_new_note")

        if st.button("등록하기", key="fb_add_button"):
            if not new_content or not new_writer:
                st.warning("⚠ 처리내용과 등록자는 반드시 입력해야 합니다.")
            else:
                new_row = {
                    "계약번호_정제": sel_cn,
                    "고객대응내용": new_content,
                    "등록자": new_writer,
                    "등록일자": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                    "비고": new_note
                }

                fb_all = pd.concat([fb_all, pd.DataFrame([new_row])], ignore_index=True)
                st.session_state["feedback_df"] = fb_all
                save_feedback(FEEDBACK_PATH, fb_all)

                st.success("✅ 처리내역이 성공적으로 등록되었습니다.")
                st.rerun()

# ============================================================
# PART 5 — 담당자 알림 발송(기업용 고도화 버전)
# ============================================================

with tab_alert:
    st.subheader("📨 담당자 알림 발송 (기업용 버전)")

    st.markdown(
        """
        비매칭(X) 시설을 담당자에게 이메일로 자동 안내합니다.<br>
        계약번호는 중복 없이 **최신 VOC 기준 1건으로 요약**되어 CSV로 첨부됩니다.
        """,
        unsafe_allow_html=True
    )

    # -----------------------------------------------------------
    # 1) 비매칭 → 담당자 매핑 테이블 생성
    # -----------------------------------------------------------
    unmatched_alert = unmatched_filtered.copy()

    alert_rows = []
    for mgr, g in unmatched_alert.groupby("구역담당자_통합"):
        if not mgr or str(mgr).strip() == "":
            continue
        cnt = g["계약번호_정제"].nunique()
        email = manager_contacts.get(mgr, {}).get("email", "")
        alert_rows.append([mgr, email, cnt])

    alert_df = pd.DataFrame(
        alert_rows, columns=["담당자", "이메일", "비매칭 계약수"]
    )

    st.markdown("### 👤 담당자별 비매칭 현황")
    st.dataframe(alert_df, use_container_width=True, height=250)

    st.markdown("---")

    # -----------------------------------------------------------
    # 2) 담당자 선택
    # -----------------------------------------------------------
    sel_mgr = st.selectbox(
        "알림을 보낼 담당자 선택",
        ["(선택)"] + alert_df["담당자"].tolist(),
        key="alert_mgr"
    )

    if sel_mgr != "(선택)":

        # 기본 이메일 자동 입력
        default_email = manager_contacts.get(sel_mgr, {}).get("email", "")
        email_input = st.text_input("수신 이메일 주소", value=default_email, key="alert_email")

        # 선택 담당자 데이터 필터
        df_mgr = unmatched_alert[unmatched_alert["구역담당자_통합"] == sel_mgr]

        if df_mgr.empty:
            st.info("📭 해당 담당자는 비매칭 시설이 없습니다.")
            st.stop()

        # -----------------------------------------------------------
        # 3) 상세 테이블 표시
        # -----------------------------------------------------------
        st.markdown(f"### 🔍 {sel_mgr} 담당자 비매칭 시설 목록")

        disp_cols = [
            "계약번호_정제",
            "상호",
            "관리지사",
            "VOC유형",
            "VOC유형소",
            "등록내용",
            "리스크등급",
            "경과일수",
            "설치주소_표시",
        ]
        disp_cols = [c for c in disp_cols if c in df_mgr.columns]

        st.dataframe(df_mgr[disp_cols], use_container_width=True, height=350)

        # -----------------------------------------------------------
        # 4) CSV 생성 (중복 제거 + 최신 VOC 1건)
        # -----------------------------------------------------------
        df_sorted = df_mgr.sort_values("접수일시", ascending=False)
        grp = df_sorted.groupby("계약번호_정제")
        latest_idx = grp["접수일시"].idxmax()

        df_latest = df_sorted.loc[latest_idx].copy()

        df_latest = df_latest[
            [
                "계약번호_정제",
                "상호",
                "관리지사",
                "구역담당자_통합",
                "VOC유형",
                "VOC유형소",
                "등록내용",
                "설치주소_표시",
                "리스크등급",
                "경과일수"
            ]
        ]

        csv_bytes = df_latest.to_csv(index=False, encoding="utf-8-sig").encode("utf-8-sig")

        st.success(f"📁 첨부 CSV 생성 완료 — {len(df_latest)}건")

        # -----------------------------------------------------------
        # 5) 이메일 본문 자동 구성
        # -----------------------------------------------------------
        subject = f"[해지VOC] {sel_mgr} 담당자 비매칭 시설 안내"

        body = (
            f"{sel_mgr} 담당자님,\n\n"
            f"현재 담당 구역에 총 {len(df_latest)}건의 비매칭 시설이 확인되었습니다.\n"
            f"첨부된 CSV 파일을 확인하시어 빠른 처리 부탁드립니다.\n\n"
            "— 해지VOC 관리자 드림 —"
        )

        # -----------------------------------------------------------
        # 6) 이메일 발송
        # -----------------------------------------------------------
        if st.button("📤 이메일 발송하기", key="alert_send"):
            try:
                msg = EmailMessage()
                msg["Subject"] = subject
                msg["From"] = f"{SENDER_NAME} <{SMTP_USER}>"
                msg["To"] = email_input
                msg.set_content(body)

                # 첨부파일 추가
                msg.add_attachment(
                    csv_bytes,
                    maintype="application",
                    subtype="octet-stream",
                    filename=f"비매칭계약_{sel_mgr}.csv"
                )

                # SMTP 발송
                with smtplib.SMTP(SMTP_HOST, SMTP_PORT) as smtp:
                    smtp.starttls()
                    smtp.login(SMTP_USER, SMTP_PASSWORD)
                    smtp.send_message(msg)

                st.success(f"✅ 이메일 발송 완료 → {email_input}")

            except Exception as e:
                st.error(f"❌ 이메일 전송 실패: {e}")

    st.markdown("---")
    st.caption("※ 비매칭(X): 해지VOC 접수 후 매칭 이력이 존재하지 않는 계약입니다.")
