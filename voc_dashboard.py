# ===========================================================
# PART 1 — 기본 설정 / UI 스타일 / 환경변수 / 데이터 로딩
# ===========================================================

import os
from datetime import datetime, date
import smtplib
from email.message import EmailMessage

import numpy as np
import pandas as pd
import streamlit as st

# Plotly (옵션)
try:
    import plotly.express as px
    HAS_PLOTLY = True
except:
    HAS_PLOTLY = False


# -----------------------------------------------------------
# 0. Streamlit 페이지 설정
# -----------------------------------------------------------
st.set_page_config(
    page_title="해지 VOC Enterprise Dashboard",
    page_icon="📊",
    layout="wide",
)

# -----------------------------------------------------------
# 1. Apple Glass + Google Material Hybrid UI 스타일 적용
# -----------------------------------------------------------
st.markdown(
    """
    <style>
    /* 전체 배경 – Apple White */
    html, body, .stApp {
        background: #f2f2f7 !important;
        color: #111;
        font-family: -apple-system, BlinkMacSystemFont, 'Inter', sans-serif;
    }

    /* Glass Card – Apple Glass Feel */
    .glass-card {
        background: rgba(255,255,255,0.55);
        border-radius: 18px;
        padding: 1.2rem 1.4rem;
        box-shadow: 0 12px 28px rgba(0,0,0,0.08);
        backdrop-filter: blur(14px);
        -webkit-backdrop-filter: blur(14px);
        border: 1px solid rgba(255,255,255,0.40);
        margin-bottom: 1rem;
    }

    /* 제목 개선 */
    h1, h2, h3 {
        font-weight: 720;
        letter-spacing: -0.02em;
    }

    /* 사이드바 개선 */
    section[data-testid="stSidebar"] {
        background: rgba(255,255,255,0.6) !important;
        backdrop-filter: blur(12px);
    }

    /* KPI 카드 */
    .kpi-card {
        background: rgba(255,255,255,0.7);
        border-radius: 16px;
        padding: 1rem 1.3rem;
        border: 1px solid rgba(255,255,255,0.55);
        box-shadow: 0 8px 22px rgba(0,0,0,0.05);
        backdrop-filter: blur(10px);
    }

    /* 표 라인 */
    .dataframe tbody tr:nth-child(odd) { background: #fafafa; }
    .dataframe tbody tr:nth-child(even) { background: #eef2ff; }

    /* plotly */
    .js-plotly-plot .plotly {
        background-color: transparent !important;
    }

    </style>
    """,
    unsafe_allow_html=True,
)

# -----------------------------------------------------------
# 2. 환경변수 (Secrets 또는 .env 기반)
# -----------------------------------------------------------
if "SMTP_HOST" in st.secrets:
    SMTP_HOST = st.secrets["SMTP_HOST"]
    SMTP_PORT = int(st.secrets["SMTP_PORT"])
    SMTP_USER = st.secrets["SMTP_USER"]
    SMTP_PASSWORD = st.secrets["SMTP_PASSWORD"]
    SENDER_NAME = st.secrets["SENDER_NAME"]
else:
    # 로컬 개발용
    try:
        from dotenv import load_dotenv
        load_dotenv()
    except:
        pass

    SMTP_HOST = os.getenv("SMTP_HOST", "")
    SMTP_PORT = int(os.getenv("SMTP_PORT", "587"))
    SMTP_USER = os.getenv("SMTP_USER", "")
    SMTP_PASSWORD = os.getenv("SMTP_PASSWORD", "")
    SENDER_NAME = os.getenv("SENDER_NAME", "해지VOC 관리자")


# -----------------------------------------------------------
# 3. 파일 경로
# -----------------------------------------------------------
MERGED_PATH = "merged.xlsx"
FEEDBACK_PATH = "feedback.csv"
CONTACT_PATH = "영업구역담당자_251204.xlsx"


# -----------------------------------------------------------
# 4. 공통 함수
# -----------------------------------------------------------
def safe_str(v):
    if pd.isna(v):
        return ""
    return str(v).strip()

def detect_column(df, names):
    """담당자/이메일 자동 탐지"""
    for n in names:  # 완전 일치 우선
        if n in df.columns:
            return n
    for col in df.columns:  # 부분 포함
        for n in names:
            if n.lower() in str(col).lower():
                return col
    return None


# -----------------------------------------------------------
# 5. 데이터 로딩 함수
# -----------------------------------------------------------
@st.cache_data
def load_voc(path):
    if not os.path.exists(path):
        st.error("❌ merged.xlsx 파일이 존재하지 않습니다.")
        return pd.DataFrame()

    df = pd.read_excel(path)

    # 계약번호 정제
    if "계약번호" in df.columns:
        df["계약번호"] = df["계약번호"].astype(str).str.replace(",", "").str.strip()
        df["계약번호_정제"] = df["계약번호"].str.replace(r"[^0-9A-Za-z]", "", regex=True)
    else:
        df["계약번호_정제"] = ""

    # 접수일시 변환
    if "접수일시" in df.columns:
        df["접수일시"] = pd.to_datetime(df["접수일시"], errors="coerce")

    return df


@st.cache_data
def load_feedback(path):
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
    if not os.path.exists(path):
        st.warning("⚠ 담당자 매핑 파일이 없습니다.")
        return pd.DataFrame(), {}

    df = pd.read_excel(path)

    nm_col = detect_column(df, ["담당자", "성명", "구역담당자"])
    em_col = detect_column(df, ["이메일", "email"])

    if not nm_col or not em_col:
        st.warning("⚠ 담당자/이메일 컬럼을 찾을 수 없습니다.")
        return df, {}

    df = df[[nm_col, em_col]].copy()
    df.columns = ["담당자", "이메일"]

    mapping = {safe_str(r["담당자"]): safe_str(r["이메일"]) for _, r in df.iterrows()}

    return df, mapping

# ===========================================================
# PART 2 — VOC 전처리 / 주소·월정료 정제 / 매칭판정 / 리스크 등급
# ===========================================================

# ---------------------------
# 1) VOC 데이터 로딩
# ---------------------------
df = load_voc(MERGED_PATH)
if df.empty:
    st.stop()

# 피드백 로딩
if "feedback_df" not in st.session_state:
    st.session_state["feedback_df"] = load_feedback(FEEDBACK_PATH)

# 담당자 매핑
contact_df, manager_contacts = load_contact_map(CONTACT_PATH)


# ---------------------------
# 2) 지사명 표준화
# ---------------------------
BRANCH_ALIAS = {
    "중앙지사": "중앙",
    "강북지사": "강북",
    "서대문지사": "서대문",
    "고양지사": "고양",
    "의정부지사": "의정부",
    "남양주지사": "남양주",
    "강릉지사": "강릉",
    "원주지사": "원주",
}

if "관리지사" in df.columns:
    df["관리지사"] = df["관리지사"].replace(BRANCH_ALIAS)
else:
    df["관리지사"] = ""

BRANCH_ORDER = ["중앙", "강북", "서대문", "고양", "의정부", "남양주", "강릉", "원주"]

def order_branches(vals):
    return [b for b in BRANCH_ORDER if b in vals]


# ---------------------------
# 3) 영업구역 / 담당자 자동 통합 컬럼
# ---------------------------
def pick_zone(row):
    for col in ["영업구역번호", "영업구역정보", "담당상세"]:
        if col in row and pd.notna(row[col]):
            return row[col]
    return ""

df["영업구역_통합"] = df.apply(pick_zone, axis=1)

def pick_manager(row):
    for col in ["구역담당자", "담당자", "처리자"]:
        if col in row and pd.notna(row[col]) and safe_str(row[col]) != "":
            return safe_str(row[col])
    return ""

df["구역담당자_통합"] = df.apply(pick_manager, axis=1)


# ---------------------------
# 4) 해지VOC / 기타출처 분리 + 매칭여부 부여
# ---------------------------
df_voc = df[df.get("출처") == "해지VOC"].copy()
df_other = df[df.get("출처") != "해지VOC"].copy()

other_contracts = set(df_other["계약번호_정제"].dropna().tolist())

df_voc["매칭여부"] = df_voc["계약번호_정제"].apply(
    lambda x: "매칭(O)" if x in other_contracts else "비매칭(X)"
)


# ---------------------------
# 5) 설치주소 통합
# ---------------------------
def merge_address(row):
    for col in ["시설_설치주소", "설치주소"]:
        if col in row and pd.notna(row[col]) and safe_str(row[col]) not in ["None", "nan", ""]:
            return safe_str(row[col])
    return np.nan

df_voc["설치주소_표시"] = df_voc.apply(merge_address, axis=1)

address_cols = [c for c in df.columns if "주소" in c]


# ---------------------------
# 6) 월정료 정제
# ---------------------------
fee_raw_col = None
for cand in ["시설_KTT월정료(조정)", "KTT월정료(조정)", "월정료"]:
    if cand in df_voc.columns:
        fee_raw_col = cand
        break

def parse_fee(v):
    """월정료 정제 + 10배 오류 보정"""
    if pd.isna(v):
        return np.nan

    s = safe_str(v).replace(",", "")
    s = "".join(ch for ch in s if ch.isdigit())
    if s == "":
        return np.nan

    f = float(s)

    # 20만 이상 숫자는 대부분 10배 오류 → 보정
    if f >= 200000:
        f = f / 10

    return f

if fee_raw_col:
    df_voc["월정료_수치"] = df_voc[fee_raw_col].apply(parse_fee)
else:
    df_voc["월정료_수치"] = np.nan

df_voc["월정료_표시"] = df_voc["월정료_수치"].apply(
    lambda x: f"{int(x):,}" if pd.notna(x) else ""
)

df_voc["월정료구간"] = df_voc["월정료_수치"].apply(
    lambda x: "10만 이상" if pd.notna(x) and x >= 100000 else (
        "10만 미만" if pd.notna(x) else "미기재"
    )
)


# ---------------------------
# 7) 리스크 등급 산정
# ---------------------------
today = date.today()

def compute_risk(row):
    dt = row.get("접수일시")
    if pd.isna(dt):
        return np.nan, "LOW"

    if isinstance(dt, datetime):
        dt = dt.date()

    days = (today - dt).days

    if days <= 3:
        level = "HIGH"
    elif days <= 10:
        level = "MEDIUM"
    else:
        level = "LOW"

    return days, level

df_voc["경과일수"], df_voc["리스크등급"] = zip(
    *df_voc.apply(compute_risk, axis=1)
)

# 비매칭 데이터
df_unmatched = df_voc[df_voc["매칭여부"] == "비매칭(X)"].copy()

# ===========================================================
# PART 3 — 글로벌 필터 / Apple Glass UI / KPI 대시보드 / 고급 시각화
# ===========================================================

# ---------------------------
# Apple Glass 스타일 개선
# ---------------------------
st.markdown("""
<style>
/* 전체 배경 톤 다운 */
html, body, .stApp {
    background: #f5f5f7 !important;
}

/* KPI 카드 */
.kpi-card {
    background: rgba(255,255,255,0.75);
    backdrop-filter: blur(12px);
    -webkit-backdrop-filter: blur(12px);
    padding: 1.2rem 1.3rem;
    border-radius: 16px;
    border: 1px solid rgba(255,255,255,0.45);
    box-shadow: 0 6px 14px rgba(0,0,0,0.08);
}

/* 표 가독성 */
.dataframe tbody tr:nth-child(odd) { background:#fafafa; }
.dataframe tbody tr:nth-child(even) { background:#eef2ff; }

/* 시각화 여백 */
.js-plotly-plot .plotly {
    background-color: transparent !important;
}

/* 4열 그리드 지사 배치용 */
.branch-grid {
    display:grid;
    grid-template-columns:repeat(4, minmax(0,1fr));
    gap:14px;
}
.branch-item {
    background:rgba(255,255,255,0.55);
    backdrop-filter:blur(10px);
    padding:0.9rem;
    border-radius:14px;
    border:1px solid rgba(255,255,255,0.4);
    text-align:center;
    font-weight:600;
    box-shadow:0 4px 10px rgba(0,0,0,0.05);
}
</style>
""", unsafe_allow_html=True)


# ===========================================================
# 🔧 Sidebar — 글로벌 필터
# ===========================================================

st.sidebar.title("🔧 글로벌 필터")

# 1) 날짜 필터
if "접수일시" in df_voc.columns and df_voc["접수일시"].notna().any():
    min_d = df_voc["접수일시"].min().date()
    max_d = df_voc["접수일시"].max().date()

    sel_date = st.sidebar.date_input(
        "접수일자",
        value=(min_d, max_d),
        min_value=min_d,
        max_value=max_d
    )
else:
    sel_date = None

# 2) 지사
branch_list = order_branches(df_voc["관리지사"].dropna().unique())
sel_branches = st.sidebar.multiselect(
    "관리지사",
    options=branch_list,
    default=branch_list
)

# 3) 리스크
risk_opts = ["HIGH", "MEDIUM", "LOW"]
sel_risk = st.sidebar.multiselect(
    "리스크등급",
    risk_opts,
    default=risk_opts
)

# 4) 매칭여부
sel_match = st.sidebar.multiselect(
    "매칭여부",
    ["매칭(O)", "비매칭(X)"],
    default=["매칭(O)", "비매칭(X)"]
)

# 5) 월정료 구간
sel_fee = st.sidebar.radio(
    "월정료 구간",
    ["전체", "10만 이상", "10만 미만"],
    index=0
)

st.sidebar.markdown("---")
st.sidebar.caption("필터는 모든 탭에 즉시 반영됩니다.")


# ===========================================================
# 🔍 글로벌 필터 적용
# ===========================================================

voc_f = df_voc.copy()

# 날짜
if sel_date:
    start, end = sel_date
    voc_f = voc_f[
        (voc_f["접수일시"] >= pd.to_datetime(start)) &
        (voc_f["접수일시"] < pd.to_datetime(end) + pd.Timedelta(days=1))
    ]

# 지사
voc_f = voc_f[voc_f["관리지사"].isin(sel_branches)]

# 리스크
voc_f = voc_f[voc_f["리스크등급"].isin(sel_risk)]

# 매칭 여부
voc_f = voc_f[voc_f["매칭여부"].isin(sel_match)]

# 월정료
if sel_fee == "10만 이상":
    voc_f = voc_f[voc_f["월정료_수치"] >= 100000]
elif sel_fee == "10만 미만":
    voc_f = voc_f[(voc_f["월정료_수치"] < 100000) & voc_f["월정료_수치"].notna()]

# 비매칭 필터만
unmatched_f = voc_f[voc_f["매칭여부"] == "비매칭(X)"]


# ===========================================================
# 📊 KPI 대시보드
# ===========================================================

st.markdown("## 📊 해지 VOC 종합 대시보드 (Apple Glass Edition)")

k1, k2, k3, k4 = st.columns(4)

with k1:
    st.markdown("<div class='kpi-card'>", unsafe_allow_html=True)
    st.metric("총 VOC 건수", f"{len(voc_f):,}")
    st.markdown("</div>", unsafe_allow_html=True)

with k2:
    st.markdown("<div class='kpi-card'>", unsafe_allow_html=True)
    st.metric("계약 수(Unique)", f"{voc_f['계약번호_정제'].nunique():,}")
    st.markdown("</div>", unsafe_allow_html=True)

with k3:
    st.markdown("<div class='kpi-card'>", unsafe_allow_html=True)
    st.metric("비매칭 계약", f"{unmatched_f['계약번호_정제'].nunique():,}")
    st.markdown("</div>", unsafe_allow_html=True)

with k4:
    st.markdown("<div class='kpi-card'>", unsafe_allow_html=True)
    st.metric("매칭 계약",
              f"{voc_f[voc_f['매칭여부']=='매칭(O)']['계약번호_정제'].nunique():,}")
    st.markdown("</div>", unsafe_allow_html=True)

st.markdown("---")

# ===========================================================
# PART 4 — 계약별 상세조회 + 활동등록(피드백) 시스템
# Apple Glass UI 강화 + 관리자 모드 추가
# ===========================================================

tab1, tab2, tab3 = st.tabs([
    "🔍 계약별 상세 조회",
    "📝 활동등록(피드백)",
    "📦 통합 이력 다운로드"
])


# ===========================================================
# TAB 1 — 계약별 상세 조회
# ===========================================================
with tab1:
    st.subheader("🔍 계약별 상세 조회 (Apple Glass Edition)")

    df_drill = voc_f.copy()
    contract_list = sorted(df_drill["계약번호_정제"].dropna().unique().tolist())

    sel_cn = st.selectbox("📌 계약번호 선택", ["(선택)"] + contract_list)

    if sel_cn != "(선택)":

        voc_hist = df_voc[df_voc["계약번호_정제"] == sel_cn].sort_values("접수일시", ascending=False)
        other_hist = df_other[df_other["계약번호_정제"] == sel_cn]

        base = voc_hist.iloc[0]

        # -------------------------------------
        # Apple Glass 카드형 기본 정보
        # -------------------------------------
        st.markdown("### 🧊 기본 정보 (Glass Card)")
        info1, info2, info3, info4 = st.columns(4)

        info1.metric("상호", base.get("상호", ""))
        info2.metric("지사", base.get("관리지사", ""))
        info3.metric("담당자", base.get("구역담당자_통합", ""))
        info4.metric("총 VOC 건수", f"{len(voc_hist)} 건")

        st.caption(f"📍 설치주소: {base.get('설치주소_표시', '')}")
        if fee_col:
            st.caption(f"💰 월정료: {base.get('월정료_표시','')} 원")

        st.markdown("---")

        # -------------------------------------
        # VOC 상세 이력
        # -------------------------------------
        st.markdown("### 📘 VOC 상세 이력 (최신순)")

        show_cols_voc = [
            "접수일시", "VOC유형", "VOC유형소", "등록내용",
            "리스크등급", "경과일수", "관리지사", "구역담당자_통합"
        ]
        show_cols_voc = [c for c in show_cols_voc if c in voc_hist.columns]

        st.dataframe(
            style_risk(voc_hist[show_cols_voc]),
            use_container_width=True, height=350
        )

        st.markdown("---")

        # -------------------------------------
        # 기타 출처 (해지시설/설변 등)
        # -------------------------------------
        st.markdown("### 📂 기타 출처 이력")
        if other_hist.empty:
            st.info("📭 기타 출처 데이터 없음")
        else:
            st.dataframe(
                other_hist,
                use_container_width=True, height=300
            )


# ===========================================================
# TAB 2 — 활동등록(피드백) 관리
# ===========================================================
with tab2:
    st.subheader("📝 해지상담 활동등록 (피드백)")

    if sel_cn == "(선택)":
        st.info("✔ 먼저 TAB 1에서 계약번호를 선택하세요.")
    else:
        fb_all = st.session_state["feedback_df"]
        fb_sel = fb_all[fb_all["계약번호_정제"] == sel_cn].sort_values("등록일자", ascending=False)

        st.markdown("### 📄 기존 처리내역")

        if fb_sel.empty:
            st.info("등록된 처리내역 없음")
        else:
            st.dataframe(fb_sel, use_container_width=True, height=320)

        # -------------------------------------
        # 신규 입력
        # -------------------------------------
        st.markdown("### ➕ 새 처리내역 등록")

        new_msg = st.text_area("고객 대응내용 입력")
        new_writer = st.text_input("등록자")
        new_note = st.text_input("비고(Optional)")

        if st.button("등록하기"):
            if not new_msg or not new_writer:
                st.warning("내용 + 등록자 필수 입력입니다.")
            else:
                new_row = {
                    "계약번호_정제": sel_cn,
                    "고객대응내용": new_msg,
                    "등록자": new_writer,
                    "등록일자": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                    "비고": new_note
                }
                fb_all = pd.concat([fb_all, pd.DataFrame([new_row])], ignore_index=True)
                st.session_state["feedback_df"] = fb_all
                save_feedback(FEEDBACK_PATH, fb_all)
                st.success("등록 완료되었습니다.")
                st.rerun()

        # -------------------------------------
        # 관리자 삭제 모드
        # -------------------------------------
        st.markdown("---")
        st.markdown("### 🗑 관리자 삭제 기능")

        admin_pw = st.text_input("관리자 비밀번호", type="password")
        ADMIN_CODE = "C3A"

        if admin_pw == ADMIN_CODE:
            del_idx = st.number_input("삭제할 행 index 입력", min_value=0, step=1)
            if st.button("행 삭제하기"):
                try:
                    fb_all = fb_all.drop(fb_sel.index[del_idx])
                    st.session_state["feedback_df"] = fb_all
                    save_feedback(FEEDBACK_PATH, fb_all)
                    st.success("삭제 완료!")
                    st.rerun()
                except:
                    st.error("해당 index가 존재하지 않습니다.")
        else:
            st.caption("🔐 삭제는 관리자만 가능합니다 (코드: C3A)")


# ===========================================================
# TAB 3 — 통합 이력 다운로드
# ===========================================================
with tab3:
    st.subheader("📦 선택 계약 통합 이력 다운로드")

    if sel_cn == "(선택)":
        st.info("계약번호를 먼저 선택하세요.")
    else:
        export_frames = []

        if not voc_hist.empty:
            t1 = voc_hist.copy()
            t1.insert(0, "구분", "VOC")
            export_frames.append(t1)

        if not other_hist.empty:
            t2 = other_hist.copy()
            t2.insert(0, "구분", "기타출처")
            export_frames.append(t2)

        fb_sel = st.session_state["feedback_df"]
        fb_sel = fb_sel[fb_sel["계약번호_정제"] == sel_cn]
        if not fb_sel.empty:
            t3 = fb_sel.copy()
            t3.insert(0, "구분", "처리내역")
            export_frames.append(t3)

        if export_frames:
            merged = pd.concat(export_frames, ignore_index=True)
            st.download_button(
                "📥 통합 CSV 다운로드",
                merged.to_csv(index=False).encode("utf-8-sig"),
                file_name=f"통합이력_{sel_cn}.csv",
                mime="text/csv"
            )
        else:
            st.info("다운로드할 이력이 없습니다.")

# ===========================================================
# PART 5 — 담당자 알림 발송 (Apple Glass Enterprise Edition)
# ===========================================================

with st.tab("📨 담당자 알림 (기업용)"):

    st.subheader("📨 담당자 알림 발송 시스템 (Apple Glass Edition)")

    st.markdown("""
        비매칭(X) 계약 건을 담당자에게 이메일로 전송합니다.<br>
        CSV 첨부파일에는 **계약번호 중복 없이 최신 VOC 1건 요약본**이 포함됩니다.
        """, unsafe_allow_html=True)

    # -----------------------------
    # 담당자별 비매칭 계약 집계
    # -----------------------------
    unmatched_alert = unmatched_f.copy()

    if unmatched_alert.empty:
        st.info("현재 글로벌 필터에서 비매칭 계약이 없습니다.")
        st.stop()

    grouped = unmatched_alert.groupby("구역담당자_통합")

    rows = []
    for mgr, g in grouped:
        if not mgr:
            continue
        count = g["계약번호_정제"].nunique()
        email = manager_contacts.get(mgr, {}).get("email", "")
        rows.append([mgr, email, count])

    alert_df = pd.DataFrame(rows, columns=["담당자", "이메일", "비매칭 계약수"])

    st.markdown("### 🧊 전체 담당자 현황 (Glass table)")
    st.dataframe(alert_df, use_container_width=True, height=260)

    st.markdown("---")

    # -----------------------------
    # 담당자 선택
    # -----------------------------
    sel_mgr = st.selectbox(
        "담당자 선택",
        ["(선택)"] + alert_df["담당자"].tolist(),
        key="alert_mgr_select"
    )

    if sel_mgr != "(선택)":

        # 담당자 이메일 자동 입력
        default_email = manager_contacts.get(sel_mgr, {}).get("email", "")
        email_input = st.text_input("📮 이메일 주소", value=default_email)

        # 담당자의 비매칭 데이터 필터
        df_mgr = unmatched_alert[
            unmatched_alert["구역담당자_통합"] == sel_mgr
        ].copy()

        st.markdown(f"### 🔍 {sel_mgr} 담당자 비매칭 계약 상세")

        show_cols = [
            "계약번호_정제", "상호", "관리지사",
            "VOC유형", "VOC유형소", "등록내용",
            "리스크등급", "경과일수"
        ]
        show_cols = [c for c in show_cols if c in df_mgr.columns]

        st.dataframe(df_mgr[show_cols], use_container_width=True, height=350)

        # -----------------------------
        # 최신 VOC 기준 1건 요약 CSV 생성
        # -----------------------------
        df_sorted = df_mgr.sort_values("접수일시", ascending=False)
        grp = df_sorted.groupby("계약번호_정제")
        idx = grp["접수일시"].idxmax()
        df_latest = df_sorted.loc[idx].copy()

        export_cols = [
            "계약번호_정제", "상호", "관리지사", "구역담당자_통합",
            "VOC유형", "VOC유형소", "등록내용",
            "설치주소_표시", "리스크등급", "경과일수"
        ]
        export_cols = [c for c in export_cols if c in df_latest.columns]

        df_latest = df_latest[export_cols]

        csv_bytes = df_latest.to_csv(index=False, encoding="utf-8-sig").encode("utf-8-sig")

        # -----------------------------
        # 이메일 본문
        # -----------------------------
        subject = f"[해지VOC] {sel_mgr} 담당자 비매칭 시설 안내"
        body = (
            f"{sel_mgr} 담당자님,\n\n"
            f"현재 총 {len(df_latest)}건의 비매칭 시설이 확인되었습니다.\n"
            f"첨부된 CSV 파일을 확인하시고, 신속히 확인 및 대응 부탁드립니다.\n\n"
            "- 해지VOC 관리자 드림 -"
        )

        # -----------------------------
        # 이메일 발송 버튼
        # -----------------------------
        if st.button("📤 이메일 발송하기", key="send_email_button"):
            if not email_input:
                st.error("이메일 주소를 입력해주세요.")
            else:
                try:
                    msg = EmailMessage()
                    msg["Subject"] = subject
                    msg["From"] = f"{SENDER_NAME} <{SMTP_USER}>"
                    msg["To"] = email_input
                    msg.set_content(body)

                    msg.add_attachment(
                        csv_bytes,
                        maintype="application",
                        subtype="octet-stream",
                        filename=f"비매칭계약_{sel_mgr}.csv",
                    )

                    with smtplib.SMTP(SMTP_HOST, SMTP_PORT) as smtp:
                        smtp.starttls()
                        smtp.login(SMTP_USER, SMTP_PASSWORD)
                        smtp.send_message(msg)

                    st.success(f"✅ 이메일 발송 완료 → {email_input}")

                except Exception as e:
                    st.error(f"❌ 이메일 전송 실패: {e}")

    st.caption("※ 비매칭 = 해지VOC 접수 후 활동내역 미등록 시설")
