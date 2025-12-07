# ================================================================
# 해지 VOC 종합 대시보드 — ENTERPRISE FINAL VERSION (Single File)
# Apple Glass UI + 지사/담당자 분석 + 계약별 드릴다운 + 활동등록 CRUD + 이메일 알림
# contact_map.xlsx 기반 자동 매핑
# ================================================================

import os
from datetime import datetime, date
import smtplib
from email.message import EmailMessage

import numpy as np
import pandas as pd
import streamlit as st

# Plotly (Optional)
try:
    import plotly.express as px
    HAS_PLOTLY = True
except:
    HAS_PLOTLY = False

# ================================================================
# 0. STREAMLIT 기본 스타일(라이트톤 + Apple Glass)
# ================================================================
st.set_page_config(page_title="해지 VOC 종합 대시보드", layout="wide")

st.markdown("""
<style>
html, body, .stApp {
    background: #f5f5f7 !important;
    color: #111 !important;
    font-family: -apple-system, BlinkMacSystemFont, "Inter", sans-serif;
}
.block-container { padding-top: 0.6rem !important; padding-bottom: 2rem !important; }
section[data-testid="stSidebar"] {
    background: #fafafa !important;
    border-right: 1px solid #e5e7eb !important;
}
textarea, input, select { border-radius: 8px !important; }
.dataframe tbody tr:nth-child(odd) { background: #f9fafb !important; }
.dataframe tbody tr:nth-child(even) { background: #eef2ff !important; }

/* KPI 카드 */
.kpi-card {
    background: rgba(255,255,255,0.75);
    backdrop-filter: blur(12px);
    border-radius: 16px;
    padding: 16px;
    border: 1px solid rgba(0,0,0,0.05);
    box-shadow: 0 3px 12px rgba(0,0,0,0.06);
}
</style>
""", unsafe_allow_html=True)

# ================================================================
# 1. SMTP 설정
# ================================================================
if "SMTP_HOST" in st.secrets:
    SMTP_HOST = st.secrets["SMTP_HOST"]
    SMTP_PORT = int(st.secrets["SMTP_PORT"])
    SMTP_USER = st.secrets["SMTP_USER"]
    SMTP_PASSWORD = st.secrets["SMTP_PASSWORD"]
    SENDER_NAME = st.secrets["SENDER_NAME"]
else:
    from dotenv import load_dotenv
    load_dotenv()
    SMTP_HOST = os.getenv("SMTP_HOST", "")
    SMTP_PORT = int(os.getenv("SMTP_PORT", "587"))
    SMTP_USER = os.getenv("SMTP_USER", "")
    SMTP_PASSWORD = os.getenv("SMTP_PASSWORD", "")
    SENDER_NAME = os.getenv("SENDER_NAME", "해지VOC 관리자")

# ================================================================
# 2. 파일 경로
# ================================================================
MERGED_PATH = "merged.xlsx"
FEEDBACK_PATH = "feedback.csv"
CONTACT_PATH = "contact_map.xlsx"

# ================================================================
# 3. 유틸 함수
# ================================================================
def safe_str(x):
    return "" if pd.isna(x) else str(x).strip()

def detect_column(df, keys):
    for k in keys:
        if k in df.columns:
            return k
    for col in df.columns:
        for k in keys:
            if k.lower() in col.lower():
                return col
    return None

# ================================================================
# 4. 데이터 로딩
# ================================================================
@st.cache_data
def load_voc(path):
    if not os.path.exists(path):
        st.error("❌ 'merged.xlsx' 파일이 없습니다.")
        return pd.DataFrame()

    df = pd.read_excel(path)

    if "계약번호" in df.columns:
        df["계약번호"] = df["계약번호"].astype(str).str.replace(",", "").str.strip()
        df["계약번호_정제"] = df["계약번호"].str.replace(r"[^0-9A-Za-z]", "", regex=True)

    if "출처" in df.columns:
        df["출처"] = df["출처"].replace({"고객리스트": "해지시설"})

    if "접수일시" in df.columns:
        df["접수일시"] = pd.to_datetime(df["접수일시"], errors="coerce")

    return df


@st.cache_data
def load_feedback(path):
    if not os.path.exists(path):
        return pd.DataFrame(columns=["계약번호_정제", "고객대응내용", "등록자", "등록일자", "비고"])
    try:
        return pd.read_csv(path, encoding="utf-8-sig")
    except:
        return pd.read_csv(path)


def save_feedback(df):
    df.to_csv(FEEDBACK_PATH, index=False, encoding="utf-8-sig")


@st.cache_data
def load_contact(path):
    if not os.path.exists(path):
        st.warning("⚠ contact_map.xlsx 파일이 없습니다.")
        return pd.DataFrame(), {}

    df = pd.read_excel(path)
    name_col = detect_column(df, ["구역담당자", "담당자", "성명"])
    email_col = detect_column(df, ["이메일", "email"])
    phone_col = detect_column(df, ["휴대폰", "전화"])

    if (not name_col) or (not email_col):
        st.warning("⚠ 담당자/이메일 컬럼을 찾지 못했습니다.")
        return df, {}

    df = df[[name_col, email_col] + ([phone_col] if phone_col else [])]
    df.columns = ["담당자", "이메일"] + (["휴대폰"] if phone_col else [])

    mapping = {
        safe_str(r["담당자"]): {
            "email": safe_str(r.get("이메일", "")),
            "phone": safe_str(r.get("휴대폰", "")),
        }
        for _, r in df.iterrows()
        if safe_str(r["담당자"]) != ""
    }

    return df, mapping


# ---------------------------------------------------
# 실제 로딩
# ---------------------------------------------------
df = load_voc(MERGED_PATH)
if df.empty:
    st.stop()

contact_df, manager_contacts = load_contact(CONTACT_PATH)

if "feedback_df" not in st.session_state:
    st.session_state["feedback_df"] = load_feedback(FEEDBACK_PATH)

# ================================================================
# 5. 전처리
# ================================================================
df["관리지사"] = df["관리지사"].replace({
    "중앙지사":"중앙","강북지사":"강북","서대문지사":"서대문","고양지사":"고양",
    "의정부지사":"의정부","남양주지사":"남양주","강릉지사":"강릉","원주지사":"원주"
})

BRANCH_ORDER = ["중앙","강북","서대문","고양","의정부","남양주","강릉","원주"]

df["구역담당자_통합"] = df.apply(
    lambda r: safe_str(r.get("구역담당자")) or safe_str(r.get("담당자")) or safe_str(r.get("처리자")),
    axis=1
)

df_voc = df[df.get("출처") == "해지VOC"].copy()
df_other = df[df.get("출처") != "해지VOC"].copy()

other_set = set(df_other["계약번호_정제"].dropna().unique())
df_voc["매칭여부"] = df_voc["계약번호_정제"].apply(lambda x: "매칭(O)" if x in other_set else "비매칭(X)")

df_voc["설치주소_표시"] = df_voc.apply(
    lambda r: safe_str(r.get("시설_설치주소")) or safe_str(r.get("설치주소")),
    axis=1
)

fee_col = None
for c in ["시설_KTT월정료(조정)", "KTT월정료(조정)", "월정료"]:
    if c in df_voc.columns:
        fee_col = c
        break

def parse_fee(x):
    if pd.isna(x): return np.nan
    s = "".join(ch for ch in str(x) if ch.isdigit())
    if not s: return np.nan
    v = float(s)
    if v >= 200000: v /= 10
    return v

if fee_col:
    df_voc["월정료_수치"] = df_voc[fee_col].apply(parse_fee)
    df_voc["월정료_표시"] = df_voc["월정료_수치"].apply(lambda v: f"{int(v):,}" if pd.notna(v) else "")
else:
    df_voc["월정료_수치"] = np.nan
    df_voc["월정료_표시"] = ""

today = date.today()
def calc_risk(row):
    dt = row.get("접수일시")
    if pd.isna(dt): return np.nan, "LOW"
    days = (today - dt.date()).days
    if days <= 3: return days, "HIGH"
    elif days <= 10: return days, "MEDIUM"
    else: return days, "LOW"

df_voc["경과일수"], df_voc["리스크등급"] = zip(*df_voc.apply(calc_risk, axis=1))
df_unmatched = df_voc[df_voc["매칭여부"] == "비매칭(X)"].copy()

display_cols = [
    "계약번호_정제","상호","관리지사","구역담당자_통합","VOC유형","VOC유형소",
    "등록내용","설치주소_표시","리스크등급","경과일수","월정료_표시"
]
display_cols = [c for c in display_cols if c in df_voc.columns]

def style_risk(dfview):
    def fmt(row):
        lvl = row["리스크등급"]
        if lvl == "HIGH": color="#fee2e2"
        elif lvl == "MEDIUM": color="#fef3c7"
        else: color="#e0f2fe"
        return [f"background-color:{color}"] * len(row)
    return dfview.style.apply(fmt, axis=1)

# ================================================================
# 6. 글로벌 필터
# ================================================================
st.sidebar.title("🔧 글로벌 필터")

if df_voc["접수일시"].notna().any():
    dmin = df_voc["접수일시"].min().date()
    dmax = df_voc["접수일시"].max().date()
    dr = st.sidebar.date_input("📅 접수일 범위", (dmin, dmax))
else:
    dr = None

branches = BRANCH_ORDER
sel_branches = st.sidebar.multiselect("🏢 지사", branches, default=branches)

sel_risk = st.sidebar.multiselect("⚠ 리스크", ["HIGH","MEDIUM","LOW"], default=["HIGH","MEDIUM","LOW"])
sel_match = st.sidebar.multiselect("🔗 매칭", ["매칭(O)","비매칭(X)"], default=["매칭(O)","비매칭(X)"])

fee_sel = st.sidebar.radio("💰 월정료", ["전체","10만 이상","10만 미만"], index=0)

voc_f = df_voc.copy()

if dr:
    s, e = dr
    voc_f = voc_f[(voc_f["접수일시"] >= pd.to_datetime(s)) &
                  (voc_f["접수일시"] < pd.to_datetime(e) + pd.Timedelta(days=1))]

voc_f = voc_f[voc_f["관리지사"].isin(sel_branches)]
voc_f = voc_f[voc_f["리스크등급"].isin(sel_risk)]
voc_f = voc_f[voc_f["매칭여부"].isin(sel_match)]

if fee_sel=="10만 이상":
    voc_f = voc_f[voc_f["월정료_수치"] >= 100000]
elif fee_sel=="10만 미만":
    voc_f = voc_f[(voc_f["월정료_수치"] < 100000) & voc_f["월정료_수치"].notna()]

unmatched_f = voc_f[voc_f["매칭여부"]=="비매칭(X)"]

# ================================================================
# 7. KPI
# ================================================================
st.markdown("## 📊 해지 VOC 종합 대시보드")

k1,k2,k3,k4=st.columns(4)
k1.metric("총 VOC 행", f"{len(voc_f):,}")
k2.metric("계약 수", f"{voc_f['계약번호_정제'].nunique():,}")
k3.metric("비매칭 계약", f"{unmatched_f['계약번호_정제'].nunique():,}")
k4.metric("매칭 계약", f"{voc_f[voc_f['매칭여부']=='매칭(O)']['계약번호_정제'].nunique():,}")

# ================================================================
# 8. 탭 구성
# ================================================================
tab_viz, tab_all, tab_un, tab_drill, tab_alert = st.tabs([
    "📊 시각화", "📘 전체 VOC", "🧯 비매칭 시설", "🔍 계약별 상세", "📨 담당자 알림"
])

# ----------------------------------------------------------------
# 📘 전체 VOC
# ----------------------------------------------------------------
with tab_all:
    st.subheader("📘 전체 VOC (계약번호 기준)")

    temp = voc_f.sort_values("접수일시", ascending=False)
    latest_idx = temp.groupby("계약번호_정제")["접수일시"].idxmax()
    df_sum = temp.loc[latest_idx]

    df_sum["접수건수"] = temp.groupby("계약번호_정제")["접수일시"].size().reindex(df_sum["계약번호_정제"]).values

    show_cols = ["계약번호_정제","상호","관리지사","구역담당자_통합","리스크등급",
                 "경과일수","매칭여부","접수건수","설치주소_표시","월정료_표시"]
    show_cols = [c for c in show_cols if c in df_sum.columns]

    st.dataframe(style_risk(df_sum[show_cols]), use_container_width=True, height=480)

# ----------------------------------------------------------------
# 🧯 비매칭 시설
# ----------------------------------------------------------------
with tab_un:
    st.subheader("🧯 비매칭 시설 목록")

    temp = unmatched_f.sort_values("접수일시", ascending=False)
    latest_idx = temp.groupby("계약번호_정제")["접수일시"].idxmax()
    df_unq = temp.loc[latest_idx]

    st.markdown(f"총 {len(df_unq)} 계약")

    show_cols_u = ["계약번호_정제","상호","관리지사","구역담당자_통합",
                   "리스크등급","경과일수","설치주소_표시"]
    show_cols_u = [c for c in show_cols_u if c in df_unq.columns]

    st.dataframe(style_risk(df_unq[show_cols_u]), use_container_width=True, height=480)

# ----------------------------------------------------------------
# 🔍 계약별 상세
# ----------------------------------------------------------------
with tab_drill:
    st.subheader("🔍 계약별 상세")

    cn_list = sorted(voc_f["계약번호_정제"].dropna().unique().tolist())
    sel_cn = st.selectbox("계약번호 선택", ["(선택)"] + cn_list)

    if sel_cn != "(선택)":
        voc_hist = df_voc[df_voc["계약번호_정제"]==sel_cn].sort_values("접수일시", ascending=False)
        other_hist = df_other[df_other["계약번호_정제"]==sel_cn]
        fb_df = st.session_state["feedback_df"]

        if not voc_hist.empty:
            base = voc_hist.iloc[0]
            st.metric("상호", base.get("상호",""))
            st.metric("지사", base.get("관리지사",""))
            st.metric("담당자", base.get("구역담당자_통합",""))

        st.markdown("### VOC 이력")
        st.dataframe(style_risk(voc_hist[display_cols]), use_container_width=True)

        st.markdown("### 기타 출처")
        st.dataframe(other_hist, use_container_width=True)

        st.markdown("### 📝 활동등록")
        fb_sel = fb_df[fb_df["계약번호_정제"]==sel_cn].sort_values("등록일자", ascending=False)

        if fb_sel.empty:
            st.info("등록된 처리내역 없음")
        else:
            st.dataframe(fb_sel, use_container_width=True)

        new_con = st.text_area("고객대응내용")
        new_writer = st.text_input("등록자")
        new_note = st.text_input("비고")

        if st.button("등록"):
            if not new_con or not new_writer:
                st.warning("내용/등록자 필수")
            else:
                new_row = {
                    "계약번호_정제": sel_cn,
                    "고객대응내용": new_con,
                    "등록자": new_writer,
                    "등록일자": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                    "비고": new_note
                }
                fb_df = pd.concat([fb_df, pd.DataFrame([new_row])], ignore_index=True)
                st.session_state["feedback_df"] = fb_df
                save_feedback(fb_df)
                st.success("등록 완료!")
                st.rerun()

# ----------------------------------------------------------------
# 📨 담당자 알림
# ----------------------------------------------------------------
with tab_alert:
    st.subheader("📨 담당자 알림 발송")

    alert_rows=[]
    for mgr, g in unmatched_f.groupby("구역담당자_통합"):
        mgr = safe_str(mgr)
        if not mgr: continue
        email = manager_contacts.get(mgr,{}).get("email","")
        cnt = g["계약번호_정제"].nunique()
        alert_rows.append([mgr,email,cnt])

    alert_df = pd.DataFrame(alert_rows, columns=["담당자","이메일","비매칭건"])

    st.dataframe(alert_df, use_container_width=True, height=300)

    sel_mgr = st.selectbox("담당자 선택", ["(선택)"] + alert_df["담당자"].tolist())
    if sel_mgr != "(선택)":
        to_email = st.text_input("메일주소", value=manager_contacts.get(sel_mgr,{}).get("email",""))

        df_mgr = unmatched_f[unmatched_f["구역담당자_통합"]==sel_mgr]
        st.dataframe(df_mgr[["계약번호_정제","상호","리스크등급","경과일수"]], use_container_width=True)

        # 최신 VOC만 첨부
        sorted_mgr=df_mgr.sort_values("접수일시",ascending=False)
        latest_idx=sorted_mgr.groupby("계약번호_정제")["접수일시"].idxmax()
        df_latest=sorted_mgr.loc[latest_idx]

        csv_bytes=df_latest.to_csv(index=False,encoding="utf-8-sig").encode("utf-8-sig")

        if st.button("📤 발송"):
            try:
                msg=EmailMessage()
                msg["Subject"]=f"[해지VOC] {sel_mgr} 담당자 비매칭 안내"
                msg["From"]=f"{SENDER_NAME} <{SMTP_USER}>"
                msg["To"]=to_email
                msg.set_content("비매칭 VOC 첨부파일을 확인해주세요.")

                msg.add_attachment(csv_bytes, maintype="application",
                    subtype="octet-stream", filename=f"비매칭_{sel_mgr}.csv")

                with smtplib.SMTP(SMTP_HOST,SMTP_PORT) as smtp:
                    smtp.starttls()
                    smtp.login(SMTP_USER,SMTP_PASSWORD)
                    smtp.send_message(msg)

                st.success("발송 완료!")
            except Exception as e:
                st.error(f"오류: {e}")
