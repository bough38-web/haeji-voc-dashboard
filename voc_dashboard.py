import os
from datetime import datetime, date

import numpy as np
import pandas as pd
import streamlit as st
import smtplib
from email.mime.text import MIMEText
from email.utils import formataddr

# ----------------------------------------------------
# 0. 기본 설정 & 메일 설정
# ----------------------------------------------------
st.set_page_config(page_title="해지VOC 담당자 안내 / 알림", layout="wide")

# 🔐 메일 서버 설정 (비밀번호는 환경변수 이용!)
SMTP_HOST = "smtp.gmail.com"
SMTP_PORT = 587
SMTP_USER = "bough38@gmail.com"
SMTP_PASSWORD = os.environ.get("SMTP_PASSWORD", "")  # ← 여기 직접 쓰지 말기
SENDER_NAME = "해지VOC 관리자"

# ----------------------------------------------------
# 1. 공통 경로 & 유틸
# ----------------------------------------------------
BASE_DIR = os.path.dirname(os.path.abspath(__file__))

def find_file(default_name: str) -> str | None:
    """
    기본 파일명이 없으면, 같은 폴더에서 비슷한 이름의 파일을 찾아서 사용
    (한글/공백/정규화 문제 방지)
    """
    candidate = os.path.join(BASE_DIR, default_name)
    if os.path.exists(candidate):
        return candidate

    # fallback: 현재 폴더 전체 스캔
    for fn in os.listdir(BASE_DIR):
        if fn.endswith(".xlsx") and default_name.split(".")[0] in fn:
            return os.path.join(BASE_DIR, fn)

    return None


MERGED_PATH = find_file("merged.xlsx")
MAPPING_PATH = find_file("영업구역담당자_251204.xlsx")

# ----------------------------------------------------
# 2. 데이터 로딩
# ----------------------------------------------------
@st.cache_data
def load_voc_data(path: str) -> pd.DataFrame:
    if not path or not os.path.exists(path):
        st.error("❌ VOC 데이터 파일 'merged.xlsx' 을(를) 찾을 수 없습니다.")
        return pd.DataFrame()

    df = pd.read_excel(path)

    # 계약번호 정제
    if "계약번호" in df.columns:
        df["계약번호_정제"] = (
            df["계약번호"]
            .astype(str)
            .str.replace(r"[^0-9A-Za-z]", "", regex=True)
            .str.strip()
        )
    else:
        df["계약번호_정제"] = ""

    # 출처 정제
    if "출처" in df.columns:
        df["출처"] = df["출처"].replace({"고객리스트": "해지시설"})

    # 접수일시
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

    return df


@st.cache_data
def load_manager_mapping(path: str) -> pd.DataFrame:
    """
    영업구역 담당자 매핑 파일 로드.
    원본 컬럼 기준:
    - 처리자1  : 담당자 이름
    - 담당상세 : 영업구역 번호
    - 소속     : 지사
    - 연략처   : 연락처(휴대폰)
    - E-MAIL  : 이메일
    """
    if not path or not os.path.exists(path):
        st.error("❌ 담당자 매핑 파일 '영업구역담당자_251204.xlsx' 을(를) 찾을 수 없습니다.")
        return pd.DataFrame()

    df = pd.read_excel(path)

    # 안전하게 컬럼명 변환
    rename_map = {}
    for col in df.columns:
        if col == "처리자1":
            rename_map[col] = "담당자"
        elif col == "소속":
            rename_map[col] = "관리지사"
        elif col == "E-MAIL":
            rename_map[col] = "이메일"
        elif col == "연략처":
            rename_map[col] = "연락처"
        elif col == "담당상세":
            rename_map[col] = "영업구역번호"

    df = df.rename(columns=rename_map)

    # 필요한 컬럼만 방어
    for need in ["담당자", "관리지사", "이메일"]:
        if need not in df.columns:
            st.error(f"❌ 매핑 파일에 '{need}' 컬럼이 없습니다. 엑셀 컬럼명을 확인해주세요.")
            return pd.DataFrame()

    df["담당자"] = df["담당자"].astype(str).str.strip()
    df["관리지사"] = df["관리지사"].astype(str).str.strip()
    df["이메일"] = df["이메일"].astype(str).str.strip()
    if "연락처" in df.columns:
        df["연락처"] = df["연락처"].astype(str).str.strip()
    if "영업구역번호" in df.columns:
        df["영업구역번호"] = df["영업구역번호"].astype(str).str.strip()

    return df


voc_df = load_voc_data(MERGED_PATH)
mgr_df = load_manager_mapping(MAPPING_PATH)

if voc_df.empty or mgr_df.empty:
    st.stop()

# ----------------------------------------------------
# 3. 해지VOC 비매칭(=해지방어 활동시설) 산출
# ----------------------------------------------------
# 기타 출처
df_other = voc_df[voc_df.get("출처") != "해지VOC"].copy()
df_voc = voc_df[voc_df.get("출처") == "해지VOC"].copy()

# 계약번호 매칭
other_union = set(
    df_other.get("계약번호_정제", pd.Series(dtype=str)).dropna().astype(str)
)
df_voc["매칭여부"] = df_voc["계약번호_정제"].astype(str).apply(
    lambda x: "매칭(O)" if x in other_union else "비매칭(X)"
)

unmatched = df_voc[df_voc["매칭여부"] == "비매칭(X)"].copy()

# 리스크 등급(간단 버전)
today = date.today()
def compute_risk(dt):
    if pd.isna(dt):
        return np.nan, "LOW"
    if not isinstance(dt, (pd.Timestamp, datetime)):
        dt = pd.to_datetime(dt, errors="coerce")
    if pd.isna(dt):
        return np.nan, "LOW"
    days = (today - dt.date()).days
    if days <= 3:
        return days, "HIGH"
    elif days <= 10:
        return days, "MEDIUM"
    return days, "LOW"

if "접수일시" in unmatched.columns:
    unmatched["경과일수"], unmatched["리스크등급"] = zip(
        *unmatched["접수일시"].map(compute_risk)
    )
else:
    unmatched["경과일수"] = np.nan
    unmatched["리스크등급"] = "LOW"

# ----------------------------------------------------
# 4. 메일 발송 함수
# ----------------------------------------------------
def send_email(to_addr: str, subject: str, body: str, cc_addr: str | None = None):
    if not SMTP_PASSWORD:
        st.error("❌ SMTP_PASSWORD 환경변수가 설정되어 있지 않습니다. 서버 환경변수로 Gmail 앱 비밀번호를 넣어주세요.")
        return False

    msg = MIMEText(body, _charset="utf-8")
    msg["Subject"] = subject
    msg["From"] = formataddr((SENDER_NAME, SMTP_USER))
    msg["To"] = to_addr
    if cc_addr:
        msg["Cc"] = cc_addr

    recipients = [to_addr]
    if cc_addr:
        recipients.append(cc_addr)

    try:
        with smtplib.SMTP(SMTP_HOST, SMTP_PORT) as server:
            server.starttls()
            server.login(SMTP_USER, SMTP_PASSWORD)
            server.sendmail(SMTP_USER, recipients, msg.as_string())
        return True
    except Exception as e:
        st.error(f"메일 발송 중 오류가 발생했습니다: {e}")
        return False

# ----------------------------------------------------
# 5. UI 구성
# ----------------------------------------------------
st.markdown("## 📨 해지VOC 담당자 안내 / 알림")

with st.expander("ℹ️ 이 페이지 설명", expanded=True):
    st.markdown(
        """
        - **목적** : 해지VOC 접수 후 **해지방어 활동이 진행되지 않은 시설**을 담당자에게 이메일로 안내하기 위한 화면입니다.  
        - **해지방어 활동시설(비매칭)** :  
          > 해지VOC가 접수되었지만, 해지시설/해지요청/설변/정지/해지파이프라인 등에서 활동 이력이 확인되지 않은 계약입니다.  
          > **신속한 확인과 현장 활동 등록이 필요합니다.**
        - 아래 단계대로 사용해주세요.  
          1. 지사 선택 → 담당자 선택  
          2. 대상 계약 목록 / 리스크 확인  
          3. 메일 제목·내용 확인 후 발송  
        """
    )

# 5-1. 지사 / 담당자 선택
left, right = st.columns([1, 2])

with left:
    st.markdown("### 1️⃣ 지사 / 담당자 선택")

    branches = ["전체"] + sorted(mgr_df["관리지사"].dropna().unique().tolist())
    sel_branch = st.selectbox("관리지사 선택", options=branches)

    if sel_branch == "전체":
        mgr_sel_df = mgr_df.copy()
    else:
        mgr_sel_df = mgr_df[mgr_df["관리지사"] == sel_branch]

    mgr_names = sorted(mgr_sel_df["담당자"].dropna().unique().tolist())
    sel_mgr_name = st.selectbox(
        "알림을 보낼 담당자 선택",
        options=mgr_names,
        index=0 if mgr_names else None,
    )

    # 선택한 담당자 정보
    sel_mgr_row = (
        mgr_sel_df[mgr_sel_df["담당자"] == sel_mgr_name].iloc[0]
        if sel_mgr_name and not mgr_sel_df.empty
        else None
    )

    if sel_mgr_row is not None:
        default_email = sel_mgr_row["이메일"]
        default_phone = sel_mgr_row.get("연락처", "")
        default_zone = sel_mgr_row.get("영업구역번호", "")

        st.markdown("---")
        st.markdown(f"**담당자:** {sel_mgr_name}")
        st.markdown(f"**이메일:** {default_email}")
        if default_phone:
            st.markdown(f"**연락처:** {default_phone}")
        if default_zone:
            st.markdown(f"**영업구역:** {default_zone}")
    else:
        default_email = ""
        default_phone = ""
        default_zone = ""

with right:
    st.markdown("### 2️⃣ 선택 담당자 기준 해지방어 활동시설")

    # 담당자 매핑 기준으로 VOC 매칭 (담당상세/영업구역번호 기준)
    target = unmatched.copy()

    # 지사 필터
    if sel_branch != "전체":
        target = target[target["관리지사"] == sel_branch]

    # 담당자 매핑: voc_df의 '영업구역_통합' 또는 '영업구역번호' 와 매핑파일의 '영업구역번호' 를 맞춰보는 방식
    voc_zone_col = None
    for c in ["영업구역_통합", "영업구역번호", "담당상세"]:
        if c in target.columns:
            voc_zone_col = c
            break

    if voc_zone_col and sel_mgr_row is not None and "영업구역번호" in sel_mgr_row:
        mgr_zone = str(sel_mgr_row["영업구역번호"])
        target = target[target[voc_zone_col].astype(str) == mgr_zone]

    # 요약
    total_contracts = target["계약번호_정제"].nunique()
    high_cnt = target[target["리스크등급"] == "HIGH"]["계약번호_정제"].nunique()
    med_cnt = target[target["리스크등급"] == "MEDIUM"]["계약번호_정제"].nunique()
    low_cnt = target[target["리스크등급"] == "LOW"]["계약번호_정제"].nunique()

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("대상 계약 수", f"{total_contracts:,}")
    c2.metric("HIGH", f"{high_cnt:,}")
    c3.metric("MEDIUM", f"{med_cnt:,}")
    c4.metric("LOW", f"{low_cnt:,}")

    st.caption("※ ‘해지방어 활동시설’ = 해지VOC 접수 후 활동 이력이 확인되지 않은 계약")

    if not target.empty:
        # 계약별 최신 VOC만 요약 표시
        target_sorted = target.sort_values("접수일시", ascending=False)
        grp = target_sorted.groupby("계약번호_정제")
        idx_latest = grp["접수일시"].idxmax()
        df_summary = target_sorted.loc[idx_latest].copy()
        df_summary["접수건수"] = grp.size().reindex(
            df_summary["계약번호_정제"]
        ).values

        show_cols = [
            "계약번호_정제",
            "상호",
            "관리지사",
            "리스크등급",
            "경과일수",
            "접수건수",
            "설치주소",
            "서비스(소)",
        ]
        show_cols = [c for c in show_cols if c in df_summary.columns]

        st.dataframe(
            df_summary[show_cols].sort_values(
                ["리스크등급", "경과일수"], ascending=[True, False]
            ),
            use_container_width=True,
            height=300,
        )

        csv_bytes = df_summary.to_csv(index=False).encode("utf-8-sig")
        st.download_button(
            "📥 대상 계약 목록 다운로드 (CSV)",
            data=csv_bytes,
            file_name=f"해지방어_활동시설_{sel_mgr_name}.csv",
            mime="text/csv",
        )
    else:
        st.info("선택한 조건에서 해지방어 활동시설(비매칭) 대상이 없습니다.")

st.markdown("---")

# ----------------------------------------------------
# 6. 메일 작성 & 발송
# ----------------------------------------------------
st.markdown("### 3️⃣ 알림 메일 작성 및 발송")

col_a, col_b = st.columns([2, 1])

with col_a:
    to_email = st.text_input(
        "받는 사람 이메일",
        value=default_email,
        help="드롭다운에서 선택한 담당자의 이메일이 기본값으로 들어갑니다. 필요 시 직접 수정하세요.",
    )
    cc_email = st.text_input(
        "참조(CC) 이메일 (선택)",
        value="",
        help="팀장/관리자 등 참조가 필요하면 입력하세요. 없으면 비워두셔도 됩니다.",
    )

    # 기본 제목/본문 템플릿
    default_subject = f"[해지VOC] {sel_branch} {sel_mgr_name} 담당 해지방어 활동시설 안내"

    example_lines = [
        f"{sel_mgr_name} 담당님,",
        "",
        "해지VOC 접수 후 해지방어 활동 이력이 확인되지 않은 시설 목록을 안내드립니다.",
        "",
        f"- 대상 계약 수 : {total_contracts:,}건",
        f"- HIGH : {high_cnt:,}건 / MEDIUM : {med_cnt:,}건 / LOW : {low_cnt:,}건",
        "",
        "각 계약에 대해 해지방어 활동 진행 후,",
        "대시보드나 관련 시스템에 **활동 내역을 등록**해주시기 바랍니다.",
        "",
        "※ 상세 계약 목록은 첨부된 CSV 파일 또는 대시보드의 '해지방어 활동시설' 화면에서 확인 가능합니다.",
        "",
        "감사합니다.",
        "",
        "해지VOC 운영담당 드림",
    ]
    default_body = "\n".join(example_lines)

    subject = st.text_input("메일 제목", value=default_subject)
    body = st.text_area("메일 내용", value=default_body, height=260)

with col_b:
    st.markdown("#### 발송 전 확인 체크리스트")
    st.checkbox("받는 사람 이메일 주소 확인 완료", value=True)
    st.checkbox("해지방어 활동시설 대상 계약 내역 확인 완료", value=True)
    st.checkbox("메일 제목 / 내용을 검토했습니다", value=True)

    st.markdown("---")
    send_btn = st.button("📧 알림 메일 발송")

    if send_btn:
        if not to_email.strip():
            st.warning("받는 사람 이메일 주소를 입력해주세요.")
        elif not subject.strip():
            st.warning("메일 제목을 입력해주세요.")
        elif not body.strip():
            st.warning("메일 내용을 입력해주세요.")
        else:
            ok = send_email(
                to_addr=to_email.strip(),
                subject=subject.strip(),
                body=body.strip(),
                cc_addr=cc_email.strip() or None,
            )
            if ok:
                st.success("✅ 메일이 정상적으로 발송되었습니다.")

st.caption(
    "※ Gmail 2단계 인증 + 앱 비밀번호가 필요합니다. "
    "서버/로컬 환경 변수 `SMTP_PASSWORD` 에 Gmail 앱 비밀번호를 설정해 주세요."
)
