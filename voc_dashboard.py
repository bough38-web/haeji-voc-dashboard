import os
from datetime import datetime, date

import numpy as np
import pandas as pd
import streamlit as st

# ----------------------------------------------------
# 0. 기본 설정 & 라이트톤 스타일
# ----------------------------------------------------
st.set_page_config(page_title="해지 VOC 종합 대시보드", layout="wide")

st.markdown(
    """
    <style>
    .stApp {
        background-color: #f8fafc;
        color: #111827;
        font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif;
    }
    [data-testid="stHeader"] {
        background-color: #f8fafc;
    }
    .dataframe tbody tr:nth-child(odd) {
        background-color: #f9fafb;
    }
    .dataframe tbody tr:nth-child(even) {
        background-color: #eef2ff;
    }
    </style>
    """,
    unsafe_allow_html=True,
)

# ----------------------------------------------------
# 1. 파일 경로
# ----------------------------------------------------
MERGED_PATH = "merged.xlsx"
FEEDBACK_PATH = "feedback.csv"  # (필요 없으면 나중에 제거해도 됨)

# ----------------------------------------------------
# 2. 데이터 로딩
# ----------------------------------------------------
@st.cache_data
def load_merged(path: str) -> pd.DataFrame:
    if not os.path.exists(path):
        st.error("❌ 'merged.xlsx' 파일이 존재하지 않습니다. 저장소 루트 경로를 확인하세요.")
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

    # 출처 정제 (예전 '고객리스트' → '해지시설')
    if "출처" in df.columns:
        df["출처"] = df["출처"].replace({"고객리스트": "해지시설"})

    # 접수일시 → datetime
    if "접수일시" in df.columns:
        df["접수일시"] = pd.to_datetime(df["접수일시"], errors="coerce")

    # 관리지사 축약
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

    # 구역담당자_통합 생성 (구역담당자 / 담당자 / 처리자 우선순위)
    mgr_priority = ["구역담당자", "담당자", "처리자"]

    def pick_manager(row):
        for c in mgr_priority:
            if c in row and pd.notna(row[c]) and str(row[c]).strip() != "":
                return row[c]
        return ""

    df["구역담당자_통합"] = df.apply(pick_manager, axis=1)

    # 시설_KTT월정료(조정) 숫자/구간/표시 정제
    if "시설_KTT월정료(조정)" in df.columns:
        def parse_fee(x):
            if pd.isna(x):
                return np.nan
            s = str(x).replace(",", "").strip()
            if s == "" or s.lower() in ["none", "nan"]:
                return np.nan
            # 숫자만 남기기
            s = "".join(ch for ch in s if ch.isdigit())
            if s == "":
                return np.nan
            try:
                return float(s)
            except Exception:
                return np.nan

        df["시설_월정료_수치"] = df["시설_KTT월정료(조정)"].apply(parse_fee)

        # 10만 이상/미만 구간
        def fee_band(v):
            if pd.isna(v):
                return "미기재"
            if v >= 100000:
                return "10만 이상"
            return "10만 미만"

        df["월정료구간"] = df["시설_월정료_수치"].apply(fee_band)

        # 화면 표시용(천단위 콤마, 소수점 제거)
        def fmt_fee(v):
            if pd.isna(v):
                return ""
            try:
                return f"{int(v):,}"
            except Exception:
                return ""

        df["시설_KTT월정료(조정)"] = df["시설_월정료_수치"].apply(fmt_fee)
    else:
        df["시설_월정료_수치"] = np.nan
        df["월정료구간"] = "미기재"

    # 기존 월정료_표시 컬럼은 중복이므로 제거
    if "월정료_표시" in df.columns:
        df = df.drop(columns=["월정료_표시"])

    return df


@st.cache_data
def load_feedback(path: str) -> pd.DataFrame:
    if os.path.exists(path):
        try:
            return pd.read_csv(path, encoding="utf-8-sig")
        except Exception:
            return pd.read_csv(path)
    else:
        return pd.DataFrame(
            columns=["계약번호_정제", "고객대응내용", "등록자", "등록일자", "비고"]
        )


def save_feedback(path: str, fb_df: pd.DataFrame) -> None:
    fb_df.to_csv(path, index=False, encoding="utf-8-sig")


# ---------------- 실제 로딩 ----------------
df_all = load_merged(MERGED_PATH)
if df_all.empty:
    st.stop()

if "feedback_df" not in st.session_state:
    st.session_state["feedback_df"] = load_feedback(FEEDBACK_PATH)

BRANCH_ORDER = ["중앙", "강북", "서대문", "고양", "의정부", "남양주", "강릉", "원주"]


def sort_branch(series):
    uniq = [s for s in series if pd.notna(s)]
    uniq = list(dict.fromkeys(uniq))  # 순서 유지 중복 제거
    return sorted([s for s in uniq if s in BRANCH_ORDER], key=BRANCH_ORDER.index)


# ----------------------------------------------------
# 3. VOC / 기타 출처 분리 및 매칭여부 계산
# ----------------------------------------------------
df_voc_rows = df_all[df_all.get("출처") == "해지VOC"].copy()
df_other_rows = df_all[df_all.get("출처") != "해지VOC"].copy()

# 기타 출처 계약번호 집합
other_contracts = set(
    df_other_rows["계약번호_정제"].dropna().astype(str).tolist()
)

# 매칭여부 (행 단위)
df_voc_rows["매칭여부"] = np.where(
    df_voc_rows["계약번호_정제"].astype(str).isin(other_contracts),
    "매칭(O)",
    "비매칭(X)",
)

# 리스크 계산
today = date.today()


def compute_risk(dt):
    if pd.isna(dt):
        return np.nan, "LOW"
    if not isinstance(dt, (datetime, pd.Timestamp)):
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


df_voc_rows["경과일수"], df_voc_rows["리스크등급"] = zip(
    *df_voc_rows["접수일시"].map(compute_risk)
)

# 통합 설치주소: 시설_설치주소 우선, 없으면 기존 설치주소
if "시설_설치주소" in df_voc_rows.columns or "설치주소" in df_voc_rows.columns:
    df_voc_rows["설치주소_통합"] = df_voc_rows["시설_설치주소"].fillna(
        df_voc_rows.get("설치주소")
    )
else:
    df_voc_rows["설치주소_통합"] = ""

# ----------------------------------------------------
# 4. 계약번호 기준 요약 테이블 생성
# ----------------------------------------------------
# 필터를 적용하기 전에 기본 구조를 만드는 함수
def make_contract_summary(voc_df: pd.DataFrame) -> pd.DataFrame:
    """
    voc_df : 해지VOC 행단위 데이터 (필터 후)
    계약번호_정제 기준으로 최근 VOC만 남기고, 접수건수/매칭여부/리스크 등 요약
    """
    if voc_df.empty:
        return voc_df.head(0).copy()

    voc_df = voc_df.copy()
    voc_df = voc_df.sort_values("접수일시", ascending=False)

    grp = voc_df.groupby("계약번호_정제", dropna=True)
    # 최신 VOC 행 index
    latest_idx = grp["접수일시"].idxmax()
    df_latest = voc_df.loc[latest_idx].copy()

    # 접수건수
    df_latest["접수건수"] = grp.size().reindex(df_latest["계약번호_정제"]).values

    return df_latest


# ----------------------------------------------------
# 5. 스타일링 함수 (리스크색)
# ----------------------------------------------------
def style_risk(df_view: pd.DataFrame):
    if "리스크등급" not in df_view.columns:
        return df_view

    def _row(row):
        lv = row.get("리스크등급", "")
        if lv == "HIGH":
            bg = "#fee2e2"
        elif lv == "MEDIUM":
            bg = "#fef3c7"
        else:
            bg = "#e0f2fe"
        return [f"background-color: {bg};"] * len(row)

    return df_view.style.apply(_row, axis=1)


# ----------------------------------------------------
# 6. 글로벌 필터 (사이드바)
# ----------------------------------------------------
st.sidebar.header("🔧 글로벌 필터")

# 날짜 범위
if "접수일시" in df_voc_rows.columns and df_voc_rows["접수일시"].notna().any():
    min_d = df_voc_rows["접수일시"].min().date()
    max_d = df_voc_rows["접수일시"].max().date()
    date_range = st.sidebar.date_input(
        "접수일자 범위",
        value=(min_d, max_d),
        min_value=min_d,
        max_value=max_d,
        key="global_date_range",
    )
else:
    date_range = None

# 지사
branches_all = sort_branch(df_voc_rows["관리지사"].dropna().unique())
sel_branches = st.sidebar.multiselect(
    "관리지사(복수 선택)",
    options=branches_all,
    default=branches_all,
)

# 리스크
risk_all = ["HIGH", "MEDIUM", "LOW"]
sel_risk = st.sidebar.multiselect(
    "리스크등급",
    options=risk_all,
    default=risk_all,
)

# 매칭여부
match_all = ["매칭(O)", "비매칭(X)"]
sel_match = st.sidebar.multiselect(
    "매칭여부",
    options=match_all,
    default=match_all,
)

# 월정료 구간
fee_band_choice = st.sidebar.radio(
    "시설_KTT월정료(조정) 구간",
    options=["전체", "10만 미만", "10만 이상"],
    index=0,
    key="global_fee_band",
)

st.sidebar.markdown("---")
st.sidebar.caption(
    f"마지막 갱신: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}"
)

# ----------------------------------------------------
# 7. 글로벌 필터 적용 (행 단위 → 계약 요약)
# ----------------------------------------------------
voc_filtered_rows = df_voc_rows.copy()

# 날짜
if (
    date_range
    and isinstance(date_range, tuple)
    and len(date_range) == 2
):
    start_d, end_d = date_range
    if isinstance(start_d, date) and isinstance(end_d, date):
        voc_filtered_rows = voc_filtered_rows[
            (voc_filtered_rows["접수일시"] >= pd.to_datetime(start_d))
            & (
                voc_filtered_rows["접수일시"]
                < pd.to_datetime(end_d) + pd.Timedelta(days=1)
            )
        ]

# 지사
if sel_branches:
    voc_filtered_rows = voc_filtered_rows[
        voc_filtered_rows["관리지사"].isin(sel_branches)
    ]

# 리스크
if sel_risk:
    voc_filtered_rows = voc_filtered_rows[
        voc_filtered_rows["리스크등급"].isin(sel_risk)
    ]

# 매칭여부
if sel_match:
    voc_filtered_rows = voc_filtered_rows[
        voc_filtered_rows["매칭여부"].isin(sel_match)
    ]

# 월정료 구간
if fee_band_choice != "전체":
    if "월정료구간" in voc_filtered_rows.columns:
        voc_filtered_rows = voc_filtered_rows[
            voc_filtered_rows["월정료구간"] == fee_band_choice
        ]

# 계약요약(계약번호 기준)
df_contract = make_contract_summary(voc_filtered_rows)
df_contract = df_contract.reset_index(drop=True)

# 비매칭 계약요약
df_contract_unmatched = df_contract[
    df_contract["매칭여부"] == "비매칭(X)"
].copy()

# ----------------------------------------------------
# 8. 상단 KPI
# ----------------------------------------------------
st.markdown("## 📊 해지 VOC 종합 대시보드")

total_voc_rows = len(voc_filtered_rows)
total_contracts = df_contract["계약번호_정제"].nunique()
unmatched_contracts = df_contract_unmatched["계약번호_정제"].nunique()
matched_contracts = df_contract[
    df_contract["매칭여부"] == "매칭(O)"
]["계약번호_정제"].nunique()

k1, k2, k3, k4 = st.columns(4)
k1.metric("VOC 접수건수(행 기준)", f"{total_voc_rows:,}")
k2.metric("VOC 계약 수(계약 기준)", f"{total_contracts:,}")
k3.metric("비매칭(X) 계약 수", f"{unmatched_contracts:,}")
k4.metric("매칭(O) 계약 수", f"{matched_contracts:,}")

st.markdown("---")

# ----------------------------------------------------
# 9. 탭 구성
# ----------------------------------------------------
tab1, tab2, tab3 = st.tabs(
    [
        "📘 계약 기준 VOC 요약",
        "🚨 비매칭(X) 계약 상세",
        "📋 글로벌 VOC 리스트 (행 단위)",
    ]
)

# ----------------------------------------------------
# TAB 1 — 계약 기준 VOC 요약
# ----------------------------------------------------
with tab1:
    st.subheader("📘 계약 기준 VOC 요약 (최신 VOC 1건 + 접수건수)")

    if df_contract.empty:
        st.info("필터 조건에 해당하는 계약이 없습니다.")
    else:
        # 검색
        s1, s2, s3 = st.columns(3)
        q_cn = s1.text_input("계약번호 검색(부분)", key="t1_cn")
        q_name = s2.text_input("상호 검색(부분)", key="t1_name")
        q_addr = s3.text_input("시설_설치주소 검색(부분)", key="t1_addr")

        df_view = df_contract.copy()

        if q_cn:
            df_view = df_view[
                df_view["계약번호_정제"]
                .astype(str)
                .str.contains(q_cn.strip())
            ]
        if q_name and "상호" in df_view.columns:
            df_view = df_view[
                df_view["상호"].astype(str).str.contains(q_name.strip())
            ]
        if q_addr and "시설_설치주소" in df_view.columns:
            df_view = df_view[
                df_view["시설_설치주소"]
                .astype(str)
                .str.contains(q_addr.strip())
            ]

        # 표시 컬럼 후보 (None/NaN만 있는 컬럼은 자동 제외)
        candidate_cols = [
            "계약번호_정제",
            "상호",
            "관리지사",
            "구역담당자_통합",
            "리스크등급",
            "경과일수",
            "매칭여부",
            "접수건수",
            "시설_설치주소",
            "시설_KTT월정료(조정)",
            "월정료구간",
            "시설_계약상태(중)",
            "시설_서비스(소)",
        ]
        show_cols = []
        for c in candidate_cols:
            if c in df_view.columns:
                col_val = df_view[c]
                # None/NaN/빈문자만 있으면 제외
                non_null = col_val.dropna().astype(str)
                non_null = non_null[non_null.str.lower() != "none"]
                non_null = non_null[non_null != ""]
                if not non_null.empty:
                    show_cols.append(c)

        st.markdown(f"📌 표시 계약 수: **{len(df_view):,}건**")
        st.dataframe(
            style_risk(df_view[show_cols]),
            use_container_width=True,
            height=500,
        )

# ----------------------------------------------------
# TAB 2 — 비매칭(X) 계약 상세
# ----------------------------------------------------
with tab2:
    st.subheader("🚨 비매칭(X) 계약 기준 상세")

    if df_contract_unmatched.empty:
        st.info("현재 필터 조건에서 비매칭(X) 계약이 없습니다.")
    else:
        # 요약 리스트
        candidate_cols_u = [
            "계약번호_정제",
            "상호",
            "관리지사",
            "구역담당자_통합",
            "리스크등급",
            "경과일수",
            "매칭여부",
            "접수건수",
            "시설_설치주소",
            "시설_KTT월정료(조정)",
            "월정료구간",
            "시설_계약상태(중)",
            "시설_서비스(소)",
        ]
        show_cols_u = [
            c
            for c in candidate_cols_u
            if c in df_contract_unmatched.columns
        ]

        st.markdown(
            f"📌 필터 적용 후 비매칭(X) 계약 수: **{len(df_contract_unmatched):,}건**"
        )
        st.dataframe(
            style_risk(df_contract_unmatched[show_cols_u]),
            use_container_width=True,
            height=320,
        )

        st.markdown("---")
        st.markdown("### 📂 특정 계약 상세 이력")

        cn_list = (
            df_contract_unmatched["계약번호_정제"]
            .astype(str)
            .sort_values()
            .tolist()
        )
        sel_cn = st.selectbox(
            "상세 이력을 볼 계약 선택",
            options=["(선택)"] + cn_list,
            key="t2_cn_select",
        )

        if sel_cn != "(선택)":
            # 선택 계약의 VOC 행 이력 (필터 적용 후 데이터 기준)
            voc_hist = voc_filtered_rows[
                voc_filtered_rows["계약번호_정제"].astype(str) == sel_cn
            ].copy()
            voc_hist = voc_hist.sort_values("접수일시", ascending=False)

            # 기타 출처 이력 (해지시설/요청/설변/정지/해지파이프라인)
            other_hist = df_other_rows[
                df_other_rows["계약번호_정제"].astype(str) == sel_cn
            ].copy()

            base_info = (
                df_contract_unmatched[
                    df_contract_unmatched["계약번호_정제"].astype(str) == sel_cn
                ].iloc[0]
            )

            col1, col2, col3 = st.columns(3)
            col1.metric("상호", str(base_info.get("상호", "")))
            col2.metric("관리지사", str(base_info.get("관리지사", "")))
            col3.metric(
                "구역담당자",
                str(base_info.get("구역담당자_통합", "")),
            )

            c2_1, c2_2, c2_3 = st.columns(3)
            c2_1.metric("접수건수", f"{int(base_info.get('접수건수', 0)):,}건")
            c2_2.metric("리스크등급", str(base_info.get("리스크등급", "")))
            c2_3.metric("매칭여부", str(base_info.get("매칭여부", "")))

            st.caption(
                f"📍 시설_설치주소: {str(base_info.get('시설_설치주소', ''))}"
            )
            st.caption(
                f"💰 시설_KTT월정료(조정): {str(base_info.get('시설_KTT월정료(조정)', ''))}"
            )

            st.markdown("---")
            left, right = st.columns(2)

            with left:
                st.markdown("#### 📘 VOC 이력 (행 단위)")
                if voc_hist.empty:
                    st.info("VOC 이력이 없습니다.")
                else:
                    # 행 기준 표시 컬럼
                    row_cols = [
                        "계약번호_정제",
                        "출처",
                        "관리지사",
                        "구역담당자_통합",
                        "리스크등급",
                        "경과일수",
                        "매칭여부",
                        "접수일시",
                        "VOC유형",
                        "VOC유형중",
                        "VOC유형소",
                        "해지상세",
                        "등록내용",
                        "시설_설치주소",
                        "시설_KTT월정료(조정)",
                        "월정료구간",
                    ]
                    row_cols = [
                        c for c in row_cols if c in voc_hist.columns
                    ]
                    st.dataframe(
                        style_risk(voc_hist[row_cols]),
                        use_container_width=True,
                        height=320,
                    )

            with right:
                st.markdown("#### 📂 기타 출처 이력")
                if other_hist.empty:
                    st.info("기타 출처 데이터가 없습니다.")
                else:
                    st.dataframe(
                        other_hist,
                        use_container_width=True,
                        height=320,
                    )

            st.markdown("---")
            st.markdown("#### 📝 고객대응 / 현장 처리내역")

            fb_all = st.session_state["feedback_df"]
            fb_sel = fb_all[
                fb_all["계약번호_정제"].astype(str) == sel_cn
            ].copy()
            fb_sel = fb_sel.sort_values("등록일자", ascending=False)

            if fb_sel.empty:
                st.info("등록된 처리 이력이 없습니다.")
            else:
                st.dataframe(
                    fb_sel,
                    use_container_width=True,
                    height=220,
                )

            st.markdown("##### ✏️ 새 처리내용 등록")
            fb1, fb2 = st.columns([3, 1])
            new_fb = fb1.text_area("고객대응 / 현장 처리내용", key="t2_fb_content")
            new_user = fb2.text_input("등록자", key="t2_fb_user")
            new_note = fb2.text_input("비고", key="t2_fb_note")

            if st.button("💾 처리내역 저장", key="t2_fb_save"):
                if not new_fb.strip():
                    st.warning("처리내용을 입력하세요.")
                elif not new_user.strip():
                    st.warning("등록자를 입력하세요.")
                else:
                    new_row = pd.DataFrame(
                        [
                            {
                                "계약번호_정제": sel_cn,
                                "고객대응내용": new_fb.strip(),
                                "등록자": new_user.strip(),
                                "등록일자": datetime.now().strftime(
                                    "%Y-%m-%d %H:%M:%S"
                                ),
                                "비고": new_note.strip(),
                            }
                        ]
                    )
                    st.session_state["feedback_df"] = pd.concat(
                        [st.session_state["feedback_df"], new_row],
                        ignore_index=True,
                    )
                    save_feedback(FEEDBACK_PATH, st.session_state["feedback_df"])
                    st.success("저장 완료! 화면을 새로고침합니다.")
                    st.experimental_rerun()

# ----------------------------------------------------
# TAB 3 — 글로벌 VOC 리스트 (행 단위)
# ----------------------------------------------------
with tab3:
    st.subheader("📋 글로벌 VOC 리스트 (행 단위)")

    if voc_filtered_rows.empty:
        st.info("필터 조건에 해당하는 VOC 행 데이터가 없습니다.")
    else:
        st.markdown(
            f"📌 필터 적용 후 VOC 행 수: **{len(voc_filtered_rows):,}건**"
        )

        # 표시 컬럼 후보
        candidate_cols_row = [
            "계약번호_정제",
            "상호",
            "관리지사",
            "구역담당자_통합",
            "리스크등급",
            "경과일수",
            "매칭여부",
            "접수일시",
            "VOC유형",
            "VOC유형중",
            "VOC유형소",
            "해지상세",
            "등록내용",
            "시설_설치주소",
            "시설_KTT월정료(조정)",
            "월정료구간",
        ]
        row_cols = [
            c for c in candidate_cols_row if c in voc_filtered_rows.columns
        ]

        st.dataframe(
            style_risk(voc_filtered_rows[row_cols]),
            use_container_width=True,
            height=550,
        )

        st.download_button(
            "📥 현재 필터 기준 VOC 행 데이터 내려받기 (CSV)",
            voc_filtered_rows[row_cols]
            .to_csv(index=False)
            .encode("utf-8-sig"),
            file_name="VOC_글로벌행_필터적용.csv",
            mime="text/csv",
        )
