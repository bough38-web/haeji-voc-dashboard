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
    /* 데이터프레임 줄무늬 */
    .dataframe tbody tr:nth-child(odd) {
        background-color: #f9fafb;
    }
    .dataframe tbody tr:nth-child(even) {
        background-color: #eef2ff;
    }
    h2, h3, h4 {
        margin-top: 0.8rem;
        margin-bottom: 0.4rem;
    }
    </style>
    """,
    unsafe_allow_html=True,
)

# ----------------------------------------------------
# 1. 파일 경로
# ----------------------------------------------------
MERGED_PATH = "merged.xlsx"        # GitHub 루트에 위치
FEEDBACK_PATH = "feedback.csv"     # 계약번호별 피드백 저장용


# ----------------------------------------------------
# 2. 데이터 로딩
# ----------------------------------------------------
@st.cache_data
def load_data(path: str) -> pd.DataFrame:
    if not os.path.exists(path):
        st.error("❌ 'merged.xlsx' 파일을 찾을 수 없습니다. 저장소 루트에 위치하는지 확인하세요.")
        return pd.DataFrame()

    df = pd.read_excel(path)

    # 숫자형 컬럼 콤마 제거
    for col in ["계약번호", "고객번호"]:
        if col in df.columns:
            df[col] = (
                df[col]
                .astype(str)
                .str.replace(",", "")
                .str.strip()
            )

    # 출처 정제
    df["출처"] = df["출처"].replace({"고객리스트": "해지시설"})

    # 계약번호 정제
    df["계약번호_정제"] = (
        df["계약번호"]
        .astype(str)
        .str.replace(r"[^0-9A-Za-z]", "", regex=True)
        .str.strip()
    )

    # 접수일시 → datetime
    if "접수일시" in df.columns:
        df["접수일시"] = pd.to_datetime(df["접수일시"], errors="coerce")

    return df


def load_feedback(path: str) -> pd.DataFrame:
    """계약번호 단위 피드백 저장용 CSV 로드"""
    if os.path.exists(path):
        try:
            fb = pd.read_csv(path, encoding="utf-8-sig")
        except Exception:
            fb = pd.read_csv(path)
    else:
        fb = pd.DataFrame(
            columns=["계약번호_정제", "고객대응내용", "등록자", "등록일자", "비고"]
        )
    return fb


def save_feedback(path: str, fb_df: pd.DataFrame) -> None:
    fb_df.to_csv(path, index=False, encoding="utf-8-sig")


df = load_data(MERGED_PATH)
if df.empty:
    st.stop()

# 세션에 피드백 적재
if "feedback_df" not in st.session_state:
    st.session_state["feedback_df"] = load_feedback(FEEDBACK_PATH)


# ----------------------------------------------------
# 3. 지사명 축약 & 정렬 순서
# ----------------------------------------------------
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

BRANCH_ORDER = ["중앙", "강북", "서대문", "고양", "의정부", "남양주", "강릉", "원주"]


def sort_branch(series):
    return sorted(
        [s for s in series if s in BRANCH_ORDER],
        key=lambda x: BRANCH_ORDER.index(x),
    )


# ----------------------------------------------------
# 4. 통합 구역/담당자 / 주소 컬럼
# ----------------------------------------------------
def make_zone(row):
    if "영업구역번호" in row and pd.notna(row["영업구역번호"]):
        return row["영업구역번호"]
    if "담당상세" in row and pd.notna(row["담당상세"]):
        return row["담당상세"]
    return ""


df["영업구역_통합"] = df.apply(make_zone, axis=1)

mgr_priority = ["구역담당자", "담당자", "처리자"]


def pick_manager(row):
    for c in mgr_priority:
        if c in row and pd.notna(row[c]) and str(row[c]).strip() != "":
            return row[c]
    return ""


df["구역담당자_통합"] = df.apply(pick_manager, axis=1)

# 주소 컬럼 자동 탐색
address_cols = [c for c in df.columns if "주소" in c]


# ----------------------------------------------------
# 5. 출처 분리 + 매칭 계산
# ----------------------------------------------------
df_voc = df[df["출처"] == "해지VOC"].copy()
df_other = df[df["출처"] != "해지VOC"].copy()

other_sets = {
    src: set(df_other[df_other["출처"] == src]["계약번호_정제"].dropna())
    for src in ["해지시설", "해지요청", "설변", "정지", "해지파이프라인"]
}
other_union = set().union(*other_sets.values()) if other_sets else set()

df_voc["매칭여부"] = df_voc["계약번호_정제"].apply(
    lambda x: "매칭(O)" if x in other_union else "비매칭(X)"
)


# ----------------------------------------------------
# 6. 리스크 등급/경과일 계산 (요청 기준)
#  - 최근 3일 : HIGH
#  - 3일 초과 ~ 10일 이하 : MEDIUM
#  - 10일 초과 : LOW
# ----------------------------------------------------
today = date.today()


def compute_risk(row):
    dt = row.get("접수일시")
    if pd.isna(dt):
        return np.nan, "LOW"  # 날짜 없으면 낮음으로 처리

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

df_unmatched = df_voc[df_voc["매칭여부"] == "비매칭(X)"].copy()

# ----------------------------------------------------
# 7. 공통 표시 컬럼
# ----------------------------------------------------
fixed_order = [
    "상호",
    "계약번호_정제",
    "매칭여부",
    "리스크등급",
    "경과일수",
    "출처",
    "관리지사",
    "영업구역번호",
    "영업구역_통합",
    "구역담당자_통합",
    "처리자",
    "담당유형",
    "처리유형",
    "처리내용",
    "접수일시",
    "서비스개시일",
    "계약종료일",
    "서비스중",
    "서비스소",
    "VOC유형",
    "VOC유형중",
    "VOC유형소",
    "해지상세",
    "등록내용",
]
display_cols = [c for c in fixed_order if c in df_voc.columns]


# ----------------------------------------------------
# 8. 스타일링 (리스크 등급 색상 강조)
# ----------------------------------------------------
def style_risk(df_view: pd.DataFrame):
    if "리스크등급" not in df_view.columns:
        return df_view

    def _row_style(row):
        level = row.get("리스크등급", "")
        if level == "HIGH":
            bg = "#fee2e2"  # red-100
        elif level == "MEDIUM":
            bg = "#fef3c7"  # amber-100
        else:
            bg = "#e0f2fe"  # sky-100
        return [f"background-color: {bg};"] * len(row)

    return df_view.style.apply(_row_style, axis=1)


# ----------------------------------------------------
# 9. 사이드바 글로벌 필터
# ----------------------------------------------------
st.sidebar.title("🔧 글로벌 필터")

# 날짜 범위 필터
if "접수일시" in df_voc.columns and df_voc["접수일시"].notna().any():
    min_d = df_voc["접수일시"].min().date()
    max_d = df_voc["접수일시"].max().date()
    dr = st.sidebar.date_input(
        "접수일자 범위",
        value=(min_d, max_d),
        min_value=min_d,
        max_value=max_d,
        key="global_date_range",
    )
else:
    dr = None

# 지사 필터
branches_all = sort_branch(df_voc["관리지사"].dropna().unique())
sel_branches = st.sidebar.multiselect(
    "관리지사(복수 선택)",
    options=branches_all,
    default=branches_all,
    key="global_branches",
)

# 리스크 등급 필터
risk_all = ["HIGH", "MEDIUM", "LOW"]
sel_risk = st.sidebar.multiselect(
    "리스크등급",
    options=risk_all,
    default=risk_all,
    key="global_risk",
)

# 매칭여부 필터
match_all = ["매칭(O)", "비매칭(X)"]
sel_match = st.sidebar.multiselect(
    "매칭여부",
    options=match_all,
    default=match_all,
    key="global_match",
)

st.sidebar.markdown("---")
st.sidebar.caption(
    f"마지막 갱신: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}"
)

# ----------------------------------------------------
# 10. 글로벌 필터 적용
# ----------------------------------------------------
voc_filtered_global = df_voc.copy()

# 날짜
if dr and isinstance(dr, tuple) and len(dr) == 2:
    start_d, end_d = dr
    if isinstance(start_d, date) and isinstance(end_d, date):
        voc_filtered_global = voc_filtered_global[
            (voc_filtered_global["접수일시"] >= pd.to_datetime(start_d))
            & (voc_filtered_global["접수일시"] < pd.to_datetime(end_d) + pd.Timedelta(days=1))
        ]

# 지사
if sel_branches:
    voc_filtered_global = voc_filtered_global[
        voc_filtered_global["관리지사"].isin(sel_branches)
    ]

# 리스크
if sel_risk:
    voc_filtered_global = voc_filtered_global[
        voc_filtered_global["리스크등급"].isin(sel_risk)
    ]

# 매칭여부
if sel_match:
    voc_filtered_global = voc_filtered_global[
        voc_filtered_global["매칭여부"].isin(sel_match)
    ]

unmatched_global = voc_filtered_global[
    voc_filtered_global["매칭여부"] == "비매칭(X)"
].copy()

# ----------------------------------------------------
# 11. 상단 KPI 카드
# ----------------------------------------------------
st.markdown("## 📊 해지 VOC 종합 대시보드")

total_rows = len(voc_filtered_global)
unique_cn = voc_filtered_global["계약번호_정제"].nunique()
unmatched_contracts = (
    voc_filtered_global[voc_filtered_global["매칭여부"] == "비매칭(X)"]["계약번호_정제"]
    .nunique()
)
matched_contracts = (
    voc_filtered_global[voc_filtered_global["매칭여부"] == "매칭(O)"]["계약번호_정제"]
    .nunique()
)

k1, k2, k3, k4 = st.columns(4)
k1.metric("VOC 접수건수", f"{total_rows:,}")
k2.metric("VOC 계약 수(유니크)", f"{unique_cn:,}")
k3.metric("비매칭(X) 계약 수", f"{unmatched_contracts:,}")
k4.metric("매칭(O) 계약 수", f"{matched_contracts:,}")

st.markdown("---")

# ----------------------------------------------------
# 12. 탭 구성
# ----------------------------------------------------
tab1, tab2, tab3, tab4 = st.tabs(
    ["📘 VOC 전체(계약 기준)", "🚨 비매칭(활동대상)", "📊 지사/담당자 현황", "🔍 계약별 드릴다운"]
)

# ====================================================
# TAB 1 — VOC 전체 (계약번호 기준 요약)
# ====================================================
with tab1:
    st.subheader("📘 VOC 전체 (계약번호 기준 요약)")

    # 지사 / 담당자 버튼식 필터
    row1_col1, row1_col2 = st.columns([2, 3])

    branches_for_tab1 = ["전체"] + sort_branch(
        voc_filtered_global["관리지사"].dropna().unique()
    )
    selected_branch_tab1 = row1_col1.radio(
        "지사 선택",
        options=branches_for_tab1,
        horizontal=True,
        key="tab1_branch_radio",
    )

    temp_for_mgr = voc_filtered_global.copy()
    if selected_branch_tab1 != "전체":
        temp_for_mgr = temp_for_mgr[
            temp_for_mgr["관리지사"] == selected_branch_tab1
        ]

    mgr_options_tab1 = (
        ["전체"]
        + sorted(
            temp_for_mgr["구역담당자_통합"]
            .dropna()
            .astype(str)
            .unique()
            .tolist()
        )
        if "구역담당자_통합" in temp_for_mgr.columns
        else ["전체"]
    )

    selected_mgr_tab1 = row1_col2.radio(
        "담당자 선택",
        options=mgr_options_tab1,
        horizontal=True,
        key="tab1_mgr_radio",
    )

    # 검색 입력
    s1, s2, s3 = st.columns(3)
    q_cn = s1.text_input("계약번호 검색(부분)", key="tab1_cn")
    q_name = s2.text_input("상호 검색(부분)", key="tab1_name")
    q_addr = s3.text_input("주소 검색(부분)", key="tab1_addr")

    temp = voc_filtered_global.copy()

    if selected_branch_tab1 != "전체":
        temp = temp[temp["관리지사"] == selected_branch_tab1]
    if selected_mgr_tab1 != "전체":
        temp = temp[temp["구역담당자_통합"].astype(str) == selected_mgr_tab1]

    if q_cn:
        temp = temp[
            temp["계약번호_정제"].astype(str).str.contains(q_cn.strip())
        ]
    if q_name and "상호" in temp.columns:
        temp = temp[
            temp["상호"].astype(str).str.contains(q_name.strip())
        ]
    if q_addr and address_cols:
        cond = False
        for col in address_cols:
            cond |= temp[col].astype(str).str.contains(q_addr.strip())
        temp = temp[cond]

    # 계약번호 기준 요약 (최신 VOC + 접수건수)
    if temp.empty:
        st.info("조건에 맞는 VOC 데이터가 없습니다.")
    else:
        temp_sorted = temp.sort_values("접수일시", ascending=False)
        grp = temp_sorted.groupby("계약번호_정제")
        idx_latest = grp["접수일시"].idxmax()
        df_summary = temp_sorted.loc[idx_latest].copy()
        df_summary["접수건수"] = grp.size().reindex(df_summary["계약번호_정제"]).values

        summary_cols = [
            "계약번호_정제",
            "상호",
            "관리지사",
            "구역담당자_통합",
            "리스크등급",
            "경과일수",
            "매칭여부",
            "접수건수",
        ]
        summary_cols = [c for c in summary_cols if c in df_summary.columns]

        st.markdown(f"**표시 계약 수:** {len(df_summary):,} 건")
        st.dataframe(
            style_risk(df_summary[summary_cols]),
            use_container_width=True,
            height=480,
        )

# ====================================================
# TAB 2 — 비매칭(X) 활동대상 (계약번호 기준)
# ====================================================
with tab2:
    st.subheader("🚨 비매칭(X) 활동대상 (계약번호 기준)")

    if unmatched_global.empty:
        st.info("현재 글로벌 필터 조건에서 비매칭(X) 계약이 없습니다.")
    else:
        # 지사 / 담당자 버튼식 필터
        u_col1, u_col2 = st.columns([2, 3])

        branches_u = ["전체"] + sort_branch(
            unmatched_global["관리지사"].dropna().unique()
        )
        selected_branch_u = u_col1.radio(
            "지사 선택",
            options=branches_u,
            horizontal=True,
            key="tab2_branch_radio",
        )

        temp_u_for_mgr = unmatched_global.copy()
        if selected_branch_u != "전체":
            temp_u_for_mgr = temp_u_for_mgr[
                temp_u_for_mgr["관리지사"] == selected_branch_u
            ]

        mgr_options_u = (
            ["전체"]
            + sorted(
                temp_u_for_mgr["구역담당자_통합"]
                .dropna()
                .astype(str)
                .unique()
                .tolist()
            )
            if "구역담당자_통합" in temp_u_for_mgr.columns
            else ["전체"]
        )

        selected_mgr_u = u_col2.radio(
            "담당자 선택",
            options=mgr_options_u,
            horizontal=True,
            key="tab2_mgr_radio",
        )

        # 추가 검색 (계약번호/상호)
        us1, us2 = st.columns(2)
        uq_cn = us1.text_input("계약번호 검색(부분)", key="tab2_cn")
        uq_name = us2.text_input("상호 검색(부분)", key="tab2_name")

        temp_u = unmatched_global.copy()
        if selected_branch_u != "전체":
            temp_u = temp_u[temp_u["관리지사"] == selected_branch_u]
        if selected_mgr_u != "전체":
            temp_u = temp_u[temp_u["구역담당자_통합"].astype(str) == selected_mgr_u]

        if uq_cn:
            temp_u = temp_u[
                temp_u["계약번호_정제"].astype(str).str.contains(uq_cn.strip())
            ]
        if uq_name and "상호" in temp_u.columns:
            temp_u = temp_u[
                temp_u["상호"].astype(str).str.contains(uq_name.strip())
            ]

        if temp_u.empty:
            st.info("조건에 맞는 비매칭(X) 계약이 없습니다.")
        else:
            temp_u_sorted = temp_u.sort_values("접수일시", ascending=False)
            grp_u = temp_u_sorted.groupby("계약번호_정제")
            idx_latest_u = grp_u["접수일시"].idxmax()
            df_u_summary = temp_u_sorted.loc[idx_latest_u].copy()
            df_u_summary["접수건수"] = grp_u.size().reindex(
                df_u_summary["계약번호_정제"]
            ).values

            summary_cols_u = [
                "계약번호_정제",
                "상호",
                "관리지사",
                "구역담당자_통합",
                "리스크등급",
                "경과일수",
                "접수건수",
            ]
            summary_cols_u = [c for c in summary_cols_u if c in df_u_summary.columns]

            st.markdown(
                f"⚠ 활동대상 비매칭(X) 계약 수: **{len[df_u_summary]:,} 건**"
            )
            st.dataframe(
                style_risk(df_u_summary[summary_cols_u]),
                use_container_width=True,
                height=420,
            )

            # 🔽 계약번호 상세 이력 보기
            st.markdown("---")
            st.markdown("### 📂 선택한 계약번호 상세 VOC 이력")

            u_contract_list = df_u_summary["계약번호_정제"].astype(str).tolist()
            sel_u_contract = st.selectbox(
                "상세 VOC 이력을 볼 계약 선택",
                options=["(선택)"] + u_contract_list,
                key="tab2_select_contract",
            )

            if sel_u_contract != "(선택)":
                voc_detail = temp_u[
                    temp_u["계약번호_정제"].astype(str) == sel_u_contract
                ].copy()
                voc_detail = voc_detail.sort_values("접수일시", ascending=False)

                st.markdown(
                    f"#### 🔍 `{sel_u_contract}` VOC 상세 이력 ({len(voc_detail)}건)"
                )
                st.dataframe(
                    style_risk(voc_detail[display_cols]),
                    use_container_width=True,
                    height=350,
                )

            # 내려받기 (행 단위 전체)
            st.download_button(
                "📥 비매칭(X) 원천 VOC 행 기준 다운로드 (CSV)",
                temp_u.to_csv(index=False).encode("utf-8-sig"),
                file_name="비매칭_활동대상_원천행.csv",
                mime="text/csv",
            )

# ====================================================
# TAB 3 — 지사/담당자 시각화
# ====================================================
with tab3:
    st.subheader("📊 지사 / 담당자별 비매칭 리스크 현황")

    if unmatched_global.empty:
        st.info("비매칭(X) 데이터가 없습니다.")
    else:
        c1, c2, c3 = st.columns(3)

        # 지사별 비매칭 계약 수
        bc = (
            unmatched_global.groupby("관리지사")["계약번호_정제"]
            .nunique()
            .rename("비매칭계약수")
        )
        bc = bc[bc.index.isin(BRANCH_ORDER)].reindex(BRANCH_ORDER).dropna()

        with c1:
            st.markdown("#### 🏢 지사별 비매칭 계약 수")
            st.bar_chart(bc, use_container_width=True)

        # 담당자별 TOP 15
        mc = (
            unmatched_global.groupby("구역담당자_통합")["계약번호_정제"]
            .nunique()
            .rename("비매칭계약수")
            .sort_values(ascending=False)
        )
        mc = mc[mc.index.astype(str).str.strip() != ""].head(15)

        with c2:
            st.markdown("#### 👤 담당자별 비매칭 TOP 15")
            st.bar_chart(mc, use_container_width=True)

        # 리스크 분포
        rc = (
            unmatched_global["리스크등급"]
            .value_counts()
            .reindex(["HIGH", "MEDIUM", "LOW"])
            .fillna(0)
        )

        with c3:
            st.markdown("#### 🔥 리스크 등급 분포 (비매칭, 계약 단위)")
            st.bar_chart(rc, use_container_width=True)

        st.markdown("---")

        # 일별 비매칭 추이
        if "접수일시" in unmatched_global.columns:
            trend = (
                unmatched_global.assign(접수일=unmatched_global["접수일시"].dt.date)
                .groupby("접수일")["계약번호_정제"]
                .nunique()
                .rename("비매칭계약수")
                .sort_index()
            )
            st.markdown("#### 📈 일별 비매칭 계약 추이")
            st.line_chart(trend, use_container_width=True)

# ====================================================
# TAB 4 — 계약별 드릴다운 (계약번호 단위로 그룹)
# ====================================================
with tab4:
    st.subheader("🔍 계약번호 기준 통합 드릴다운")

    base_all = voc_filtered_global.copy()

    # 매칭여부 필터 추가
    match_choice = st.radio(
        "매칭여부 선택",
        options=["전체", "매칭(O)", "비매칭(X)"],
        horizontal=True,
        key="tab4_match_radio",
    )

    drill_base = base_all.copy()
    if match_choice == "매칭(O)":
        drill_base = drill_base[drill_base["매칭여부"] == "매칭(O)"]
    elif match_choice == "비매칭(X)":
        drill_base = drill_base[drill_base["매칭여부"] == "비매칭(X)"]

    # 지사 / 담당자 필터
    d1, d2 = st.columns([2, 3])
    branches_d = ["전체"] + sort_branch(drill_base["관리지사"].dropna().unique())
    sel_branch_d = d1.radio(
        "지사 선택",
        options=branches_d,
        horizontal=True,
        key="tab4_branch_radio",
    )

    tmp_mgr_d = drill_base.copy()
    if sel_branch_d != "전체":
        tmp_mgr_d = tmp_mgr_d[tmp_mgr_d["관리지사"] == sel_branch_d]

    mgr_options_d = (
        ["전체"]
        + sorted(
            tmp_mgr_d["구역담당자_통합"]
            .dropna()
            .astype(str)
            .unique()
            .tolist()
        )
        if "구역담당자_통합" in tmp_mgr_d.columns
        else ["전체"]
    )

    sel_mgr_d = d2.radio(
        "담당자 선택",
        options=mgr_options_d,
        horizontal=True,
        key="tab4_mgr_radio",
    )

    # 계약번호 / 상호 검색
    dd1, dd2 = st.columns(2)
    dq_cn = dd1.text_input("계약번호 검색(부분)", key="tab4_cn")
    dq_name = dd2.text_input("상호 검색(부분)", key="tab4_name")

    drill = drill_base.copy()
    if sel_branch_d != "전체":
        drill = drill[drill["관리지사"] == sel_branch_d]
    if sel_mgr_d != "전체":
        drill = drill[drill["구역담당자_통합"].astype(str) == sel_mgr_d]

    if dq_cn:
        drill = drill[
            drill["계약번호_정제"].astype(str).str.contains(dq_cn.strip())
        ]
    if dq_name and "상호" in drill.columns:
        drill = drill[
            drill["상호"].astype(str).str.contains(dq_name.strip())
        ]

    if drill.empty:
        st.info("조건에 맞는 계약이 없습니다. 필터를 조정해보세요.")
    else:
        drill_sorted = drill.sort_values("접수일시", ascending=False)
        g = drill_sorted.groupby("계약번호_정제")
        idx_latest_d = g["접수일시"].idxmax()
        df_d_summary = drill_sorted.loc[idx_latest_d].copy()
        df_d_summary["접수건수"] = g.size().reindex(
            df_d_summary["계약번호_정제"]
        ).values

        sum_cols_d = [
            "계약번호_정제",
            "상호",
            "관리지사",
            "구역담당자_통합",
            "리스크등급",
            "경과일수",
            "매칭여부",
            "접수건수",
        ]
        sum_cols_d = [c for c in sum_cols_d if c in df_d_summary.columns]

        st.markdown("#### 📋 계약 요약 (최신 VOC 기준, 계약번호당 1행)")
        st.dataframe(
            style_risk(df_d_summary[sum_cols_d]),
            use_container_width=True,
            height=260,
        )

        # 선택 계약번호
        cn_list = df_d_summary["계약번호_정제"].astype(str).tolist()

        def format_cn(cn_value: str) -> str:
            row = df_d_summary[
                df_d_summary["계약번호_정제"].astype(str) == str(cn_value)
            ].iloc[0]
            name = row.get("상호", "")
            branch = row.get("관리지사", "")
            cnt = row.get("접수건수", 0)
            return f"{cn_value} | {name} | {branch} | 접수 {int(cnt)}건"

        sel_cn = st.selectbox(
            "상세를 볼 계약 선택",
            options=cn_list,
            format_func=format_cn,
            key="tab4_cn_selectbox",
        )

        if sel_cn:
            # VOC 이력 (해당 계약 모든 이력)
            voc_hist = df_voc[
                df_voc["계약번호_정제"].astype(str) == str(sel_cn)
            ].copy()
            voc_hist = voc_hist.sort_values("접수일시", ascending=False)

            # 기타 출처 이력
            other_hist = df_other[
                df_other["계약번호_정제"].astype(str) == str(sel_cn)
            ].copy()

            # 기본 정보
            base_info = voc_hist.iloc[0] if not voc_hist.empty else None

            st.markdown(f"### 🔎 선택된 계약번호: `{sel_cn}`")

            if base_info is not None:
                info_col1, info_col2, info_col3 = st.columns(3)
                info_col1.metric("상호", str(base_info.get("상호", "")))
                info_col2.metric("관리지사", str(base_info.get("관리지사", "")))
                info_col3.metric(
                    "구역담당자",
                    str(base_info.get("구역담당자_통합", base_info.get("처리자", ""))),
                )

                m2_1, m2_2, m2_3 = st.columns(3)
                m2_1.metric("접수건수", f"{len(voc_hist):,}건")
                m2_2.metric("리스크등급", str(base_info.get("리스크등급", "")))
                m2_3.metric("매칭여부", str(base_info.get("매칭여부", "")))

            st.markdown("---")

            c_left, c_right = st.columns(2)

            with c_left:
                st.markdown("#### 📘 VOC 이력 (전체)")
                if voc_hist.empty:
                    st.info("VOC 이력이 없습니다.")
                else:
                    st.dataframe(
                        style_risk(voc_hist[display_cols]),
                        use_container_width=True,
                        height=320,
                    )

            with c_right:
                st.markdown("#### 📂 기타 출처 이력 (해지시설/요청/설변/정지/파이프라인)")
                if other_hist.empty:
                    st.info("기타 출처 데이터가 없습니다.")
                else:
                    st.dataframe(
                        other_hist,
                        use_container_width=True,
                        height=320,
                    )

            st.markdown("---")

            # 피드백 이력 & 입력
            st.markdown("#### 📝 고객대응 / 현장 처리내역")

            fb_all = st.session_state["feedback_df"]
            fb_sel = fb_all[
                fb_all["계약번호_정제"].astype(str) == str(sel_cn)
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
            new_fb = fb1.text_area("고객대응 / 현장 처리내용", key="fb_content")
            new_user = fb2.text_input("등록자", key="fb_user")
            new_note = fb2.text_input("비고", key="fb_note")

            if st.button("💾 처리내역 저장", key="fb_save_btn"):
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
                    # CSV 저장
                    save_feedback(FEEDBACK_PATH, st.session_state["feedback_df"])
                    st.success("처리내역이 저장되었습니다.")
                    st.rerun()

            st.markdown("---")

            # 다운로드
            dcol1, dcol2 = st.columns(2)

            if not voc_hist.empty:
                dcol1.download_button(
                    "📥 선택 계약 VOC 이력 다운로드 (CSV)",
                    voc_hist.to_csv(index=False).encode("utf-8-sig"),
                    file_name=f"VOC이력_{sel_cn}.csv",
                    mime="text/csv",
                )

            # VOC + 기타 + 피드백 통합 내려받기 (구분 컬럼 추가)
            export_frames = []

            if not voc_hist.empty:
                v_exp = voc_hist.copy()
                v_exp.insert(0, "구분", "VOC")
                export_frames.append(v_exp)

            if not other_hist.empty:
                o_exp = other_hist.copy()
                o_exp.insert(0, "구분", "기타출처")
                export_frames.append(o_exp)

            if not fb_sel.empty:
                f_exp = fb_sel.copy()
                f_exp.insert(0, "구분", "피드백")
                export_frames.append(f_exp)

            if export_frames:
                merged_export = pd.concat(export_frames, ignore_index=True)
                dcol2.download_button(
                    "📥 선택 계약 통합 이력 다운로드 (CSV)",
                    merged_export.to_csv(index=False).encode("utf-8-sig"),
                    file_name=f"통합이력_{sel_cn}.csv",
                    mime="text/csv",
                )
