import os
from datetime import datetime, date

import numpy as np
import pandas as pd
import streamlit as st

# ----------------------------------------------------
# 0. 기본 설정
# ----------------------------------------------------
st.set_page_config(page_title="해지 VOC 종합 대시보드", layout="wide")

# ----------------------------------------------------
# 글로벌 라이트 테마용 간단 스타일 (다크모드여도 표시는 밝게)
# ----------------------------------------------------
st.markdown(
    """
    <style>
    .stApp {
        background-color: #f8fafc;
        color: #111827;
    }
    [data-testid="stHeader"] {
        background-color: #f8fafc;
    }
    /* 데이터프레임 기본 배경 톤 다운 */
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
# 1. 데이터 로딩 (GitHub 내 merged.xlsx 사용)
# ----------------------------------------------------
@st.cache_data
def load_data() -> pd.DataFrame:
    file_path = "merged.xlsx"  # GitHub repo 루트에 있어야 함

    if not os.path.exists(file_path):
        st.error(f"❌ GitHub 저장소에서 'merged.xlsx' 파일을 찾을 수 없습니다.")
        return pd.DataFrame()

    df = pd.read_excel(file_path)

    # 숫자형 컬럼 콤마 제거 (계약번호/고객번호)
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


df = load_data()
if df.empty:
    st.stop()

# ----------------------------------------------------
# 2. 지사명 축약 & 정렬 순서
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
# 3. 통합 구역/담당자 컬럼 생성
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
# 4. 출처 분리 + 매칭 계산
# ----------------------------------------------------
df_voc = df[df["출처"] == "해지VOC"].copy()
df_other = df[df["출처"] != "해지VOC"].copy()

other_sets = {
    src: set(df_other[df_other["출처"] == src]["계약번호_정제"].dropna())
    for src in ["해지시설", "해지요청", "설변", "정지", "해지파이프라인"]
}

other_union = set().union(*other_sets.values())

# VOC ∧ 기타 출처 있음 → 매칭(O)
df_voc["매칭여부"] = df_voc["계약번호_정제"].apply(
    lambda x: "매칭(O)" if x in other_union else "비매칭(X)"
)


# ----------------------------------------------------
# 5. 리스크 등급/경과일 계산
# ----------------------------------------------------
today = date.today()


def compute_risk(row):
    dt = row.get("접수일시")
    if pd.isna(dt):
        return np.nan, "MEDIUM"

    days = (today - dt.date()).days

    if days <= 7:
        level = "HIGH"
    elif days <= 30:
        level = "MEDIUM"
    else:
        level = "LOW"

    hs = str(row.get("해지상세", "") or "")
    if any(k in hs for k in ["즉시", "강성", "불만"]):
        if level == "MEDIUM":
            level = "HIGH"

    return days, level


df_voc["경과일수"], df_voc["리스크등급"] = zip(
    *df_voc.apply(lambda r: compute_risk(r), axis=1)
)

df_unmatched = df_voc[df_voc["매칭여부"] == "비매칭(X)"].copy()

# ----------------------------------------------------
# 6. 공통 표시 컬럼 정의
# ----------------------------------------------------
exclude_cols = {
    "기타출처",
    "담당상세",
    "구역담당자_통합",
    "계약번호",
    "고객번호",
    "고객번호_정제",
    "고객명",
    "설치주소",
    "청구주소",
    "주소",
}

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
# 7. 스타일링 (리스크 등급 색상 강조)
# ----------------------------------------------------
def style_risk(df_view: pd.DataFrame):
    if "리스크등급" not in df_view.columns:
        return df_view

    def _row_style(row):
        level = row.get("리스크등급", "")
        if level == "HIGH":
            bg = "#ffe5e5"
        elif level == "MEDIUM":
            bg = "#fff6e5"
        else:
            bg = "#e5f7ff"
        return [f"background-color: {bg};"] * len(row)

    return df_view.style.apply(_row_style, axis=1)


# ----------------------------------------------------
# 8. 사이드바 - 글로벌 필터 (엔터프라이즈식)
# ----------------------------------------------------
st.sidebar.title("🔧 글로벌 필터")

# 날짜 범위
if "접수일시" in df_voc.columns and df_voc["접수일시"].notna().any():
    min_d = df_voc["접수일시"].min().date()
    max_d = df_voc["접수일시"].max().date()
    dr = st.sidebar.date_input(
        "접수일자 범위",
        value=(min_d, max_d),
        min_value=min_d,
        max_value=max_d,
    )
else:
    dr = None

# 지사
branches_all = sort_branch(df_voc["관리지사"].dropna().unique())
sel_branches = st.sidebar.multiselect(
    "관리지사(복수 선택)",
    options=branches_all,
    default=branches_all,
)

# 리스크 등급
risk_all = ["HIGH", "MEDIUM", "LOW"]
sel_risk = st.sidebar.multiselect(
    "리스크등급",
    options=risk_all,
    default=risk_all,
)

# 매칭여부 (전역)
match_all = ["매칭(O)", "비매칭(X)"]
sel_match = st.sidebar.multiselect(
    "매칭여부",
    options=match_all,
    default=match_all,
)

st.sidebar.markdown("---")
st.sidebar.caption(
    f"마지막 갱신: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}"
)

# ----------------------------------------------------
# 9. 글로벌 필터 적용
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

# 매칭 여부
if sel_match:
    voc_filtered_global = voc_filtered_global[
        voc_filtered_global["매칭여부"].isin(sel_match)
    ]

unmatched_global = voc_filtered_global[
    voc_filtered_global["매칭여부"] == "비매칭(X)"
].copy()

# ----------------------------------------------------
# 10. 헤더 & 상단 KPI
# ----------------------------------------------------
st.markdown("## 📊 해지 VOC 종합 대시보드 (엔터프라이즈 버전)")

c1, c2, c3, c4 = st.columns(4)

total_voc = len(voc_filtered_global)
unique_cn = voc_filtered_global["계약번호_정제"].nunique()
unmatched_cnt = (voc_filtered_global["매칭여부"] == "비매칭(X)").sum()
matched_cnt = (voc_filtered_global["매칭여부"] == "매칭(O)").sum()

c1.metric("전체 VOC 건수", f"{total_voc:,}")
c2.metric("유니크 계약 수", f"{unique_cn:,}")
c3.metric("비매칭(X) 활동대상", f"{unmatched_cnt:,}")
c4.metric("매칭(O) 건수", f"{matched_cnt:,}")

st.markdown("---")

# ----------------------------------------------------
# 11. 탭 구성
# ----------------------------------------------------
tab1, tab2, tab3, tab4 = st.tabs(
    ["📘 VOC 전체", "🚨 비매칭(활동대상)", "📊 지사/담당자 현황", "🔍 계약번호 드릴다운"]
)

# ====================================================
# TAB 1 — VOC 전체 조회
# ====================================================
with tab1:
    st.subheader("📘 VOC 전체 조회 (글로벌 필터 적용 상태)")

    c1, c2, c3 = st.columns(3)
    key_cn = c1.text_input("계약번호 검색")
    key_name = c2.text_input("상호 검색")
    key_addr = c3.text_input("주소 검색")

    temp = voc_filtered_global.copy()

    if key_cn:
        key = key_cn.strip()
        temp = temp[temp["계약번호_정제"].astype(str).str.contains(key)]

    if key_name and "상호" in temp.columns:
        key = key_name.strip()
        temp = temp[temp["상호"].astype(str).str.contains(key)]

    if key_addr and address_cols:
        key = key_addr.strip()
        cond = False
        for col in address_cols:
            cond |= temp[col].astype(str).str.contains(key)
        temp = temp[cond]

    temp = temp.sort_values("접수일시", ascending=False)

    st.write(f"표시 건수: **{len(temp):,} 건**")

    st.dataframe(
        style_risk(temp[display_cols]),
        use_container_width=True,
        height=520,
    )

# ====================================================
# TAB 2 — 비매칭(X) = 활동대상
# ====================================================
with tab2:
    st.subheader("🚨 비매칭(X) 활동대상 리스크 리스트")

    if unmatched_global.empty:
        st.info("현재 글로벌 필터 조건에서 비매칭(X) 데이터가 없습니다.")
    else:
        b1, b2, b3 = st.columns(3)

        # 지사 선택 (글로벌 필터 이후 남아있는 지사)
        ub_branches = sort_branch(
            unmatched_global["관리지사"].dropna().unique()
        )
        sel_b = b1.multiselect(
            "지사 선택 (추가 필터)",
            options=ub_branches,
            default=ub_branches,
        )

        # 담당자 선택
        tmp_branch = unmatched_global.copy()
        if sel_b:
            tmp_branch = tmp_branch[tmp_branch["관리지사"].isin(sel_b)]

        mgr_opts = sorted(
            tmp_branch["구역담당자_통합"]
            .dropna()
            .astype(str)
            .unique()
            .tolist()
        )
        sel_mgr = b2.multiselect(
            "담당자 선택",
            options=mgr_opts,
            default=mgr_opts,
        )

        # 리스크 추가 필터
        risk_opts = ["HIGH", "MEDIUM", "LOW"]
        sel_r2 = b3.multiselect(
            "리스크 등급(추가 필터)",
            options=risk_opts,
            default=risk_opts,
        )

        temp = unmatched_global.copy()
        if sel_b:
            temp = temp[temp["관리지사"].isin(sel_b)]
        if sel_mgr:
            temp = temp[temp["구역담당자_통합"].astype(str).isin(sel_mgr)]
        if sel_r2:
            temp = temp[temp["리스크등급"].isin(sel_r2)]

        temp = temp.sort_values("접수일시", ascending=False)

        st.write(f"⚠ 활동대상 비매칭 건수: **{len(temp):,} 건**")

        st.dataframe(
            style_risk(temp[display_cols]),
            use_container_width=True,
            height=450,
        )

        st.download_button(
            "📥 현재 필터 기준 비매칭 리스트 다운로드 (CSV)",
            temp.to_csv(index=False).encode("utf-8-sig"),
            file_name="비매칭_활동대상_리스트.csv",
            mime="text/csv",
        )

# ====================================================
# TAB 3 — 지사/담당자 현황 (시각화)
# ====================================================
with tab3:
    st.subheader("📊 지사 / 담당자별 비매칭 리스크 현황")

    if unmatched_global.empty:
        st.info("비매칭(X) 데이터가 없습니다.")
    else:
        col1, col2, col3 = st.columns(3)

        # 지사별 비매칭
        bc = (
            unmatched_global.groupby("관리지사")["계약번호_정제"]
            .nunique()
            .rename("비매칭건수")
        )
        bc = bc[bc.index.isin(BRANCH_ORDER)].reindex(BRANCH_ORDER).dropna()

        with col1:
            st.markdown("#### 🏢 지사별 비매칭 계약 수")
            st.bar_chart(bc)

        # 담당자별 TOP 15
        mc = (
            unmatched_global.groupby("구역담당자_통합")["계약번호_정제"]
            .nunique()
            .rename("비매칭건수")
            .sort_values(ascending=False)
        )
        mc = mc[mc.index.astype(str).str.strip() != ""].head(15)

        with col2:
            st.markdown("#### 👤 담당자별 비매칭 TOP 15")
            st.bar_chart(mc)

        # 리스크 등급 비율
        rc = (
            unmatched_global["리스크등급"]
            .value_counts()
            .reindex(["HIGH", "MEDIUM", "LOW"])
            .fillna(0)
        )

        with col3:
            st.markdown("#### 🔥 리스크 등급 분포 (비매칭)")
            st.bar_chart(rc)

        st.markdown("---")

        # 일별 비매칭 추이
        if "접수일시" in unmatched_global.columns:
            trend = (
                unmatched_global.assign(접수일=unmatched_global["접수일시"].dt.date)
                .groupby("접수일")["계약번호_정제"]
                .nunique()
                .rename("비매칭건수")
                .sort_index()
            )

            st.markdown("#### 📈 일별 비매칭 추이")
            st.line_chart(trend)

# ====================================================
# TAB 4 — 계약번호 / 상호 단위 드릴다운
# ====================================================
with tab4:
    st.subheader("🔍 계약번호 / 상호 기준 통합 드릴다운")

    # 4-1. 글로벌 필터가 적용된 VOC를 기반으로 추가 필터 폼
    base = voc_filtered_global.copy()

    # 지사 목록 / 담당자 목록
    base_branches = sort_branch(base["관리지사"].dropna().unique())
    base_mgrs = (
        base["구역담당자_통합"]
        .dropna()
        .astype(str)
        .unique()
        .tolist()
        if "구역담당자_통합" in base.columns
        else []
    )

    with st.form("drill_form"):
        f1, f2 = st.columns(2)
        sel_br = f1.multiselect(
            "지사 선택",
            options=base_branches,
            default=base_branches,
        )

        # 선택된 지사 기준 담당자 옵션 축소
        tmp_for_mgr = base.copy()
        if sel_br:
            tmp_for_mgr = tmp_for_mgr[tmp_for_mgr["관리지사"].isin(sel_br)]

        mgr_options = (
            tmp_for_mgr["구역담당자_통합"]
            .dropna()
            .astype(str)
            .unique()
            .tolist()
            if "구역담당자_통합" in tmp_for_mgr.columns
            else []
        )
        sel_mgr = f2.multiselect(
            "담당자 선택",
            options=sorted(mgr_options),
            default=sorted(mgr_options),
        )

        g1, g2, g3 = st.columns([1, 1, 0.7])
        q_cn = g1.text_input("계약번호 검색 (부분 입력 가능)")
        q_name = g2.text_input("상호 검색 (부분 입력 가능)")
        submitted = g3.form_submit_button("검색 실행 🔎")

    # 4-2. 폼 기준으로 데이터 필터링
    drill_df = base.copy()
    if sel_br:
        drill_df = drill_df[drill_df["관리지사"].isin(sel_br)]
    if sel_mgr:
        drill_df = drill_df[drill_df["구역담당자_통합"].astype(str).isin(sel_mgr)]

    if q_cn:
        key = q_cn.strip()
        drill_df = drill_df[
            drill_df["계약번호_정제"].astype(str).str.contains(key)
        ]
    if q_name and "상호" in drill_df.columns:
        key = q_name.strip()
        drill_df = drill_df[
            drill_df["상호"].astype(str).str.contains(key)
        ]

    if drill_df.empty:
        st.info("조건에 맞는 계약이 없습니다. 필터를 조정해 보세요.")
    else:
        # 4-3. 계약번호 단위 요약 테이블 (계약번호당 1행)
        latest_voc = (
            drill_df.sort_values("접수일시", ascending=False)
            .groupby("계약번호_정제")
            .first()
            .reset_index()
        )

        summary_cols = [
            c for c in [
                "계약번호_정제",
                "상호",
                "관리지사",
                "구역담당자_통합",
                "리스크등급",
                "경과일수",
                "매칭여부",
            ]
            if c in latest_voc.columns
        ]

        st.markdown("#### 📋 필터링된 계약 요약 (계약번호당 1행, 최신 VOC 기준)")
        st.dataframe(
            style_risk(latest_voc[summary_cols]),
            use_container_width=True,
            height=260,
        )

        # 4-4. 요약 테이블에서 상세 볼 계약번호 선택
        cn_options = latest_voc["계약번호_정제"].astype(str).tolist()

        def _format_cn(cn_value: str) -> str:
            try:
                row = latest_voc[latest_voc["계약번호_정제"].astype(str) == str(cn_value)].iloc[0]
                name = row.get("상호", "")
                branch = row.get("관리지사", "")
                return f"{cn_value} | {name} | {branch}"
            except Exception:
                return str(cn_value)

        sel_cn = st.selectbox(
            "상세를 볼 계약 선택 (계약번호 | 상호 | 지사)",
            options=cn_options,
            format_func=_format_cn,
        )

        # 4-5. 선택된 계약번호에 대한 VOC/기타 출처 이력
        if sel_cn:
            voc_detail = df_voc[df_voc["계약번호_정제"].astype(str) == str(sel_cn)].copy()
            voc_detail = voc_detail.sort_values("접수일시", ascending=False)

            others_detail = df_other[df_other["계약번호_정제"].astype(str) == str(sel_cn)].copy()

            st.markdown(f"### 🔎 선택한 계약번호: `{sel_cn}`")

            # 최신 VOC 1건 요약
            if voc_detail.empty:
                st.info("선택된 계약번호의 VOC 이력이 없습니다.")
            else:
                st.markdown("#### ✅ 최신 VOC 요약 (1건)")
                st.dataframe(
                    style_risk(voc_detail[display_cols].head(1)),
                    use_container_width=True,
                    height=130,
                )

            col_a, col_b = st.columns(2)

            with col_a:
                st.markdown("##### 📘 VOC 전체 이력 (시간순)")
                if voc_detail.empty:
                    st.info("VOC 이력이 없습니다.")
                else:
                    st.dataframe(
                        style_risk(voc_detail[display_cols]),
                        use_container_width=True,
                        height=320,
                    )

            with col_b:
                st.markdown("##### 📂 기타 출처 이력 (해지시설/요청/설변/정지/파이프라인)")
                if others_detail.empty:
                    st.info("기타 출처 데이터가 없습니다.")
                else:
                    st.dataframe(
                        others_detail,
                        use_container_width=True,
                        height=320,
                    )
