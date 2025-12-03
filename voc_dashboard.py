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
# TAB 2 — 비매칭(X) = 활동대상  (버튼 방식 / 리스크필터 제거)
# ====================================================
with tab2:
    st.subheader("🚨 비매칭(X) 활동대상 — 지사 / 담당자 선택 방식")

    if unmatched_global.empty:
        st.info("현재 글로벌 필터 조건에서 비매칭(X) 데이터가 없습니다.")
    else:

        # -------------------------------
        # 지사 버튼 생성
        # -------------------------------
        branches_raw = sort_branch(
            unmatched_global["관리지사"].dropna().unique()
        )

        branch_buttons = ["전체"] + branches_raw

        selected_branch = st.radio(
            "지사 선택",
            options=branch_buttons,
            horizontal=True
        )

        # 지사 필터 적용
        temp = unmatched_global.copy()
        if selected_branch != "전체":
            temp = temp[temp["관리지사"] == selected_branch]

        # -------------------------------
        # 담당자 버튼 (지사 선택 시 동적 생성)
        # -------------------------------
        mgr_list = (
            temp["구역담당자_통합"]
            .dropna()
            .astype(str)
            .unique()
            .tolist()
        )

        mgr_buttons = ["전체"] + sorted(mgr_list)

        selected_mgr = st.radio(
            "담당자 선택",
            options=mgr_buttons,
            horizontal=True
        )

        # 담당자 필터 적용
        if selected_mgr != "전체":
            temp = temp[temp["구역담당자_통합"] == selected_mgr]

        # -------------------------------
        # 결과 정렬 + 표시
        # -------------------------------
        temp = temp.sort_values("접수일시", ascending=False)

        st.write(f"총 활동대상 : **{len(temp):,} 건**")

        st.dataframe(
            style_risk(temp[display_cols]),
            use_container_width=True,
            height=450,
        )

        # 다운로드 기능
        st.download_button(
            "📥 현재 조건 비매칭 리스트 다운로드 (CSV)",
            temp.to_csv(index=False).encode("utf-8-sig"),
            file_name="비매칭_활동대상.csv",
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
# TAB 4 — 계약번호 드릴다운 (지사/담당자/검색 개선)
# ====================================================
with tab4:
    st.subheader("🔍 계약번호 기반 VOC + 기타 출처 통합 조회 (전문가 버전)")

    # ----------------------------------------------------
    # 1) 지사 선택 (버튼)
    # ----------------------------------------------------
    branches_raw = sort_branch(df_voc["관리지사"].dropna().unique())
    branch_buttons = ["전체"] + branches_raw

    selected_branch = st.radio(
        "지사 선택",
        options=branch_buttons,
        horizontal=True
    )

    temp = df_voc.copy()
    if selected_branch != "전체":
        temp = temp[temp["관리지사"] == selected_branch]

    # ----------------------------------------------------
    # 2) 담당자 선택 (동적 생성)
    # ----------------------------------------------------
    mgr_list = (
        temp["구역담당자_통합"]
        .dropna()
        .astype(str)
        .unique()
        .tolist()
    )

    mgr_buttons = ["전체"] + sorted(mgr_list)

    selected_mgr = st.radio(
        "담당자 선택",
        options=mgr_buttons,
        horizontal=True
    )

    temp2 = temp.copy()
    if selected_mgr != "전체":
        temp2 = temp2[temp2["구역담당자_통합"] == selected_mgr]

    # ----------------------------------------------------
    # 3) 계약번호 / 상호 검색 입력
    # ----------------------------------------------------
    c1, c2, c3 = st.columns([1.2, 1.2, 0.7])

    input_cn = c1.text_input("계약번호 (일부 입력 가능)")
    input_name = c2.text_input("상호 (일부 입력 가능)")

    search_clicked = c3.button("🔍 검색")

    # ----------------------------------------------------
    # 4) 검색 실행
    # ----------------------------------------------------
    result_df = temp2.copy()

    if search_clicked:
        # 계약번호 검색
        if input_cn.strip():
            key = input_cn.strip()
            result_df = result_df[
                result_df["계약번호_정제"].astype(str).str.contains(key, na=False)
            ]

        # 상호 검색
        if input_name.strip() and "상호" in result_df.columns:
            key = input_name.strip()
            result_df = result_df[
                result_df["상호"].astype(str).str.contains(key, na=False)
            ]

        # 검색결과가 1개 이상이면 계약번호 목록 표시
        found_cn_list = (
            result_df["계약번호_정제"].dropna().astype(str).unique().tolist()
        )

        if len(found_cn_list) == 0:
            st.warning("검색 조건과 일치하는 계약번호가 없습니다.")
            st.stop()

        # 자동으로 하나만 남으면 바로 조회
        if len(found_cn_list) == 1:
            sel_cn = found_cn_list[0]
        else:
            sel_cn = st.selectbox("계약번호 선택", found_cn_list)

    else:
        sel_cn = None

    # ----------------------------------------------------
    # 5) 최종 조회 및 VOC / 기타출처 결과 표시
    # ----------------------------------------------------
    if sel_cn:
        st.markdown(f"### 📌 조회된 계약번호: `{sel_cn}`")

        # VOC 상세 (글로벌 필터는 무시하고 temp2 기준)
        voc_detail = df_voc[df_voc["계약번호_정제"] == sel_cn].copy()
        voc_detail = voc_detail.sort_values("접수일시", ascending=False)

        # 기타 출처 (전체 df 기준 조회)
        others_detail = df_other[df_other["계약번호_정제"] == sel_cn].copy()

        c1, c2 = st.columns(2)

        # VOC
        with c1:
            st.markdown("#### 📘 해지 VOC 이력")
            if voc_detail.empty:
                st.info("VOC 이력 없음")
            else:
                st.dataframe(
                    style_risk(voc_detail[display_cols]),
                    use_container_width=True,
                    height=350,
                )

        # 기타 출처
        with c2:
            st.markdown("#### 📂 기타 출처 이력 (해지시설/요청/설변/정지/파이프라인)")
            if others_detail.empty:
                st.info("기타 출처 데이터 없음")
            else:
                st.dataframe(
                    others_detail,
                    use_container_width=True,
                    height=350,
                )
