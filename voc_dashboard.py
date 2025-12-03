import os
from datetime import datetime, date

import numpy as np
import pandas as pd
import streamlit as st

# ----------------------------------------------------
# 0. 기본 설정 + 라이트 스타일 강제
# ----------------------------------------------------
st.set_page_config(page_title="해지 VOC 종합 대시보드", layout="wide")

# 다크 모드여도 항상 흰 배경 + 검정 글씨로 보이도록 CSS 적용
st.markdown(
    """
    <style>
    .stApp {
        background-color: #ffffff !important;
        color: #000000 !important;
    }
    .stDataFrame th, .stDataFrame td {
        color: #000000 !important;
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

# 주소 컬럼 자동 탐색 (검색에만 사용)
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

# VOC ∧ 기타 출처 있음 → 매칭(O), 없으면 비매칭(X)
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
# 6. 피드백(고객대응/현장처리) in-memory 저장 구조
# ----------------------------------------------------
FEEDBACK_COLS = ["계약번호_정제", "고객대응사항", "등록자", "등록일자", "비고"]

if "feedback_df" not in st.session_state:
    st.session_state["feedback_df"] = pd.DataFrame(columns=FEEDBACK_COLS)


def attach_feedback(df_in: pd.DataFrame) -> pd.DataFrame:
    """계약번호_정제 기준으로 최신 피드백 1건을 붙여준다."""
    fb = st.session_state.get("feedback_df", pd.DataFrame(columns=FEEDBACK_COLS))
    if fb.empty or "계약번호_정제" not in df_in.columns:
        return df_in

    fb_sorted = fb.sort_values("등록일자")
    fb_last = fb_sorted.groupby("계약번호_정제").last().reset_index()

    merged = df_in.merge(fb_last, on="계약번호_정제", how="left")
    return merged


# ----------------------------------------------------
# 7. 공통 표시 컬럼 정의
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

# 피드백 컬럼도 표시 순서에 추가
for c in ["고객대응사항", "등록자", "등록일자", "비고"]:
    if c not in display_cols:
        display_cols.append(c)

# ----------------------------------------------------
# 8. 스타일링 (리스크 등급 색상 강조)
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
# 9. 사이드바 - 글로벌 필터
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

# 매칭 여부
if sel_match:
    voc_filtered_global = voc_filtered_global[
        voc_filtered_global["매칭여부"].isin(sel_match)
    ]

unmatched_global = voc_filtered_global[
    voc_filtered_global["매칭여부"] == "비매칭(X)"
].copy()

# ----------------------------------------------------
# 11. 헤더 & 상단 KPI
# ----------------------------------------------------
st.markdown("## 📊 해지 VOC 종합 대시보드 (엔터프라이즈 버전)")

k1, k2, k3, k4 = st.columns(4)

total_voc = len(voc_filtered_global)
unique_cn = voc_filtered_global["계약번호_정제"].nunique()
unmatched_cnt = (voc_filtered_global["매칭여부"] == "비매칭(X)").sum()
matched_cnt = (voc_filtered_global["매칭여부"] == "매칭(O)").sum()

k1.metric("전체 VOC 건수", f"{total_voc:,}")
k2.metric("유니크 계약 수", f"{unique_cn:,}")
k3.metric("비매칭(X) 활동대상", f"{unmatched_cnt:,}")
k4.metric("매칭(O) 건수", f"{matched_cnt:,}")

st.markdown("---")

# ----------------------------------------------------
# 12. 탭 구성
# ----------------------------------------------------
tab1, tab2, tab3, tab4 = st.tabs(
    ["📘 VOC 전체", "🚨 비매칭(활동대상)", "📊 지사/담당자 시각화", "🔍 계약번호 통합 드릴다운"]
)

# ====================================================
# TAB 1 — VOC 전체 조회
# ====================================================
with tab1:
    st.subheader("📘 VOC 전체 조회 (글로벌 필터 적용 상태)")

    c1, c2, c3 = st.columns(3)
    key_cn = c1.text_input("계약번호 검색", key="tab1_cn")
    key_name = c2.text_input("상호 검색", key="tab1_name")
    key_addr = c3.text_input("주소 검색", key="tab1_addr")

    temp = voc_filtered_global.copy()

    if key_cn:
        key = key_cn.strip()
        temp = temp[temp["계약번호_정제"].astype(str).str.contains(key, na=False)]

    if key_name and "상호" in temp.columns:
        key = key_name.strip()
        temp = temp[temp["상호"].astype(str).str.contains(key, na=False)]

    if key_addr and address_cols:
        key = key_addr.strip()
        cond = False
        for col in address_cols:
            cond |= temp[col].astype(str).str.contains(key, na=False)
        temp = temp[cond]

    temp = temp.sort_values("접수일시", ascending=False)

    temp_view = attach_feedback(temp)
    cols_to_show = [c for c in display_cols if c in temp_view.columns]

    st.write(f"표시 건수: **{len(temp_view):,} 건**")

    st.dataframe(
        style_risk(temp_view[cols_to_show]),
        use_container_width=True,
        height=520,
    )

# ====================================================
# TAB 2 — 비매칭(X) = 활동대상 (피드백 입력 포함)
# ====================================================
with tab2:
    st.subheader("🚨 비매칭(X) 활동대상 — 지사 / 담당자 기준 리스트")

    if unmatched_global.empty:
        st.info("현재 글로벌 필터 조건에서 비매칭(X) 데이터가 없습니다.")
    else:
        # 1) 지사 선택
        branches_raw = sort_branch(unmatched_global["관리지사"].dropna().unique())
        branch_buttons = ["전체"] + branches_raw

        selected_branch = st.radio(
            "지사 선택",
            options=branch_buttons,
            horizontal=True,
            key="tab2_branch",
        )

        temp2 = unmatched_global.copy()
        if selected_branch != "전체":
            temp2 = temp2[temp2["관리지사"] == selected_branch]

        # 2) 담당자 선택
        mgr_list = (
            temp2["구역담당자_통합"]
            .dropna()
            .astype(str)
            .unique()
            .tolist()
        )
        mgr_buttons = ["전체"] + sorted(mgr_list)

        selected_mgr = st.radio(
            "담당자 선택",
            options=mgr_buttons,
            horizontal=True,
            key="tab2_mgr",
        )

        if selected_mgr != "전체":
            temp2 = temp2[temp2["구역담당자_통합"] == selected_mgr]

        temp2 = temp2.sort_values("접수일시", ascending=False)

        # 3) 피드백 붙인 테이블
        temp2_view = attach_feedback(temp2)
        cols_to_show = [c for c in display_cols if c in temp2_view.columns]

        st.write(f"총 활동대상: **{len(temp2_view):,} 건**")

        st.dataframe(
            style_risk(temp2_view[cols_to_show]),
            use_container_width=True,
            height=360,
        )

        st.download_button(
            "📥 비매칭 활동대상 다운로드 (CSV)",
            temp2_view.to_csv(index=False).encode("utf-8-sig"),
            file_name="비매칭_활동대상.csv",
            mime="text/csv",
        )

        st.markdown("---")

        # 4) 피드백 입력 폼
        st.markdown("### ✏️ 고객 대응 / 현장 처리 내용 등록")

        cn_options = (
            temp2["계약번호_정제"]
            .dropna()
            .astype(str)
            .unique()
            .tolist()
        )

        if not cn_options:
            st.info("선택된 지사/담당자에 해당하는 계약이 없습니다.")
        else:
            f1, f2 = st.columns([1, 2])

            sel_cn_fb = f1.selectbox(
                "계약번호 선택",
                options=cn_options,
                key="tab2_feedback_cn",
            )

            # 선택된 계약 요약 정보
            info_row = temp2[temp2["계약번호_정제"] == sel_cn_fb].iloc[0]
            f1.write(f"**상호:** {info_row.get('상호', '')}")
            f1.write(f"**지사:** {info_row.get('관리지사', '')}")
            f1.write(f"**담당자:** {info_row.get('구역담당자_통합', '')}")

            with f2:
                txt_feedback = st.text_area(
                    "고객대응 / 현장 처리내용",
                    key="tab2_feedback_text",
                    placeholder="고객 통화내용, 현장 방문 처리내용, 후속 계획 등을 입력하세요.",
                )
                col_f1, col_f2 = st.columns(2)
                reg_user = col_f1.text_input("등록자", key="tab2_feedback_user")
                remark = col_f2.text_input("비고", key="tab2_feedback_remark")

                if st.button("💾 피드백 저장", key="tab2_feedback_save"):
                    if not txt_feedback.strip():
                        st.warning("고객대응/처리내용을 입력해주세요.")
                    elif not reg_user.strip():
                        st.warning("등록자를 입력해주세요.")
                    else:
                        new_row = pd.DataFrame(
                            [
                                {
                                    "계약번호_정제": sel_cn_fb,
                                    "고객대응사항": txt_feedback.strip(),
                                    "등록자": reg_user.strip(),
                                    "등록일자": datetime.now().strftime(
                                        "%Y-%m-%d %H:%M:%S"
                                    ),
                                    "비고": remark.strip(),
                                }
                            ]
                        )
                        st.session_state["feedback_df"] = pd.concat(
                            [st.session_state["feedback_df"], new_row],
                            ignore_index=True,
                        )
                        st.success("피드백이 저장되었습니다. (현재 세션 기준으로 즉시 반영됩니다.)")

# ====================================================
# TAB 3 — 지사/담당자 시각화 (5개 시각화)
# ====================================================
with tab3:
    st.subheader("📊 지사 / 담당자 / 리스크 시각화")

    if unmatched_global.empty:
        st.info("비매칭(X) 데이터가 없습니다.")
    else:
        col1, col2 = st.columns(2)

        # 1) 지사별 비매칭 건수
        branch_counts = (
            unmatched_global.groupby("관리지사")["계약번호_정제"]
            .nunique()
            .rename("비매칭건수")
        )
        branch_counts = branch_counts[
            branch_counts.index.isin(BRANCH_ORDER)
        ].reindex(BRANCH_ORDER).fillna(0)

        with col1:
            st.markdown("#### 1️⃣ 🏢 지사별 비매칭 계약 수")
            st.bar_chart(branch_counts)

        # 2) 담당자별 비매칭 TOP 15
        mgr_counts = (
            unmatched_global.groupby("구역담당자_통합")["계약번호_정제"]
            .nunique()
            .rename("비매칭건수")
            .sort_values(ascending=False)
        )
        mgr_counts = mgr_counts[
            mgr_counts.index.astype(str).str.strip() != ""
        ].head(15)

        with col2:
            st.markdown("#### 2️⃣ 👤 담당자별 비매칭 TOP 15")
            st.bar_chart(mgr_counts)

        st.markdown("---")

        col3, col4 = st.columns(2)

        # 3) 리스크 등급 분포
        risk_dist = (
            unmatched_global["리스크등급"]
            .value_counts()
            .reindex(["HIGH", "MEDIUM", "LOW"])
            .fillna(0)
        )

        with col3:
            st.markdown("#### 3️⃣ 🔥 리스크 등급 분포")
            st.bar_chart(risk_dist)

        # 4) 일별 비매칭 추이
        trend = None
        if "접수일시" in unmatched_global.columns:
            trend = (
                unmatched_global.assign(접수일=unmatched_global["접수일시"].dt.date)
                .groupby("접수일")["계약번호_정제"]
                .nunique()
                .rename("비매칭건수")
                .sort_index()
            )

            with col4:
                st.markdown("#### 4️⃣ 📈 일별 비매칭 추이")
                st.line_chart(trend)

        # 5) 누적 비매칭 추세
        if trend is not None:
            cum = trend.cumsum()
            st.markdown("#### 5️⃣ 📊 누적 비매칭 계약 수 추세")
            st.area_chart(cum)

# ====================================================
# TAB 4 — 계약번호 단위 통합 드릴다운
# ====================================================
with tab4:
    st.subheader("🔍 계약번호 기준 통합 VOC 이력 조회")

    # 1) 검색 조건
    colA, colB = st.columns(2)
    search_cn = colA.text_input("계약번호(일부 가능)", key="tab4_search_cn")
    search_name = colB.text_input("상호(일부 가능)", key="tab4_search_name")

    search_df = df.copy()

    if search_cn.strip():
        search_df = search_df[
            search_df["계약번호_정제"].astype(str).str.contains(search_cn.strip(), na=False)
        ]
    if search_name.strip():
        search_df = search_df[
            search_df["상호"].astype(str).str.contains(search_name.strip(), na=False)
        ]

    cn_candidates = (
        search_df["계약번호_정제"]
        .dropna()
        .astype(str)
        .unique()
        .tolist()
    )

    if not cn_candidates:
        st.info("검색 조건에 해당하는 계약번호가 없습니다.")
    else:
        sel_cn = st.selectbox(
            "조회할 계약번호 선택",
            options=cn_candidates,
            key="tab4_cn_select",
        )

        # ---------------------------------------------------------
        # 2) VOC 상세 (최신 접수일시 우선)
        # ---------------------------------------------------------
        voc_detail = df_voc[df_voc["계약번호_정제"] == sel_cn].copy()
        voc_detail = voc_detail.sort_values("접수일시", ascending=False)

        # ---------------------------------------------------------
        # 3) 기타 출처 상세
        # ---------------------------------------------------------
        others_detail = df_other[df_other["계약번호_정제"] == sel_cn].copy()

        # ---------------------------------------------------------
        # 4) 피드백 / 현장 처리 내역 (계약번호 단위 전체)
        # ---------------------------------------------------------
        fb_all = st.session_state.get("feedback_df", pd.DataFrame(columns=FEEDBACK_COLS))
        fb_rows = fb_all[fb_all["계약번호_정제"] == sel_cn].copy()
        fb_rows = fb_rows.sort_values("등록일자", ascending=False)

        if not fb_rows.empty:
            feedback_text = "\n".join(
                [
                    f"[{row['등록일자']}] ({row['등록자']})\n"
                    f"고객대응: {row['고객대응사항']}\n"
                    f"비고: {row['비고']}\n"
                    for _, row in fb_rows.iterrows()
                ]
            )
        else:
            feedback_text = "등록된 고객 대응/현장 처리 이력이 없습니다."

        # ---------------------------------------------------------
        # 5) 화면 표시
        # ---------------------------------------------------------
        st.markdown(f"### 📄 계약번호 `{sel_cn}` 통합 이력")

        c1, c2 = st.columns(2)

        # VOC 이력
        with c1:
            st.markdown("#### 📘 해지VOC 접수 이력 (최신순)")
            if voc_detail.empty:
                st.info("해지VOC 이력 없음")
            else:
                voc_view = attach_feedback(voc_detail)
                cols_v = [c for c in display_cols if c in voc_view.columns]
                st.dataframe(
                    style_risk(voc_view[cols_v]),
                    use_container_width=True,
                    height=300,
                )

        # 기타 출처 이력
        with c2:
            st.markdown("#### 📂 기타 출처 이력 (해지시설/요청/설변/정지/파이프라인)")
            if others_detail.empty:
                st.info("기타 출처 데이터 없음")
            else:
                st.dataframe(
                    others_detail,
                    use_container_width=True,
                    height=300,
                )

        st.markdown("---")

        # 통합 고객대응 / 현장 처리 이력
        st.markdown("#### 📝 고객대응 / 현장 처리 통합 이력")
        st.text_area(
            "통합 고객대응·현장 처리 이력",
            value=feedback_text,
            height=230,
            disabled=True,
        )

        st.markdown("### ✏️ 새로운 고객대응 / 현장 처리 내용 등록")

        fb1, fb2 = st.columns([2, 1])

        new_feedback = fb1.text_area(
            "고객대응 / 현장 처리내역 입력",
            placeholder="고객과의 통화, 방문 처리내용, 향후 조치 계획 등을 입력하세요.",
            key="tab4_new_feedback",
        )
        new_user = fb2.text_input("등록자", key="tab4_new_user")
        new_remark = fb2.text_input("비고", key="tab4_new_remark")

        if st.button("💾 처리내역 저장", key="tab4_save_feedback"):
            if not new_feedback.strip():
                st.warning("내용을 입력하세요.")
            elif not new_user.strip():
                st.warning("등록자를 입력하세요.")
            else:
                new_row = pd.DataFrame(
                    [
                        {
                            "계약번호_정제": sel_cn,
                            "고객대응사항": new_feedback.strip(),
                            "등록자": new_user.strip(),
                            "등록일자": datetime.now().strftime(
                                "%Y-%m-%d %H:%M:%S"
                            ),
                            "비고": new_remark.strip(),
                        }
                    ]
                )
                st.session_state["feedback_df"] = pd.concat(
                    [st.session_state["feedback_df"], new_row],
                    ignore_index=True,
                )
                st.success("처리내역이 저장되었습니다. 위 통합 이력에 즉시 반영됩니다.")
