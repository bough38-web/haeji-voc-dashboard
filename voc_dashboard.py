# ------------------------------------------------
# 🔹 적층 세로 막대그래프 + 애니메이션 지원
# ------------------------------------------------
def force_stacked_bar_animated(
    df: pd.DataFrame,
    x: str,
    y_cols: list[str],
    anim_col: str,
    height: int = 280
):
    """
    Plotly 적층 세로 막대그래프 (애니메이션 적용)
    df: 데이터프레임
    x: x축 컬럼명
    y_cols: 적층될 수치 컬럼 리스트 ["HIGH","MEDIUM","LOW"]
    anim_col: 애니메이션 기준 컬럼명 (예: '접수일', '관리지사', '구역담당자_통합', ...)
    """

    if df.empty or not y_cols or anim_col not in df.columns:
        st.info("애니메이션을 표시할 데이터가 부족합니다.")
        return

    # Plotly 애니메이션 bar chart
    if HAS_PLOTLY:
        fig = px.bar(
            df,
            x=x,
            y=y_cols,
            color=None,
            animation_frame=anim_col,
            barmode="stack",
            text_auto=True,
            height=height,
        )

        fig.update_layout(
            margin=dict(l=40, r=20, t=40, b=40),
            transition={"duration": 500},
        )

        st.plotly_chart(fig, use_container_width=True)
    else:
        st.warning("Plotly가 설치되어야 애니메이션 그래프를 표시할 수 있습니다.")
