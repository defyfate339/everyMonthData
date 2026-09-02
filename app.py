# -*- coding: utf-8 -*-
"""
粤徽交付中心八月数据看板 — Streamlit 可视化应用
数据源: 本机 D:/YueHuiProject/粤徽交付中心八月数据看板.xlsx（不存在时回退到仓库内同名文件，兼容云端部署）

运行: streamlit run app.py
部署: 推送到 GitHub 后在 Streamlit Community Cloud 一键部署
"""

import io
import os

import pandas as pd
import plotly.express as px
import streamlit as st

# ---------------------------------------------------------------------------
# 常量与配置
# ---------------------------------------------------------------------------
LOCAL_PATH = r"D:/YueHuiProject/粤徽交付中心八月数据看板.xlsx"
DATA_FILE = "粤徽交付中心八月数据看板.xlsx"  # 仓库内数据文件（云端部署用）
METRICS = ["新增微信", "预约", "到场", "合格", "在职"]
METRIC_COLORS = {"新增微信": "#4A90D9", "预约": "#9B59B6", "到场": "#F39C12",
                 "合格": "#27AE60", "在职": "#E74C3C"}


def resolve_data_path():
    """优先本机路径，其次仓库内数据文件（云端可用）"""
    if os.path.exists(LOCAL_PATH):
        return LOCAL_PATH
    repo_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), DATA_FILE)
    if os.path.exists(repo_path):
        return repo_path
    st.error(f"未找到数据文件（{LOCAL_PATH} 或仓库内 {DATA_FILE}）")
    st.stop()


DATA_PATH = resolve_data_path()

st.set_page_config(page_title="粤徽交付中心 · 八月数据看板",
                   page_icon="📊", layout="wide")

# 深色主题微调
st.markdown("""
<style>
.block-container {padding-top: 1.2rem;}
.metric-card {
    background: linear-gradient(135deg, #1A1F29 0%, #232A38 100%);
    border: 1px solid #2E3A4D; border-radius: 12px;
    padding: 16px 20px; margin: 4px 0;
}
.metric-card .m-label {font-size: 13px; color: #8FA3BF;}
.metric-card .m-value {font-size: 30px; font-weight: 700; color: #E8EDF5;}
.metric-card .m-sub {font-size: 12px; color: #6B7C93;}
</style>
""", unsafe_allow_html=True)


# ---------------------------------------------------------------------------
# 数据加载与解析（兼容 openpyxl 无法解析的样式，使用 calamine 引擎）
# ---------------------------------------------------------------------------
@st.cache_data(show_spinner="正在解析 Excel 数据...")
def load_and_parse():
    df = pd.read_excel(DATA_PATH, sheet_name=0, engine="calamine", header=None)

    # 1) 日期列定位: 第6行(row5)为日期、且第7行(row6)对应为指标名（每组5列起始）
    #    —— 注意：row5 还有注释日期"（截至）2026-08-24"(col5)，
    #       它后面不是指标名（是"数据汇总"），必须排除；真正的 8/1 从 N 列开始
    METRIC_SET = set(METRICS)
    date_cols, dates = [], []
    for c in range(df.shape[1]):
        v = df.iloc[5, c]
        if isinstance(v, (pd.Timestamp,)) or hasattr(v, "strftime"):
            if str(df.iloc[6, c]).strip() in METRIC_SET:  # 二级表头必须是指标名
                date_cols.append(c)
                dates.append(pd.Timestamp(v))
    if not date_cols:
        st.error("未能在第 6 行找到日期标签，请确认文件结构（8月1日~8月31日 每组5列）")
        st.stop()

    # 2) 员工行定位: 姓名(col3)、岗位(col4) 均非空，且组别(col1) 非"合计"
    #    （合计行 col3 为人数数字，如"12"，需排除）
    def _is_empty(v):
        return v is None or (isinstance(v, float) and pd.isna(v)) or str(v).strip() == ""

    emp_rows = [r for r in range(7, df.shape[0])
                if not _is_empty(df.iloc[r, 3]) and str(df.iloc[r, 3]).strip() != "姓名"
                and not _is_empty(df.iloc[r, 4])
                and str(df.iloc[r, 1]).strip() != "合计"]
    if not emp_rows:
        st.error("未找到员工明细行")
        st.stop()

    # 3) 主信息表（整月汇总 + 基础信息）
    main_records = []
    for r in emp_rows:
        main_records.append({
            "所属部门": df.iloc[r, 0],
            "组别": df.iloc[r, 1],
            "管理": df.iloc[r, 2],
            "姓名": str(df.iloc[r, 3]).strip(),
            "岗位": df.iloc[r, 4],
            "BOSS账号数量": _num(df.iloc[r, 5]),
            **{m: _num(df.iloc[r, 6 + i]) for i, m in enumerate(METRICS)},
            "目标在职": _num(df.iloc[r, 11]),
            "达成率": df.iloc[r, 12] if pd.notna(df.iloc[r, 12]) else None,
        })
    main_df = pd.DataFrame(main_records)

    # 4) 每日明细长表
    long_records = []
    for r in emp_rows:
        name = str(df.iloc[r, 3]).strip()
        for k, c0 in enumerate(date_cols):
            d = dates[k]
            for j, m in enumerate(METRICS):
                v = df.iloc[r, c0 + j]
                long_records.append({
                    "日期": d, "姓名": name,
                    "组别": df.iloc[r, 1], "管理": df.iloc[r, 2], "岗位": df.iloc[r, 4],
                    "指标": m, "数值": _num(v),
                })
    detail_df = pd.DataFrame(long_records)

    # 宽表: 行=日期, 列=指标, 值按人数合计（用于每日团队趋势）
    pivot_daily = detail_df.groupby(["日期", "指标"], as_index=False)["数值"].sum()
    return main_df, detail_df, pivot_daily, dates


def _num(v):
    """安全转数值，NaN/None -> 0"""
    try:
        if v is None or (isinstance(v, float) and pd.isna(v)):
            return 0.0
        return float(v)
    except (TypeError, ValueError):
        return 0.0


# ---------------------------------------------------------------------------
# 图表辅助
# ---------------------------------------------------------------------------
def trend_fig(df, y_col, title, color=None):
    """折线图（默认按姓名分色）"""
    fig = px.line(df, x="日期", y=y_col, color=color if color else "姓名",
                  markers=True, title=title,
                  )
    fig.update_layout(template="plotly_dark", height=420,
                      legend=dict(orientation="h", y=-0.25),
                      margin=dict(l=10, r=10, t=50, b=10))
    fig.update_xaxes(dtick="D1", tickformat="%m-%d")
    return fig


def bar_fig(df, x, y, color=None, title=""):
    fig = px.bar(df, x=x, y=y, color=color if color else None,
                 text=y, title=title, barmode="group",
                 )
    fig.update_traces(textposition="outside", texttemplate="%{text:.0f}")
    fig.update_layout(template="plotly_dark", height=420,
                      margin=dict(l=10, r=10, t=50, b=10))
    return fig


def heatmap_fig(df, metric):
    """人 × 日期 热力图（横轴每天显示日期刻度）"""
    pivot = df[df["指标"] == metric].pivot_table(
        index="姓名", columns="日期", values="数值", aggfunc="sum")
    pivot = pivot.sort_index(ascending=False)
    # 列顺序按日期升序，确保 px.imshow 的 X 轴数值轴按时间排列
    pivot = pivot[sorted(pivot.columns)]
    fig = px.imshow(pivot, text_auto=".0f", aspect="auto",
                    color_continuous_scale="Blues",
                    labels=dict(x="日期", y="姓名", color=metric))
    # px.imshow 的 X 轴是日期/数值轴，tickvals 必须传 Timestamp，不能用整数索引
    cols = list(pivot.columns)
    tick_labels = [pd.Timestamp(c).strftime("%m-%d") for c in cols]
    fig.update_xaxes(type="date", tickmode="array",
                     tickvals=cols, ticktext=tick_labels, tickangle=-45)
    fig.update_layout(template="plotly_dark", height=520,
                      title=f"{metric} · 人 × 日 热力图",
                      margin=dict(l=10, r=10, t=50, b=10))
    return fig


# ---------------------------------------------------------------------------
# 主流程
# ---------------------------------------------------------------------------
def 主分支():
    st.title("📊 粤徽交付中心 · 八月数据看板")

    # 侧边栏
    with st.sidebar:
        st.header("⚙️ 筛选条件")
        main_df, detail_df, pivot_daily, dates = load_and_parse()

        date_min, date_max = min(dates).date(), max(dates).date()
        d_start, d_end = st.date_input("日期范围",
                                       value=(date_min, date_max),
                                       min_value=date_min, max_value=date_max)
        if isinstance(d_start, tuple):
            d_start, d_end = d_start
        if d_start is None or d_end is None:
            st.warning("请选择完整的日期范围")
            st.stop()

        groups = st.multiselect("组别", main_df["组别"].dropna().unique().tolist(),
                                default=main_df["组别"].dropna().unique().tolist())
        persons = st.multiselect("人员", main_df["姓名"].tolist(),
                                 default=main_df["姓名"].tolist())
        metrics = st.multiselect("指标", METRICS, default=METRICS)
        if not groups or not persons or not metrics:
            st.warning("请至少选择一个组别 / 人员 / 指标")
            st.stop()
        st.caption(f"数据覆盖 {date_min} ~ {date_max}，共 {len(main_df)} 人")

    # 过滤
    sel = detail_df[(detail_df["日期"].dt.date >= d_start) &
                    (detail_df["日期"].dt.date <= d_end) &
                    (detail_df["组别"].isin(groups)) &
                    (detail_df["姓名"].isin(persons)) &
                    (detail_df["指标"].isin(metrics))]

    m_sel = main_df[main_df["姓名"].isin(persons) & main_df["组别"].isin(groups)]

    if sel.empty:
        st.info("所选范围内没有数据")
        st.stop()

    # ---- Tab 1: 总览 ----
    tab1, tab2, tab3, tab4 = st.tabs(["📊 总览", "👥 每人明细", "📈 趋势分析", "📋 数据表"])

    with tab1:
        # KPI 卡片
        cols = st.columns(len(metrics))
        for col, m in zip(cols, metrics):
            total = int(sel[sel["指标"] == m]["数值"].sum())
            avg = total / max(len(sel[sel["指标"] == m]["日期"].unique()), 1)
            with col:
                st.markdown(
                    f'<div class="metric-card"><div class="m-label">{m}</div>'
                    f'<div class="m-value">{total:,}</div>'
                    f'<div class="m-sub">日均 {avg:,.1f}</div></div>',
                    unsafe_allow_html=True)
        st.write("")
        st.caption(f"统计区间：{d_start} ~ {d_end} ｜ 人数：{len(m_sel)} ｜ 组别：{'、'.join(groups)}")

        c1, c2 = st.columns(2)
        with c1:
            # 组别 × 指标 合计条形图
            g = sel.groupby(["组别", "指标"], as_index=False)["数值"].sum()
            st.plotly_chart(bar_fig(g, "组别", "数值", color="指标",
                                    title="按组别汇总（区间合计）"),
                            width="stretch")
        with c2:
            st.markdown("**每人汇总（区间合计 · 分指标）**")
            # 每人合计：每个指标一张独立柱状图
            p = sel.groupby(["姓名", "指标"], as_index=False)["数值"].sum()
            sub_cols = st.columns(2)
            for i, m in enumerate(metrics):
                pm = p[p["指标"] == m].sort_values("数值", ascending=False)
                fig = px.bar(pm, x="姓名", y="数值", text="数值", title=m,
                             color_discrete_sequence=[METRIC_COLORS.get(m, "#4A90D9")])
                fig.update_traces(textposition="outside", texttemplate="%{text:.0f}")
                fig.update_layout(template="plotly_dark", height=300,
                                  margin=dict(l=10, r=10, t=50, b=10),
                                  xaxis_tickangle=-45, showlegend=False)
                with sub_cols[i % 2]:
                    st.plotly_chart(fig, width="stretch")

    # ---- Tab 2: 每人明细 ----
    with tab2:
        # 区间汇总表（含整月信息）
        agg = sel.groupby(["姓名", "组别", "岗位", "管理"], as_index=False).apply(
            lambda x: pd.Series({m: int(x[x["指标"] == m]["数值"].sum()) for m in metrics}),
            include_groups=False).reset_index()
        m_merge = m_sel[["姓名", "BOSS账号数量", "目标在职", "达成率"]].drop_duplicates("姓名")
        agg = agg.merge(m_merge, on="姓名", how="left")
        agg = agg.sort_values(agg.columns[3], key=lambda s: s.fillna("")).reset_index(drop=True)

        st.subheader("区间合计（每人）")
        fmt = {m: "{:,.0f}" for m in metrics}
        fmt["BOSS账号数量"] = "{:,.0f}"
        fmt["目标在职"] = "{:,.0f}"
        fmt["达成率"] = "{:.2f}"
        st.dataframe(agg.style
                     .background_gradient(subset=metrics, cmap="Blues")
                     .format(fmt, na_rep="0"),
                     width="stretch", height=380)

        st.write("")
        for m in metrics:
            st.plotly_chart(heatmap_fig(sel, m), width="stretch")

    # ---- Tab 3: 趋势分析 ----
    with tab3:
        t1, t2 = st.columns(2)
        with t1:
            sel_m = st.selectbox("指标（趋势图）", metrics, key="trend_metric")
        with t2:
            view = st.radio("维度", ["每人", "全组合计"], horizontal=True, key="trend_view")
        sub = sel[sel["指标"] == sel_m]
        if view == "每人":
            st.plotly_chart(trend_fig(sub, "数值", f"{sel_m} · 每日趋势（按人）"),
                            width="stretch")
        else:
            daily = sub.groupby("日期", as_index=False)["数值"].sum()
            fig = px.bar(daily, x="日期", y="数值", text="数值",
                         title=f"{sel_m} · 全组每日合计",
                         color_discrete_sequence=[METRIC_COLORS.get(sel_m, "#4A90D9")])
            fig.update_traces(textposition="outside", texttemplate="%{text:.0f}")
            fig.update_layout(template="plotly_dark", height=420,
                              margin=dict(l=10, r=10, t=50, b=10))
            fig.update_xaxes(dtick="D1", tickformat="%m-%d")
            st.plotly_chart(fig, width="stretch")

        # 多指标对比（全组合计）
        daily_all = sel.groupby(["日期", "指标"], as_index=False)["数值"].sum()
        fig2 = px.line(daily_all, x="日期", y="数值", color="指标", markers=True,
                       title="全组 · 多指标每日合计对比",
                       )
        fig2.update_layout(template="plotly_dark", height=420,
                           legend=dict(orientation="h", y=-0.25),
                           margin=dict(l=10, r=10, t=50, b=10))
        fig2.update_xaxes(dtick="D1", tickformat="%m-%d")
        st.plotly_chart(fig2, width="stretch")

    # ---- Tab 4: 数据表 ----
    with tab4:
        st.subheader("每日明细（长表）")
        show = sel.sort_values(["日期", "组别", "姓名"])
        st.dataframe(show, width="stretch", height=420)

        # 透视表下载
        pivot_out = sel.pivot_table(index=["日期", "姓名", "组别", "岗位"],
                                    columns="指标", values="数值", aggfunc="sum").reset_index()
        buf = io.BytesIO()
        with pd.ExcelWriter(buf, engine="openpyxl") as writer:
            pivot_out.to_excel(writer, index=False, sheet_name="区间汇总")
        st.download_button("📥 下载区间汇总 Excel", buf.getvalue(),
                           file_name=f"粤徽看板_{d_start}_{d_end}.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


if __name__ == "__main__":
    main()
