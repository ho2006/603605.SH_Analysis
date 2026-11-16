# -*- coding: utf-8 -*-
"""
Plotly 高清研报图表生成器（PNG 静态图 + SVG 矢量图双轨输出）
================================================================
为「5 财务绩效分析」生成 6 幅深度可视化图表：
- fig1_growth_truth         : 增长的真伪——营收、归母净利与扣非净利剪刀差（含4.45亿并购重估收益拆解）
- fig2_brand_expense        : 产品生命周期分化与流量内卷黑洞（分品牌增速 vs 销售/研发费用率及流量ROE）
- fig3_peer_comparison      : 珀莱雅 vs 贝泰妮横向四维对标（毛利率、研发投入率、扣非增速反转、商誉占净资产比）
- fig4_dupont_decomposition : 杜邦分析双轨拆解（名义ROE vs 扣非真实ROE及三因子驱动演变）
- fig5_eva_goodwill         : 穿透会计利润表象：会计归母净利润 vs 经济增加值（EVA）与资本成本比较
- fig6_cashflow_bsc         : 现金流质量恶化（OCF、净现比、存货积压）与平衡计分卡（BSC）四维评价雷达图

每幅图表同时输出：
- assets/charts/<name>.png (供 docxtpl Word 嵌入，scale=2 高清静态图)
- assets/charts/<name>.svg (供 HTML + WeasyPrint PDF 嵌入，矢量无损)
"""

from __future__ import annotations

import json
from pathlib import Path

import pandas as pd
import plotly.graph_objects as go
from plotly.subplots import make_subplots

BASE_DIR = Path(__file__).resolve().parents[1]  # 仓库根目录（src/ 的上一级），避免硬编码本机绝对路径
JSON_PATH = BASE_DIR / "data" / "computed_metrics.json"
OUT_DIR = BASE_DIR / "assets" / "charts"
OUT_DIR.mkdir(parents=True, exist_ok=True)

# 机构研报级配色方案
C_NAVY = "#1B365D"      # 深海蓝（主品牌/珀莱雅/主柱）
C_BLUE = "#3B7DDD"      # 研报蓝
C_TEAL = "#2E8B57"      # 翠绿（贝泰妮/正向/修复）
C_RED = "#C0392B"       # 琥珀红（扣非下滑/风险警示）
C_GOLD = "#D9822B"      # 暖金（并购重估泡沫/非经常性损益）
C_SLATE = "#6C7A89"     # 石板灰
C_BG = "#FFFFFF"
C_GRID = "#E6ECF2"

FONT_FAMILY = "Microsoft YaHei, SimHei, PingFang SC, Noto Sans CJK SC, sans-serif"


def apply_base_layout(fig: go.Figure, title: str, subtitle: str, height: int = 560, width: int = 1040) -> go.Figure:
    fig.update_layout(
        title=dict(
            text=f"<b>{title}</b><br><span style='font-size:12px;color:#5D6D7E'>{subtitle}</span>",
            x=0.02,
            xanchor="left",
            y=0.97,
            font=dict(family=FONT_FAMILY, size=15.5, color="#17202A"),
        ),
        font=dict(family=FONT_FAMILY, size=11, color="#2C3E50"),
        plot_bgcolor=C_BG,
        paper_bgcolor=C_BG,
        width=width,
        height=height,
        margin=dict(l=64, r=64, t=122, b=58),
        legend=dict(
            orientation="h",
            yanchor="bottom",
            y=1.08,
            xanchor="center",
            x=0.5,
            bgcolor="rgba(248,250,252,0.92)",
            bordercolor="#DCE3EA",
            borderwidth=0.6,
            font=dict(size=10.2),
        ),
    )
    # 适当提升子图标题 y 位置，避免与图例重叠或过紧
    for ann in fig.layout.annotations:
        if ann.text and ann.text.startswith("("):
            ann.y = ann.y + 0.015 if ann.y < 0.9 else 1.02
    fig.update_xaxes(showgrid=False, linecolor="#BFC9CA", linewidth=1, tickfont=dict(size=10.5))
    fig.update_yaxes(showgrid=True, gridcolor=C_GRID, gridwidth=0.8, zeroline=True, zerolinecolor="#ABB2B9", zerolinewidth=1)
    return fig


def save_fig(fig: go.Figure, name: str, width: int = 1040, height: int = 560) -> None:
    png_path = OUT_DIR / f"{name}.png"
    svg_path = OUT_DIR / f"{name}.svg"
    fig.write_image(str(png_path), width=width, height=height, scale=2)
    fig.write_image(str(svg_path), width=width, height=height)
    print(f"[CHART] Generated {name}.png & {name}.svg")


def chart_1_growth_truth(df: pd.DataFrame) -> None:
    proya = df[(df["company"] == "珀莱雅") & (df["period"] != "2026Q3单季")].copy()
    proya_all = df[df["company"] == "珀莱雅"].copy()

    fig = make_subplots(
        rows=1, cols=2,
        column_widths=[0.54, 0.46],
        specs=[[{"secondary_y": True}, {"secondary_y": False}]],
        subplot_titles=("(a) 利润结构拆解：扣非利润与并购重估泡沫（亿元）", "(b) 增速背离：2026营收、归母与扣非净利剪刀差（%）"),
        horizontal_spacing=0.14,
    )

    # 左图：堆叠柱状图（扣非净利润 + 非经常性损益）+ 营收折线；右图：增速背离分组柱状图
    periods = proya["period"].tolist()
    sub_g = proya_all[proya_all["period"].isin(["2024A", "2025A", "2026Q1-Q3", "2026Q3单季"])]
    gp = sub_g["period"].tolist()

    # 使用 barmode="group" + offsetgroup/base 同时实现左图堆叠效果与右图分组效果
    fig.add_trace(go.Bar(
        x=periods, y=proya["deducted_np"], name="扣非归母净利润（亿元）",
        marker_color=C_NAVY, offsetgroup="left_stack",
        text=[f"{v:.2f}" for v in proya["deducted_np"]], textposition="inside",
        textfont=dict(color="white", size=10),
    ), row=1, col=1, secondary_y=False)
    fig.add_trace(go.Bar(
        x=periods, y=proya["non_recurring_np"], base=proya["deducted_np"], name="非经常性损益（2026含4.45亿重估）",
        marker_color=C_GOLD, offsetgroup="left_stack",
        text=[f"+{v:.2f}" if v >= 0.4 else "" for v in proya["non_recurring_np"]],
        textposition="inside", textfont=dict(color="white", size=10),
    ), row=1, col=1, secondary_y=False)
    fig.add_trace(go.Scatter(
        x=periods, y=proya["revenue"], name="营业收入（右轴，亿元）",
        mode="lines+markers+text", line=dict(color=C_BLUE, width=2.8), marker=dict(size=7),
        text=[f"{v:.1f}" for v in proya["revenue"]],
        textposition=["top center", "top center", "top center", "top center", "bottom center"],
        textfont=dict(color=C_BLUE, size=10),
    ), row=1, col=1, secondary_y=True)

    fig.add_trace(go.Bar(
        x=gp, y=sub_g["rev_yoy"], name="营收同比(%)",
        marker_color=C_BLUE, offsetgroup="g1",
        text=[f"{v:+.1f}%" for v in sub_g["rev_yoy"]], textposition="outside", textfont=dict(size=9),
    ), row=1, col=2)
    fig.add_trace(go.Bar(
        x=gp, y=sub_g["np_yoy"], name="归母净利同比(%)",
        marker_color=C_GOLD, offsetgroup="g2",
        text=[f"{v:+.1f}%" for v in sub_g["np_yoy"]], textposition="outside", textfont=dict(size=9),
    ), row=1, col=2)
    fig.add_trace(go.Bar(
        x=gp, y=sub_g["dnp_yoy"], name="扣非净利同比(%)",
        marker_color=C_RED, offsetgroup="g3",
        text=[f"{v:+.1f}%" for v in sub_g["dnp_yoy"]], textposition="outside", textfont=dict(size=9, color=C_RED),
    ), row=1, col=2)

    fig.update_layout(barmode="group")
    apply_base_layout(fig, "图1  珀莱雅“增长的真伪”：并购收益虚增与扣非利润大幅背离（2022–2026Q3）",
                      "数据来源：珀莱雅定期报告、CNEquity 本地 DuckDB 数据湖；单位：亿元 / %", height=560)
    fig.update_yaxes(title_text="利润额（亿元）", range=[0, 19.5], row=1, col=1, secondary_y=False)
    fig.update_yaxes(title_text="营业收入（亿元）", range=[0, 135], showgrid=False, row=1, col=1, secondary_y=True)
    fig.update_yaxes(title_text="同比增速（%）", range=[-82, 48], row=1, col=2)
    fig.update_annotations(font=dict(size=12, color="#1B365D"))
    save_fig(fig, "fig1_growth_truth", height=560)


def chart_2_brand_expense(df: pd.DataFrame, df_brands: pd.DataFrame) -> None:
    proya = df[df["company"] == "珀莱雅"].copy()
    fig = make_subplots(
        rows=1, cols=2,
        column_widths=[0.44, 0.56],
        specs=[[{"secondary_y": False}, {"secondary_y": True}]],
        subplot_titles=("(a) 内生核心失速：分品牌同比增速分化（%）", "(b) 流量内卷：销售费用率攀升与流量ROE衰减"),
        horizontal_spacing=0.14,
    )

    # 左图：分品牌 2024A, 2025A, 2026H1 同比增速
    focus_brands = ["珀莱雅（主品牌）", "彩棠（Timage）", "Off&Relax（OR）", "原色波塔（INSBAHA）"]
    short_names = {"珀莱雅（主品牌）": "主品牌珀莱雅", "彩棠（Timage）": "彩棠(第二曲线)", "Off&Relax（OR）": "OR洗护", "原色波塔（INSBAHA）": "原色波塔"}
    colors_p = {"2024A": C_SLATE, "2025A": C_BLUE, "2026H1": C_RED}
    for per in ["2024A", "2025A", "2026H1"]:
        sub = df_brands[(df_brands["period"] == per) & (df_brands["brand"].isin(focus_brands))].set_index("brand").reindex(focus_brands)
        fig.add_trace(go.Bar(
            x=[short_names[b] for b in focus_brands],
            y=sub["yoy"].tolist(),
            name=f"{per}增速",
            marker_color=colors_p[per],
            text=[f"{v:+.1f}%" for v in sub["yoy"]],
            textposition="outside",
            textfont=dict(size=9),
        ), row=1, col=1)

    # 右图：销售费用率 vs 研发费用率（左轴），流量扣非ROE = deducted_np / selling_exp（右轴）
    periods_all = proya["period"].tolist()
    fig.add_trace(go.Scatter(
        x=periods_all, y=proya["selling_exp_ratio"], name="销售费用率（左轴,%）",
        mode="lines+markers+text", line=dict(color=C_RED, width=3), marker=dict(size=7),
        text=[f"{v:.1f}%" for v in proya["selling_exp_ratio"]], textposition="top center",
        textfont=dict(color=C_RED, size=9.5),
    ), row=1, col=2, secondary_y=False)

    fig.add_trace(go.Scatter(
        x=periods_all, y=proya["rd_exp_ratio"], name="研发费用率（左轴,%）",
        mode="lines+markers+text", line=dict(color=C_SLATE, width=2.2, dash="dash"), marker=dict(size=6),
        text=[f"{v:.2f}%" for v in proya["rd_exp_ratio"]], textposition="top center",
        textfont=dict(color=C_SLATE, size=9),
    ), row=1, col=2, secondary_y=False)

    fig.add_trace(go.Bar(
        x=periods_all, y=proya["traffic_dnp_roi"], name="流量净回报率=扣非净利/销售费用（右轴,%）",
        marker_color="rgba(27, 54, 93, 0.25)", marker_line=dict(color=C_NAVY, width=1.2),
        text=[f"{v:.1f}%" for v in proya["traffic_dnp_roi"]], textposition="inside",
        textfont=dict(color=C_NAVY, size=9.5),
    ), row=1, col=2, secondary_y=True)

    fig.update_layout(barmode="group")
    apply_base_layout(fig, "图2  产品生命周期分化与“流量内卷”：大单品承压、费用率飙升与流量ROE下行",
                      "注：2026H1主品牌（-7.19%）与彩棠（-21.93%）同步下滑；右图流量净回报率=扣非归母净利润÷销售费用", height=560)
    fig.update_yaxes(title_text="分品牌同比增速（%）", range=[-45, 260], row=1, col=1)
    fig.update_yaxes(title_text="费用率（%）", range=[0, 64], row=1, col=2, secondary_y=False)
    fig.update_yaxes(title_text="流量净回报率（%）", range=[0, 42], showgrid=False, row=1, col=2, secondary_y=True)
    fig.update_annotations(font=dict(size=12, color="#1B365D"))
    save_fig(fig, "fig2_brand_expense", height=560)


def chart_3_peer_comparison(df: pd.DataFrame) -> None:
    sub = df[df["period"].isin(["2023A", "2024A", "2025A", "2026Q1-Q3"])].copy()
    p = sub[sub["company"] == "珀莱雅"].set_index("period")
    b = sub[sub["company"] == "贝泰妮"].set_index("period")
    pers = ["2023A", "2024A", "2025A", "2026Q1-Q3"]

    fig = make_subplots(
        rows=2, cols=2,
        subplot_titles=(
            "(a) 毛利率对比（%）：贝泰妮稳健 vs 珀莱雅Q3稀释",
            "(b) 研发费用率对比（%）：壁垒差距显著",
            "(c) 扣非归母净利润同比增速（%）：剪刀差反转",
            "(d) 商誉占归母净资产比重对比（%）",
        ),
        vertical_spacing=0.18,
        horizontal_spacing=0.11,
    )

    # (a) 毛利率
    fig.add_trace(go.Scatter(x=pers, y=p.loc[pers, "gross_margin_rep"], name="珀莱雅", mode="lines+markers+text",
                             line=dict(color=C_NAVY, width=2.8), text=[f"{v:.2f}%" for v in p.loc[pers, "gross_margin_rep"]],
                             textposition="bottom center", textfont=dict(size=9.5, color=C_NAVY)), row=1, col=1)
    fig.add_trace(go.Scatter(x=pers, y=b.loc[pers, "gross_margin_rep"], name="贝泰妮", mode="lines+markers+text",
                             line=dict(color=C_TEAL, width=2.8), text=[f"{v:.2f}%" for v in b.loc[pers, "gross_margin_rep"]],
                             textposition="top center", textfont=dict(size=9.5, color=C_TEAL)), row=1, col=1)

    # (b) 研发费用率
    fig.add_trace(go.Bar(x=pers, y=p.loc[pers, "rd_exp_ratio"], name="珀莱雅", showlegend=False,
                         marker_color=C_NAVY, text=[f"{v:.2f}%" for v in p.loc[pers, "rd_exp_ratio"]], textposition="outside", textfont=dict(size=9)), row=1, col=2)
    fig.add_trace(go.Bar(x=pers, y=b.loc[pers, "rd_exp_ratio"], name="贝泰妮", showlegend=False,
                         marker_color=C_TEAL, text=[f"{v:.2f}%" for v in b.loc[pers, "rd_exp_ratio"]], textposition="outside", textfont=dict(size=9)), row=1, col=2)

    # (c) 扣非净利增速
    fig.add_trace(go.Bar(x=pers, y=p.loc[pers, "dnp_yoy"], name="珀莱雅", showlegend=False,
                         marker_color=C_RED, text=[f"{v:+.1f}%" for v in p.loc[pers, "dnp_yoy"]], textposition="outside", textfont=dict(size=9)), row=2, col=1)
    fig.add_trace(go.Bar(x=pers, y=b.loc[pers, "dnp_yoy"], name="贝泰妮", showlegend=False,
                         marker_color=C_TEAL, text=[f"{v:+.1f}%" for v in b.loc[pers, "dnp_yoy"]], textposition="outside", textfont=dict(size=9)), row=2, col=1)

    # (d) 商誉占归母净资产比重
    fig.add_trace(go.Bar(x=pers, y=p.loc[pers, "goodwill_to_equity_pct"], name="珀莱雅商誉/净资产", showlegend=False,
                         marker_color=C_GOLD, text=[f"{v:.2f}%" for v in p.loc[pers, "goodwill_to_equity_pct"]], textposition="outside", textfont=dict(size=9)), row=2, col=2)
    fig.add_trace(go.Bar(x=pers, y=b.loc[pers, "goodwill_to_equity_pct"], name="贝泰妮商誉/净资产", showlegend=False,
                         marker_color=C_SLATE, text=[f"{v:.2f}%" for v in b.loc[pers, "goodwill_to_equity_pct"]], textposition="outside", textfont=dict(size=9)), row=2, col=2)

    fig.update_layout(barmode="group")
    apply_base_layout(fig, "图3  珀莱雅 vs 贝泰妮横向四维对标：研发护城河、周期错位与商誉迁移（2023–2026Q3）",
                      "说明：贝泰妮经历2023–2024悦江并购阵痛后扣非增速转正（+48.2%→+21.4%），珀莱雅则步入高商誉与扣非下行通道", height=620)
    fig.update_layout(margin=dict(t=132))
    fig.update_yaxes(range=[67, 77.5], row=1, col=1)
    fig.update_yaxes(range=[0, 7.2], row=1, col=2)
    fig.update_yaxes(range=[-78, 68], row=2, col=1)
    fig.update_yaxes(range=[0, 18.5], row=2, col=2)
    fig.update_annotations(font=dict(size=11.5, color="#1B365D"))
    save_fig(fig, "fig3_peer_comparison", height=620)


def chart_4_dupont_decomposition(df: pd.DataFrame) -> None:
    sub = df[df["period"].isin(["2023A", "2024A", "2025A", "2026Q1-Q3"])].copy()
    p = sub[sub["company"] == "珀莱雅"].set_index("period")
    b = sub[sub["company"] == "贝泰妮"].set_index("period")
    pers = ["2023A", "2024A", "2025A", "2026Q1-Q3"]

    fig = make_subplots(
        rows=1, cols=2,
        column_widths=[0.50, 0.50],
        subplot_titles=("(a) 珀莱雅双轨 ROE 对比：名义虚高 vs 扣非真实下行（%）", "(b) 杜邦三因子演变（净利率 / 年化周转率 / 权益乘数）"),
        specs=[[{"secondary_y": False}, {"secondary_y": True}]],
        horizontal_spacing=0.14,
    )

    fig.add_trace(go.Bar(
        x=pers, y=p.loc[pers, "dupont_roe_gaap_ann"], name="珀莱雅名义年化ROE(GAAP)",
        marker_color=C_GOLD, text=[f"{v:.2f}%" for v in p.loc[pers, "dupont_roe_gaap_ann"]], textposition="outside", textfont=dict(size=9.5),
    ), row=1, col=1)
    fig.add_trace(go.Bar(
        x=pers, y=p.loc[pers, "dupont_roe_deducted_ann"], name="珀莱雅扣非年化ROE(真实)",
        marker_color=C_NAVY, text=[f"{v:.2f}%" for v in p.loc[pers, "dupont_roe_deducted_ann"]], textposition="outside", textfont=dict(size=9.5),
    ), row=1, col=1)
    fig.add_trace(go.Scatter(
        x=pers, y=b.loc[pers, "dupont_roe_deducted_ann"], name="贝泰妮扣非年化ROE",
        mode="lines+markers+text", line=dict(color=C_TEAL, width=2.6), marker=dict(size=7),
        text=[f"{v:.2f}%" for v in b.loc[pers, "dupont_roe_deducted_ann"]], textposition="top center", textfont=dict(size=9.5, color=C_TEAL),
    ), row=1, col=1)

    # 右图：珀莱雅三因子分解
    fig.add_trace(go.Scatter(
        x=pers, y=p.loc[pers, "npm_deducted_pct"], name="扣非销售净利率（左轴,%）",
        mode="lines+markers+text", line=dict(color=C_RED, width=2.8),
        text=[f"{v:.2f}%" for v in p.loc[pers, "npm_deducted_pct"]],
        textposition=["top center", "top center", "top center", "top right"],
        textfont=dict(size=9.5, color=C_RED),
    ), row=1, col=2, secondary_y=False)

    fig.add_trace(go.Scatter(
        x=pers, y=p.loc[pers, "asset_turnover_ann"], name="年化总资产周转率（右轴,次）",
        mode="lines+markers+text", line=dict(color=C_BLUE, width=2.5),
        text=[f"{v:.2f}x" for v in p.loc[pers, "asset_turnover_ann"]], textposition="bottom center", textfont=dict(size=9.5, color=C_BLUE),
    ), row=1, col=2, secondary_y=True)

    fig.add_trace(go.Scatter(
        x=pers, y=p.loc[pers, "equity_multiplier"], name="权益乘数（右轴,倍）",
        mode="lines+markers+text", line=dict(color=C_SLATE, width=2.2, dash="dot"),
        text=[f"{v:.2f}x" for v in p.loc[pers, "equity_multiplier"]], textposition="top center", textfont=dict(size=9.5, color=C_SLATE),
    ), row=1, col=2, secondary_y=True)

    fig.update_layout(barmode="group")
    apply_base_layout(fig, "图4  双轨杜邦三因子拆解：珀莱雅会计泡沫（10.96pct）与真实经营驱动退化",
                      "注：2026Q1-Q3 珀莱雅名义年化ROE高达26.83%，剔除并购重估后真实扣非年化ROE跌至15.87%，周转率跌破1.1x", height=560)
    fig.update_yaxes(title_text="年化 ROE（%）", range=[0, 38], row=1, col=1)
    fig.update_yaxes(title_text="扣非净利率（%）", range=[6, 17.5], row=1, col=2, secondary_y=False)
    fig.update_yaxes(title_text="周转率(次) / 权益乘数(倍)", range=[0.8, 2.05], showgrid=False, row=1, col=2, secondary_y=True)
    fig.update_annotations(font=dict(size=12, color="#1B365D"))
    save_fig(fig, "fig4_dupont_decomposition", height=560)


def chart_5_eva_goodwill(df: pd.DataFrame) -> None:
    proya = df[df["company"] == "珀莱雅"].copy()
    pers = proya["period"].tolist()

    fig = make_subplots(
        rows=1, cols=2,
        column_widths=[0.58, 0.42],
        specs=[[{"secondary_y": False}, {"secondary_y": False}]],
        subplot_titles=("(a) 会计归母净利 vs 真实 EVA（亿元）", "(b) ROIC 年化回报率 vs WACC 资本成本率（%）"),
        horizontal_spacing=0.10,
    )

    fig.add_trace(go.Bar(
        x=pers, y=proya["net_profit_parent"], name="会计归母净利润（亿元）",
        marker_color=C_BLUE, text=[f"{v:.2f}" for v in proya["net_profit_parent"]], textposition="outside", textfont=dict(size=9),
    ), row=1, col=1)
    fig.add_trace(go.Bar(
        x=pers, y=proya["capital_charge"], name="投入资本成本 TC×WACC（亿元）",
        marker_color=C_SLATE, text=[f"{v:.2f}" for v in proya["capital_charge"]], textposition="outside", textfont=dict(size=9),
    ), row=1, col=1)
    fig.add_trace(go.Bar(
        x=pers, y=proya["eva"], name="经济增加值 EVA（亿元）",
        marker_color=[C_RED if v < 0 else C_TEAL for v in proya["eva"]],
        text=[f"{v:.2f}" for v in proya["eva"]], textposition="outside", textfont=dict(size=9.5, color="#17202A"),
    ), row=1, col=1)

    fig.add_trace(go.Scatter(
        x=pers, y=proya["roic_ann_pct"], name="年化真实 ROIC（%）",
        mode="lines+markers+text", line=dict(color=C_NAVY, width=2.8), marker=dict(size=7),
        text=[f"{v:.1f}%" for v in proya["roic_ann_pct"]], textposition="top center", textfont=dict(size=9.5, color=C_NAVY),
    ), row=1, col=2)
    fig.add_trace(go.Scatter(
        x=pers, y=proya["wacc_pct"], name="WACC 资本成本线（%）",
        mode="lines+markers", line=dict(color=C_RED, width=2.2, dash="dash"), marker=dict(size=5),
    ), row=1, col=2)

    fig.update_layout(barmode="group")
    apply_base_layout(fig, "图5  EVA 真实价值穿透：并购虚增账面利润与边际资本价值毁损（2022–2026Q3）",
                      "注：2026Q1-Q3 会计归母净利12.56亿（+22.46%），但剔除并购泡沫并扣除72.53亿投入资本对应的4.18亿期间资本成本后，EVA仅3.49亿且Q3单季转负（-0.53亿）", height=560)
    fig.update_yaxes(title_text="金额（亿元）", range=[-3.5, 19], row=1, col=1)
    fig.update_yaxes(title_text="资本回报与成本率（%）", range=[0, 29], row=1, col=2)
    fig.update_annotations(font=dict(size=12, color="#1B365D"))
    save_fig(fig, "fig5_eva_goodwill", height=560)


def chart_6_cashflow_bsc(df: pd.DataFrame, df_bsc: pd.DataFrame) -> None:
    proya = df[df["company"] == "珀莱雅"].copy()
    pers = proya["period"].tolist()

    fig = make_subplots(
        rows=1, cols=2,
        column_widths=[0.50, 0.50],
        specs=[[{"secondary_y": True}, {"type": "polar"}]],
        subplot_titles=("(a) 珀莱雅经营现金流、净现比与存货积压", "(b) 平衡计分卡（BSC）四维健康度评价（0-100分）"),
        horizontal_spacing=0.14,
    )

    fig.add_trace(go.Bar(
        x=pers, y=proya["ocf"], name="经营现金流净额（左轴,亿元）",
        marker_color=[C_RED if v < 5 else C_NAVY for v in proya["ocf"]],
        text=[f"{v:.2f}" for v in proya["ocf"]], textposition="outside", textfont=dict(size=9),
    ), row=1, col=1, secondary_y=False)

    fig.add_trace(go.Scatter(
        x=pers, y=proya["inventory"], name="期末存货余额（左轴,亿元）",
        mode="lines+markers+text", line=dict(color=C_GOLD, width=2.6), marker=dict(size=7),
        text=[f"{v:.2f}" for v in proya["inventory"]],
        textposition=["bottom center", "bottom center", "bottom center", "bottom center", "top center", "top center"],
        textfont=dict(size=9, color=C_GOLD),
    ), row=1, col=1, secondary_y=False)

    fig.add_trace(go.Scatter(
        x=pers, y=proya["ocf_to_np_ratio"], name="净现比=OCF/归母净利（右轴）",
        mode="lines+markers", line=dict(color=C_BLUE, width=2.2, dash="dot"), marker=dict(size=6),
    ), row=1, col=1, secondary_y=True)

    # 右图：BSC 雷达图（对比 珀莱雅2023A, 珀莱雅2026Q1-Q3, 贝泰妮2026Q1-Q3）
    categories = ["财务价值与EVA", "客户与流量效率", "内部运营与<br>单品周期", "研发创新与并购治理"]
    bsc_map = {
        "珀莱雅 (2023A 黄金增长期)": (C_BLUE, "珀莱雅 2023A（86.30分）"),
        "珀莱雅 (2026Q1-Q3 并购粉饰期)": (C_RED, "珀莱雅 2026Q1-Q3（52.60分）"),
        "贝泰妮 (2026Q1-Q3 控费修复期)": (C_TEAL, "贝泰妮 2026Q1-Q3（71.85分）"),
    }
    for ent, (col, label) in bsc_map.items():
        r = df_bsc[df_bsc["entity"] == ent].iloc[0]
        vals = [r["financial_score"], r["customer_traffic_score"], r["internal_product_score"], r["learning_ma_score"]]
        vals_closed = vals + [vals[0]]
        cats_closed = categories + [categories[0]]
        fig.add_trace(go.Scatterpolar(
            r=vals_closed, theta=cats_closed, fill="toself",
            name=label, line=dict(color=col, width=2.4), opacity=0.72,
        ), row=1, col=2)

    apply_base_layout(fig, "图6  现金流与盈利质量评价及平衡计分卡（BSC）四维健康度诊断",
                      "说明：2026Q1-Q3珀莱雅OCF骤降61.06%至4.67亿元（Q3单季流出4.84亿），净现比跌至0.37，BSC综合评分由86.30分降至52.60分", height=580)
    fig.update_layout(margin=dict(r=78, t=128))
    fig.update_yaxes(title_text="金额（亿元）", range=[-8, 26], row=1, col=1, secondary_y=False)
    fig.update_yaxes(title_text="净现比（倍）", range=[-7, 2.5], showgrid=False, row=1, col=1, secondary_y=True)
    fig.update_polars(domain=dict(x=[0.57, 0.94], y=[0.08, 0.85]),
                      radialaxis=dict(visible=True, range=[0, 100], tickfont=dict(size=8.5)),
                      angularaxis=dict(tickfont=dict(size=10)))
    fig.update_annotations(font=dict(size=12, color="#1B365D"))
    save_fig(fig, "fig6_cashflow_bsc", height=580)


def main() -> None:
    payload = json.loads(JSON_PATH.read_text(encoding="utf-8"))
    df = pd.DataFrame(payload["metrics"])
    df_brands = pd.DataFrame(payload["brands"])
    df_bsc = pd.DataFrame(payload["bsc"])

    chart_1_growth_truth(df)
    chart_2_brand_expense(df, df_brands)
    chart_3_peer_comparison(df)
    chart_4_dupont_decomposition(df)
    chart_5_eva_goodwill(df)
    chart_6_cashflow_bsc(df, df_bsc)
    print("[OK] All 6 charts exported as PNG and SVG.")


if __name__ == "__main__":
    main()
