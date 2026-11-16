# -*- coding: utf-8 -*-
"""
Word（docxtpl + python-docx + PNG）与 PDF（HTML + WeasyPrint + SVG）双格式报告生成器
========================================================================================
职责：
1. 读取本地 DuckDB / JSON 计算结果，构建面向商业分析文章「三、珀莱雅 —— 5 财务绩效分析」
   的完整结构化上下文数据与深度学术研报正文。
2. Word 输出链路：
   - 使用 python-docx 生成结构化、带学术三线表样式与 Jinja2 占位符的模板
     templates/report_template.docx；
   - 使用 docxtpl (DocxTemplate + InlineImage) 填充指标数据与 6 幅高清 PNG 静态图，
     生成交付文件：`珀莱雅财务绩效分析报告_5_财务绩效分析.docx`。
3. PDF 输出链路：
   - 构建内嵌 6 幅 SVG 矢量图、KaTeX/MathML 风格排版与三线表 CSS 的 HTML 报告
     `build/report.html`；
   - 自动注入本地 WeasyPrint Pango/GObject DLL 目录后调用 WeasyPrint 渲染交付文件：
     `珀莱雅财务绩效分析报告_5_财务绩效分析.pdf`。
"""

from __future__ import annotations

import json
import os
import re
import sys
from pathlib import Path

from docx import Document
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor
from docxtpl import DocxTemplate, InlineImage

BASE_DIR = Path(__file__).resolve().parents[1]  # 仓库根目录（src/ 的上一级），避免硬编码本机绝对路径
DATA_JSON = BASE_DIR / "data" / "computed_metrics.json"
CHARTS_DIR = BASE_DIR / "assets" / "charts"
TEMPLATE_DIR = BASE_DIR / "templates"
BUILD_DIR = BASE_DIR / "build"

DOCX_TEMPLATE_PATH = TEMPLATE_DIR / "report_template.docx"
DOCX_OUT_PATH = BASE_DIR / "珀莱雅财务绩效分析报告_5_财务绩效分析.docx"
HTML_OUT_PATH = BUILD_DIR / "report.html"
PDF_OUT_PATH = BASE_DIR / "珀莱雅财务绩效分析报告_5_财务绩效分析.pdf"

# 注入 Windows WeasyPrint 原生 Pango/Cairo DLL 目录（优先使用当前 conda 环境的 Library/bin）
CONDA_DLL_DIR = Path(sys.prefix) / "Library" / "bin"
WEASYPRINT_DLL_DIR = BASE_DIR / "tools" / "weasyprint_win" / "onedir" / "weasyprint" / "_internal"
if CONDA_DLL_DIR.exists() and any(CONDA_DLL_DIR.glob("pango*.dll")):
    os.environ["WEASYPRINT_DLL_DIRECTORIES"] = str(CONDA_DLL_DIR)
elif WEASYPRINT_DLL_DIR.exists():
    os.environ["WEASYPRINT_DLL_DIRECTORIES"] = str(WEASYPRINT_DLL_DIR)

import weasyprint  # noqa: E402


# ==============================================================================
# 工具函数：格式化表格与数据提取
# ==============================================================================

def load_context() -> dict:
    raw = json.loads(DATA_JSON.read_text(encoding="utf-8"))
    df_m = {f"{r['company']}_{r['period']}": r for r in raw["metrics"]}

    def fmt(val: float, dec: int = 2, sign: bool = False) -> str:
        if val is None:
            return "—"
        return f"{val:+.{dec}f}" if sign else f"{val:.{dec}f}"

    # 表5-1：珀莱雅纵向核心财务指标（2022A–2026Q3单季）
    proya_periods = ["2022A", "2023A", "2024A", "2025A", "2026Q1-Q3", "2026Q3单季"]
    t51_headers = ["核心财务与经营指标", "2022A", "2023A", "2024A", "2025A", "2026Q1-Q3", "2026Q3单季"]
    t51_specs = [
        ("营业收入（亿元）", "revenue", 2, False),
        ("营收同比增速（%）", "rev_yoy", 2, True),
        ("综合毛利率（%）", "gross_margin_rep", 2, False),
        ("归母净利润（亿元）", "net_profit_parent", 2, False),
        ("归母净利同比（%）", "np_yoy", 2, True),
        ("扣非归母净利润（亿元）", "deducted_np", 2, False),
        ("扣非净利同比（%）", "dnp_yoy", 2, True),
        ("非经常性损益（亿元）", "non_recurring_np", 2, True),
        ("销售费用（亿元）", "selling_exp", 2, False),
        ("销售费用率（%）", "selling_exp_ratio", 2, False),
        ("研发费用率（%）", "rd_exp_ratio", 2, False),
        ("流量净回报率（扣非/销售费用,%）", "traffic_dnp_roi", 2, False),
        ("经营现金流净额（亿元）", "ocf", 2, False),
        ("期末商誉（亿元）", "goodwill", 2, False),
        ("期末存货（亿元）", "inventory", 2, False),
    ]
    t51_rows = []
    for label, key, dec, sign in t51_specs:
        row = [label] + [fmt(df_m[f"珀莱雅_{p}"][key], dec, sign) for p in proya_periods]
        t51_rows.append(row)

    # 表5-2：珀莱雅分品牌收入与增速结构表
    brand_order = ["珀莱雅（主品牌）", "彩棠（Timage）", "Off&Relax（OR）", "悦芙媞（UTIQUE）", "原色波塔（INSBAHA）"]
    b_map = {(r["brand"], r["period"]): r for r in raw["brands"]}
    t52_headers = ["品牌及定位", "2023A营收(亿)", "2024A营收(亿)", "2025A营收(亿)", "2025A同比", "2026H1营收(亿)", "2026H1同比", "2026H1占比"]
    t52_rows = []
    for b in brand_order:
        r23 = b_map.get((b, "2023A"))
        r24 = b_map.get((b, "2024A"))
        r25 = b_map.get((b, "2025A"))
        r26 = b_map.get((b, "2026H1"))
        t52_rows.append([
            b,
            fmt(r23["revenue"]) if r23 else "—",
            fmt(r24["revenue"]) if r24 else "—",
            fmt(r25["revenue"]) if r25 else "—",
            f"{r25['yoy']:+.2f}%" if r25 else "—",
            fmt(r26["revenue"]) if r26 else "—",
            f"{r26['yoy']:+.2f}%" if r26 else "—",
            f"{r26['share_pct']:.2f}%" if r26 else "—",
        ])
    t52_rows.append(["花知晓（2026.06并表）", "未并表", "未并表", "独立17.26", "外延高增", "2026年6月起并表", "贡献Q3营收跳涨", "商誉+9.24亿"])

    # 表5-3：珀莱雅 vs 贝泰妮横向财务对标表（2023A–2026Q1-Q3）
    t53_headers = ["对标维度与指标", "公司", "2023A", "2024A", "2025A", "2026Q1-Q3"]
    peer_pers = ["2023A", "2024A", "2025A", "2026Q1-Q3"]
    peer_specs = [
        ("营业收入（亿元）", "revenue", 2, False),
        ("营收同比增速（%）", "rev_yoy", 2, True),
        ("综合毛利率（%）", "gross_margin_rep", 2, False),
        ("扣非归母净利润（亿元）", "deducted_np", 2, False),
        ("扣非净利同比增速（%）", "dnp_yoy", 2, True),
        ("销售费用率（%）", "selling_exp_ratio", 2, False),
        ("研发费用率（%）", "rd_exp_ratio", 2, False),
        ("流量净回报率（扣非/销售费用,%）", "traffic_dnp_roi", 2, False),
        ("商誉占归母净资产比（%）", "goodwill_to_equity_pct", 2, False),
        ("加权平均ROE（披露口径,%）", "roe_reported", 2, False),
    ]
    t53_rows = []
    for label, key, dec, sign in peer_specs:
        for comp in ["珀莱雅", "贝泰妮"]:
            t53_rows.append([label if comp == "珀莱雅" else "〃", comp] + [fmt(df_m[f"{comp}_{p}"][key], dec, sign) for p in peer_pers])

    # 表5-4：双轨杜邦三因子拆解表（珀莱雅 vs 贝泰妮）
    t54_headers = ["主体与期间", "名义净利率(%)", "扣非净利率(%)", "年化资产周转率(次)", "权益乘数(倍)", "名义年化ROE(%)", "扣非年化ROE(%)", "泡沫差值(pct)"]
    t54_rows = []
    for comp in ["珀莱雅", "贝泰妮"]:
        for p in ["2023A", "2024A", "2025A", "2026Q1-Q3"]:
            r = df_m[f"{comp}_{p}"]
            t54_rows.append([
                f"{comp} {p}",
                fmt(r["npm_gaap_pct"]),
                fmt(r["npm_deducted_pct"]),
                f"{r['asset_turnover_ann']:.2f}x",
                f"{r['equity_multiplier']:.2f}x",
                fmt(r["dupont_roe_gaap_ann"]),
                fmt(r["dupont_roe_deducted_ann"]),
                fmt(r["roe_bubble_spread"], 2, True),
            ])

    # 表5-5：珀莱雅 EVA（经济增加值）真实价值穿透测算表
    t55_headers = ["测算项目", "2022A", "2023A", "2024A", "2025A", "2026Q1-Q3", "2026Q3单季"]
    t55_specs = [
        ("归母净利润（GAAP,亿元）", "net_profit_parent", 2, False),
        ("真实NOPAT（扣非调整,亿元）", "nopat_core", 2, False),
        ("期末投入资本TC（亿元）", "invested_capital", 2, False),
        ("加权平均资本成本WACC（%）", "wacc_pct", 2, False),
        ("期间资本成本TC×WACC（亿元）", "capital_charge", 2, False),
        ("经济增加值EVA（亿元）", "eva", 2, True),
        ("年化投入资本回报率ROIC（%）", "roic_ann_pct", 2, False),
        ("EVA利差（ROIC-WACC,%）", "eva_spread_pct", 2, True),
    ]
    t55_rows = [[lbl] + [fmt(df_m[f"珀莱雅_{p}"][k], dec, s) for p in proya_periods] for lbl, k, dec, s in t55_specs]

    # 表5-6：现金流与盈利质量及 BSC 平衡计分卡评价表
    t56_headers = ["质量评价指标", "珀莱雅2023A", "珀莱雅2024A", "珀莱雅2025A", "珀莱雅2026Q1-Q3", "珀莱雅2026Q3单季", "贝泰妮2026Q1-Q3"]
    t56_cols = [("珀莱雅", "2023A"), ("珀莱雅", "2024A"), ("珀莱雅", "2025A"), ("珀莱雅", "2026Q1-Q3"), ("珀莱雅", "2026Q3单季"), ("贝泰妮", "2026Q1-Q3")]
    t56_specs = [
        ("经营现金流净额OCF（亿元）", "ocf", 2, False),
        ("收现比（销售收现/营收）", "cash_to_rev_ratio", 2, False),
        ("净现比（OCF/归母净利）", "ocf_to_np_ratio", 2, False),
        ("扣非净现比（OCF/扣非净利）", "ocf_to_dnp_ratio", 2, False),
        ("应计利润/平均总资产（%）", "accrual_ratio_pct", 2, True),
        ("期末存货余额（亿元）", "inventory", 2, False),
        ("存货周转天数DIO（天）", "dio_days", 1, False),
        ("流动负债现金覆盖率（倍）", "ocf_to_current_liab", 2, False),
    ]
    t56_rows = [[lbl] + [fmt(df_m[f"{c}_{p}"][k], dec, s) for c, p in t56_cols] for lbl, k, dec, s in t56_specs]

    t57_headers = ["评价对象与阶段", "财务价值(30%)", "客户流量(25%)", "内部运营(25%)", "学习治理(20%)", "BSC综合得分", "核心诊断结论"]
    t57_rows = [
        [
            r["entity"],
            f"{r['financial_score']:.1f}",
            f"{r['customer_traffic_score']:.1f}",
            f"{r['internal_product_score']:.1f}",
            f"{r['learning_ma_score']:.1f}",
            f"{r['composite_score']:.2f}",
            r["key_diagnosis"],
        ]
        for r in raw["bsc"]
    ]

    return {
        "t51": {"title": "表 5-1  珀莱雅（603605.SH）纵向核心财务绩效指标汇总（2022A–2026Q3）", "headers": t51_headers, "rows": t51_rows},
        "t52": {"title": "表 5-2  珀莱雅分品牌收入结构与增长分化（2023A–2026H1 及花知晓并表）", "headers": t52_headers, "rows": t52_rows},
        "t53": {"title": "表 5-3  珀莱雅（603605.SH）与贝泰妮（300957.SZ）横向核心指标对标", "headers": t53_headers, "rows": t53_rows},
        "t54": {"title": "表 5-4  双轨杜邦三因子拆解对比：名义 GAAP 口径 vs 扣非真实经营口径", "headers": t54_headers, "rows": t54_rows},
        "t55": {"title": "表 5-5  珀莱雅 EVA（经济增加值）真实价值创造穿透测算表", "headers": t55_headers, "rows": t55_rows},
        "t56": {"title": "表 5-6  珀莱雅与贝泰妮现金流质量、营运效率与盈利含金量对比", "headers": t56_headers, "rows": t56_rows},
        "t57": {"title": "表 5-7  平衡计分卡（BSC）四维健康度综合评价矩阵（0–100分制）", "headers": t57_headers, "rows": t57_rows},
    }


EXECUTIVE_SUMMARY = (
    "2026 年前三季度，珀莱雅账面营业收入同比增长 7.46% 至 76.27 亿元（Q3 单季 +29.76%），归母净利润同比增长 22.46% 至 12.56 亿元。"
    "穿透非经常性损益后，扣非归母净利润同比暴跌 25.50% 至 7.43 亿元（Q3 单季同比剧降 65.20%）。"
    "账面繁荣主要源于分步收购花知晓（持股升至 51% 并自 2026 年 6 月并表）产生的 4.45 亿元公允价值重估收益与外延并表。"
    "原有核心盘（主品牌 2025 年 -10.39%、2026H1 -7.19%，彩棠 2026H1 -21.93%）已步入大单品生命周期老化通道，"
    "销售费用率飙升至 53.23%，拖累流量净回报率跌至 6.56%（单季），经营现金流净额骤降 61.06%（Q3 单季净流出 4.84 亿元）。"
    "双轨杜邦与 EVA 穿透测算表明，珀莱雅扣非年化 ROE 已由 31.21% 跌落至 15.87%，2026Q3 单季 EVA 直接转负为 -0.53 亿元，"
    "同时伴随着 9.24 亿元高额商誉（占净资产 14.57%）悬顶风险，公司亟需告别“流量内卷+并购粉饰”，向研发驱动与 EVA 价值创造回归。"
)


def get_report_sections() -> list[dict]:
    """编写详实严谨、学术规范与商业洞见兼备的「5 财务绩效分析」完整文章结构。"""
    return [
        {
            "level": 1,
            "heading": "5 财务绩效分析",
            "paragraphs": [
                "本章综合运用文献研究法、案例分析法与财务比率分析法，依托 CNEquity 数据湖架构与 DuckDB 本地财务数据库，对珀莱雅化妆品股份有限公司（603605.SH）2022–2025 年年度报告及 2026 年第三季度报告进行全链路实证拆解，并引入行业核心对标标的云南贝泰妮生物科技集团股份有限公司（300957.SZ）进行横向对比。在指标体系定义上，本文参考 FinanceToolkit 的标准财务比率框架，构建营业成本率与费用率、现金流质量比率，以及杜邦恒等式三因子拆解。在此基础上，进一步构建“GAAP 名义口径—扣非真实口径”双轨杜邦分析模型与 EVA（经济增加值）穿透测算模型，剔除并购花知晓带来的一次性会计重估收益干扰，评估珀莱雅的真实资本价值创造能力。最后结合平衡计分卡（Balanced Scorecard, BSC）四维管理思想，将财务绩效指标与产品生命周期、流量营销边际效率结合，同时纳入研发投入与并购整合风险，全面剖析珀莱雅这一国货美妆龙头在“营收账面增长、扣非利润大幅下滑”表象下的高质量增长困境。",
            ],
        },
        {
            "level": 2,
            "heading": "5.1 核心财务指标纵向趋势分析（近3年及2026年最新进展）",
            "paragraphs": [
                "从 2022 至 2026 年第三季度的纵向财务轨迹来看，珀莱雅的演化历经三个阶段：高速内生红利期（2022–2024）、规模触顶拐点期（2025）以及并购并表掩盖主业滑坡期（2026）。如表 5-1 及图 1 所示，2022–2024 年公司营业收入由 63.85 亿元跃升至 107.78 亿元，成为首家营收突破百亿大关的国货美妆上市公司，同期扣非归母净利润由 7.89 亿元增至 15.22 亿元，主营业务呈现出高度的业绩一致性（2023–2024 年非经常性损益占比均不足 2%，2022 年亦仅 3.43%）。进入 2025 年后，珀莱雅营收首次出现负增长（105.97 亿元，同比下降 1.68%），扣非归母净利润同步回落 3.23% 至 14.73 亿元，标志着依托大单品红利实现自然扩张的周期已正式触及天花板。",
                "2026 年三季报披露，珀莱雅前三季度实现营业收入 76.27 亿元（同比增长 7.46%），其中第三季度单季营收更是高达 22.53 亿元（同比大幅跳涨 29.76%），同期归母净利润达 12.56 亿元（同比增长 22.46%）。尽管账面营收与归母净利润同比回升，扣除非经常性损益后的净利润却出现剧烈的“剪刀差背离”。2026 年前三季度扣非归母净利润仅录得 7.43 亿元，同比大幅下滑 25.50%。第三季度单季扣非归母净利润更暴跌至 0.79 亿元，同比断崖式下降 65.20%（单季归母净利润亦同比下降 61.18% 至 0.88 亿元）。",
            ],
            "table_key": "t51",
            "chart": {"file": "fig1_growth_truth", "caption": "图 1  珀莱雅“增长的真伪”：并购收益虚增与扣非利润大幅背离（2022–2026Q3）"},
            "sub_paragraphs": [
                "【切入点一：增长的真伪——并购并表制造的“会计性繁荣”】\n"
                "造成 2026 年营收与归母净利润账面增长、而扣非归母净利润大幅坍塌的核心原因，在于公司分步收购彩妆品牌“花知晓（Flower Knows）”并将其纳入合并报表，由此产生了会计调节效应。2025 年珀莱雅先以 4.28 亿元受让花知晓 38.45% 股权，2026 年 5 月再出资 3.51 亿元增持股权至 51% 实现控股，并自 2026 年 6 月起将其正式并表。花知晓独立运营主体在 2025 全年已具备 17.26 亿元营收与 2.80 亿元净利润的较大体量，其并表直接拉动珀莱雅 2026Q3 单季营收实现 29.76% 的高增速。依据《企业会计准则第 33 号——合并财务报表》，分步实现非同一控制下企业合并时，原持有的 38.45% 股权在购买日需按公允价值重新计量，由此产生高达 4.45 亿元的一次性非经常性投资收益，直接贡献了 2026Q1-Q3 归母净利润的 35.43%。2026 年前三季度 12.56 亿元归母净利润中，非经常性损益高达 5.13 亿元（占比 40.84%），账面高增主要由上述并购重估投资收益与外延并表撑起。",
                "【切入点二：产品生命周期——核心大单品老化与内生品牌失速】\n"
                "剥离外延并购后，表 5-2 清晰揭示了珀莱雅原有品牌矩阵的内生疲态。占集团收入基本盘近七成的主品牌“珀莱雅”营收已连续两个报告期负增长，由 2024 年峰值 85.81 亿元萎缩至 2025 年 76.89 亿元（-10.39%），2026H1 进一步同比下滑 7.19% 至 36.92 亿元，主营收入占比从 2022 年的 82.44% 滑落至 68.79%。早年间公司凭借“早 C 晚 A”（双抗、红宝石）及“源力修护”系列确立了大单品壁垒，如今正面临国货成分红利消退、竞品同质化跟进与迭代青黄不接的生命周期困境。更严峻的是，曾被视作“第二增长曲线”的专业彩妆品牌彩棠（Timage）亦在 2026H1 出现 21.93% 的显著负增长（营收 5.51 亿元）。尽管 Off&Relax（OR，2026H1 +70.81%）与原色波塔（+222.23%）尚处高速成长期，但其绝对体量仍不足以对冲主品牌与彩棠核心盘失血，且新品牌尚处于烧钱拓客期，未能转化为实质性扣非利润贡献。",
                "【切入点三：流量内卷困境——销售费用率突破53%与流量 ROE 断崖】\n"
                "自然流量红利枯竭、大单品复购承压后，珀莱雅被迫通过“砸钱买流量”维系规模体量。纵向数据显示，公司销售费用率已由 2022 年的 43.63% 逐年攀升至 2024 年的 47.88%、2025 年的 49.63%，并在 2026Q1-Q3 突破 53.23%（绝对额达 40.60 亿元），2026Q3 单季销售费用率更高达 53.44%（若按财报口径加计成长品牌前置推广费，单季推广费用占收入比已逼近行业极限）。与之形成鲜明反差的是，研发费用率长期徘徊在 1.95%–2.05% 的低位水平（2026Q1-Q3 仅 1.97%，约 1.50 亿元）。我们将单位销售费用创造的扣非归母净利润定义为“流量净回报率（流量 ROE）”。该指标在 2023 年尚维持在 29.56% 的健康水平，但至 2026Q1-Q3 断崖式暴跌至 18.30%，2026Q3 单季更是跌至仅 6.56%，即每投入 100 元销售费用，最终仅能换取 6.56 元主业扣非净利润。受低毛利彩妆并表及结构变化影响，公司 2026Q3 单季毛利率同比大幅稀释至 69.78%（前三季度 72.98%），进一步压缩了高投放下的盈利安全垫。",
            ],
            "table_key_2": "t52",
            "chart_2": {"file": "fig2_brand_expense", "caption": "图 2  产品生命周期分化与“流量内卷”：大单品承压、费用率飙升与流量 ROE 下行"},
        },
        {
            "level": 2,
            "heading": "5.2 与贝泰妮横向对标分析",
            "paragraphs": [
                "为客观剥离化妆品行业宏观消费周期波动的影响，本节选取功效性护肤龙头贝泰妮（300957.SZ）作为横向对标对象。贝泰妮定位高度可比，同样在 2023–2025 年间经历过大单品瓶颈与外延并购洗礼。通过对比两家企业的毛利护城河和研发营销资源配置，以及扣非利润周期走势与并购商誉后果，可以更加精准地识别珀莱雅当前财务失速的结构性成因。",
                "如表 5-3 及图 3 所示，在 2023–2024 年，珀莱雅与贝泰妮呈现出截然相反的经营周期。彼时珀莱雅凭借全渠道精细化运营与大单品放量，扣非净利润维持 +48.91% 与 +29.60% 的高增长，而贝泰妮则因核心品牌薇诺娜增速放缓、大促费效比失衡及并购标的拖累，扣非净利润连续两年深度下跌（-35.20%、-61.12%），并于 2024 年探底至 2.40 亿元。自 2025 年起至 2026 年前三季度，两家企业的盈利剪刀差发生戏剧性反转。贝泰妮扣非净利润实现 2025 年 +48.22%（3.56 亿元）、2026Q1-Q3 +21.38%（2.23 亿元）的连续强劲修复，珀莱雅却掉头步入负增长通道（2025 年 -3.23%，2026Q1-Q3 -25.50%，Q3 单季 -65.20%）。横向深度拆解揭示了三大核心分化：",
            ],
            "table_key": "t53",
            "chart": {"file": "fig3_peer_comparison", "caption": "图 3  珀莱雅 vs 贝泰妮横向四维对标：研发护城河、周期错位与商誉迁移（2023–2026Q3）"},
            "sub_paragraphs": [
                "第一，毛利率与研发壁垒的实质性差距。尽管两者销售费用率在 2026Q1-Q3 均处于 53% 以上的高位（珀莱雅 53.23% vs 贝泰妮 53.56%），但资源配置的战略重心迥异。贝泰妮依托“敏感肌功效皮肤学”专业定位，综合毛利率常年稳定在 73.7%–75.2% 区间（2026Q1-Q3 达 74.65%，高于珀莱雅 1.67 个百分点），且 Q3 单季综合毛利率仍维持在 74.71%。更为突出的是，贝泰妮研发费用率长期保持在 4.42%–5.88% 区间（2026Q1-Q3 为 5.17%，投入 1.88 亿元），约为珀莱雅（1.97%，投入 1.50 亿元）的 2.6 倍。珀莱雅长期“重营销、轻研发”（2025 年研发费用 2.17 亿元仅为销售费用 52.59 亿元的 4.13%），在行业步入强监管与成分技术迭代期后，直接导致其无法依靠独家专利原料构筑产品接续壁垒，只能被迫陷入更高烈度的流量价格战。",
                "第二，运营效率与控费调整的周期错位。贝泰妮在经历了 2023–2024 年的失速阵痛后，于 2025 年果断推行“放弃低效规模、聚焦盈利质量”的收缩策略，主动削减低 ROI 促销活动（2025 年营收虽收缩 6.57%，但销售费用同比下降 6.63%，经营现金流同比大增 40.54% 至 9.77 亿元），成功止住扣非利润失血。反观珀莱雅，在 2025 年主品牌出现 -10.39% 下滑后，管理层在多品牌扩张（OR、原色波塔及惊时）与交接班转型期的规模诉求下，选择在 2026 年继续加码形象宣传与投流预算（2026Q1-Q3 销售费用同比大增 15.16% 至 40.60 亿元），结果陷入“越投流、主业扣非利润越薄”的恶性内卷。",
                "第三，外延并购与商誉减值的前车之鉴（悦江投资 vs 花知晓）。这是两者横向对比中最具警示意义的资本教训。贝泰妮早在 2023 年 9 月即以 5.355 亿元现金收购悦江投资（旗下拥有 Za 姬芮、泊美）51% 股权试图破局大众彩妆与多品牌矩阵，当时形成 4.14 亿元商誉。随后悦江投资在大众美妆白热化竞争中连续失利，2023–2025 年三年累计净利润未达 2.35 亿元对赌承诺，既拖累贝泰妮 2024–2025 年连续计提商誉减值（商誉余额由 4.14 亿元缩减至 2.93 亿元），也触发了原股东的定向股权补偿机制（贝泰妮持股比例被动升至 69.28%）。反观珀莱雅，2025 年前账面商誉一直为 0，但在 2026 年以合计 7.79 亿元分步控股花知晓后，资产负债表骤然新增高达 9.24 亿元的商誉，占其 2026Q3 末归母净资产（63.41 亿元）的 14.57%，绝对规模与净资产占比均远超贝泰妮当年的悦江投资（7.06%）。一旦花知晓这一强营销属性彩妆品牌在并表后出现高基数回落或潮流退坡，珀莱雅将面临远超贝泰妮当年的商誉减值“爆雷”风险。",
            ],
        },
        {
            "level": 2,
            "heading": "5.3 杜邦分析：盈利驱动拆解与 EVA 真实价值穿透",
            "paragraphs": [
                "传统财务报告往往以单一的披露净资产收益率（ROE）评价股东回报，极易被一次性并购重估损益与财务杠杆放大所迷惑。依据杜邦分析恒等式（参考 FinanceToolkit 标准三因子模型）：",
                "ROE = 销售净利率（NPM） × 总资产周转率（ATO） × 权益乘数（EM）",
                "本节创新性地构建了“GAAP 名义口径”与“扣非真实经营口径”双轨杜邦拆解体系（其中资产与净资产取期初期末均值，2026Q1-Q3 按线性年化系数 4/3 折算以便跨期可比。线性年化未计入 Q4“双十一”大促季节性放量效应，全年实际周转率与 ROE 取决于四季度旺季兑现程度），以精准剥离花知晓并购带来的会计泡沫。测算结果见表 5-4 与图 4。",
            ],
            "table_key": "t54",
            "chart": {"file": "fig4_dupont_decomposition", "caption": "图 4  双轨杜邦三因子拆解：珀莱雅会计泡沫（10.96pct）与真实经营驱动退化"},
            "sub_paragraphs": [
                "双轨杜邦三因子拆解显示，2023–2024 年是珀莱雅股东回报的巅峰期，其扣非年化 ROE 分别高达 29.82% 与 31.21%，且驱动引擎由“高销售净利率（13.18%→14.12%）+ 极高资产周转率（1.36x→1.45x）”构成，形成了轻资产高周转消费品护城河，与贝泰妮同期的低周转高研发模式形成降维打击。进入 2026Q1-Q3，在名义账面口径下，珀莱雅名义年化 ROE 仍录得 26.83%（甚至较 2025 年的 25.96% 微升 0.87 个百分点）。切换至扣非真实经营口径后，其真实年化 ROE 已断崖式跌落至 15.87%，名义与真实 ROE 之间出现了高达 +10.96 个百分点的巨大“会计泡沫差值”。",
                "进一步将扣非真实 ROE 的下滑拆解至三大驱动因子，可以清晰追踪盈利能力的全面退化：\n"
                "（1）销售净利率（NPM）遭遇腰斩式挤压：由 2024 年的 14.12%、2025 年的 13.90% 暴跌至 2026Q1-Q3 的 9.74%（2026Q3 单季更是跌至仅 3.51%），单因子即下拉年化 ROE 超过 7.5 个百分点，直接反映了销售费用率飙升与低毛利彩妆并表对销售净利率的严重侵蚀。\n"
                "（2）总资产周转率（ATO）从 1.45x 跌破 1.1x：由 2024 年 1.45 次、2025 年 1.32 次下滑至 2026Q1-Q3 年化 1.09 次。周转效率恶化的根源在于资产端发生了“重资产化与虚胖化”。2026Q3 末总资产激增 20.84% 至 102.56 亿元，其中 9.24 亿元为并购沉淀的无产出实体商誉，另有 11.75 亿元为多品牌备货积压的存货（较年初 6.35 亿元暴增 85.04%）。\n"
                "（3）权益乘数（EM）被动抬升：由 2025 年的 1.39 倍回升至 2026Q1-Q3 的 1.50 倍（总负债由 22.21 亿元增至 33.87 亿元，资产负债率由 26.17% 升至 33.02%），是杜邦三因子中唯一向上拉动 ROE 的杠杆因素。这意味着珀莱雅的盈利驱动模式已经发生本质退化，从“内生净利率与周转率双轮驱动”，退化为“主业净利率与周转率双杀、仅靠并购债务杠杆与一次性公允价值重估勉力支撑账面 ROE”。",
                "【引入 EVA（经济增加值）穿透并购表象下的真实价值创造】\n"
                "由于杜邦体系本身属于会计利润衍生指标，并未扣除股东投入资本的机会成本，本节进一步引入经济增加值模型 EVA = NOPAT - TC × WACC。其中，税后净营业利润（NOPAT）以剔除 4.45 亿元分步并购重估泡沫后的扣非归母净利润为基础并加回税后利息费用。投入资本总额（TC）定义为期末归母净资产与有息负债（短期借款 0.80 亿元 + 可转债 8.32 亿元）之和，全额纳入并购支付及新增 9.24 亿元商誉形成的资本占用。参考 10 年期国债无风险利率（2.10%）与市场风险溢价（6.00%），结合美妆行业 Beta（1.05）和高新技术企业 15% 所得税率，测算得出珀莱雅动态 WACC 维持在 7.43%–7.69% 区间。完整穿透结果如表 5-5 与图 5 所示。",
            ],
            "table_key_2": "t55",
            "chart_2": {"file": "fig5_eva_goodwill", "caption": "图 5  EVA 真实价值穿透：并购虚增账面利润与边际资本价值毁损（2022–2026Q3）"},
            "tail_paragraphs": [
                "EVA 穿透测算显示，2022–2024 年期间，珀莱雅真实 EVA 由 4.87 亿元高速攀升至 10.73 亿元，年化投入资本回报率（ROIC）高达 25.05%，超过 WACC 资本成本率达 +17.36 个百分点，是名副其实的“超级价值创造者”。2025 年真实 EVA 已回落至 9.60 亿元（-10.53%）。及至 2026Q1-Q3，尽管账面归母净利润高达 12.56 亿元（同比 +22.46%），但随着投入资本 TC 因并购扩张膨胀至 72.53 亿元（前三季度折算资本成本达 4.18 亿元），剔除并购会计重估收益后的真实 EVA 骤降至 3.49 亿元（年化约 4.65 亿元，较 2024 年峰值腰斩逾半）。更为严峻的是，在花知晓完全并表后的 2026Q3 单季，珀莱雅年化真实 ROIC 仅剩 4.78%，跌破 7.69% 的 WACC 资本成本率（EVA 利差 -2.91 个百分点），单季 EVA 直接转负至 -0.53 亿元。从资本经济学本质上看，在主业护城河松动之际，通过高溢价并购换取规模无法弥补内生现金流与盈利的坍塌，巨额资本占用反而拉低 ROIC，在边际上造成了真实的股东价值毁损（Value Destruction）。",
            ],
        },
        {
            "level": 2,
            "heading": "5.4 现金流与盈利质量评价（结合平衡计分卡 BSC 思想）",
            "paragraphs": [
                "盈利质量的高低最终取决于会计利润能否顺利转化为真金白银的自由现金流。如表 5-6 与图 6(a) 所示，2022–2023 年以及 2025 年，珀莱雅的经营活动现金流量净额（OCF）均稳健高于归母净利润（2023 年净现比 1.23，2025 年通过优化账期与库存回笼实现 OCF 21.93 亿元、净现比高达 1.46）。进入 2026 年后，伴随流量买量成本飙升与多品牌并表备货，公司现金流与盈利质量出现了系统性的同步恶化。",
                "2026Q1-Q3 珀莱雅实现收现比 1.11（销售商品收到的现金 84.59 亿元），表面回款尚属正常，但经营活动现金流量净额（OCF）却暴跌 61.06% 至仅 4.67 亿元。若结合半年报数据（2026H1 OCF 为 9.51 亿元）进行单季倒推，2026Q3 单季 OCF 竟然净流出高达 4.84 亿元。对应 2026Q1-Q3 GAAP 净现比骤降至 0.37（扣非净现比 0.63），2026Q3 单季净现比跌至 -5.50。同时，公司期末存货由 2025 年末的 6.35 亿元激增 85.04% 至 2026Q3 末的 11.75 亿元，存货周转天数（DIO）从 2025 年的 83.5 天大幅拉长至 119.9 天。应计利润占总资产比由 2025 年的 -8.68%（现金优于利润）恶化为 2026Q1-Q3 的 +8.42%（利润严重虚胖于现金流）。究其原因，一方面是主品牌新品迭代与花知晓、OR 及原色波塔等成长品牌前置支付了巨额形象宣传推广现金（现金流出的刚性前置），另一方面是并购并表及双十一备货导致大量资金被沉淀在周转趋缓的存货库房之中，形成“账面有利润，库房堆存货，账上少现金”的典型低质量增长特征。",
            ],
            "table_key": "t56",
            "chart": {"file": "fig6_cashflow_bsc", "caption": "图 6  现金流与盈利质量评价及平衡计分卡（BSC）四维健康度诊断"},
            "sub_paragraphs": [
                "【基于平衡计分卡（BSC）四维联动的高质量增长困境全面剖析】\n"
                "单纯的财务指标下滑往往是前端运营、客户生态与创新能力失灵的滞后结果。为彻底回答“国货美妆能否摆脱流量内卷、实现高质量增长”这一核心命题，本节遵循平衡计分卡（BSC）因果链条，构建四维综合评价矩阵，包括财务维度（30%）和客户与流量维度（25%），以及内部运营与产品周期维度（25%）和学习成长与并购治理维度（20%）（见表 5-7 与图 6(b)）。",
                "1. 财务维度（评分由 92.0 降至 56.0）：并购重估粉饰了 GAAP 净利，但 EVA 转负（Q3 单季 -0.53 亿元）、扣非 ROE 腰斩与 OCF 暴跌 61.06% 暴露了财务造血功能的实质性衰退。\n"
                "2. 客户与流量维度（评分由 88.0 降至 48.0）：流量买量呈现显著的边际报酬递减，销售费用率攀升至 53.23% 以上而流量净回报率跌至 6.56%（单季），反映品牌自然心智与客户复购黏性正在被平台高昂的竞价流采买成本吞噬。\n"
                "3. 内部运营与产品生命周期维度（评分由 86.0 降至 52.0）：主品牌大单品生命周期步入成熟后期，第二曲线彩棠失速（-21.93%），叠加多品牌 SKU 膨胀导致存货周转天数飙升至 119.9 天，供应链敏捷度与库存健康度面临严峻考验。\n"
                "4. 学习成长与并购治理维度（评分由 76.0 降至 54.0）：核心症结在于“2% 研发费用率天花板”与“9.24 亿元高商誉悬顶”的双重夹击。在二代接班与平台化多品牌转型期，研发投入长期显著落后于贝泰妮（5.17%），导致公司缺乏底层原料与技术代差，不得不以高溢价外延并购（花知晓）购买时间窗口，进而承担类似贝泰妮收购悦江投资的整合与商誉减值风险。",
            ],
            "table_key_2": "t57",
            "tail_paragraphs": [
                "综合上述四维诊断（珀莱雅 BSC 总分由 2023 年巅峰期的 86.30 分滑落至 2026Q1-Q3 的 52.60 分，显著低于已完成控费与研发聚焦修复的贝泰妮 71.85 分），依赖“高额投流买规模 + 溢价并购粉饰报表”的双轮扩张路径已然走到尽头。珀莱雅若要避免重蹈流量黑洞与商誉减值覆辙、真正迈向高质量发展，必须借鉴战略重构经验，完成三项根本性转向：其一，建立以“扣非 ROIC 与 EVA”为核心的内部预算考核机制，砍掉流量净回报率低于资本成本线的无效投流。其二，重塑研发战略优先级，将研发费用率由 2% 阶梯式提升至 4%–5% 区间，依靠自有核心活性成分与配方专利完成大单品的代际跃迁，打破同质化内卷。其三，加强对花知晓等并购标的的供应链、渠道复用与独立造血整合，严格控制存货周转与现金流净现比回升至 1.0 以上安全阈值，用内生性经济利润逐步消化 9.24 亿元高悬商誉。",
            ],
        },
    ]


# ==============================================================================
# 1. Word 生成引擎：python-docx 模板构建 + docxtpl 填充
# ==============================================================================

def _set_cell_shading(cell, fill_hex: str) -> None:
    tcPr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), fill_hex)
    tcPr.append(shd)


def _set_cell_borders(cell, top=None, bottom=None) -> None:
    tcPr = cell._tc.get_or_add_tcPr()
    tcBorders = OxmlElement("w:tcBorders")
    for side in ("top", "bottom", "left", "right"):
        el = OxmlElement(f"w:{side}")
        spec = top if side == "top" else (bottom if side == "bottom" else None)
        if spec:
            el.set(qn("w:val"), "single")
            el.set(qn("w:sz"), str(spec["sz"]))
            el.set(qn("w:space"), "0")
            el.set(qn("w:color"), spec["color"])
        else:
            el.set(qn("w:val"), "none")
        tcBorders.append(el)
    tcPr.append(tcBorders)


def _add_styled_paragraph(doc: Document, text: str, style_name: str | None = None,
                          bold: bool = False, size_pt: float = 10.5,
                          color_rgb: tuple[int, int, int] = (34, 47, 62),
                          space_before: float = 4, space_after: float = 6,
                          align=WD_ALIGN_PARAGRAPH.JUSTIFY, first_line_indent_cm: float = 0.0):
    p = doc.add_paragraph(style=style_name)
    p.alignment = align
    pf = p.paragraph_format
    pf.space_before = Pt(space_before)
    pf.space_after = Pt(space_after)
    pf.line_spacing = 1.38
    if first_line_indent_cm:
        pf.first_line_indent = Cm(first_line_indent_cm)
    run = p.add_run(text)
    run.bold = bold
    run.font.name = "Microsoft YaHei"
    run._element.rPr.rFonts.set(qn("w:eastAsia"), "Microsoft YaHei")
    run.font.size = Pt(size_pt)
    run.font.color.rgb = RGBColor(*color_rgb)
    return p


def _add_three_line_table_jinja(doc: Document, table_var: str, col_count: int) -> None:
    """在 python-docx 模板中构建学术三线表并注入 docxtpl Jinja2 行循环占位符。"""
    p_title = _add_styled_paragraph(
        doc, f"{{{{ {table_var}.title }}}}", bold=True, size_pt=10.5,
        color_rgb=(27, 54, 93), space_before=8, space_after=4, align=WD_ALIGN_PARAGRAPH.CENTER
    )
    p_title.paragraph_format.keep_with_next = True

    table = doc.add_table(rows=3, cols=col_count)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = True

    # Row 0: 表头循环占位（docxtpl 标准语法：{% for h in t.headers %} 写在列或用固定列索引）
    for j in range(col_count):
        cell = table.cell(0, j)
        _set_cell_shading(cell, "1B365D")
        _set_cell_borders(cell, top={"sz": 12, "color": "1B365D"}, bottom={"sz": 8, "color": "1B365D"})
        cp = cell.paragraphs[0]
        cp.alignment = WD_ALIGN_PARAGRAPH.LEFT if j == 0 else WD_ALIGN_PARAGRAPH.CENTER
        cp.paragraph_format.space_before = Pt(3)
        cp.paragraph_format.space_after = Pt(3)
        r = cp.add_run(f"{{{{ {table_var}.headers[{j}] }}}}")
        r.bold = True
        r.font.name = "Microsoft YaHei"
        r._element.rPr.rFonts.set(qn("w:eastAsia"), "Microsoft YaHei")
        r.font.size = Pt(8.5)
        r.font.color.rgb = RGBColor(255, 255, 255)

    # Row 1: Jinja2 {%tr for row in ... %} 循环起始行 + 数据单元格
    for j in range(col_count):
        cell = table.cell(1, j)
        _set_cell_borders(cell, top=None, bottom={"sz": 4, "color": "D5DCE4"})
        cp = cell.paragraphs[0]
        cp.alignment = WD_ALIGN_PARAGRAPH.LEFT if j == 0 else WD_ALIGN_PARAGRAPH.CENTER
        cp.paragraph_format.space_before = Pt(2.5)
        cp.paragraph_format.space_after = Pt(2.5)
        prefix = f"{{%tr for row in {table_var}.rows %}}" if j == 0 else ""
        r = cp.add_run(f"{prefix}{{{{ row[{j}] }}}}")
        r.font.name = "Microsoft YaHei"
        r._element.rPr.rFonts.set(qn("w:eastAsia"), "Microsoft YaHei")
        r.font.size = Pt(8.2)
        r.font.color.rgb = RGBColor(33, 47, 61)

    # Row 2: Jinja2 {%tr endfor %} 结束行（同时作为三线表底线，实际渲染后闭合）
    for j in range(col_count):
        cell = table.cell(2, j)
        _set_cell_borders(cell, top={"sz": 10, "color": "1B365D"}, bottom=None)
        cp = cell.paragraphs[0]
        cp.paragraph_format.space_before = Pt(0)
        cp.paragraph_format.space_after = Pt(0)
        if j == 0:
            r = cp.add_run("{%tr endfor %}")
            r.font.size = Pt(1)


def build_docx_template(ctx: dict, sections: list[dict]) -> None:
    """利用 python-docx 生成标准化的 docxtpl 模板文件 templates/report_template.docx。"""
    TEMPLATE_DIR.mkdir(parents=True, exist_ok=True)
    doc = Document()

    # 页面设置（A4，边距适中）
    for section in doc.sections:
        section.page_width = Cm(21.0)
        section.page_height = Cm(29.7)
        section.top_margin = Cm(2.2)
        section.bottom_margin = Cm(2.2)
        section.left_margin = Cm(2.0)
        section.right_margin = Cm(2.0)
        # 页眉页脚
        header = section.header
        hp = header.paragraphs[0]
        hp.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        hr = hp.add_run("三、珀莱雅（603605.SH）商业深度分析  |  5 财务绩效分析")
        hr.font.name = "Microsoft YaHei"
        hr._element.rPr.rFonts.set(qn("w:eastAsia"), "Microsoft YaHei")
        hr.font.size = Pt(8.5)
        hr.font.color.rgb = RGBColor(127, 140, 141)

    # 报告主标题栏
    _add_styled_paragraph(
        doc, "{{ report_super_title }}", bold=True, size_pt=11,
        color_rgb=(59, 125, 221), space_before=0, space_after=2, align=WD_ALIGN_PARAGRAPH.LEFT
    )
    _add_styled_paragraph(
        doc, "{{ report_title }}", bold=True, size_pt=18,
        color_rgb=(27, 54, 93), space_before=0, space_after=6, align=WD_ALIGN_PARAGRAPH.LEFT
    )
    _add_styled_paragraph(
        doc, "{{ report_meta }}", bold=False, size_pt=9.5,
        color_rgb=(93, 109, 126), space_before=0, space_after=10, align=WD_ALIGN_PARAGRAPH.LEFT
    )

    # 核心命题摘要框（1x1 高亮表格）
    box = doc.add_table(rows=1, cols=1)
    box.alignment = WD_TABLE_ALIGNMENT.CENTER
    cell = box.cell(0, 0)
    _set_cell_shading(cell, "F2F6FA")
    _set_cell_borders(cell, top={"sz": 12, "color": "1B365D"}, bottom={"sz": 8, "color": "1B365D"})
    bp = cell.paragraphs[0]
    bp.paragraph_format.space_before = Pt(6)
    bp.paragraph_format.space_after = Pt(6)
    bp.paragraph_format.line_spacing = 1.35
    br = bp.add_run("{{ executive_summary }}")
    br.font.name = "Microsoft YaHei"
    br._element.rPr.rFonts.set(qn("w:eastAsia"), "Microsoft YaHei")
    br.font.size = Pt(9.8)
    br.font.color.rgb = RGBColor(27, 54, 93)

    # 动态渲染各章节 Jinja2 占位符、三线表与 InlineImage 图表占位符
    chart_idx = 0
    for s_idx, sec in enumerate(sections):
        if sec["level"] == 1:
            _add_styled_paragraph(
                doc, f"{{{{ sec_{s_idx}_heading }}}}", bold=True, size_pt=15,
                color_rgb=(27, 54, 93), space_before=14, space_after=6, align=WD_ALIGN_PARAGRAPH.LEFT
            )
        else:
            _add_styled_paragraph(
                doc, f"{{{{ sec_{s_idx}_heading }}}}", bold=True, size_pt=13,
                color_rgb=(27, 54, 93), space_before=12, space_after=5, align=WD_ALIGN_PARAGRAPH.LEFT
            )

        for p_idx in range(len(sec.get("paragraphs", []))):
            _add_styled_paragraph(
                doc, f"{{{{ sec_{s_idx}_p_{p_idx} }}}}", size_pt=10.5,
                first_line_indent_cm=0.74
            )

        if "table_key" in sec:
            tk = sec["table_key"]
            _add_three_line_table_jinja(doc, tk, len(ctx[tk]["headers"]))

        if "chart" in sec:
            chart_idx += 1
            p_img = _add_styled_paragraph(
                doc, f"{{{{ img_{sec['chart']['file']} }}}}",
                space_before=8, space_after=2, align=WD_ALIGN_PARAGRAPH.CENTER
            )
            p_img.paragraph_format.keep_with_next = True
            _add_styled_paragraph(
                doc, f"{{{{ cap_{sec['chart']['file']} }}}}", bold=True, size_pt=9.2,
                color_rgb=(93, 109, 126), space_before=2, space_after=8, align=WD_ALIGN_PARAGRAPH.CENTER
            )

        for sp_idx in range(len(sec.get("sub_paragraphs", []))):
            _add_styled_paragraph(
                doc, f"{{{{ sec_{s_idx}_sp_{sp_idx} }}}}", size_pt=10.5,
                first_line_indent_cm=0.74
            )

        if "table_key_2" in sec:
            tk2 = sec["table_key_2"]
            _add_three_line_table_jinja(doc, tk2, len(ctx[tk2]["headers"]))

        if "chart_2" in sec:
            p_img2 = _add_styled_paragraph(
                doc, f"{{{{ img_{sec['chart_2']['file']} }}}}",
                space_before=8, space_after=2, align=WD_ALIGN_PARAGRAPH.CENTER
            )
            p_img2.paragraph_format.keep_with_next = True
            _add_styled_paragraph(
                doc, f"{{{{ cap_{sec['chart_2']['file']} }}}}", bold=True, size_pt=9.2,
                color_rgb=(93, 109, 126), space_before=2, space_after=8, align=WD_ALIGN_PARAGRAPH.CENTER
            )

        for tp_idx in range(len(sec.get("tail_paragraphs", []))):
            _add_styled_paragraph(
                doc, f"{{{{ sec_{s_idx}_tp_{tp_idx} }}}}", size_pt=10.5,
                first_line_indent_cm=0.74
            )

    doc.save(str(DOCX_TEMPLATE_PATH))
    print(f"[WORD] Template created: {DOCX_TEMPLATE_PATH}")


def render_docx_with_docxtpl(ctx: dict, sections: list[dict]) -> None:
    """加载模板并使用 docxtpl 填充数据、文本与静态 PNG 图表，输出最终 Word 报告。"""
    tpl = DocxTemplate(str(DOCX_TEMPLATE_PATH))
    render_ctx = {
        **ctx,
        "report_super_title": "商业深度分析专题报告  ·  国货美妆高质量发展诊断",
        "report_title": "三、珀莱雅（603605.SH）：5 财务绩效分析",
        "report_meta": "研究标的：珀莱雅（603605.SH）｜ 横向对标：贝泰妮（300957.SZ）｜ 数据引擎：CNEquity + DuckDB + pandas + Plotly",
        "executive_summary": f"【核心命题与研究结论摘要】\n{EXECUTIVE_SUMMARY}",
    }

    for s_idx, sec in enumerate(sections):
        render_ctx[f"sec_{s_idx}_heading"] = sec["heading"]
        for p_idx, p_txt in enumerate(sec.get("paragraphs", [])):
            render_ctx[f"sec_{s_idx}_p_{p_idx}"] = p_txt
        for sp_idx, sp_txt in enumerate(sec.get("sub_paragraphs", [])):
            render_ctx[f"sec_{s_idx}_sp_{sp_idx}"] = sp_txt
        for tp_idx, tp_txt in enumerate(sec.get("tail_paragraphs", [])):
            render_ctx[f"sec_{s_idx}_tp_{tp_idx}"] = tp_txt
        for ck in ("chart", "chart_2"):
            if ck in sec:
                fname = sec[ck]["file"]
                png_path = CHARTS_DIR / f"{fname}.png"
                render_ctx[f"img_{fname}"] = InlineImage(tpl, str(png_path), width=Cm(16.2))
                render_ctx[f"cap_{fname}"] = sec[ck]["caption"]

    tpl.render(render_ctx)
    tpl.save(str(DOCX_OUT_PATH))
    print(f"[WORD] Final DOCX generated: {DOCX_OUT_PATH}")


# ==============================================================================
# 2. PDF 生成引擎：内联 SVG 矢量图 HTML + WeasyPrint 渲染
# ==============================================================================

def _load_inline_svg(chart_name: str) -> str:
    svg_path = CHARTS_DIR / f"{chart_name}.svg"
    svg_text = svg_path.read_text(encoding="utf-8")
    # 清理 XML 声明头并确保 SVG 自适应页面宽度
    svg_text = re.sub(r"<\?xml.*?\?>", "", svg_text, flags=re.DOTALL).strip()
    svg_text = re.sub(r'(<svg\b[^>]*?)\s+width="[^"]*"', r'\1 width="100%"', svg_text, count=1)
    svg_text = re.sub(r'(<svg\b[^>]*?)\s+height="[^"]*"', r'\1 height="auto"', svg_text, count=1)
    return svg_text


def _render_html_table(table_dict: dict) -> str:
    headers_html = "".join(f"<th>{h}</th>" for h in table_dict["headers"])
    rows_html = []
    for row in table_dict["rows"]:
        tds = []
        for idx, cell in enumerate(row):
            cls = "left" if idx == 0 else "num"
            # 对关键负向指标或警示行做柔和视觉标红
            if isinstance(cell, str) and (cell.startswith("-") and "%" in cell or cell in ("-4.84", "-0.53", "-5.50", "9.24")):
                cls += " neg"
            tds.append(f'<td class="{cls}">{cell}</td>')
        rows_html.append(f"<tr>{''.join(tds)}</tr>")
    return f"""
    <div class="table-block">
      <div class="table-title">{table_dict['title']}</div>
      <table class="academic-table">
        <thead><tr>{headers_html}</tr></thead>
        <tbody>{''.join(rows_html)}</tbody>
      </table>
    </div>
    """


def build_html_and_pdf(ctx: dict, sections: list[dict]) -> None:
    BUILD_DIR.mkdir(parents=True, exist_ok=True)

    body_parts = []
    for sec in sections:
        tag = "h1" if sec["level"] == 1 else "h2"
        body_parts.append(f"<{tag}>{sec['heading']}</{tag}>")

        for p in sec.get("paragraphs", []):
            body_parts.append(f"<p>{p.replace(chr(10), '<br/>')}</p>")

        if "table_key" in sec:
            body_parts.append(_render_html_table(ctx[sec["table_key"]]))

        if "chart" in sec:
            svg_code = _load_inline_svg(sec["chart"]["file"])
            body_parts.append(f"""
            <div class="figure-block">
              <div class="svg-wrapper">{svg_code}</div>
              <div class="figure-caption">{sec['chart']['caption']}</div>
            </div>
            """)

        for sp in sec.get("sub_paragraphs", []):
            formatted = sp.replace("\n", "<br/>")
            if formatted.startswith("【"):
                formatted = re.sub(r"^(【[^】]+】)", r"<strong>\1</strong>", formatted)
            body_parts.append(f"<p>{formatted}</p>")

        if "table_key_2" in sec:
            body_parts.append(_render_html_table(ctx[sec["table_key_2"]]))

        if "chart_2" in sec:
            svg_code2 = _load_inline_svg(sec["chart_2"]["file"])
            body_parts.append(f"""
            <div class="figure-block">
              <div class="svg-wrapper">{svg_code2}</div>
              <div class="figure-caption">{sec['chart_2']['caption']}</div>
            </div>
            """)

        for tp in sec.get("tail_paragraphs", []):
            body_parts.append(f"<p>{tp.replace(chr(10), '<br/>')}</p>")

    html_content = f"""<!DOCTYPE html>
<html lang="zh-CN">
<head>
<meta charset="utf-8"/>
<title>三、珀莱雅：5 财务绩效分析</title>
<style>
  @page {{
    size: A4;
    margin: 20mm 16mm 20mm 16mm;
    @top-left {{
      content: "三、珀莱雅（603605.SH）商业深度分析专题";
      font-family: "Microsoft YaHei", "SimHei", sans-serif;
      font-size: 8pt;
      color: #6C7A89;
      border-bottom: 0.5pt solid #D5DCE4;
      padding-bottom: 3pt;
    }}
    @top-right {{
      content: "5 财务绩效分析";
      font-family: "Microsoft YaHei", "SimHei", sans-serif;
      font-size: 8pt;
      color: #1B365D;
      border-bottom: 0.5pt solid #D5DCE4;
      padding-bottom: 3pt;
    }}
    @bottom-center {{
      content: "第 " counter(page) " 页 / 共 " counter(pages) " 页";
      font-family: "Microsoft YaHei", "SimHei", sans-serif;
      font-size: 8pt;
      color: #6C7A89;
    }}
  }}
  @page :first {{
    @top-left {{ content: none; }}
    @top-right {{ content: none; }}
  }}
  body {{
    font-family: "Microsoft YaHei", "SimSun", sans-serif;
    font-size: 9.8pt;
    line-height: 1.62;
    color: #212F3D;
    margin: 0;
  }}
  .hero-header {{
    border-left: 4pt solid #1B365D;
    padding: 4pt 0 4pt 12pt;
    margin-bottom: 12pt;
  }}
  .hero-kicker {{
    font-size: 9pt;
    font-weight: bold;
    color: #3B7DDD;
    letter-spacing: 0.5pt;
  }}
  .hero-title {{
    font-size: 17.5pt;
    font-weight: bold;
    color: #1B365D;
    margin: 3pt 0 4pt 0;
  }}
  .hero-meta {{
    font-size: 8.5pt;
    color: #5D6D7E;
  }}
  .summary-box {{
    background-color: #F2F6FA;
    border-top: 1.5pt solid #1B365D;
    border-bottom: 1pt solid #1B365D;
    padding: 9pt 11pt;
    margin-bottom: 14pt;
    font-size: 9pt;
    line-height: 1.55;
    color: #1B365D;
  }}
  h1 {{
    font-size: 14pt;
    color: #1B365D;
    border-bottom: 1.5pt solid #1B365D;
    padding-bottom: 4pt;
    margin: 14pt 0 8pt 0;
    page-break-after: avoid;
  }}
  h2 {{
    font-size: 11.8pt;
    color: #1B365D;
    background: #F4F7FA;
    border-left: 3pt solid #3B7DDD;
    padding: 4pt 8pt;
    margin: 14pt 0 8pt 0;
    page-break-after: avoid;
  }}
  p {{
    text-align: justify;
    text-indent: 2em;
    margin: 0 0 7pt 0;
  }}
  .table-block {{
    margin: 10pt 0 12pt 0;
    page-break-inside: avoid;
  }}
  .table-title {{
    font-size: 9.2pt;
    font-weight: bold;
    color: #1B365D;
    text-align: center;
    margin-bottom: 4pt;
    page-break-after: avoid;
  }}
  table.academic-table {{
    width: 100%;
    border-collapse: collapse;
    font-size: 7.8pt;
    border-top: 1.5pt solid #1B365D;
    border-bottom: 1.5pt solid #1B365D;
  }}
  table.academic-table thead th {{
    background-color: #1B365D;
    color: #FFFFFF;
    font-weight: bold;
    padding: 4.5pt 4pt;
    text-align: center;
    border-bottom: 1pt solid #1B365D;
  }}
  table.academic-table thead th:first-child {{
    text-align: left;
  }}
  table.academic-table tbody td {{
    padding: 3.8pt 4pt;
    border-bottom: 0.4pt solid #DCE3EA;
  }}
  table.academic-table tbody tr:nth-child(even) {{
    background-color: #F9FBFD;
  }}
  table.academic-table td.left {{
    text-align: left;
    font-weight: 500;
  }}
  table.academic-table td.num {{
    text-align: center;
    font-variant-numeric: tabular-nums;
  }}
  table.academic-table td.neg {{
    color: #C0392B;
    font-weight: bold;
  }}
  .figure-block {{
    margin: 10pt 0 12pt 0;
    padding: 6pt;
    border: 0.6pt solid #DCE3EA;
    background: #FFFFFF;
    page-break-inside: avoid;
  }}
  .svg-wrapper svg {{
    display: block;
    width: 100%;
    height: auto;
  }}
  .figure-caption {{
    font-size: 8.6pt;
    font-weight: bold;
    color: #5D6D7E;
    text-align: center;
    margin-top: 4pt;
  }}
</style>
</head>
<body>
  <div class="hero-header">
    <div class="hero-kicker">商业深度分析专题报告  ·  国货美妆高质量发展诊断</div>
    <div class="hero-title">三、珀莱雅（603605.SH）：5 财务绩效分析</div>
    <div class="hero-meta">研究标的：珀莱雅（603605.SH）｜ 横向对标：贝泰妮（300957.SZ）｜ 底座管线：CNEquity + DuckDB + pandas + Plotly SVG + WeasyPrint</div>
  </div>
  <div class="summary-box">
    <strong>【核心命题与研究结论摘要】</strong><br/>
    {EXECUTIVE_SUMMARY}
  </div>
  {''.join(body_parts)}
</body>
</html>
"""
    HTML_OUT_PATH.write_text(html_content, encoding="utf-8")
    print(f"[HTML] Built HTML with inline SVGs: {HTML_OUT_PATH}")

    weasyprint.HTML(filename=str(HTML_OUT_PATH), base_url=str(BASE_DIR)).write_pdf(str(PDF_OUT_PATH))
    print(f"[PDF] Final PDF rendered via WeasyPrint: {PDF_OUT_PATH}")


def main() -> None:
    ctx = load_context()
    sections = get_report_sections()
    build_docx_template(ctx, sections)
    render_docx_with_docxtpl(ctx, sections)
    build_html_and_pdf(ctx, sections)


if __name__ == "__main__":
    main()
