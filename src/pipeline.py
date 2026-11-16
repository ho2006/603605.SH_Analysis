# -*- coding: utf-8 -*-
"""
CNEquity 本地数据湖落地 + DuckDB 持久化 + pandas 财务绩效全指标计算引擎
========================================================================
核心职责：
1. 依托 CNEquity 的数据湖契约规范（DATASET_SCHEMAS & 统一证券代码 603605.SH / 300957.SZ）
   一次性将珀莱雅（603605.SH）与贝泰妮（300957.SZ）2022A–2026Q3 经审计年报与三季报核心财务数据、
   分品牌收入结构（brand_revenue）与重大并购商誉事件（ma_events）落地为本地 DuckDB 数据库及 Parquet 数据湖。
2. 使用 pandas 执行全量财务绩效计算（严格遵循：FinanceToolkit 仅作比率定义参考，不进入执行链路）：
   - 5.1 纵向趋势与费用率 / 流量 ROE 指标
   - 5.2 珀莱雅 vs 贝泰妮横向对标矩阵
   - 5.3 双轨杜邦三因子拆解（名义 GAAP ROE vs 扣非真实 ROE）与 EVA（经济增加值）真实价值穿透测算
   - 5.4 现金流与盈利质量评价指标及平衡计分卡（BSC）四维量化评分矩阵
3. 持久化计算结果至 DuckDB 衍生视图、Excel 汇总工作簿及 JSON（供 Plotly 出图及 Word/PDF 渲染管线消费）。
"""

from __future__ import annotations

import json
from datetime import datetime, timezone
from pathlib import Path

import duckdb
import pandas as pd

# 引入 CNEquity 域模型契约规范（验证本地数据湖 Schema 与代码格式一致性）
from cnequity.domain.schemas import DATASET_SCHEMAS
from cnequity.domain.symbols import format_symbol

BASE_DIR = Path(__file__).resolve().parents[1]  # 仓库根目录（src/ 的上一级），避免硬编码本机绝对路径
DATA_DIR = BASE_DIR / "data"
PARQUET_DIR = DATA_DIR / "parquet"
DUCKDB_PATH = DATA_DIR / "cnequity_lake.duckdb"
EXCEL_OUT = DATA_DIR / "financial_metrics_summary.xlsx"
JSON_OUT = DATA_DIR / "computed_metrics.json"


# ==============================================================================
# FinanceToolkit 比率公式定义参考（仅作语义与公式标准对照，不引入执行链路）
# ------------------------------------------------------------------------------
# Reference: FinanceToolkit (ratios/profitability.py, efficiency.py, solvency.py)
# - Gross Margin Ratio       = (Revenue - COGS) / Revenue
# - Net Profit Margin        = Net Income / Revenue
# - Operating Expense Ratio  = Operating Expenses / Revenue
# - Return on Equity (ROE)   = Net Income / Average Shareholders Equity
# - Return on Assets (ROA)   = Net Income / Average Total Assets
# - Asset Turnover Ratio     = Revenue / Average Total Assets
# - Equity Multiplier        = Average Total Assets / Average Shareholders Equity
# - DuPont Identity (3-step) = Net Profit Margin * Asset Turnover * Equity Multiplier
# - Operating Cash Flow Ratio= Operating Cash Flow / Current Liabilities
# - Cash Flow to Net Income  = Operating Cash Flow / Net Income (净现比)
# - Days Inventory Outstand. = (Average Inventory / COGS) * 365
# - Days Sales Outstanding   = (Average Accounts Receivable / Revenue) * 365
# - Economic Value Added(EVA)= NOPAT - (Invested Capital * WACC)
# ==============================================================================


def build_raw_datasets() -> dict[str, pd.DataFrame]:
    """构建经过权威财报核验的珀莱雅与贝泰妮 2022A–2026Q3 核心财务数据集（单位：亿元人民币）。"""

    proya_code = format_symbol("603605", "SH")
    botanee_code = format_symbol("300957", "SZ")

    # 1. 核心合并财务报表数据（单位：亿元）
    # 数据来源：珀莱雅（603605.SH）与贝泰妮（300957.SZ）2022-2025年度报告、2026年半年度报告及2026年第三季度报告
    financials_rows = [
        # ---------------- 珀莱雅 603605.SH ----------------
        {
            "symbol": proya_code, "company": "珀莱雅", "period": "2022A", "period_order": 1, "annualize_factor": 1.0, "days_in_period": 365,
            "revenue": 63.85, "rev_yoy": 37.82, "cogs": 19.35, "gross_margin_rep": 69.70,
            "selling_exp": 27.86, "admin_exp": 3.27, "rd_exp": 1.28, "fin_exp": -0.41, "interest_exp": 0.16, "tax_exp": 2.23,
            "net_profit_parent": 8.17, "np_yoy": 41.88, "deducted_np": 7.89, "dnp_yoy": 43.80,
            "ma_reval_gain": 0.00, "ocf": 11.11, "ocf_yoy": 33.90, "cash_received": 70.88,
            "total_assets": 57.78, "total_liab": 22.41, "current_assets": 41.47, "current_liab": 14.28,
            "equity_parent": 35.24, "goodwill": 0.00, "inventory": 6.69, "accounts_receivable": 1.02, "accounts_payable": 6.10,
            "short_debt": 0.00, "long_debt": 0.00, "bonds_payable": 7.28, "roe_reported": 25.95,
            "announce_date": "2023-04-20",
        },
        {
            "symbol": proya_code, "company": "珀莱雅", "period": "2023A", "period_order": 2, "annualize_factor": 1.0, "days_in_period": 365,
            "revenue": 89.05, "rev_yoy": 39.45, "cogs": 26.77, "gross_margin_rep": 69.93,
            "selling_exp": 39.72, "admin_exp": 4.55, "rd_exp": 1.74, "fin_exp": -0.59, "interest_exp": 0.18, "tax_exp": 2.65,
            "net_profit_parent": 11.94, "np_yoy": 46.06, "deducted_np": 11.74, "dnp_yoy": 48.91,
            "ma_reval_gain": 0.00, "ocf": 14.69, "ocf_yoy": 32.19, "cash_received": 93.29,
            "total_assets": 73.23, "total_liab": 29.23, "current_assets": 53.67, "current_liab": 23.16,
            "equity_parent": 43.50, "goodwill": 0.00, "inventory": 7.97, "accounts_receivable": 3.45, "accounts_payable": 7.45,
            "short_debt": 0.00, "long_debt": 0.00, "bonds_payable": 7.53, "roe_reported": 29.94,
            "announce_date": "2024-04-18",
        },
        {
            "symbol": proya_code, "company": "珀莱雅", "period": "2024A", "period_order": 3, "annualize_factor": 1.0, "days_in_period": 365,
            "revenue": 107.78, "rev_yoy": 21.04, "cogs": 30.84, "gross_margin_rep": 71.39,
            "selling_exp": 51.61, "admin_exp": 3.66, "rd_exp": 2.10, "fin_exp": -0.36, "interest_exp": 0.31, "tax_exp": 3.04,
            "net_profit_parent": 15.52, "np_yoy": 30.00, "deducted_np": 15.22, "dnp_yoy": 29.60,
            "ma_reval_gain": 0.00, "ocf": 11.07, "ocf_yoy": -24.63, "cash_received": 109.05,
            "total_assets": 75.30, "total_liab": 20.44, "current_assets": 56.13, "current_liab": 14.76,
            "equity_parent": 54.02, "goodwill": 0.00, "inventory": 6.61, "accounts_receivable": 5.18, "accounts_payable": 7.20,
            "short_debt": 0.00, "long_debt": 0.00, "bonds_payable": 7.80, "roe_reported": 32.53,
            "announce_date": "2025-04-24",
        },
        {
            "symbol": proya_code, "company": "珀莱雅", "period": "2025A", "period_order": 4, "annualize_factor": 1.0, "days_in_period": 365,
            "revenue": 105.97, "rev_yoy": -1.68, "cogs": 28.33, "gross_margin_rep": 73.26,
            "selling_exp": 52.59, "admin_exp": 4.04, "rd_exp": 2.17, "fin_exp": -0.29, "interest_exp": 0.32, "tax_exp": 3.28,
            "net_profit_parent": 14.98, "np_yoy": -3.50, "deducted_np": 14.73, "dnp_yoy": -3.23,
            "ma_reval_gain": 0.00, "ocf": 21.93, "ocf_yoy": 98.12, "cash_received": 113.70,
            "total_assets": 84.87, "total_liab": 22.21, "current_assets": 61.36, "current_liab": 13.74,
            "equity_parent": 61.41, "goodwill": 0.00, "inventory": 6.35, "accounts_receivable": 0.70, "accounts_payable": 7.98,
            "short_debt": 0.80, "long_debt": 0.00, "bonds_payable": 8.04, "roe_reported": 25.80,
            "announce_date": "2026-04-22",
        },
        {
            "symbol": proya_code, "company": "珀莱雅", "period": "2026Q1-Q3", "period_order": 5, "annualize_factor": 4.0 / 3.0, "days_in_period": 273,
            "revenue": 76.27, "rev_yoy": 7.46, "cogs": 20.61, "gross_margin_rep": 72.98,
            "selling_exp": 40.60, "admin_exp": 3.82, "rd_exp": 1.50, "fin_exp": 0.03, "interest_exp": 0.28, "tax_exp": 1.75,
            "net_profit_parent": 12.56, "np_yoy": 22.46, "deducted_np": 7.43, "dnp_yoy": -25.50,
            "ma_reval_gain": 4.45, "ocf": 4.67, "ocf_yoy": -61.06, "cash_received": 84.59,
            "total_assets": 102.56, "total_liab": 33.87, "current_assets": 67.75, "current_liab": 24.69,
            "equity_parent": 63.41, "goodwill": 9.24, "inventory": 11.75, "accounts_receivable": 2.95, "accounts_payable": 10.50,
            "short_debt": 0.80, "long_debt": 0.00, "bonds_payable": 8.32, "roe_reported": 19.05,
            "announce_date": "2026-10-29",
        },
        {
            # 2026Q3单季由2026三季报直接披露及半年报差额推算：
            # 2026H1 OCF=9.51亿 -> Q3单季OCF=4.67-9.51=-4.84亿（上年同期Q3单季OCF=-0.95亿，净流出扩大409.47%）；2026H1 销售费用=28.56亿 -> Q3单季=12.04亿
            "symbol": proya_code, "company": "珀莱雅", "period": "2026Q3单季", "period_order": 6, "annualize_factor": 4.0, "days_in_period": 92,
            "revenue": 22.53, "rev_yoy": 29.76, "cogs": 6.81, "gross_margin_rep": 69.78,
            "selling_exp": 12.04, "admin_exp": 1.45, "rd_exp": 0.51, "fin_exp": 0.01, "interest_exp": 0.09, "tax_exp": 0.45,
            "net_profit_parent": 0.88, "np_yoy": -61.18, "deducted_np": 0.79, "dnp_yoy": -65.20,
            "ma_reval_gain": 0.00, "ocf": -4.84, "ocf_yoy": -409.47, "cash_received": 25.10,
            "total_assets": 102.56, "total_liab": 33.87, "current_assets": 67.75, "current_liab": 24.69,
            "equity_parent": 63.41, "goodwill": 9.24, "inventory": 11.75, "accounts_receivable": 2.95, "accounts_payable": 10.50,
            "short_debt": 0.80, "long_debt": 0.00, "bonds_payable": 8.32, "roe_reported": 1.39,
            "announce_date": "2026-10-29",
        },

        # ---------------- 贝泰妮 300957.SZ ----------------
        {
            "symbol": botanee_code, "company": "贝泰妮", "period": "2022A", "period_order": 1, "annualize_factor": 1.0, "days_in_period": 365,
            "revenue": 50.14, "rev_yoy": 24.65, "cogs": 12.43, "gross_margin_rep": 75.21,
            "selling_exp": 20.48, "admin_exp": 3.43, "rd_exp": 2.55, "fin_exp": -0.14, "interest_exp": 0.04, "tax_exp": 1.63,
            "net_profit_parent": 10.51, "np_yoy": 21.82, "deducted_np": 9.54, "dnp_yoy": 17.37,
            "ma_reval_gain": 0.00, "ocf": 7.69, "ocf_yoy": -33.20, "cash_received": 54.68,
            "total_assets": 67.19, "total_liab": 11.17, "current_assets": 50.10, "current_liab": 10.80,
            "equity_parent": 55.52, "goodwill": 0.00, "inventory": 6.71, "accounts_receivable": 2.70, "accounts_payable": 3.20,
            "short_debt": 0.00, "long_debt": 0.00, "bonds_payable": 0.00, "roe_reported": 20.48,
            "announce_date": "2023-03-30",
        },
        {
            "symbol": botanee_code, "company": "贝泰妮", "period": "2023A", "period_order": 2, "annualize_factor": 1.0, "days_in_period": 365,
            "revenue": 55.22, "rev_yoy": 10.14, "cogs": 14.41, "gross_margin_rep": 73.90,
            "selling_exp": 26.10, "admin_exp": 4.13, "rd_exp": 3.00, "fin_exp": -0.21, "interest_exp": 0.05, "tax_exp": 1.10,
            "net_profit_parent": 7.57, "np_yoy": -28.02, "deducted_np": 6.18, "dnp_yoy": -35.20,
            "ma_reval_gain": 0.00, "ocf": 6.10, "ocf_yoy": -20.67, "cash_received": 60.03,
            "total_assets": 75.07, "total_liab": 14.91, "current_assets": 49.80, "current_liab": 13.80,
            "equity_parent": 58.61, "goodwill": 4.14, "inventory": 9.04, "accounts_receivable": 5.62, "accounts_payable": 4.10,
            "short_debt": 1.04, "long_debt": 0.00, "bonds_payable": 0.00, "roe_reported": 13.21,
            "announce_date": "2024-04-25",
        },
        {
            "symbol": botanee_code, "company": "贝泰妮", "period": "2024A", "period_order": 3, "annualize_factor": 1.0, "days_in_period": 365,
            "revenue": 57.36, "rev_yoy": 3.87, "cogs": 15.06, "gross_margin_rep": 73.74,
            "selling_exp": 28.66, "admin_exp": 5.13, "rd_exp": 3.37, "fin_exp": -0.01, "interest_exp": 0.12, "tax_exp": 1.14,
            "net_profit_parent": 5.03, "np_yoy": -33.53, "deducted_np": 2.40, "dnp_yoy": -61.12,
            "ma_reval_gain": 0.00, "ocf": 6.95, "ocf_yoy": 13.80, "cash_received": 62.46,
            "total_assets": 83.04, "total_liab": 21.16, "current_assets": 54.60, "current_liab": 18.20,
            "equity_parent": 60.50, "goodwill": 3.23, "inventory": 6.12, "accounts_receivable": 6.91, "accounts_payable": 3.90,
            "short_debt": 3.15, "long_debt": 0.00, "bonds_payable": 0.00, "roe_reported": 8.53,
            "announce_date": "2025-04-24",
        },
        {
            "symbol": botanee_code, "company": "贝泰妮", "period": "2025A", "period_order": 4, "annualize_factor": 1.0, "days_in_period": 365,
            "revenue": 53.59, "rev_yoy": -6.57, "cogs": 13.69, "gross_margin_rep": 74.46,
            "selling_exp": 26.76, "admin_exp": 4.95, "rd_exp": 2.37, "fin_exp": 0.11, "interest_exp": 0.14, "tax_exp": 1.00,
            "net_profit_parent": 5.06, "np_yoy": 0.53, "deducted_np": 3.56, "dnp_yoy": 48.22,
            "ma_reval_gain": 0.00, "ocf": 9.77, "ocf_yoy": 40.54, "cash_received": 63.58,
            "total_assets": 83.41, "total_liab": 20.37, "current_assets": 58.68, "current_liab": 15.56,
            "equity_parent": 61.57, "goodwill": 2.93, "inventory": 5.28, "accounts_receivable": 5.59, "accounts_payable": 3.60,
            "short_debt": 4.81, "long_debt": 0.00, "bonds_payable": 0.00, "roe_reported": 8.20,
            "announce_date": "2026-04-23",
        },
        {
            "symbol": botanee_code, "company": "贝泰妮", "period": "2026Q1-Q3", "period_order": 5, "annualize_factor": 4.0 / 3.0, "days_in_period": 273,
            "revenue": 36.37, "rev_yoy": 4.99, "cogs": 9.22, "gross_margin_rep": 74.65,
            "selling_exp": 19.48, "admin_exp": 2.91, "rd_exp": 1.88, "fin_exp": 0.16, "interest_exp": 0.11, "tax_exp": 0.60,
            "net_profit_parent": 3.10, "np_yoy": 13.95, "deducted_np": 2.23, "dnp_yoy": 21.38,
            "ma_reval_gain": 0.00, "ocf": 3.17, "ocf_yoy": -28.32, "cash_received": 42.17,
            "total_assets": 79.73, "total_liab": 18.43, "current_assets": 54.55, "current_liab": 13.25,
            "equity_parent": 60.46, "goodwill": 2.93, "inventory": 7.58, "accounts_receivable": 4.42, "accounts_payable": 3.40,
            "short_debt": 3.49, "long_debt": 2.78, "bonds_payable": 0.00, "roe_reported": 5.07,
            "announce_date": "2026-10-28",
        },
    ]
    df_fin = pd.DataFrame(financials_rows)

    # 2. CNEquity 标准长表格式 financial_statement_items（遵循 DATASET_SCHEMAS['financial_statement_items'] 契约）
    long_rows = []
    now_utc = datetime.now(timezone.utc)
    metric_cols = [
        ("REVENUE", "income"), ("COGS", "income"), ("SELLING_EXP", "income"),
        ("ADMIN_EXP", "income"), ("RD_EXP", "income"), ("NET_PROFIT_PARENT", "income"),
        ("DEDUCTED_NP", "income"), ("TOTAL_ASSETS", "balance"), ("TOTAL_LIAB", "balance"),
        ("EQUITY_PARENT", "balance"), ("GOODWILL", "balance"), ("INVENTORY", "balance"),
        ("ACCOUNTS_RECEIVABLE", "balance"), ("OCF", "cashflow"), ("CASH_RECEIVED", "cashflow"),
    ]
    for _, row in df_fin.iterrows():
        for code, st_type in metric_cols:
            col_key = code.lower()
            long_rows.append({
                "symbol": row["symbol"],
                "report_period": row["period"],
                "statement_type": st_type,
                "item_code": code,
                "item_value": float(row[col_key]) * 1e8,  # 还原为人民币元，与 CNEquity PIT 存储精度对齐
                "announce_date": pd.to_datetime(row["announce_date"]).date(),
                "source": "cninfo_annual_q3_verified",
                "data_version": "2026.11.v1",
                "fetched_at": now_utc,
            })
    df_cnequity_items = pd.DataFrame(long_rows)

    # 3. 珀莱雅分品牌收入结构表（brand_revenue，对应 business_segments 语义）
    brand_rows = [
        {"symbol": proya_code, "brand": "珀莱雅（主品牌）", "category": "核心护肤", "period": "2022A", "revenue": 52.64, "yoy": 37.46, "share_pct": 82.44},
        {"symbol": proya_code, "brand": "珀莱雅（主品牌）", "category": "核心护肤", "period": "2023A", "revenue": 71.77, "yoy": 36.36, "share_pct": 80.60},
        {"symbol": proya_code, "brand": "珀莱雅（主品牌）", "category": "核心护肤", "period": "2024A", "revenue": 85.81, "yoy": 19.55, "share_pct": 79.62},
        {"symbol": proya_code, "brand": "珀莱雅（主品牌）", "category": "核心护肤", "period": "2025A", "revenue": 76.89, "yoy": -10.39, "share_pct": 72.56},
        {"symbol": proya_code, "brand": "珀莱雅（主品牌）", "category": "核心护肤", "period": "2026H1", "revenue": 36.92, "yoy": -7.19, "share_pct": 68.79},

        {"symbol": proya_code, "brand": "彩棠（Timage）", "category": "专业彩妆", "period": "2022A", "revenue": 5.72, "yoy": 132.04, "share_pct": 8.96},
        {"symbol": proya_code, "brand": "彩棠（Timage）", "category": "专业彩妆", "period": "2023A", "revenue": 10.01, "yoy": 75.06, "share_pct": 11.24},
        {"symbol": proya_code, "brand": "彩棠（Timage）", "category": "专业彩妆", "period": "2024A", "revenue": 11.91, "yoy": 19.04, "share_pct": 11.05},
        {"symbol": proya_code, "brand": "彩棠（Timage）", "category": "专业彩妆", "period": "2025A", "revenue": 12.55, "yoy": 5.37, "share_pct": 11.84},
        {"symbol": proya_code, "brand": "彩棠（Timage）", "category": "专业彩妆", "period": "2026H1", "revenue": 5.51, "yoy": -21.93, "share_pct": 10.26},

        {"symbol": proya_code, "brand": "Off&Relax（OR）", "category": "头皮洗护", "period": "2022A", "revenue": 1.26, "yoy": 50.00, "share_pct": 1.97},
        {"symbol": proya_code, "brand": "Off&Relax（OR）", "category": "头皮洗护", "period": "2023A", "revenue": 2.15, "yoy": 70.63, "share_pct": 2.41},
        {"symbol": proya_code, "brand": "Off&Relax（OR）", "category": "头皮洗护", "period": "2024A", "revenue": 3.68, "yoy": 71.14, "share_pct": 3.41},
        {"symbol": proya_code, "brand": "Off&Relax（OR）", "category": "头皮洗护", "period": "2025A", "revenue": 7.44, "yoy": 102.19, "share_pct": 7.02},
        {"symbol": proya_code, "brand": "Off&Relax（OR）", "category": "头皮洗护", "period": "2026H1", "revenue": 4.77, "yoy": 70.81, "share_pct": 8.87},

        {"symbol": proya_code, "brand": "悦芙媞（UTIQUE）", "category": "年轻护肤", "period": "2022A", "revenue": 1.87, "yoy": 188.27, "share_pct": 2.93},
        {"symbol": proya_code, "brand": "悦芙媞（UTIQUE）", "category": "年轻护肤", "period": "2023A", "revenue": 3.03, "yoy": 61.82, "share_pct": 3.40},
        {"symbol": proya_code, "brand": "悦芙媞（UTIQUE）", "category": "年轻护肤", "period": "2024A", "revenue": 3.32, "yoy": 9.41, "share_pct": 3.08},
        {"symbol": proya_code, "brand": "悦芙媞（UTIQUE）", "category": "年轻护肤", "period": "2025A", "revenue": 3.71, "yoy": 11.80, "share_pct": 3.50},
        {"symbol": proya_code, "brand": "悦芙媞（UTIQUE）", "category": "年轻护肤", "period": "2026H1", "revenue": 1.75, "yoy": 5.02, "share_pct": 3.26},

        {"symbol": proya_code, "brand": "原色波塔（INSBAHA）", "category": "新锐彩妆", "period": "2024A", "revenue": 1.14, "yoy": 120.00, "share_pct": 1.06},
        {"symbol": proya_code, "brand": "原色波塔（INSBAHA）", "category": "新锐彩妆", "period": "2025A", "revenue": 2.56, "yoy": 125.38, "share_pct": 2.42},
        {"symbol": proya_code, "brand": "原色波塔（INSBAHA）", "category": "新锐彩妆", "period": "2026H1", "revenue": 3.13, "yoy": 222.23, "share_pct": 5.82},
    ]
    df_brands = pd.DataFrame(brand_rows)

    # 4. 并购与商誉风险事件表（ma_events）
    ma_rows = [
        {
            "symbol": proya_code, "company": "珀莱雅", "target": "花知晓（Flower Knows）",
            "deal_timeline": "2025年首购38.45%（4.28亿元）；2026年5月增持至51%控股并自6月并表（3.51亿元）",
            "total_consideration_bn": 7.79, "stake_pct": 51.00, "goodwill_added_bn": 9.24,
            "revaluation_gain_bn": 4.45, "target_2025_rev_bn": 17.26, "target_2025_np_bn": 2.80,
            "status_note": "非同一控制下分步并购产生4.45亿元公允价值重估一次性收益，推高账面归母净利+22.46%，但形成9.24亿元商誉（占净资产14.57%），并伴随低毛利彩妆并表稀释Q3毛利率至69.78%、推高推广费。",
        },
        {
            "symbol": botanee_code, "company": "贝泰妮", "target": "悦江投资（Za姬芮、泊美）",
            "deal_timeline": "2023年9月现金收购51%股权（5.355亿元）；2023–2025业绩对赌期满未达标触发股权补偿",
            "total_consideration_bn": 5.355, "stake_pct": 69.28, "goodwill_added_bn": 4.14,
            "revaluation_gain_bn": 0.00, "target_2025_rev_bn": 4.80, "target_2025_np_bn": 0.35,
            "status_note": "2023–2025三年累计净利润未达2.35亿元承诺底线，2024–2025连续计提商誉及无形资产减值（商誉由4.14亿降至2.93亿），触发原股东以转增资本公积方式进行股权补偿，贝泰妮持股升至69.28%。",
        },
    ]
    df_ma = pd.DataFrame(ma_rows)

    return {
        "financials_wide": df_fin,
        "financial_statement_items": df_cnequity_items,
        "brand_revenue": df_brands,
        "ma_events": df_ma,
    }


def persist_to_duckdb_and_parquet(datasets: dict[str, pd.DataFrame]) -> None:
    """将全部数据一次性写入本地 DuckDB 数据库及 Parquet 数据湖目录。"""
    PARQUET_DIR.mkdir(parents=True, exist_ok=True)
    if DUCKDB_PATH.exists():
        DUCKDB_PATH.unlink()

    con = duckdb.connect(str(DUCKDB_PATH))
    try:
        for table_name, df in datasets.items():
            pq_path = PARQUET_DIR / f"{table_name}.parquet"
            df.to_parquet(pq_path, index=False)
            con.register(f"_tmp_{table_name}", df)
            con.execute(f"CREATE OR REPLACE TABLE {table_name} AS SELECT * FROM _tmp_{table_name}")
            con.unregister(f"_tmp_{table_name}")

        # 校验 CNEquity financial_statement_items 契约字段完整性
        expected_cols = list(DATASET_SCHEMAS["financial_statement_items"].keys())
        actual_cols = [r[0] for r in con.execute("DESCRIBE financial_statement_items").fetchall()]
        assert actual_cols == expected_cols, f"Schema mismatch: {actual_cols} != {expected_cols}"
    finally:
        con.close()


def compute_all_metrics() -> dict[str, pd.DataFrame]:
    """从 DuckDB 读取数据，使用 pandas 完成全部比率、杜邦、EVA、现金流质量及 BSC 测算。"""
    con = duckdb.connect(str(DUCKDB_PATH), read_only=True)
    try:
        df = con.execute("SELECT * FROM financials_wide ORDER BY symbol DESC, period_order ASC").df()
        df_brands = con.execute("SELECT * FROM brand_revenue").df()
        df_ma = con.execute("SELECT * FROM ma_events").df()
    finally:
        con.close()

    # 按公司与报告期排序，计算期初期末平均资产、平均归母净资产、平均存货与平均应收账款
    df = df.sort_values(["company", "period_order"]).reset_index(drop=True)

    # 提取年度与前三季度连续序列（2026Q3单季资产负债表期初取2026H1或近似期初，此处单独处理）
    avg_assets_list = []
    avg_equity_list = []
    avg_inv_list = []
    avg_ar_list = []
    for idx, row in df.iterrows():
        comp = row["company"]
        p_ord = row["period_order"]
        if p_ord == 1:
            # 2022A 采用期初期末均值估计（2021年末珀莱雅总资产46.30亿/净资产28.91亿，贝泰妮总资产53.85亿/净资产48.10亿）
            if comp == "珀莱雅":
                prev_a, prev_e, prev_inv, prev_ar = 46.30, 28.91, 4.65, 1.32
            else:
                prev_a, prev_e, prev_inv, prev_ar = 53.85, 48.10, 4.61, 2.25
        elif p_ord == 6:
            # 2026Q3单季，资产负债表取当期均值近似（与2026Q1-Q3期末一致）
            prev_a = row["total_assets"]
            prev_e = row["equity_parent"]
            prev_inv = row["inventory"]
            prev_ar = row["accounts_receivable"]
        else:
            prev_row = df[(df["company"] == comp) & (df["period_order"] == p_ord - 1)].iloc[0]
            prev_a = prev_row["total_assets"]
            prev_e = prev_row["equity_parent"]
            prev_inv = prev_row["inventory"]
            prev_ar = prev_row["accounts_receivable"]

        avg_assets_list.append((prev_a + row["total_assets"]) / 2.0)
        avg_equity_list.append((prev_e + row["equity_parent"]) / 2.0)
        avg_inv_list.append((prev_inv + row["inventory"]) / 2.0)
        avg_ar_list.append((prev_ar + row["accounts_receivable"]) / 2.0)

    df["avg_assets"] = avg_assets_list
    df["avg_equity"] = avg_equity_list
    df["avg_inventory"] = avg_inv_list
    df["avg_ar"] = avg_ar_list

    # --------------------------------------------------------------------------
    # 模块一：盈利能力、费用率与流量内卷效率（5.1 & 5.2）
    # --------------------------------------------------------------------------
    df["gross_margin_calc"] = (df["revenue"] - df["cogs"]) / df["revenue"] * 100.0
    df["selling_exp_ratio"] = df["selling_exp"] / df["revenue"] * 100.0
    df["admin_exp_ratio"] = df["admin_exp"] / df["revenue"] * 100.0
    df["rd_exp_ratio"] = df["rd_exp"] / df["revenue"] * 100.0
    df["period_exp_ratio"] = (df["selling_exp"] + df["admin_exp"] + df["rd_exp"] + df["fin_exp"]) / df["revenue"] * 100.0
    df["rd_to_selling_ratio"] = df["rd_exp"] / df["selling_exp"] * 100.0

    # 流量 ROE（营销投入边际产出）：单位销售费用创造的营收 & 单位销售费用创造的扣非净利
    df["traffic_rev_roi"] = df["revenue"] / df["selling_exp"]
    df["traffic_dnp_roi"] = df["deducted_np"] / df["selling_exp"] * 100.0
    df["non_recurring_np"] = df["net_profit_parent"] - df["deducted_np"]
    df["non_recurring_share_pct"] = df["non_recurring_np"] / df["net_profit_parent"] * 100.0
    df["debt_ratio"] = df["total_liab"] / df["total_assets"] * 100.0
    df["goodwill_to_equity_pct"] = df["goodwill"] / df["equity_parent"] * 100.0
    df["goodwill_to_assets_pct"] = df["goodwill"] / df["total_assets"] * 100.0

    # --------------------------------------------------------------------------
    # 模块二：双轨杜邦三因子拆解（名义 GAAP ROE vs 扣非真实 ROE）（5.3）
    # Reference: FinanceToolkit DuPont Identity = Net Profit Margin * Asset Turnover * Equity Multiplier
    # --------------------------------------------------------------------------
    df["npm_gaap_pct"] = df["net_profit_parent"] / df["revenue"] * 100.0
    df["npm_deducted_pct"] = df["deducted_np"] / df["revenue"] * 100.0
    df["asset_turnover_period"] = df["revenue"] / df["avg_assets"]
    df["asset_turnover_ann"] = df["asset_turnover_period"] * df["annualize_factor"]
    df["equity_multiplier"] = df["avg_assets"] / df["avg_equity"]

    # 期间原始杜邦 ROE 与 年化杜邦 ROE
    df["dupont_roe_gaap_period"] = (df["npm_gaap_pct"] / 100.0) * df["asset_turnover_period"] * df["equity_multiplier"] * 100.0
    df["dupont_roe_deducted_period"] = (df["npm_deducted_pct"] / 100.0) * df["asset_turnover_period"] * df["equity_multiplier"] * 100.0
    df["dupont_roe_gaap_ann"] = (df["npm_gaap_pct"] / 100.0) * df["asset_turnover_ann"] * df["equity_multiplier"] * 100.0
    df["dupont_roe_deducted_ann"] = (df["npm_deducted_pct"] / 100.0) * df["asset_turnover_ann"] * df["equity_multiplier"] * 100.0
    df["roe_bubble_spread"] = df["dupont_roe_gaap_ann"] - df["dupont_roe_deducted_ann"]

    # --------------------------------------------------------------------------
    # 模块三：EVA（经济增加值）真实价值穿透测算（5.3）
    # EVA = NOPAT - Invested Capital * WACC * (period_fraction)
    # 穿透原则：剔除并购花知晓带来的一次性公允价值重估非经常性收益，以扣非净利+税后利息构建真实 NOPAT；
    #          投入资本 TC = 期末归母净资产 + 有息负债（短期借款+长期借款+应付债券），全额纳入并购新增商誉9.24亿元资本占用。
    # WACC 参数设定（参考中国十年期国债收益率Rf=2.10%，市场风险溢价ERP=6.00%，美妆行业Beta=1.05 -> Ke=8.40%；
    #              税前债务成本Kd=3.20%，高新技术企业所得税率15% -> 税后Kd=2.72%）：
    #              结合资本结构，珀莱雅动态 WACC ≈ 7.43%–7.69%，贝泰妮动态 WACC ≈ 7.87%–8.40%。
    # --------------------------------------------------------------------------
    rf = 0.021
    erp = 0.060
    beta = 1.05
    ke = rf + beta * erp  # 8.40%
    kd_after_tax = 0.032 * (1 - 0.15)  # 2.72%

    df["interest_bearing_debt"] = df["short_debt"] + df["long_debt"] + df["bonds_payable"]
    df["invested_capital"] = df["equity_parent"] + df["interest_bearing_debt"]
    # 采用期末投入资本（若为2026Q1-Q3，将投入资本成本按 3/4 年折算以与前三季度 NOPAT 匹配；2026Q3单季按 1/4 年折算）
    # 以确保期末并购新增的 9.24 亿商誉与并购资金占用不被期初零商誉均值稀释
    df["wacc_pct"] = (
        (df["equity_parent"] / df["invested_capital"]) * ke
        + (df["interest_bearing_debt"] / df["invested_capital"]) * kd_after_tax
    ) * 100.0

    # 真实 NOPAT（还原真实主营经营税后利润）= 扣非归母净利润 + 利息费用*(1-15%)
    df["nopat_core"] = df["deducted_np"] + df["interest_exp"] * (1.0 - 0.15)
    df["capital_charge"] = df["invested_capital"] * (df["wacc_pct"] / 100.0) / df["annualize_factor"]
    df["eva"] = df["nopat_core"] - df["capital_charge"]
    df["eva_ann"] = df["eva"] * df["annualize_factor"]
    df["roic_ann_pct"] = (df["nopat_core"] * df["annualize_factor"]) / df["invested_capital"] * 100.0
    df["eva_spread_pct"] = df["roic_ann_pct"] - df["wacc_pct"]
    df["gaap_np_minus_eva"] = df["net_profit_parent"] - df["eva"]

    # --------------------------------------------------------------------------
    # 模块四：现金流质量与营运效率评价（5.4）
    # --------------------------------------------------------------------------
    df["cash_to_rev_ratio"] = df["cash_received"] / df["revenue"]  # 收现比
    df["ocf_to_np_ratio"] = df["ocf"] / df["net_profit_parent"]     # 净现比（GAAP口径）
    df["ocf_to_dnp_ratio"] = df["ocf"] / df["deducted_np"]          # 净现比（扣非口径）
    df["accrual_ratio_pct"] = (df["net_profit_parent"] - df["ocf"]) / df["avg_assets"] * 100.0  # 应计利润占总资产比
    df["ocf_to_current_liab"] = df["ocf"] / df["current_liab"]      # 现金流动负债比
    df["inventory_turnover"] = (df["cogs"] * df["annualize_factor"]) / df["avg_inventory"]
    df["dio_days"] = df["days_in_period"] / (df["cogs"] / df["avg_inventory"])
    df["ar_turnover"] = (df["revenue"] * df["annualize_factor"]) / df["avg_ar"]
    df["dso_days"] = df["days_in_period"] / (df["revenue"] / df["avg_ar"])

    # --------------------------------------------------------------------------
    # 模块五：平衡计分卡（BSC）四维量化评价矩阵（0–100分制）（5.4）
    # --------------------------------------------------------------------------
    bsc_rows = [
        {
            "entity": "珀莱雅 (2023A 黄金增长期)",
            "financial_score": 92.0,
            "customer_traffic_score": 88.0,
            "internal_product_score": 86.0,
            "learning_ma_score": 76.0,
            "composite_score": round(0.30 * 92 + 0.25 * 88 + 0.25 * 86 + 0.20 * 76, 2),
            "key_diagnosis": "大单品矩阵放量+高周转高ROE（29.94%），EVA高达8.03亿，但研发费用率仅1.95%埋下后继隐忧。",
        },
        {
            "entity": "珀莱雅 (2025A 拐点承压期)",
            "financial_score": 74.0,
            "customer_traffic_score": 66.0,
            "internal_product_score": 64.0,
            "learning_ma_score": 68.0,
            "composite_score": round(0.30 * 74 + 0.25 * 66 + 0.25 * 64 + 0.20 * 68, 2),
            "key_diagnosis": "主品牌营收下滑10.39%，整体营收首现负增长（-1.68%），销售费用率逼近50%，流量边际效益递减。",
        },
        {
            "entity": "珀莱雅 (2026Q1-Q3 并购粉饰期)",
            "financial_score": 56.0,
            "customer_traffic_score": 48.0,
            "internal_product_score": 52.0,
            "learning_ma_score": 54.0,
            "composite_score": round(0.30 * 56 + 0.25 * 48 + 0.25 * 52 + 0.20 * 54, 2),
            "key_diagnosis": "花知晓并表贡献账面营收与4.45亿重估收益，但扣非净利暴跌25.5%（Q3跌65.2%）、OCF骤降61.06%、9.24亿商誉悬顶、存货飙至11.75亿。",
        },
        {
            "entity": "贝泰妮 (2024A 转型阵痛谷底)",
            "financial_score": 48.0,
            "customer_traffic_score": 55.0,
            "internal_product_score": 68.0,
            "learning_ma_score": 72.0,
            "composite_score": round(0.30 * 48 + 0.25 * 55 + 0.25 * 68 + 0.20 * 72, 2),
            "key_diagnosis": "大促投入失衡与悦江投资对赌承压导致扣非净利探底2.40亿（-61.12%），但研发投入率维持5.88%高位。",
        },
        {
            "entity": "贝泰妮 (2026Q1-Q3 控费修复期)",
            "financial_score": 68.0,
            "customer_traffic_score": 67.0,
            "internal_product_score": 74.0,
            "learning_ma_score": 81.0,
            "composite_score": round(0.30 * 68 + 0.25 * 67 + 0.25 * 74 + 0.20 * 81, 2),
            "key_diagnosis": "主动收缩低效规模、聚焦高毛利（74.65%）与强研发（5.17%），扣非净利连续两年强劲修复（2025 +48.22%，2026Q1-Q3 +21.38%）。",
        },
    ]
    df_bsc = pd.DataFrame(bsc_rows)

    # 将衍生计算结果回写 DuckDB 供审计与 SQL 检索
    con = duckdb.connect(str(DUCKDB_PATH))
    try:
        con.register("_tmp_metrics", df)
        con.execute("CREATE OR REPLACE TABLE computed_financial_metrics AS SELECT * FROM _tmp_metrics")
        con.unregister("_tmp_metrics")

        con.register("_tmp_bsc", df_bsc)
        con.execute("CREATE OR REPLACE TABLE bsc_evaluation_matrix AS SELECT * FROM _tmp_bsc")
        con.unregister("_tmp_bsc")
    finally:
        con.close()

    # 导出 Excel 多工作表汇总文件
    with pd.ExcelWriter(EXCEL_OUT, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="全量财务与杜邦EVA指标", index=False)
        df_brands.to_excel(writer, sheet_name="珀莱雅分品牌收入结构", index=False)
        df_ma.to_excel(writer, sheet_name="并购与商誉事件对比", index=False)
        df_bsc.to_excel(writer, sheet_name="BSC平衡计分卡评价", index=False)

    # 导出 JSON 供图表与报告模板使用
    payload = {
        "metrics": df.round(4).to_dict(orient="records"),
        "brands": df_brands.round(4).to_dict(orient="records"),
        "ma_events": df_ma.round(4).to_dict(orient="records"),
        "bsc": df_bsc.round(4).to_dict(orient="records"),
    }
    JSON_OUT.write_text(json.dumps(payload, ensure_ascii=False, indent=2, default=str), encoding="utf-8")

    return {
        "metrics": df,
        "brands": df_brands,
        "ma_events": df_ma,
        "bsc": df_bsc,
    }


def main() -> None:
    datasets = build_raw_datasets()
    persist_to_duckdb_and_parquet(datasets)
    results = compute_all_metrics()

    df = results["metrics"]
    proya = df[df["company"] == "珀莱雅"][[
        "period", "revenue", "rev_yoy", "net_profit_parent", "deducted_np",
        "selling_exp_ratio", "rd_exp_ratio", "traffic_dnp_roi",
        "npm_deducted_pct", "asset_turnover_ann", "equity_multiplier",
        "dupont_roe_gaap_ann", "dupont_roe_deducted_ann", "eva", "roic_ann_pct",
        "ocf", "ocf_to_dnp_ratio"
    ]]
    print("=== 珀莱雅（603605.SH）关键计算结果 ===")
    print(proya.to_string(index=False))

    botanee = df[df["company"] == "贝泰妮"][[
        "period", "revenue", "rev_yoy", "net_profit_parent", "deducted_np",
        "selling_exp_ratio", "rd_exp_ratio", "traffic_dnp_roi",
        "dupont_roe_deducted_ann", "eva", "ocf", "ocf_to_dnp_ratio"
    ]]
    print("\n=== 贝泰妮（300957.SZ）横向对标关键计算结果 ===")
    print(botanee.to_string(index=False))
    print(f"\n[OK] DuckDB saved to: {DUCKDB_PATH}")
    print(f"[OK] Excel saved to : {EXCEL_OUT}")
    print(f"[OK] JSON saved to  : {JSON_OUT}")


if __name__ == "__main__":
    main()
