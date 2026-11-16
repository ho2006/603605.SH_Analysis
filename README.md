# 603605.SH_Analysis

珀莱雅（603605.SH）商业深度分析 ——「5 财务绩效分析」：CNEquity + DuckDB + pandas 指标引擎 → Plotly 图表 → Word（docxtpl）/ PDF（WeasyPrint）双格式报告。横向对标：贝泰妮（300957.SZ）。

> 数据审计结论和置信度见 [AUDIT.md](AUDIT.md)。

## 目录

| 路径 | 内容 |
|---|---|
| `src/pipeline.py` | 构建 2022A–2026Q3 数据集 → DuckDB / Parquet → 计算比率、双轨杜邦、EVA、现金流质量、BSC → 输出 `data/computed_metrics.json` 和 `data/financial_metrics_summary.xlsx` |
| `src/charts.py` | 读取 JSON，导出 6 幅 PNG + SVG 图表到 `assets/charts/` |
| `src/build_reports.py` | 生成 docxtpl 模板 `templates/report_template.docx`，渲染 DOCX，再生成 HTML 并通过 WeasyPrint 输出 PDF |
| `data/` | DuckDB、Parquet、计算结果 JSON 和 Excel |
| `configs/cnequity.toml` | CNEquity 数据湖配置（`data.root` 需要改成本机路径） |
| `珀莱雅财务绩效分析报告_5_财务绩效分析.docx/.pdf` | 交付报告 |

## 环境（仅使用 conda）

所有 Python 都在 conda 环境里运行，**不要直接调用系统 python**。

```powershell
conda env create -f environment.yml        # 首次创建环境 proya-fin
conda env update -f environment.yml --prune  # 依赖有变更时更新
```

## 运行

```powershell
pwsh scripts/run_all.ps1                  # 顺序执行 pipeline -> charts -> reports
pwsh scripts/run_all.ps1 -Step charts     # 只执行其中一步

# 也可以逐个执行
conda run -n proya-fin python src/pipeline.py
conda run -n proya-fin python src/charts.py
conda run -n proya-fin python src/build_reports.py
```

注意：
- Plotly 静态导出依赖 kaleido。kaleido ≥ 1.0 需要本机装有 Chrome / Chromium
- conda-forge 的 `weasyprint` 已自带 pango / cairo，不需要 `tools/weasyprint_win`（该目录没有提交）
- `data/cnequity_lake/`（东方财富原始抓取缓存）没有提交，需要时用 `cne` 重新 backfill
