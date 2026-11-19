#!/usr/bin/env python3
"""生成《第五章 珀莱雅财务绩效分析》压缩正式稿。

表内数字来自用户修改稿（与原 PDF 逐格核对），图使用 assets/charts 的 PNG。
正文汉字只统计分析段落，不含标题、题注、表格和文末来源。
"""

from __future__ import annotations

import json
import re
import sys
from pathlib import Path

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Emu, Pt, RGBColor, Twips

ROOT = Path(__file__).resolve().parents[1]
TABLES_PATH = ROOT / "data" / "chapter5_tables.json"
PDF_PATH = ROOT / "珀莱雅财务绩效分析报告_5_财务绩效分析.pdf"
CHARTS = [
    ROOT / "assets/charts/fig1_growth_truth.png",
    ROOT / "assets/charts/fig2_brand_expense.png",
    ROOT / "assets/charts/fig3_peer_comparison.png",
    ROOT / "assets/charts/fig4_dupont_decomposition.png",
    ROOT / "assets/charts/fig5_eva_goodwill.png",
    ROOT / "assets/charts/fig6_cashflow_bsc.png",
]
OUT_DOCX = ROOT / "deliverables" / "珀莱雅第五章财务绩效分析_压缩版.docx"

FONT = "Noto Sans CJK SC"
BODY_MIN, BODY_MAX = 2270, 2775
CONTENT_WIDTH = 9706  # twips, Letter 12240 minus left/right 1267

# 分析段落。顺序即成稿顺序；题注、表、图在 main() 里按位置插入。
BODY = [
    "2022年至2024年，珀莱雅营业收入由63.85亿元增至107.78亿元，扣非归母净利润由7.89亿元增至15.22亿元，两者同向上升。2024年营业收入超过100亿元。原报告将其写为首家达到这一规模的国货美妆上市公司。2023年和2024年非经常性损益占比均不足2%，2022年为3.43%。2022年至2025年，非经常性损益每年不高于0.30亿元。2025年营业收入105.97亿元，同比下降1.68%，是这一期间营收首次下降的一年。扣非归母净利润14.73亿元，同比下降3.23%。2026年前三季度营业收入76.27亿元，同比增长7.46%，第三季度单季22.53亿元，同比增长29.76%。归母净利润12.56亿元，同比增长22.46%。扣非归母净利润7.43亿元，同比下降25.50%。第三季度单季扣非归母净利润0.79亿元，同比下降65.20%，单季归母净利润0.88亿元，同比下降61.18%。综合毛利率由2022年的69.70%升至2025年的73.26%，2026年前三季度为72.98%，第三季度单季回到69.78%，与花知晓彩妆并表在同一季度。〔1〕〔2〕〔3〕〔4〕〔6〕",
    "2025年，公司以4.28亿元受让花知晓38.45%股权。2026年5月再出资3.51亿元，持股达到51%，2026年6月起纳入合并报表。花知晓2025年营业收入17.26亿元，净利润2.80亿元。并表之后，2026年第三季度单季营业收入同比增长29.76%。按照《企业会计准则第33号——合并财务报表》，购买日前已持有的股权按公允价值重新计量，形成4.45亿元公允价值重估收益，占2026年前三季度归母净利润12.56亿元的35.43%。同期非经常性损益5.13亿元，占归母净利润40.84%。账面营业收入和归母净利润的上升，来自这笔公允价值重估收益和并表。扣非归母净利润在同一期间下降。〔4〕〔6〕〔7〕",
    "主品牌较早的增量来自早C晚A和源力修护等大单品。2025年以后这些单品不再带动收入。主品牌2024年营业收入85.81亿元，2025年76.89亿元，同比下降10.39%。2026年上半年36.92亿元，同比下降7.19%，占主营业务收入68.79%。2022年该占比为82.44%。彩棠2026年上半年营业收入5.51亿元，同比下降21.93%。Off&Relax同比增长70.81%，原色波塔同比增长222.23%。这两个品牌增加的收入仍小于主品牌和彩棠减少的收入，也还没有形成扣非利润。〔1〕〔4〕〔5〕",
    "销售费用率由2022年的43.63%升至2024年的47.88%，再升至2025年的49.63%。2026年前三季度为53.23%，销售费用40.60亿元。第三季度单季销售费用率53.44%。销售费用的同比增幅高于同期营业收入增幅，费用率继续上升。研发费用率维持在1.95%至2.05%，2026年前三季度为1.97%，金额约1.50亿元。扣非归母净利润除以销售费用，2023年为29.56%，2026年前三季度为18.30%，第三季度单季为6.56%。〔1〕〔2〕〔3〕〔4〕〔6〕",
    "本节以贝泰妮为对标。两家公司在2023年至2025年都遇到核心单品放缓，也都做了对外收购，可以把行业需求变化和本公司的费用、研发和并购分开看。2023年与2024年，珀莱雅扣非净利润同比增长48.91%和29.60%，贝泰妮同比下降35.20%和61.12%，2024年贝泰妮扣非净利润为2.40亿元。2025年贝泰妮扣非净利润同比增长48.22%，达到3.56亿元。2026年前三季度同比增长21.38%，达到2.23亿元。同期珀莱雅扣非净利润2025年下降3.23%，2026年前三季度下降25.50%，第三季度单季下降65.20%。〔4〕〔6〕〔8〕〔9〕〔10〕",
    "2026年前三季度销售费用率，珀莱雅为53.23%，贝泰妮为53.56%。贝泰妮综合毛利率在2023年至2025年处于73.74%至74.46%，2026年前三季度为74.65%，高于珀莱雅1.67个百分点，第三季度单季为74.71%。贝泰妮研发费用率在2023年至2026年前三季度处于4.42%至5.88%，2026年前三季度为5.17%，投入1.88亿元。珀莱雅同期为1.97%，投入1.50亿元。贝泰妮的研发费用率约为珀莱雅的2.6倍。2025年珀莱雅研发费用2.17亿元，相当于当年销售费用52.59亿元的4.13%。〔4〕〔6〕〔8〕〔9〕〔10〕",
    "2025年贝泰妮营业收入下降6.57%，销售费用下降6.63%，经营活动现金流同比增长40.54%，达到9.77亿元。珀莱雅主品牌营业收入在2025年下降10.39%之后，2026年前三季度销售费用仍同比增长15.16%，达到40.60亿元。〔4〕〔6〕〔9〕",
    "2023年9月，贝泰妮以5.355亿元收购悦江投资51%股权，形成商誉4.14亿元。悦江投资2023年至2025年累计净利润未达到2.35亿元的承诺，商誉余额减至2.93亿元，贝泰妮持股比例升至69.28%。收购当年商誉占净资产7.06%，到2026年前三季度降为4.85%。珀莱雅2025年末商誉为0。2026年分步取得花知晓控股权，出资合计7.79亿元。第三季度末商誉9.24亿元，占归母净资产63.41亿元的14.57%。花知晓若在并表后收入回落，需要减值的商誉绝对额高于贝泰妮当年的悦江投资。〔6〕〔7〕〔9〕〔11〕〔12〕",
    "净资产收益率等于销售净利率、总资产周转率和权益乘数的乘积。本节同时给出名义口径和扣非口径。资产与净资产使用期初和期末的平均数。2026年前三季度乘以4/3做线性年化，第三季度单季乘以4。两种年化都没有放入第四季度促销的季节性，全年周转率和净资产收益率取决于第四季度。2026年第三季度的平均数采用期末资产负债表，没有改用半年报的期初数。〔13〕",
    "2023年与2024年，扣非年化净资产收益率分别为29.82%和31.21%。销售净利率由13.18%升至14.12%，资产周转率由1.36次升至1.45次。2026年前三季度名义年化净资产收益率为26.83%，2025年为25.96%。名义销售净利率为16.47%，扣非销售净利率为9.74%。扣非年化净资产收益率为15.87%，低于名义口径10.96个百分点。〔6〕〔13〕",
    "销售净利率由2024年的14.12%、2025年的13.90%降到2026年前三季度的9.74%，第三季度单季为3.51%。原测算中，这一因子对应的年化净资产收益率下降超过7.5个百分点。总资产周转率由2024年的1.45次、2025年的1.32次降到2026年前三季度年化1.09次。2026年第三季度末总资产102.56亿元，同比增长20.84%。其中商誉9.24亿元，存货11.75亿元，较年初6.35亿元增长85.04%。权益乘数由2025年的1.39倍升至1.50倍。总负债由22.21亿元增至33.87亿元，资产负债率由26.17%升至33.02%。这意味着净利率和周转率不再支撑净资产收益率，账面收益率靠权益乘数和公允价值重估收益维持。〔6〕〔13〕",
    "经济增加值等于税后净营业利润减去投入资本与资本成本的乘积。税后净营业利润以扣非归母净利润为底，剔除4.45亿元公允价值重估收益，再加回税后利息。投入资本取期末归母净资产加有息负债。有息负债为短期借款0.80亿元与可转债8.32亿元，商誉9.24亿元计入投入资本。无风险利率用十年期国债2.10%，市场风险溢价6.00%，行业贝塔1.05，所得税率15%。资本成本因而落在7.43%至7.69%。2022年至2024年，经济增加值由4.87亿元增至10.73亿元。2024年投入资本回报率25.05%，比资本成本高17.36个百分点。2025年经济增加值9.60亿元，下降10.53%。2026年前三季度投入资本72.53亿元，资本成本4.18亿元，经济增加值3.49亿元。按4/3年化约为4.65亿元，低于2024年的10.73亿元的一半。第三季度单季投入资本回报率4.78%，低于当期资本成本7.69%，利差为负2.91个百分点，单季经济增加值为负0.53亿元。〔6〕〔13〕",
    "2022年经营活动现金流11.11亿元，高于当年归母净利润8.17亿元。2023年经营活动现金流14.69亿元，净现比1.23。2024年净现比为0.71。2025年经营活动现金流21.93亿元，净现比1.46。2026年前三季度销售商品收到的现金84.59亿元，收现比1.11，经营活动现金流4.67亿元，同比下降61.06%。2026年上半年经营活动现金流9.51亿元。前三季度减去上半年之后，第三季度单季净流出4.84亿元。前三季度净现比0.37，扣非净现比0.63，单季净现比为负5.50。流动负债现金覆盖率由2025年的1.60倍降到2026年前三季度的0.19倍，单季为负0.20倍。〔1〕〔2〕〔3〕〔4〕〔5〕〔6〕〔13〕",
    "存货由2025年末的6.35亿元增至2026年第三季度末的11.75亿元，增幅85.04%。存货周转天数由2025年的83.5天增至2026年前三季度的119.9天，第三季度单季为158.7天。应计利润占平均总资产的比重，由2025年的负8.68%变为2026年前三季度的正8.42%。现金流出用于形象宣传推广，以及并表和促销前的备货。〔4〕〔6〕〔13〕",
    "平衡计分卡里，财务权重30%，客户与流量权重25%。内部运营权重25%，学习与治理权重20%。四维分数来自人工评定。审计记录没有给出计算公式。珀莱雅综合分2023年为86.30分，2025年为68.30分，2026年前三季度为52.60分。财务维度由92.0分降到56.0分，与扣非利润和经营现金流的下降对应。客户与流量维度由88.0分降到48.0分，与销售费用率53.23%和单季流量净回报率6.56%对应。内部运营维度由86.0分降到52.0分。彩棠2026年上半年收入下降21.93%，存货周转天数增至119.9天。学习与治理维度由76.0分降到54.0分，对应约2%的研发费用率和9.24亿元商誉。贝泰妮2024年综合分59.55分，2026年前三季度为71.85分。〔6〕〔13〕",
    "提高扣非利润和现金流，有三项安排。其一，内部预算改按扣非投入资本回报率和经济增加值考核，停止回报低于资本成本的销售费用。其二，把研发费用率由约2%提高到4%至5%，用自有成分和配方更新大单品。其三，整合花知晓的供应和渠道，使净现比回到1.0以上，以经营利润承担9.24亿元商誉。〔13〕",
]

SOURCES = [
    (
        "〔1〕珀莱雅化妆品股份有限公司《2022年年度报告》，“公司简介和主要财务指标”及“管理层讨论与分析”。营业收入、扣非归母净利润和经营活动现金流取自主要会计数据，由元折合亿元。分品牌占比以主营业务收入为分母。该节按品牌拆分表列示珀莱雅品牌占比82.74%，正文82.44%沿用原报告，两数不一致，本章不改写正文。正文“早C晚A”是对大单品的概括，该节写的是双抗、红宝石和源力系列，没有使用这一叫法。",
        "https://static.cninfo.com.cn/finalpage/2023-04-21/1216488038.PDF",
    ),
    (
        "〔2〕珀莱雅化妆品股份有限公司《2023年年度报告》，“公司简介和主要财务指标”及“合并利润表”。2023年收入、利润、费用和现金流取该报告。销售费用率按合并销售费用除以营业收入计算，百分点由原始金额分别四舍五入。",
        "https://static.cninfo.com.cn/finalpage/2024-04-19/1219664274.PDF",
    ),
    (
        "〔3〕珀莱雅化妆品股份有限公司《2024年年度报告》，“公司简介和主要财务指标”及“合并利润表”。2024年收入、利润、费用和现金流取该报告。销售费用率按合并销售费用除以营业收入计算。净现比按经营活动现金流除以归母净利润计算。",
        "https://static.cninfo.com.cn/finalpage/2025-04-25/1223280349.PDF",
    ),
    (
        "〔4〕珀莱雅化妆品股份有限公司《2025年年度报告》，“管理层讨论与分析”及“合并财务报表项目注释”。主品牌收入同比为年度主营业务分品牌口径。按品牌拆分表列示珀莱雅品牌76.89亿元、同比下降10.39%，正文85.81亿元由该金额和降幅回推。投资支付的现金列示花知晓股权投资款427,975,763元。2025年末商誉无余额。表5-1中2022年至2025年非经常性损益每年不高于0.30亿元，沿用原报告，本次未逐笔重抄各年明细页码。",
        "https://static.cninfo.com.cn/finalpage/2026-04-22/1225138279.PDF",
    ),
    (
        "〔5〕珀莱雅化妆品股份有限公司《2026年半年度报告》，“管理层讨论与分析”。2026年上半年分品牌收入以主营业务收入为分母，同比以上年同期为基准。按品牌拆分表列示珀莱雅品牌36.92亿元，同比下降7.19%，占比68.79%。彩棠5.51亿元，同比下降21.93%。OR同比增长70.81%。原色波塔同比增长222.23%。经营活动现金流950,841,180.71元，供前三季度减去上半年时使用。",
        "https://static.cninfo.com.cn/finalpage/2026-08-25/1225496527.PDF",
    ),
    (
        "〔6〕珀莱雅化妆品股份有限公司《2026年第三季度报告》，“主要财务数据”及“合并利润表”，并参见“合并资产负债表”和“非经常性损益项目和金额”。销售费用率按合并销售费用除以营业收入计算，毛利率按营业收入减营业成本后除以营业收入计算，比例及百分点均依据原始金额分别四舍五入。销售商品收到的现金8,459,484,688.31元，经营活动现金流466,699,905.42元。4.45亿元是445,360,152.15元四舍五入，列在“其他符合非经常性损益定义的损益项目”，说明为分步合并中购买日前股权按购买日公允价值重新计量。非经常性损益合计512,823,232.22元。正文35.43%和40.84%按四舍五入后的亿元金额相除，不是报表单列比例。商誉923,541,784.87元，归母净资产6,340,869,593.99元，14.57%为二者相除。短期借款0.80亿元是2025年末余额，2026年9月30日短期借款无余额。可转债取期末应付债券831,811,409.90元。重估收益记在年初至报告期末，报告没有写明6月并表。",
        "https://static.cninfo.com.cn/finalpage/2026-10-30/1225640965.PDF",
    ),
    (
        "〔7〕珀莱雅化妆品股份有限公司《关于收购深圳花知晓电子商务有限公司部分股权的公告》（2026年5月22日），“交易概况”及“交易标的主要财务信息”。交易前持股38.4521%，本次受让12.5479%，价款351,142,798元，交割后持股51%。2025年营业收入1,726,350,299.48元，净利润279,841,201.10元。公告另写2025年8月29日增资认购款17,612.0590万元，与〔4〕的投资款427,975,763元不是同一表述。公告约定交割后纳入合并报表，并写明尚需经营者集中申报，没有写6月并表日期。本次在上交所公告标题中未另见交割完成公告。",
        "https://static.cninfo.com.cn/finalpage/2026-05-22/1225322125.PDF",
    ),
    (
        "〔8〕云南贝泰妮生物科技集团股份有限公司《2024年年度报告》，“公司简介和主要财务指标”及“管理层讨论与分析”。2022年至2024年营业收入和扣非净利润取三年对比栏。2023年扣非同比下降35.20%，按栏内618,058,791.53元和953,787,022.86元计算。2024年销售毛利率约73.74%，扣非净利润同比下降61.12%，均为公司披露口径。",
        "https://disc.static.szse.cn/download/disc/disk03/finalpage/2025-04-24/ae3f9647-3f1a-4a5a-b953-b898849fb0c4.PDF",
    ),
    (
        "〔9〕云南贝泰妮生物科技集团股份有限公司《2025年年度报告》，“公司简介和主要财务指标”及“合并财务报表项目注释”之“商誉”。2025年营业收入、扣非净利润、销售费用和经营现金流取该报告。扣非同比48.22%和经营现金流同比40.54%为公司披露。商誉原值413,740,760.61元，累计减值准备121,239,495.00元，期末账面292,501,265.61元。销售费用率按合并销售费用除以营业收入计算，报表未单列该比率。收购当年商誉占净资产7.06%，按商誉原值除以2023年末归母净资产5,860,994,371.02元计算。",
        "https://disc.static.szse.cn/disc/disk03/finalpage/2026-04-27/43bb9343-8d91-4c05-800d-1365f8bfb5f6.PDF",
    ),
    (
        "〔10〕云南贝泰妮生物科技集团股份有限公司《2026年第三季度报告》，“主要会计数据”及“合并利润表”。2026年前三季度扣非净利润同比21.38%为公司披露。销售费用率按合并销售费用除以营业收入计算。毛利率按营业收入减营业成本后除以营业收入计算。研发费用率按研发费用除以营业收入计算。比例依据原始金额四舍五入。期末商誉仍为292,501,265.61元，归母净资产6,046,013,130.98元。正文4.85%按2.93亿元除以60.46亿元四舍五入，直接相除约为4.84%。",
        "https://static.cninfo.com.cn/finalpage/2026-10-27/1225624886.PDF",
    ),
    (
        "〔11〕云南贝泰妮生物科技集团股份有限公司《关于对外投资收购股权及增资的公告》（2023年9月28日），“对外投资概述”及“交易方案”。股权转让款48,550万元加增资5,000万元，合计5.355亿元。完成后通过海南贝泰妮持股51%。业绩承诺为2023年至2025年累计归母净利润不低于2.35亿元。本公告不列商誉，商誉金额见〔9〕。",
        "https://disc.static.szse.cn/disc/disk03/finalpage/2023-09-28/28e221eb-8813-4db3-8d0a-aafd05aa5c00.PDF",
    ),
    (
        "〔12〕云南贝泰妮生物科技集团股份有限公司《关于悦江（广州）投资有限公司实施业绩补偿情况的公告》（2026年6月15日），“业绩承诺约定及完成情况”及“业绩补偿的执行方式”。定向转增后持股由51%增至69.28%，公告写明原始投资成本53,550万元。2023年至2025年累计净利润实现数10,241.06万元，低于2.35亿元承诺。",
        "https://disc.static.szse.cn/download/disc/disk03/finalpage/2026-06-15/948340e8-d806-4b81-a6e2-b4dd42af072b.PDF",
    ),
    (
        "〔13〕本章测算与平衡计分卡。杜邦比率、经济增加值和现金流比率都不是报表单列项目。收现比、净现比及存货周转天数同样由报表金额计算。净资产收益率等于销售净利率、总资产周转率和权益乘数的乘积。流量净回报率按扣非归母净利润除以销售费用计算。2026年前三季度线性年化系数为4/3，第三季度单季为4倍，都不含第四季度季节性。2026年第三季度平均数用期末资产负债表。经济增加值等于税后净营业利润减去投入资本与资本成本的乘积。税后净营业利润以扣非归母净利润为底，剔除4.45亿元公允价值重估收益后再加回税后利息。投入资本取期末归母净资产加有息负债，商誉计入投入资本。资本成本用无风险利率2.10%，市场风险溢价6.00%，贝塔1.05，所得税率15%，结果在7.43%至7.69%。销售净利率对年化净资产收益率的影响超过7.5个百分点，照录原测算，不另作分解。4.65亿元是3.49亿元按4/3年化。第三季度单季经营现金流由前三季度4.67亿元减去上半年9.51亿元。平衡计分卡权重为财务30%，客户与流量25%，内部运营25%，学习与治理20%。四维分数为本章人工评定，没有计算公式，也没有外部原文，因而不附查阅链接。",
        None,
    ),
]


def han_count(text: str) -> int:
    return len(re.findall(r"[\u4e00-\u9fff]", text))


def set_run_font(run, size_pt: float, bold: bool | None = None, color: str | None = None) -> None:
    run.font.name = FONT
    run.font.size = Pt(size_pt)
    if bold is not None:
        run.bold = bold
    if color:
        run.font.color.rgb = RGBColor.from_string(color)
    rpr = run._element.get_or_add_rPr()
    rfonts = rpr.find(qn("w:rFonts"))
    if rfonts is None:
        rfonts = OxmlElement("w:rFonts")
        rpr.append(rfonts)
    for attr in ("w:ascii", "w:hAnsi", "w:eastAsia", "w:cs"):
        rfonts.set(qn(attr), FONT)


def shade_cell(cell, fill: str) -> None:
    tc = cell._tc
    tc_pr = tc.get_or_add_tcPr()
    shd = tc_pr.find(qn("w:shd"))
    if shd is None:
        shd = OxmlElement("w:shd")
        tc_pr.append(shd)
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), fill)


def set_cell_border(cell) -> None:
    tc = cell._tc
    tc_pr = tc.get_or_add_tcPr()
    tc_borders = tc_pr.find(qn("w:tcBorders"))
    if tc_borders is None:
        tc_borders = OxmlElement("w:tcBorders")
        tc_pr.append(tc_borders)
    for edge in ("top", "left", "bottom", "right"):
        el = tc_borders.find(qn(f"w:{edge}"))
        if el is None:
            el = OxmlElement(f"w:{edge}")
            tc_borders.append(el)
        el.set(qn("w:val"), "single")
        el.set(qn("w:sz"), "4")
        el.set(qn("w:space"), "0")
        el.set(qn("w:color"), "D9D9D9")


def set_repeat_header(row) -> None:
    tr = row._tr
    tr_pr = tr.get_or_add_trPr()
    mark = tr_pr.find(qn("w:tblHeader"))
    if mark is None:
        mark = OxmlElement("w:tblHeader")
        tr_pr.append(mark)


def set_cant_split(row) -> None:
    tr = row._tr
    tr_pr = tr.get_or_add_trPr()
    if tr_pr.find(qn("w:cantSplit")) is None:
        tr_pr.append(OxmlElement("w:cantSplit"))


def set_table_width(table, widths: list[int]) -> None:
    tbl = table._tbl
    tbl_pr = tbl.tblPr
    tbl_w = tbl_pr.find(qn("w:tblW"))
    if tbl_w is None:
        tbl_w = OxmlElement("w:tblW")
        tbl_pr.append(tbl_w)
    tbl_w.set(qn("w:w"), str(sum(widths)))
    tbl_w.set(qn("w:type"), "dxa")
    layout = tbl_pr.find(qn("w:tblLayout"))
    if layout is None:
        layout = OxmlElement("w:tblLayout")
        tbl_pr.append(layout)
    layout.set(qn("w:type"), "fixed")
    grid = tbl.find(qn("w:tblGrid"))
    if grid is not None:
        for child in list(grid):
            grid.remove(child)
    else:
        grid = OxmlElement("w:tblGrid")
        tbl_pr.addnext(grid)
    for w in widths:
        gc = OxmlElement("w:gridCol")
        gc.set(qn("w:w"), str(w))
        grid.append(gc)
    for row in table.rows:
        for idx, cell in enumerate(row.cells):
            tc = cell._tc
            tc_pr = tc.get_or_add_tcPr()
            tc_w = tc_pr.find(qn("w:tcW"))
            if tc_w is None:
                tc_w = OxmlElement("w:tcW")
                tc_pr.append(tc_w)
            tc_w.set(qn("w:w"), str(widths[min(idx, len(widths) - 1)]))
            tc_w.set(qn("w:type"), "dxa")


def column_widths(n: int) -> list[int]:
    if n <= 1:
        return [CONTENT_WIDTH]
    first = int(CONTENT_WIDTH * 0.22)
    other = (CONTENT_WIDTH - first) // (n - 1)
    widths = [first] + [other] * (n - 2)
    widths.append(CONTENT_WIDTH - sum(widths))
    return widths


def add_page_number(paragraph) -> None:
    run = paragraph.add_run("珀莱雅商业案例分析  /  ")
    set_run_font(run, 9, color="536579")

    def field_run(kind: str, text: str | None = None):
        r = paragraph.add_run()
        set_run_font(r, 9, color="536579")
        if kind == "instr":
            instr = OxmlElement("w:instrText")
            instr.set(qn("xml:space"), "preserve")
            instr.text = text or ""
            r._r.append(instr)
        else:
            fld = OxmlElement("w:fldChar")
            fld.set(qn("w:fldCharType"), kind)
            r._r.append(fld)

    field_run("begin")
    field_run("instr", " PAGE ")
    field_run("end")


def configure_styles(doc: Document) -> None:
    normal = doc.styles["Normal"]
    normal.font.name = FONT
    normal.font.size = Pt(11)
    normal.font.color.rgb = RGBColor.from_string("202632")
    rpr = normal.element.get_or_add_rPr()
    rfonts = rpr.find(qn("w:rFonts"))
    if rfonts is None:
        rfonts = OxmlElement("w:rFonts")
        rpr.append(rfonts)
    for attr in ("w:ascii", "w:hAnsi", "w:eastAsia", "w:cs"):
        rfonts.set(qn(attr), FONT)
    pf = normal.paragraph_format
    pf.space_before = Pt(0)
    pf.space_after = Pt(6)
    pf.line_spacing = 1.270833
    pf.first_line_indent = None
    pf.alignment = WD_ALIGN_PARAGRAPH.LEFT

    h1 = doc.styles["Heading 1"]
    h1.font.name = FONT
    h1.font.size = Pt(16)
    h1.font.bold = True
    h1.font.color.rgb = RGBColor.from_string("000000")
    h1.font.italic = False
    rpr = h1.element.get_or_add_rPr()
    rfonts = rpr.find(qn("w:rFonts"))
    if rfonts is None:
        rfonts = OxmlElement("w:rFonts")
        rpr.append(rfonts)
    for attr in ("w:ascii", "w:hAnsi", "w:eastAsia", "w:cs"):
        rfonts.set(qn(attr), FONT)
    pf = h1.paragraph_format
    pf.space_before = Pt(14)
    pf.space_after = Pt(7)
    pf.line_spacing = 1.0
    pf.keep_with_next = True
    pf.page_break_before = False

    cap = doc.styles["Caption"]
    cap.font.name = FONT
    cap.font.size = Pt(9.5)
    cap.font.bold = True
    cap.font.color.rgb = RGBColor.from_string("30465C")
    rpr = cap.element.get_or_add_rPr()
    rfonts = rpr.find(qn("w:rFonts"))
    if rfonts is None:
        rfonts = OxmlElement("w:rFonts")
        rpr.append(rfonts)
    for attr in ("w:ascii", "w:hAnsi", "w:eastAsia", "w:cs"):
        rfonts.set(qn(attr), FONT)
    pf = cap.paragraph_format
    pf.space_before = Pt(4)
    pf.space_after = Pt(5)
    pf.line_spacing = 1.0
    pf.keep_with_next = True


def add_body(doc: Document, text: str):
    p = doc.add_paragraph(style="Normal")
    p.paragraph_format.space_after = Pt(6)
    p.paragraph_format.line_spacing = 1.270833
    p.paragraph_format.first_line_indent = Cm(0)
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    run = p.add_run(text)
    set_run_font(run, 11, color="202632")
    return p


def add_heading(doc: Document, text: str):
    p = doc.add_paragraph(text, style="Heading 1")
    for run in p.runs:
        set_run_font(run, 16, bold=True, color="000000")
    p.paragraph_format.keep_with_next = True
    return p


def add_caption(doc: Document, text: str):
    p = doc.add_paragraph(text, style="Caption")
    for run in p.runs:
        set_run_font(run, 9.5, bold=True, color="30465C")
    p.paragraph_format.keep_with_next = True
    p.paragraph_format.space_before = Pt(4)
    p.paragraph_format.space_after = Pt(5)
    return p


def display_url(url: str) -> str:
    """在斜线后插入零宽空格，长链接可在版心内换行，字符本身仍是原 URL。"""
    prefix = "https://"
    if url.startswith(prefix):
        head, rest = prefix, url[len(prefix):]
    else:
        head, rest = "", url
    return head + rest.replace("/", "/\u200b")


def add_hyperlink(paragraph, text: str, url: str) -> None:
    r_id = paragraph.part.relate_to(
        url,
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink",
        is_external=True,
    )
    hyperlink = OxmlElement("w:hyperlink")
    hyperlink.set(qn("r:id"), r_id)
    run = OxmlElement("w:r")
    r_pr = OxmlElement("w:rPr")
    r_style = OxmlElement("w:rStyle")
    r_style.set(qn("w:val"), "Hyperlink")
    r_pr.append(r_style)
    color = OxmlElement("w:color")
    color.set(qn("w:val"), "304F70")
    r_pr.append(color)
    size = OxmlElement("w:sz")
    size.set(qn("w:val"), "18")
    r_pr.append(size)
    size_cs = OxmlElement("w:szCs")
    size_cs.set(qn("w:val"), "18")
    r_pr.append(size_cs)
    r_fonts = OxmlElement("w:rFonts")
    for attr in ("w:ascii", "w:hAnsi", "w:eastAsia", "w:cs"):
        r_fonts.set(qn(attr), FONT)
    r_pr.append(r_fonts)
    run.append(r_pr)
    node = OxmlElement("w:t")
    node.text = text
    run.append(node)
    hyperlink.append(run)
    paragraph._p.append(hyperlink)


def add_source(doc: Document, text: str, url: str | None):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(5)
    p.paragraph_format.line_spacing = 1.2791666666666666
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    run = p.add_run(text + (" " if url else ""))
    set_run_font(run, 9, color="202632")
    if url:
        add_hyperlink(p, "查阅原文", url)
        tail = p.add_run(" (" + display_url(url) + ")")
        set_run_font(tail, 9, color="202632")
    return p


def add_picture(doc: Document, path: Path):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(6)
    run = p.add_run()
    run.add_picture(str(path), width=Cm(16.2))
    return p


def add_table(doc: Document, rows: list[list[str]], widths: list[int] | None = None):
    n = max(len(r) for r in rows)
    table = doc.add_table(rows=len(rows), cols=n)
    table.alignment = WD_ALIGN_PARAGRAPH.CENTER
    table.autofit = False
    widths = widths or column_widths(n)
    for r_idx, row_data in enumerate(rows):
        row = table.rows[r_idx]
        set_cant_split(row)
        if r_idx == 0:
            set_repeat_header(row)
        for c_idx in range(n):
            value = row_data[c_idx] if c_idx < len(row_data) else ""
            cell = row.cells[c_idx]
            cell.text = ""
            para = cell.paragraphs[0]
            para.paragraph_format.space_before = Pt(0)
            para.paragraph_format.space_after = Pt(0)
            para.paragraph_format.line_spacing = 1.0
            run = para.add_run(value)
            if r_idx == 0:
                set_run_font(run, 9.5, bold=True, color="FFFFFF")
                shade_cell(cell, "233B53")
            else:
                set_run_font(run, 9.5, bold=False, color="202632")
                shade_cell(cell, "FFFFFF")
            set_cell_border(cell)
    set_table_width(table, widths)
    # 表后留出与正文段后相近的空隙
    spacer = doc.add_paragraph()
    spacer.paragraph_format.space_before = Pt(0)
    spacer.paragraph_format.space_after = Pt(2)
    return table


def configure_section(section) -> None:
    section.page_width = Twips(12240)
    section.page_height = Twips(15840)
    section.left_margin = Twips(1267)
    section.right_margin = Twips(1267)
    section.top_margin = Twips(1094)
    section.bottom_margin = Twips(1008)
    section.header_distance = Twips(720)
    section.footer_distance = Twips(720)
    section.different_first_page_header_footer = False
    header = section.header
    header.is_linked_to_previous = False
    for p in header.paragraphs:
        p.text = ""
    footer = section.footer
    footer.is_linked_to_previous = False
    p = footer.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    p.paragraph_format.space_after = Pt(0)
    # 清掉默认空 run
    for r in list(p.runs):
        r._element.getparent().remove(r._element)
    add_page_number(p)


def build() -> Document:
    tables = json.loads(TABLES_PATH.read_text(encoding="utf-8"))
    if len(tables) != 7:
        raise SystemExit(f"需要7张数据表，实际 {len(tables)}")
    doc = Document()
    configure_styles(doc)
    configure_section(doc.sections[0])

    title = doc.add_paragraph(style="Title")
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    title.paragraph_format.space_before = Pt(0)
    title.paragraph_format.space_after = Pt(14)
    title.paragraph_format.line_spacing = 1.0
    title.paragraph_format.keep_with_next = True
    run = title.add_run("第五章 珀莱雅财务绩效分析")
    set_run_font(run, 18, bold=True, color="000000")

    add_heading(doc, "5.1 核心财务指标的纵向变化")
    add_body(doc, BODY[0])
    add_caption(doc, "表 5-1  珀莱雅纵向核心财务绩效指标（2022A–2026Q3）〔1〕〔2〕〔3〕〔4〕〔6〕")
    add_table(doc, tables[0])
    add_caption(doc, "图 1  营业收入、归母净利润与扣非利润（2022–2026Q3）〔1〕〔2〕〔3〕〔4〕〔6〕")
    add_picture(doc, CHARTS[0])
    add_body(doc, BODY[1])
    add_body(doc, BODY[2])
    add_caption(doc, "表 5-2  珀莱雅分品牌收入（2023A–2026H1）〔2〕〔3〕〔4〕〔5〕")
    add_table(doc, tables[1])
    add_caption(doc, "图 2  分品牌收入与销售费用率〔4〕〔5〕〔6〕")
    add_picture(doc, CHARTS[1])
    add_body(doc, BODY[3])

    add_heading(doc, "5.2 与贝泰妮的横向对标")
    add_body(doc, BODY[4])
    add_caption(doc, "表 5-3  珀莱雅与贝泰妮核心指标对标〔4〕〔6〕〔8〕〔9〕〔10〕")
    add_table(doc, tables[2])
    add_caption(doc, "图 3  珀莱雅与贝泰妮的毛利率、费用与利润及商誉〔6〕〔9〕〔10〕〔11〕〔12〕")
    add_picture(doc, CHARTS[2])
    add_body(doc, BODY[5])
    add_body(doc, BODY[6])
    add_body(doc, BODY[7])

    add_heading(doc, "5.3 杜邦分析与经济增加值")
    add_body(doc, BODY[8])
    add_caption(doc, "表 5-4  名义口径与扣非口径的杜邦拆解〔13〕")
    add_table(doc, tables[3])
    add_caption(doc, "图 4  双轨杜邦三因子〔13〕")
    add_picture(doc, CHARTS[3])
    add_body(doc, BODY[9])
    add_body(doc, BODY[10])
    add_body(doc, BODY[11])
    add_caption(doc, "表 5-5  经济增加值测算〔6〕〔13〕")
    add_table(doc, tables[4])
    add_caption(doc, "图 5  经济增加值与投入资本〔6〕〔13〕")
    add_picture(doc, CHARTS[4])

    add_heading(doc, "5.4 现金流、盈利质量与平衡计分卡")
    add_body(doc, BODY[12])
    add_caption(doc, "表 5-6  现金流、周转与盈利含金量〔1〕〔5〕〔6〕〔13〕")
    add_table(doc, tables[5])
    add_caption(doc, "图 6  现金流与平衡计分卡〔6〕〔13〕")
    add_picture(doc, CHARTS[5])
    add_body(doc, BODY[13])
    add_body(doc, BODY[14])
    add_caption(doc, "表 5-7  平衡计分卡四维评分〔13〕")
    # 诊断列加宽，避免跨页时把一行拆成两截
    label, metric, score = 1500, 780, 900
    diagnosis = CONTENT_WIDTH - label - metric * 5 - score
    add_table(doc, tables[6], [label, metric, metric, metric, metric, score, diagnosis])
    add_body(doc, BODY[15])

    add_heading(doc, "注释与数据来源")
    for note, url in SOURCES:
        add_source(doc, note, url)
    return doc


def recount_docx(path: Path) -> dict:
    doc = Document(str(path))
    body_han = 0
    body_n = 0
    for p in doc.paragraphs:
        style = p.style.name if p.style is not None else ""
        text = p.text.strip()
        if not text:
            continue
        if style == "Normal" and not text.startswith("〔"):
            body_han += han_count(text)
            body_n += 1
    drawings = doc.element.body.findall(".//" + qn("w:drawing"))
    return {
        "body_han": body_han,
        "body_paragraphs": body_n,
        "tables": len(doc.tables),
        "drawings": len(drawings),
    }


def pdf_norm(path: Path) -> str:
    from pypdf import PdfReader

    text = "\n".join((page.extract_text() or "") for page in PdfReader(str(path)).pages)
    return re.sub(r"\s+", "", text).replace("—", "-").replace("–", "-").replace("−", "-")


def check_tables_against_pdf(pdf_text: str, tables: list) -> list[str]:
    missing = []
    for ti, table in enumerate(tables, start=1):
        for row in table[1:]:
            for cell in row:
                raw = cell.strip()
                if raw in {"", "—", "〃", "未并表"}:
                    continue
                token = raw.replace(" ", "").replace("—", "-").replace("–", "-").replace("−", "-")
                # 长句诊断格在 PDF 中可能被断行，改核对其内部数字
                if len(token) > 24:
                    for num in re.findall(r"-?\d+\.\d+%?|\d+%|\d+\.\d+", token):
                        if num.replace("+", "") not in pdf_text and num not in pdf_text:
                            missing.append(f"表5-{ti} 数字 {num} 不在原PDF")
                    continue
                if token.lstrip("+") not in pdf_text and token not in pdf_text:
                    missing.append(f"表5-{ti} 单元格缺失: {raw}")
    return missing


def main() -> int:
    count = han_count("".join(BODY))
    print(f"body_han_source={count}")
    if not (BODY_MIN <= count <= BODY_MAX):
        print(f"正文汉字 {count} 不在 {BODY_MIN}–{BODY_MAX}", file=sys.stderr)
        return 2
    for chart in CHARTS:
        if not chart.exists():
            print(f"缺少图 {chart}", file=sys.stderr)
            return 2
    tables = json.loads(TABLES_PATH.read_text(encoding="utf-8"))
    if PDF_PATH.exists():
        missing = check_tables_against_pdf(pdf_norm(PDF_PATH), tables)
        if missing:
            print("PDF核对失败", file=sys.stderr)
            print("\n".join(missing[:30]), file=sys.stderr)
            return 3
        print(f"pdf_cells_ok tables={len(tables)}")
    OUT_DOCX.parent.mkdir(parents=True, exist_ok=True)
    doc = build()
    doc.save(OUT_DOCX)
    stats = recount_docx(OUT_DOCX)
    print(json.dumps(stats, ensure_ascii=False))
    if stats["tables"] != 7 or stats["drawings"] != 6:
        print("图表数量不对", file=sys.stderr)
        return 4
    if not (BODY_MIN <= stats["body_han"] <= BODY_MAX):
        print("成稿正文字数越界", file=sys.stderr)
        return 4
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
