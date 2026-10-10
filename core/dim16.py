# -*- coding: utf-8 -*-
"""
16 维度同行对比分析（云端 / 无 GUI 版）。

移植自桌面版 FinancialAnalyzer + FinancialAnalysisApp，逻辑等价：
- `class FinancialAnalyzer`（桌面版 2130~2688 行）逐字移植：16 个 `analyze_dim_N()`、
  `get_val()`、`get_prev_yr()`、`get_dividend_rate()`、`build_df()`、`run_all_dimensions()`。
  判断规则、阈值、判定文案、输出列名
  `['年度','分析指标','公司','计算结果','得分','判定结果','投资建议','备注','原始数据']` 均未改动。
- `FinancialAnalysisApp.execute_all_folders()`（桌面版 5964~6175 行）的扫描/归类/表头探测/
  逐行业出表逻辑，拆分为 `scan_companies()` / `analyze_industry()` / `export_industry_excel()` /
  `run_all()` 四个纯函数；`format_excel_sheet()`（桌面版 2689~2715 行）等价移植为 `_format_excel_sheet()`。

去 UI 化改动（仅这些，且不触碰任何计算逻辑）：
1. 所有 `self.log(...)` / `messagebox` / `self.after(...)` / `threading` / `ttk` 调用改为
   `log(...)` 回调（经 `core.common.make_logger`，同时打到 stdout，云端可见）；
2. 桌面版用「全局工作目录下的 finance_analyzer_config_v5.json」（`load_config_b()`）取参数，
   云端改为由调用方传入 `cfg`（缺项按 `core.config.default_16dim_config()` 补齐，
   与桌面版 `load_config_b()` 的「FULL_CONFIG 打底 + 用户覆盖」语义一致）；
3. 桌面版把「单文件 / 单维度异常」交给外层 `except` 直接中断整轮分析，云端按云端容错要求改为
   「只记日志、跳过该项、继续跑其余行业/维度」（不改变任何成功路径上的结果）；
4. `start_process()` 的路径兜底（输出目录为空时取输入目录的上级）不再需要，由调用方显式给出。

输出 sheet 名与桌面版完全一致（共 17 个）：
    '0_综合得分排行' + '维度1_总资产分析' ... '维度16_分红情况'
另提供可选（默认关闭）的 `group_by_tab=True`：把 16 个维度按 `FULL_CONFIG[...]['tab']` 的
范围合并为 '维度 1-4' / '维度 5-8' / '维度 9-12' / '维度 13-16' 四个 sheet
（桌面版的 '维度 1-4' 等只是「参数配置界面」的分页名，其 16 维度**结果文件**始终是每维度一个 sheet；
 该分组为云端可选附加能力，默认不启用，以保证与桌面版逐格一致）。

输入容错（云端新增；对旧版文件（表头在第 2 行）的解析结果与桌面版逐格一致，已用黄金输出比对验证）：
5. 《统一整合输出_*.xlsx》两种版式都能吃下：
   A) 新版：第 1 行即表头（列如 ['项目','20221231',...]）；
   B) 旧版：第 1 行是合并标题（如“伊利股份 - 合并资产负债表”），第 2 行才是表头。
   表头定位顺序为「桌面版规则（前 15 行扫描）优先，探测不到时再退回解析版规则」，
   因此旧版文件走的仍是桌面版那条路径，逐格结果不变；新版文件即使年份列写成
   2024 / 2024.0 / 2024-12-31 / 2024年 也能被识别并统一成 f"{年份}1231"。
   （等效于「先用 header=0 读一次」的做法，但复用了桌面版同一套列归一化，不另开分支。）
6. 《分红情况》同一「年份」出现多行（中期分红 + 年度分红）时，`get_dividend_rate()`
   取「有值的行」累加；只有一行时与桌面版逐字一致（仍取首行、解析失败仍回落 0.0）。
   `分红率` 形如 '69.60%' 的字符串同样可解析。

依赖：标准库、pandas、numpy、openpyxl，以及本项目的 core.config / core.contracts / core.common。
"""
from __future__ import annotations

import os
import re
from typing import Any, Callable, Dict, List, Optional, Tuple

import pandas as pd

try:  # 作为包导入（from core import dim16）
    from . import contracts
    from .config import FULL_CONFIG, default_16dim_config
    from .common import make_logger
except ImportError:  # pragma: no cover - 兼容直接运行本文件 / 以 core 为顶层包导入
    import sys as _sys

    _sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
    from core import contracts  # type: ignore
    from core.config import FULL_CONFIG, default_16dim_config  # type: ignore
    from core.common import make_logger  # type: ignore

__all__ = [
    "FinancialAnalyzer",
    "scan_companies",
    "analyze_industry",
    "export_industry_excel",
    "run_all",
    "group_sheets_by_tab",
    "DIM_OUTPUT_COLUMNS",
    "DIMENSION_ORDER",
]

LogFn = Callable[[str], None]

#: 与原桌面版 `build_df()` 完全一致的输出列名
DIM_OUTPUT_COLUMNS: List[str] = ['年度', '分析指标', '公司', '计算结果', '得分', '判定结果', '投资建议', '备注',
                                 '原始数据']

#: 综合得分排行 sheet 名（桌面版 `final_results = {"0_综合得分排行": df_summary}`）
SUMMARY_SHEET = "0_综合得分排行"

#: 16 个维度的 sheet 名与计算方法名（顺序与桌面版 `run_all_dimensions()` 的字典顺序逐字一致）
DIMENSION_ORDER: List[Tuple[str, str]] = [
    ("维度1_总资产分析", "analyze_dim_1"),
    ("维度2_负债分析", "analyze_dim_2"),
    ("维度3_竞争优势分析", "analyze_dim_3"),
    ("维度4_产品竞争力", "analyze_dim_4"),
    ("维度5_维持竞争成本", "analyze_dim_5"),
    ("维度6_主业专注度", "analyze_dim_6"),
    ("维度7_爆雷风险", "analyze_dim_7"),
    ("维度8_行业地位", "analyze_dim_8"),
    ("维度9_毛利率分析", "analyze_dim_9"),
    ("维度10_成本管控", "analyze_dim_10"),
    ("维度11_销售难易度", "analyze_dim_11"),
    ("维度12_主营盈利能力", "analyze_dim_12"),
    ("维度13_利润质量", "analyze_dim_13"),
    ("维度14_整体盈利能力", "analyze_dim_14"),
    ("维度15_增长潜力", "analyze_dim_15"),
    ("维度16_分红情况", "analyze_dim_16"),
]

def _tab_ranges() -> List[Tuple[int, int, str]]:
    """从 `FULL_CONFIG` 的 `tab` 字段推导四表分组名（d1_* -> '维度 1-4' ...）。"""
    out: List[Tuple[int, int, str]] = []
    for lo, hi in ((1, 4), (5, 8), (9, 12), (13, 16)):
        tab = next((str(v.get("tab", "")) for k, v in FULL_CONFIG.items()
                    if re.match(rf"d{lo}_", k) and v.get("tab")), f"维度 {lo}-{hi}")
        out.append((lo, hi, tab))
    return out


#: 可选的「维度 1-4 / 5-8 / 9-12 / 13-16」四表分组（分组名取自 FULL_CONFIG 的 tab 字段）
TAB_RANGES: List[Tuple[int, int, str]] = _tab_ranges()


# =====================================================================
# 函数库：APP2核心 (原B程序) —— 逐字移植（桌面版 2130~2688 行）
# =====================================================================
class FinancialAnalyzer:
    def __init__(self, group_data, config):
        self.group_data = group_data
        self.companies = list(group_data.keys())
        self.cfg = config
        self.years = set()
        for comp in self.companies:
            for sheet, df in group_data[comp].items():
                for col in df.columns:
                    # 兼容类似 2023、2023年、20231231、2023-12-31 等各种格式
                    m = re.search(r'(20\d{2})', str(col).strip())
                    if m: self.years.add(m.group(1) + "年")
        self.years = sorted(list(self.years))

    def get_prev_yr(self, year):
        m = re.search(r'\d{4}', year)
        return year.replace(m.group(), str(int(m.group()) - 1)) if m else "未知上年"

    # 【核心修复1】：吸收游离在外部的 get_item_value，变为类内部方法
    def get_val(self, comp, year, sheets, items):
        df_dict = self.group_data.get(comp, {})
        if not df_dict: return 0.0

        target_df = None
        for sk in sheets:
            for k, df in df_dict.items():
                if sk in k and "合并" in k: target_df = df; break
            if target_df is not None: break
        if target_df is None:
            for sk in sheets:
                for k, df in df_dict.items():
                    if sk in k: target_df = df; break
                if target_df is not None: break

        if target_df is None or target_df.empty or not year: return 0.0

        year_num = str(year).replace('年', '').replace('度', '').strip()
        year_col = f"{year_num}1231"
        if year_col not in target_df.columns: return 0.0

        for item in items:
            item_clean = item.replace(' ', '').replace('帐', '账')
            matched = target_df[
                target_df['项目'].astype(str).str.replace(' ', '').str.replace('帐', '账') == item_clean]
            if matched.empty:
                matched = target_df[
                    target_df['项目'].astype(str).str.replace(' ', '').str.replace('帐', '账').str.contains(item_clean,
                                                                                                            na=False)]
            if not matched.empty:
                val = matched.iloc[0][year_col]
                try:
                    if isinstance(val, str): val = val.replace(',', '').replace('，', '').strip()
                    res = float(val)
                    if pd.isna(res): return 0.0
                    return res
                except:
                    pass
        return 0.0

    def get_dividend_rate(self, comp, year):
        df_dict = self.group_data.get(comp, {})
        for k, df in df_dict.items():
            if "分红" in k:
                if '年份' in df.columns and '分红率' in df.columns:
                    yr_str = str(year).replace('年', '').replace('度', '').strip()
                    matched = df[df['年份'].astype(str).str.replace('.0', '', regex=False) == yr_str]
                    if not matched.empty:
                        # 【云端容错】同一「年份」可能出现多行（中期分红 + 年度分红）。
                        # 单行时与桌面版逐字一致；多行时取「有值的行」累加（原桌面版只取第一行）。
                        if len(matched) > 1:
                            vals = []
                            for _, r in matched.iterrows():
                                v = r['分红率']
                                try:
                                    if isinstance(v, str): v = v.replace(',', '').replace('%', '').strip()
                                    fv = float(v)
                                    if pd.isna(fv): continue
                                    vals.append(fv)
                                except:
                                    continue
                            if vals:
                                return sum(vals)
                            continue
                        val = matched.iloc[0]['分红率']
                        try:
                            if isinstance(val, str): val = val.replace(',', '').replace('%', '').strip()
                            res = float(val)
                            if pd.isna(res): return 0.0
                            return res
                        except:
                            pass
        return 0.0

    def build_df(self, rows):
        return pd.DataFrame(rows,
                            columns=['年度', '分析指标', '公司', '计算结果', '得分', '判定结果', '投资建议', '备注',
                                     '原始数据'])

    def analyze_dim_1(self):
        rows = []
        for year in self.years:
            assets = {c: self.get_val(c, year, ["资产负债表"], ["资产总计", "总资产"]) for c in self.companies}
            max_asset = max(assets.values()) if assets.values() else 0
            for comp in self.companies:
                val = assets.get(comp, 0)
                score = max(0, self.cfg['d1_asset_score_base'] - round((max_asset - val) / max_asset * 10,
                                                                       2)) if max_asset > 0 else 0
                judge, adv = ("实力最强", "保留") if (max_asset > 0 and val == max_asset) else (
                    ("实力差", "淘汰") if (max_asset > 0 and val < max_asset * self.cfg['d1_asset_poor_ratio']) else (
                        "实力一般", "进一步分析"))
                rows.append([year, "总资产规模", comp, val, score, judge, adv, "", f"总资产:{val}"])
                prev_val = self.get_val(comp, self.get_prev_yr(year), ["资产负债表"], ["资产总计", "总资产"])
                if prev_val > 0:
                    rate = (val - prev_val) / prev_val
                    score2 = round(10 + (rate - 0.10) / 0.01, 2)
                    judge2, adv2 = ("成长性好", "保留") if rate >= 0.10 else ("成长性差", "淘汰")
                    rows.append([year, "总资产同比增长率", comp, f"{rate * 100:.2f}%", score2, judge2, adv2, "", ""])
                else:
                    rows.append([year, "总资产同比增长率", comp, "无上期数据", "-", "-", "-", "", ""])
        return self.build_df(rows)

    def analyze_dim_2(self):
        rows = []
        money_items = [x.strip() for x in self.cfg['f_dim2_money'].split(',') if x.strip()]
        debt_items = [x.strip() for x in self.cfg['f_dim2_debt'].split(',') if x.strip()]
        for year in self.years:
            for comp in self.companies:
                asset = self.get_val(comp, year, ["资产负债表"], ["资产总计", "总资产"])
                liab = self.get_val(comp, year, ["资产负债表"], ["负债合计", "总负债"])
                if asset > 0:
                    rate = liab / asset
                    score = round(10 + (0.60 - rate) / 0.06, 2)
                    adv = "保留" if rate < self.cfg['d2_debt_warn'] else "淘汰"
                    rows.append([year, "资产负债率", comp, f"{rate * 100:.2f}%", score, "评估完成", adv, "", ""])
                else:
                    rows.append([year, "资产负债率", comp, "缺失数据", "-", "-", "-", "", ""])
                money = sum([self.get_val(comp, year, ["资产负债表"], [x]) for x in money_items])
                debt = sum([self.get_val(comp, year, ["资产负债表"], [x]) for x in debt_items])
                diff = money - debt
                score2 = round(10.0 + (diff / (money if money > 0 else 1) * 10.0), 2)
                judge2, adv2 = ("没偿债风险", "保留") if diff >= 0 else ("有偿债风险", "淘汰")
                rows.append([year, "准货币资金减有息负债", comp, diff, score2, judge2, adv2, "", ""])
        return self.build_df(rows)

    def analyze_dim_3(self):
        rows = []
        pay_items = [x.strip() for x in self.cfg['f_dim3_pay'].split(',') if x.strip()]
        recv_items = [x.strip() for x in self.cfg['f_dim3_recv'].split(',') if x.strip()]
        for year in self.years:
            for comp in self.companies:
                pay_recv = sum([self.get_val(comp, year, ["资产负债表"], [x]) for x in pay_items])
                recv_pay = sum([self.get_val(comp, year, ["资产负债表"], [x]) for x in recv_items])
                diff = pay_recv - recv_pay
                score = round(self.cfg['d3_diff_score_base'] + (
                            diff / (pay_recv if pay_recv > 0 else 1) * self.cfg['d3_diff_score_step']), 2)
                judge, adv = ("没偿债风险", "保留") if diff >= 0 else ("有偿债风险", "淘汰")
                rows.append([year, "应付预收减应收预付差额", comp, diff, score, judge, adv, "",
                             f"应付预收:{pay_recv}, 应收预付:{recv_pay}"])
        return self.build_df(rows)

    def analyze_dim_4(self):
        rows = []
        for year in self.years:
            for comp in self.companies:
                ar = self.get_val(comp, year, ["资产负债表"], ["应收票据"]) + self.get_val(comp, year, ["资产负债表"],
                                                                                           ["应收账款"])
                ar += self.get_val(comp, year, ["资产负债表"], ["合同资产"])
                asset = self.get_val(comp, year, ["资产负债表"], ["资产总计", "总资产"])
                if asset > 0:
                    rate = ar / asset
                    score = round(10 + (0.15 - rate) / 0.015, 2)
                    if rate < self.cfg['d4_ar_1']:
                        judge = "最好公司，产品很畅销"
                    elif rate < self.cfg['d4_ar_2']:
                        judge = "优秀公司，产品畅销"
                    elif rate < self.cfg['d4_ar_3']:
                        judge = "一般公司，销售一般"
                    elif rate < self.cfg['d4_ar_4']:
                        judge = "较差公司，较难销"
                    else:
                        judge = "很差公司，很难销"
                    adv = "保留" if rate < self.cfg['d4_ar_out'] else "淘汰"
                    rows.append([year, "应收加合同占总资产", comp, f"{rate * 100:.2f}%", score, judge, adv, "",
                                 f"应收合同:{ar},总资产:{asset}"])
                else:
                    rows.append([year, "应收加合同占总资产", comp, "缺失数据", "-", "-", "-", "", ""])
        return self.build_df(rows)

    def analyze_dim_5(self):
        rows = []
        for year in self.years:
            for comp in self.companies:
                fix = self.get_val(comp, year, ["资产负债表"], ["固定资产"]) + self.get_val(comp, year, ["资产负债表"],
                                                                                            ["在建工程"])
                asset = self.get_val(comp, year, ["资产负债表"], ["资产总计", "总资产"])
                if asset > 0:
                    rate = fix / asset
                    score = round(10 + (0.40 - rate) / 0.04, 2)
                    if rate < self.cfg['d5_fix_1']:
                        judge = "轻资产型，风险较低"
                    elif rate < self.cfg['d5_fix_out']:
                        judge = "一般公司，风险一般"
                    elif rate < self.cfg['d5_fix_2']:
                        judge = "重资产型，风险较大"
                    else:
                        judge = "重资产型，风险很大"
                    adv = "保留" if rate < self.cfg['d5_fix_out'] else "淘汰"
                    rows.append([year, "固定及在建占总资产", comp, f"{rate * 100:.2f}%", score, judge, adv, "",
                                 f"固定在建:{fix},总资产:{asset}"])
                else:
                    rows.append([year, "固定及在建占总资产", comp, "缺失数据", "-", "-", "-", "", ""])
        return self.build_df(rows)

    def analyze_dim_6(self):
        rows = []
        inv_items = [x.strip() for x in self.cfg['f_dim6_inv'].split(',') if x.strip()]
        for year in self.years:
            for comp in self.companies:
                inv_assets = sum([self.get_val(comp, year, ["资产负债表"], [x]) for x in inv_items])
                asset = self.get_val(comp, year, ["资产负债表"], ["资产总计", "总资产"])
                if asset > 0:
                    rate = inv_assets / asset
                    score = round(10 + (0.10 - rate) / 0.01, 2)
                    if rate <= self.cfg['d6_inv_best']:
                        judge = "最优秀公司，最专注主业"
                    elif rate < self.cfg['d6_inv_out']:
                        judge = "优秀公司，专注主业"
                    else:
                        judge = "较差公司，不专注主业"
                    adv = "保留" if rate < self.cfg['d6_inv_out'] else "淘汰"
                    rows.append([year, "投产类资产占总资产比率", comp, f"{rate * 100:.2f}%", score, judge, adv, "",
                                 f"投产资产:{inv_assets},总资产:{asset}"])
                else:
                    rows.append([year, "投产类资产占总资产比率", comp, "缺失数据", "-", "-", "-", "", ""])
        return self.build_df(rows)

    def analyze_dim_7(self):
        rows = []
        for year in self.years:
            for comp in self.companies:
                ar = self.get_val(comp, year, ["资产负债表"], ["应收票据"]) + self.get_val(comp, year, ["资产负债表"],
                                                                                           ["应收账款"])
                inv = self.get_val(comp, year, ["资产负债表"], ["存货"])
                gw = self.get_val(comp, year, ["资产负债表"], ["商誉"])
                asset = self.get_val(comp, year, ["资产负债表"], ["资产总计", "总资产"])
                if asset > 0:
                    ar_r, inv_r, gw_r = ar / asset, inv / asset, gw / asset
                    score_ar = 5.0 + (0.05 - ar_r) / 0.005 * 0.5
                    score_inv = 5.0 + (0.15 - inv_r) / 0.015 * 0.5
                    judge1 = "爆雷风险较大" if (
                                ar_r >= self.cfg['d7_ar_risk'] and inv_r >= self.cfg['d7_inv_risk']) else "爆雷风险一般"
                    adv1 = "淘汰" if judge1 == "爆雷风险较大" else "保留"
                    score_total = round(score_ar + score_inv, 2)
                    rows.append(
                        [year, "应收和存货占比", comp, f"应收:{ar_r * 100:.2f}%,存货:{inv_r * 100:.2f}%", score_total,
                         judge1, adv1, f"收:{score_ar:.2f},存:{score_inv:.2f}", f"应收:{ar},存货:{inv},总资产:{asset}"])

                    score_gw = round(10 + (0.10 - gw_r) / 0.01, 2)
                    judge2, adv2 = ("爆雷风险较大", "淘汰") if gw_r >= self.cfg['d7_gw_risk'] else ("爆雷风险小",
                                                                                                    "保留")
                    rows.append([year, "商誉占总资产比率", comp, f"{gw_r * 100:.2f}%", score_gw, judge2, adv2, "",
                                 f"商誉:{gw},总资产:{asset}"])
                else:
                    rows.append([year, "应收和存货占比", comp, "缺失数据", "-", "-", "-", "", ""])
                    rows.append([year, "商誉占总资产比率", comp, "缺失数据", "-", "-", "-", "", ""])
        return self.build_df(rows)

    def analyze_dim_8(self):
        rows = []
        for year in self.years:
            revs = {c: self.get_val(c, year, ["利润表"], ["营业总收入", "营业收入"]) for c in self.companies}
            max_rev = max(revs.values()) if revs.values() else 0
            for comp in self.companies:
                val = revs.get(comp, 0)
                score = round(
                    max(0, self.cfg['d8_rev_score_base'] - (max_rev - val) / max_rev * self.cfg['d8_rev_score_step']),
                    2) if max_rev > 0 else 0
                if max_rev > 0 and val == max_rev:
                    judge, adv = "行业地位最高", "保留"
                elif max_rev > 0 and val <= max_rev * self.cfg['d8_rev_poor_ratio']:
                    judge, adv = "行业地位差", "淘汰"
                else:
                    judge, adv = "行业地位一般", "进一步分析"
                rows.append([year, "营业收入规模", comp, val, score, judge, adv, "", f"营收:{val}"])

                prev_val = self.get_val(comp, self.get_prev_yr(year), ["利润表"], ["营业总收入", "营业收入"])
                if prev_val > 0:
                    rate = (val - prev_val) / prev_val
                    score2 = round(10 + (rate - 0.10) / 0.01, 2)
                    judge2, adv2 = ("成长性好", "保留") if rate > self.cfg['d8_growth_good'] else ("成长性差", "淘汰")
                    rows.append([year, "营业收入同比增长率", comp, f"{rate * 100:.2f}%", score2, judge2, adv2, "",
                                 f"本期:{val},上期:{prev_val}"])
                else:
                    rows.append([year, "营业收入同比增长率", comp, "无上期数据", "-", "-", "-", "", ""])
        return self.build_df(rows)

    def analyze_dim_9(self):
        rows = []
        for year in self.years:
            for comp in self.companies:
                rev = self.get_val(comp, year, ["利润表"], ["营业收入"])
                cost = self.get_val(comp, year, ["利润表"], ["营业成本"])
                if cost > 0:
                    gm = (rev - cost) / rev
                    score = round(10 + (gm - 0.40) / 0.04, 2)
                    judge, adv = ("产品竞争力较强", "保留") if gm > self.cfg['d9_gm_out'] else ("产品竞争力较差",
                                                                                                "淘汰")
                    rows.append(
                        [year, "毛利率", comp, f"{gm * 100:.2f}%", score, judge, adv, "", f"收入:{rev},成本:{cost}"])

                    prev_rev = self.get_val(comp, self.get_prev_yr(year), ["利润表"], ["营业收入"])
                    prev_cost = self.get_val(comp, self.get_prev_yr(year), ["利润表"], ["营业成本"])
                    if prev_cost > 0:
                        prev_gm = (prev_rev - prev_cost) / prev_rev
                        vol = abs(gm - prev_gm) / prev_gm if prev_gm != 0 else 0
                        score2 = round(10 + (0.20 - vol) / 0.02, 2)
                        if vol >= self.cfg['d9_gm_vol_out']:
                            judge2, adv2 = "公司经营风险大", "淘汰"
                        elif vol < self.cfg['d9_gm_vol_safe']:
                            judge2, adv2 = "优秀公司的正常波动", "保留"
                        else:
                            judge2, adv2 = "公司经营风险小", "保留"
                        rows.append([year, "毛利率波动幅度", comp, f"{vol * 100:.2f}%", score2, judge2, adv2, "",
                                     f"本期GM:{gm * 100:.1f}%,上期GM:{prev_gm * 100:.1f}%"])
                    else:
                        rows.append([year, "毛利率波动幅度", comp, "无上期数据", "-", "-", "-", "", ""])
                else:
                    rows.append([year, "毛利率", comp, "缺失数据", "-", "-", "-", "", ""])
                    rows.append([year, "毛利率波动幅度", comp, "缺失数据", "-", "-", "-", "", ""])
        return self.build_df(rows)

    def analyze_dim_10(self):
        rows = []
        for year in self.years:
            for comp in self.companies:
                rev = self.get_val(comp, year, ["利润表"], ["营业收入"])
                cost = self.get_val(comp, year, ["利润表"], ["营业成本"])
                # 【核心修复2】：将 safe_div 彻底替换为安全的内联计算
                gm = (rev - cost) / rev if rev != 0 else 0.0

                exp = sum([self.get_val(comp, year, ["利润表"], [x]) for x in ["销售费用", "管理费用", "研发费用"]])
                fin_exp = self.get_val(comp, year, ["利润表"], ["财务费用"])
                if fin_exp > 0: exp += fin_exp

                if rev > 0 and gm > 0:
                    ratio = exp / rev
                    res = ratio / gm
                    score = round(10 + (0.60 - res) / 0.06, 2)
                    if res < self.cfg['d10_exp_gm_safe']:
                        judge, adv = "成本控制能力优秀", "保留"
                    elif res <= self.cfg['d10_exp_gm_out']:
                        judge, adv = "成本控制能力一般", "保留"
                    else:
                        judge, adv = "成本控制能力较差", "淘汰"
                    rows.append([year, "期间费用率/毛利率", comp, f"{res * 100:.2f}%", score, judge, adv, "",
                                 f"费用:{exp},营收:{rev},GM:{gm * 100:.1f}%"])
                else:
                    rows.append([year, "期间费用率/毛利率", comp, "缺失数据", "-", "-", "-", "", ""])
        return self.build_df(rows)

    def analyze_dim_11(self):
        rows = []
        for year in self.years:
            for comp in self.companies:
                sales_exp = self.get_val(comp, year, ["利润表"], ["销售费用"])
                rev = self.get_val(comp, year, ["利润表"], ["营业收入"])
                if rev > 0:
                    ratio = sales_exp / rev
                    score = round(10 + (0.30 - ratio) / 0.03, 2)
                    if ratio < self.cfg['d11_sales_exp_1']:
                        judge, adv = "产品比较容易销售，风险较小", "保留"
                    elif ratio <= self.cfg['d11_sales_exp_out']:
                        judge, adv = "产品销售难度一般，风险一般", "保留"
                    else:
                        judge, adv = "产品销售难度大，风险大", "淘汰"
                    rows.append([year, "销售费用率", comp, f"{ratio * 100:.2f}%", score, judge, adv, "",
                                 f"销售费用:{sales_exp},营收:{rev}"])

                    prev_exp = self.get_val(comp, self.get_prev_yr(year), ["利润表"], ["销售费用"])
                    prev_rev = self.get_val(comp, self.get_prev_yr(year), ["利润表"], ["营业收入"])
                    if prev_rev > 0:
                        prev_ratio = prev_exp / prev_rev
                        trend = ratio - prev_ratio
                        score2 = round(
                            self.cfg['d11_trend_score_base'] + (self.cfg['d11_trend_out'] - trend) / self.cfg[
                                'd11_trend_score_step'], 2)
                        if trend == self.cfg['d11_trend_out']:
                            judge2, adv2 = "产品销售难度不变", "保留"
                        elif trend < self.cfg['d11_trend_out']:
                            judge2, adv2 = "产品销售难度在下降", "保留"
                        else:
                            judge2, adv2 = "产品销售难度在上升", "淘汰"
                        rows.append([year, "销售费用率变动趋势", comp, f"{trend * 100:.2f}%", score2, judge2, adv2, "",
                                     f"本期率:{ratio * 100:.1f}%,上期率:{prev_ratio * 100:.1f}%"])
                    else:
                        rows.append([year, "销售费用率变动趋势", comp, "无上期数据", "-", "-", "-", "", ""])
                else:
                    rows.append([year, "销售费用率", comp, "缺失数据", "-", "-", "-", "", ""])
                    rows.append([year, "销售费用率变动趋势", comp, "缺失数据", "-", "-", "-", "", ""])
        return self.build_df(rows)

    def analyze_dim_12(self):
        rows = []
        op_plus_items = [x.strip() for x in self.cfg['f_dim12_op_plus'].split(',') if x.strip()]
        op_minus_items = [x.strip() for x in self.cfg['f_dim12_op_minus'].split(',') if x.strip()]
        for year in self.years:
            for comp in self.companies:
                rev = self.get_val(comp, year, ["利润表"], ["营业收入"])
                cost = self.get_val(comp, year, ["利润表"], ["营业成本"])
                tax = self.get_val(comp, year, ["利润表"], ["税金及附加"])
                exp = sum([self.get_val(comp, year, ["利润表"], [x]) for x in
                           ["销售费用", "管理费用", "研发费用", "财务费用"]])
                core_profit = rev - cost - tax - exp
                if rev > 0:
                    rate1 = core_profit / rev
                    score1 = round(10 + (rate1 - 0.15) / 0.015, 2)
                    judge1, adv1 = ("主业盈利能力强", "保留") if rate1 >= self.cfg['d12_core_out'] else (
                        "主业盈利能力弱", "淘汰")
                    rows.append([year, "主营利润率", comp, f"{rate1 * 100:.2f}%", score1, judge1, adv1, "",
                                 f"主营利润:{core_profit},营收:{rev}"])
                else:
                    rows.append([year, "主营利润率", comp, "缺失数据", "-", "-", "-", "", ""])

                op_plus_sum = sum([self.get_val(comp, year, ["利润表"], [x]) for x in op_plus_items])
                op_minus_sum = sum([self.get_val(comp, year, ["利润表"], [x]) for x in op_minus_items])
                op_profit_custom = op_plus_sum - op_minus_sum
                if op_profit_custom != 0:
                    rate2 = core_profit / op_profit_custom
                    score2 = round(10 + (rate2 - 0.80) / 0.08, 2)
                    judge2, adv2 = ("利润质量高", "保留") if rate2 > self.cfg['d12_profit_out'] else ("利润质量低",
                                                                                                      "淘汰")
                    rows.append(
                        [year, "主营利润/营业利润(自定义)", comp, f"{rate2 * 100:.2f}%", score2, judge2, adv2, "",
                         f"主营利润:{core_profit},计算分母:{op_profit_custom}"])
                else:
                    rows.append([year, "主营利润/营业利润(自定义)", comp, "缺失数据", "-", "-", "-", "", ""])
        return self.build_df(rows)

    def analyze_dim_13(self):
        rows = []
        for year in self.years:
            for comp in self.companies:
                non_op_in = self.get_val(comp, year, ["利润表"], ["营业外收入"])
                non_op_out = self.get_val(comp, year, ["利润表"], ["营业外支出"])
                op_profit = self.get_val(comp, year, ["利润表"], ["营业利润"])
                net_non_op = non_op_in - non_op_out
                total = op_profit + net_non_op
                if total != 0:
                    rate = net_non_op / total
                    score = round(10 + (0.05 - rate) / 0.005, 2)
                    judge, adv = ("利润质量优秀", "保留") if rate < self.cfg['d13_non_op_out'] else ("利润质量较差",
                                                                                                     "淘汰")
                    rows.append([year, "营业外净额/利润总额", comp, f"{rate * 100:.2f}%", score, judge, adv, "",
                                 f"营业外净额:{net_non_op},总额:{total}"])
                else:
                    rows.append([year, "营业外净额/利润总额", comp, "缺失数据", "-", "-", "-", "", ""])
        return self.build_df(rows)

    def analyze_dim_14(self):
        rows = []
        for year in self.years:
            nps = {c: self.get_val(c, year, ["利润表"],
                                   ["归属于母公司所有者的净利润", "归属于母公司股东的净利润", "净利润"]) for c in
                   self.companies}
            max_np = max(nps.values()) if nps.values() else 0
            for comp in self.companies:
                val = nps.get(comp, 0)
                score = round(
                    max(0, self.cfg['d14_np_score_base'] - (max_np - val) / max_np * self.cfg['d14_np_score_step']),
                    2) if max_np > 0 else 0
                if max_np > 0 and val == max_np:
                    judge, adv = "盈利能力最强", "保留"
                elif max_np > 0 and val <= max_np * self.cfg['d14_np_poor_ratio']:
                    judge, adv = "盈利能力差", "淘汰"
                else:
                    judge, adv = "盈利能力一般", "进一步分析"
                rows.append([year, "归母净利润规模", comp, val, score, judge, adv, "", f"归母净利润:{val}"])

                prev_val = self.get_val(comp, self.get_prev_yr(year), ["利润表"],
                                        ["归属于母公司所有者的净利润", "归属于母公司股东的净利润", "净利润"])
                if prev_val != 0:
                    rate = (val - prev_val) / abs(prev_val)
                    score2 = round(10 + (rate - 0.10) / 0.001, 2)
                    if rate >= self.cfg['d14_np_growth']:
                        judge2, adv2 = "盈利不但强且持续性较好", "保留"
                    elif rate >= self.cfg['d14_np_growth_out']:
                        judge2, adv2 = "盈利持续性差", "淘汰"
                    else:
                        judge2, adv2 = "已处于衰落中，持续性较差", "淘汰"
                    rows.append([year, "归母净利增长率", comp, f"{rate * 100:.2f}%", score2, judge2, adv2, "",
                                 f"本期:{val},上期:{prev_val}"])
                else:
                    rows.append([year, "归母净利增长率", comp, "无上期数据", "-", "-", "-", "", ""])
        return self.build_df(rows)

    def analyze_dim_15(self):
        rows = []
        for year in self.years:
            for comp in self.companies:
                capex = self.get_val(comp, year, ["现金流量表"], ["购建固定资产、无形资产和其他长期资产支付的现金"])
                cfo = self.get_val(comp, year, ["现金流量表"], ["经营活动产生的现金流量净额"])
                if cfo > 0:
                    rate1 = capex / cfo
                    if rate1 < 0.03:
                        score1 = round(10 - (0.03 - rate1) / 0.003, 2)
                    elif rate1 <= 1.0:
                        score1 = round(10 + (1.0 - rate1) / 0.10, 2)
                    else:
                        score1 = round(10 - (rate1 - 1.0) / 0.10, 2)

                    if rate1 < self.cfg['d15_capex_slow']:
                        judge1, adv1 = "成长太慢，回报较低", "淘汰"
                    elif rate1 < self.cfg['d15_capex_safe']:
                        judge1, adv1 = "增长潜力较大并且风险相对较小", "保留"
                    elif rate1 < self.cfg['d15_capex_out']:
                        judge1, adv1 = "扩张略高，风险一般", "保留"
                    else:
                        judge1, adv1 = "扩张激进，风险较大", "淘汰"
                    rows.append([year, "CAPEX/CFO", comp, f"{rate1 * 100:.2f}%", score1, judge1, adv1, "",
                                 f"CAPEX:{capex},CFO:{cfo}"])
                else:
                    rows.append([year, "CAPEX/CFO", comp, "缺失数据", "-", "-", "-", "", ""])

                sales_cash = self.get_val(comp, year, ["现金流量表"], ["销售商品、提供劳务收到的现金"])
                prev_cash = self.get_val(comp, self.get_prev_yr(year), ["现金流量表"], ["销售商品、提供劳务收到的现金"])
                prev_np = self.get_val(comp, self.get_prev_yr(year), ["利润表"],
                                       ["归属于母公司所有者的净利润", "归属于母公司股东的净利润", "净利润"])
                if prev_np > 0:
                    trend = (sales_cash - prev_cash) / prev_np
                    score2 = round(self.cfg['d15_cash_score_base'] + trend / self.cfg['d15_cash_score_step'], 2)
                    judge2, adv2 = ("有增长潜力", "保留") if trend > self.cfg['d15_cash_trend_out'] else (
                        "没有增长潜力", "淘汰")
                    rows.append([year, "销售现金变动趋势", comp, f"{trend * 100:.2f}%", score2, judge2, adv2, "",
                                 f"本期现金:{sales_cash},上期现金:{prev_cash},上期净利:{prev_np}"])
                else:
                    rows.append([year, "销售现金变动趋势", comp, "缺失数据", "-", "-", "-", "", ""])
        return self.build_df(rows)

    def analyze_dim_16(self):
        rows = []
        for year in self.years:
            for comp in self.companies:
                rate = self.get_dividend_rate(comp, year)
                if rate > 0:
                    rate_dec = rate / 100 if rate > 1 else rate
                    if rate_dec >= 0.70:
                        score = round(10 - (rate_dec - 0.70) / 0.07, 2)
                    elif rate_dec >= 0.30:
                        score = round(10 + (rate_dec - 0.30) / 0.03, 2)
                    else:
                        score = round(10 - (0.30 - rate_dec) / 0.03, 2)

                    if rate_dec < self.cfg['d16_div_out_low']:
                        judge, adv = "不够厚道", "淘汰"
                    elif rate_dec <= self.cfg['d16_div_out_high']:
                        judge, adv = "厚道且持续性强", "保留"
                    else:
                        judge, adv = "难以持续", "淘汰"
                    rows.append(
                        [year, "分红率", comp, f"{rate_dec * 100:.2f}%", score, judge, adv, "", f"提取分红率:{rate}"])
                else:
                    rows.append([year, "分红率", comp, "缺失数据", "-", "-", "-", "", ""])
        return self.build_df(rows)

    def run_all_dimensions(self):
        return {
            "维度1_总资产分析": self.analyze_dim_1(), "维度2_负债分析": self.analyze_dim_2(),
            "维度3_竞争优势分析": self.analyze_dim_3(), "维度4_产品竞争力": self.analyze_dim_4(),
            "维度5_维持竞争成本": self.analyze_dim_5(), "维度6_主业专注度": self.analyze_dim_6(),
            "维度7_爆雷风险": self.analyze_dim_7(), "维度8_行业地位": self.analyze_dim_8(),
            "维度9_毛利率分析": self.analyze_dim_9(), "维度10_成本管控": self.analyze_dim_10(),
            "维度11_销售难易度": self.analyze_dim_11(), "维度12_主营盈利能力": self.analyze_dim_12(),
            "维度13_利润质量": self.analyze_dim_13(), "维度14_整体盈利能力": self.analyze_dim_14(),
            "维度15_增长潜力": self.analyze_dim_15(), "维度16_分红情况": self.analyze_dim_16(),
        }


# =====================================================================
# 内部工具：配置合并 / 维度循环 / 表头探测（等价移植）
# =====================================================================
def _merge_cfg(cfg: Optional[Dict[str, Any]]) -> Dict[str, Any]:
    """与桌面版 `load_config_b()` 同语义：FULL_CONFIG 的 val 打底，用户配置覆盖。"""
    merged = default_16dim_config()
    if cfg:
        merged.update(cfg)
    return merged


def _run_dimensions(analyzer: FinancialAnalyzer, log: LogFn) -> Dict[str, pd.DataFrame]:
    """按桌面版 `run_all_dimensions()` 的同一顺序逐个维度计算。

    与桌面版唯一差别：单个维度抛异常时只记日志并跳过该维度（云端容错要求），
    成功路径上返回的 sheet 名、顺序、内容与 `run_all_dimensions()` 完全一致。
    """
    results: Dict[str, pd.DataFrame] = {}
    for sheet_name, method_name in DIMENSION_ORDER:
        try:
            results[sheet_name] = getattr(analyzer, method_name)()
        except Exception as e:  # noqa: BLE001
            log(f"  ⚠️ 维度 [{sheet_name}] 计算异常，已跳过: {type(e).__name__}: {str(e)[:150]}")
    return results


def _detect_header(df_temp: pd.DataFrame) -> int:
    """桌面版 execute_all_folders 的表头行探测（逐字移植）。返回 -1 表示未找到。"""
    header_idx = -1
    for i in range(min(15, len(df_temp))):
        row_vals = [str(x) for x in df_temp.iloc[i].values]
        valid_count = sum(1 for x in row_vals if x not in ['nan', 'None', ''])
        if valid_count >= 2 and ('报表日期' in row_vals or '年份' in row_vals or any(
                '1231' in str(x) or '12-31' in str(x) or '年度' in str(x) for x in row_vals)):
            header_idx = i
            break
    return header_idx


def _detect_header_relaxed(df_temp: pd.DataFrame) -> int:
    """云端容错（桌面版没有）：识别「表头就在第 1 行」的新版《统一整合输出》。

    仅当桌面版规则 `_detect_header()` 探测不到表头行时才会被调用，命中后仍走
    与桌面版完全相同的列归一化路径，因此不会改变任何旧版文件的解析结果。

    命中条件（第 1 行的非空单元格 ≥2 且满足其一）：
      - 出现 `项目` / `报表日期` / `年份` / `年度` 这类关键字列名；
      - 出现「年份列」：整格就是 20XX、20XX1231、20XX-12-31、20XX/12/31、20XX年 等
        （含 Excel 把整数年份读成 2024.0 的情形）。
    年份列用「整格匹配」而非「包含匹配」，避免把 202345 这类大额数据误判成年份。
    """
    if len(df_temp) == 0:
        return -1
    row_vals = [str(x) for x in df_temp.iloc[0].values]
    valid_count = sum(1 for x in row_vals if x not in ['nan', 'None', ''])
    if valid_count < 2:
        return -1
    for x in row_vals:
        s = str(x).strip()
        if not s or s in ('nan', 'None'):
            continue
        if any(mk in s for mk in ('项目', '报表日期', '年份', '年度')):
            return 0
        s_num = s[:-2] if s.endswith('.0') else s
        s_clean = re.sub(r'[-/.年月日\s]', '', s_num)
        if re.fullmatch(r'20\d{2}(\d{4})?', s_clean):
            return 0
    return -1


def _normalize_sheet(df_temp: pd.DataFrame) -> Optional[pd.DataFrame]:
    """桌面版表头归一化（逐字移植）：年份列 -> f"{year}1231"，首列 -> '项目'。

    表头行的定位顺序：先按桌面版规则（前 15 行扫描），探测不到时再退回
    `_detect_header_relaxed()`（兼容「表头在第 1 行且年份列不是 YYYYMMDD」的新版文件）。
    """
    header_idx = _detect_header(df_temp)
    if header_idx == -1:
        header_idx = _detect_header_relaxed(df_temp)
    if header_idx == -1:
        return None

    cols = [str(x).replace('.0', '').strip() for x in df_temp.iloc[header_idx].values]
    new_cols = []
    for c in cols:
        m = re.search(r'(20\d{2})', str(c))
        if m:
            new_cols.append(m.group(1) + "1231")
        else:
            new_cols.append(str(c))
    cols = new_cols

    if cols[0] in ['报表日期', 'nan', 'None']:
        cols[0] = '项目'
    elif '报表日期' in cols:
        cols[cols.index('报表日期')] = '项目'

    df_temp.columns = cols
    df = df_temp.iloc[header_idx + 1:].reset_index(drop=True)
    if '项目' not in df.columns:
        cols = list(df.columns)
        cols[0] = '项目'
        df.columns = cols
    return df


def _read_sheets(file_path: str, fname: str) -> Dict[str, pd.DataFrame]:
    """读单个数据文件（header=None 原始读取），与桌面版分支一致。

    与桌面版唯一差别：用 `with pd.ExcelFile(...)` 关闭文件句柄
    （云端长时间运行/删除临时文件时避免 Windows 文件占用，读取结果不变）。
    """
    if fname.endswith(".csv"):
        return {fname.split('.')[0]: pd.read_csv(file_path, encoding='utf-8-sig', header=None)}
    with pd.ExcelFile(file_path) as xls:
        return {sheet: pd.read_excel(xls, sheet_name=sheet, header=None) for sheet in xls.sheet_names}


def _scan_industry_rep_map(input_dir: str) -> Dict[str, str]:
    """桌面版第一遍扫描里维护的 industry_rep_map（按文件夹名 '{代表公司}_行业_{行业名}'）。"""
    rep_map: Dict[str, str] = {}
    for root_dir, dirs, files in os.walk(input_dir):
        if any(ex in root_dir for ex in contracts.SCAN_EXCLUDE_DIRS):
            continue
        folder_name = os.path.basename(root_dir)
        if "_行业_" in folder_name:
            ind = folder_name.split("_行业_")[-1].strip()
            rep = folder_name.split("_行业_")[0].split("_")[-1].strip()
            if ind and ind not in rep_map:
                rep_map[ind] = rep
    return rep_map


# =====================================================================
# 一、扫描：等价于桌面版 execute_all_folders 的两遍扫描 + 表头探测
# =====================================================================
def scan_companies(input_dir: str,
                   log: Optional[Callable[[str], None]] = None
                   ) -> Dict[str, Dict[str, Dict[str, pd.DataFrame]]]:
    """全域扫描输入目录，返回 {行业: {公司名: {sheet名: DataFrame}}}。

    与桌面版 execute_all_folders 内部构建的 all_data 完全一致：
    - 跳过 造假排雷结果/16维度对比结果/文本分析/企业分析/好价分析/海选情况/报表下载 等目录；
    - 跳过 同行排列表/公司属性表/年报提取表/多年度汇总与增长率分析/排雷报告/海选公司汇总表 等文件；
    - 公司名/代码取自文件名 `(.*?)[(（](\\d{6})[)）]`，行业优先取《公司属性表》的二级行业，
      否则取所在文件夹 `_行业_` 后的行业名，再兜底 "默认行业"；
    - 每个 sheet 自动探测表头行（前 15 行内；兼容「表头在第 1 行」的新版与
      「第 1 行合并标题、第 2 行表头」的旧版两种版式），
      年份列统一成 f"{年份}1231"，首列统一成 '项目'。
    """
    _log = make_logger(log)

    if not os.path.exists(input_dir):
        _log(f"⚠️ 输入源目录不存在，跳过分析: {input_dir}")
        return {}

    _log("📦 正在全域穿透扫描提取数据，并按【行业】进行归类...")

    all_data: Dict[str, Dict[str, Dict[str, pd.DataFrame]]] = {}
    company_industry_map: Dict[str, str] = {}

    # ---- 第一遍：读《公司属性表》，建立 代码 -> 二级行业 映射 ----
    for root_dir, dirs, files in os.walk(input_dir):
        if any(ex in root_dir for ex in contracts.SCAN_EXCLUDE_DIRS):
            continue

        for f in files:
            if "公司属性表" in f and not f.startswith("~"):
                file_path = os.path.join(root_dir, f)
                try:
                    df_attr = pd.read_excel(file_path) if f.endswith(".xlsx") else pd.read_csv(file_path,
                                                                                              encoding='gbk')
                    for _, row in df_attr.iterrows():
                        c_code = str(row.get("目标股票", "")).strip()
                        if c_code:
                            # clean_str：NaN → 空串，否则行业会变成字符串 "nan"（报告名出现 _nan_）
                            try:
                                from .universe import clean_str
                            except Exception:  # noqa: BLE001
                                clean_str = lambda x: "" if x is None else str(x).strip()  # noqa: E731
                            c_ind = clean_str(row.get("二级行业")) or clean_str(row.get("一级行业"))
                            c_ind = c_ind.replace("Ⅱ", "").replace("Ⅰ", "").strip()
                            if c_ind:
                                company_industry_map[c_code] = c_ind
                except Exception as e:  # noqa: BLE001
                    _log(f"  ⚠️ 公司属性表读取失败，已跳过: {f} ({type(e).__name__}: {str(e)[:120]})")

    # ---- 第二遍：读全部数据文件，按行业/公司/sheet 归类 ----
    for root_dir, dirs, files in os.walk(input_dir):
        if any(ex in root_dir for ex in contracts.SCAN_EXCLUDE_DIRS):
            continue

        folder_name = os.path.basename(root_dir)
        folder_ind = "默认行业"
        if "_行业_" in folder_name:
            folder_ind = folder_name.split("_行业_")[-1].strip()

        for f in files:
            if not (f.endswith(".xlsx") or f.endswith(".csv")) or f.startswith("~"): continue

            if any(x in f for x in contracts.SCAN_EXCLUDE_FILE_KEYS):
                continue

            file_path = os.path.join(root_dir, f)

            match = re.search(r'(.*?)[(（](\d{6})[)）]', f)
            comp_code = ""
            if match:
                comp_name = match.group(1).split('_')[-1].strip()
                comp_code = match.group(2)
            else:
                parts = f.split('.')[0].split('-')[0].split('_')
                comp_name = parts[0]
                for p in parts:
                    p = p.strip()
                    if not p.isdigit() and p not in ["统一整合输出", "宽表数据", "合并数据", "财报"]:
                        comp_name = p
                        break

            if not comp_name or comp_name.isdigit(): continue

            industry = company_industry_map.get(comp_code, folder_ind)
            if not industry: industry = "默认行业"

            if industry not in all_data: all_data[industry] = {}
            if comp_name not in all_data[industry]: all_data[industry][comp_name] = {}

            try:
                sheet_dict = _read_sheets(file_path, f)

                for sheet, df_temp in sheet_dict.items():
                    try:
                        df = _normalize_sheet(df_temp)
                    except Exception as e:  # noqa: BLE001
                        _log(f"  ⚠️ sheet 解析失败，已跳过: {f} / {sheet} ({type(e).__name__}: {str(e)[:120]})")
                        continue
                    if df is None:
                        continue
                    all_data[industry][comp_name][sheet] = df
            except Exception as e:  # noqa: BLE001
                _log(f"  ⚠️ 文件读取失败，已跳过: {f} ({type(e).__name__}: {str(e)[:120]})")

    if not all_data:
        _log("⚠️ 未提取到有效数据，已跳过。")
    else:
        _log(f"✅ 成功提取并归类到 {len(all_data)} 个行业，即将分行业出具报告...")
    return all_data


# =====================================================================
# 二、单行业 16 维度分析：等价于桌面版逐行业 results/summary/rep 逻辑
# =====================================================================
def _analyze_industry_full(industry: str,
                           group_data: Dict[str, Dict[str, pd.DataFrame]],
                           cfg: Optional[Dict[str, Any]],
                           log: LogFn,
                           rep_hint: str = "") -> Tuple[Dict[str, pd.DataFrame], str]:
    """返回 ({sheet名: DataFrame}, 代表公司名)。桌面版 execute_all_folders 的行业循环体。"""
    if not group_data:
        return {}, ""

    analyzer = FinancialAnalyzer(group_data, _merge_cfg(cfg))
    if not analyzer.years:
        log(f"⚠️ 行业 [{industry}] 缺乏符合财务标准的连续年份，跳过分析。")
        return {}, ""

    log(f"⚙️ 正在分析行业: {industry} (参战公司数: {len(group_data)})")
    results = _run_dimensions(analyzer, log)

    total_scores, elimination_counts = {}, {}
    for dim_name, df in results.items():
        if '公司' in df.columns:
            for _, row in df.iterrows():
                comp = row['公司']
                if '得分' in df.columns and row['得分'] != '-':
                    try:
                        total_scores[comp] = total_scores.get(comp, 0.0) + float(row['得分'])
                    except:
                        pass
                if '投资建议' in df.columns and pd.notna(row['投资建议']) and "淘汰" in str(
                        row['投资建议']):
                    elimination_counts[comp] = elimination_counts.get(comp, 0) + 1

    summary_rows = [[comp, total_scores.get(comp, 0.0), elimination_counts.get(comp, 0)] for comp in
                    total_scores.keys()]
    df_summary = pd.DataFrame(summary_rows, columns=["公司", "综合总得分", "淘汰总次数"]).sort_values(
        by="综合总得分", ascending=False)

    final_results: Dict[str, pd.DataFrame] = {SUMMARY_SHEET: df_summary}
    final_results.update(results)

    rep_comp = rep_hint
    if not rep_comp or rep_comp not in group_data:
        if not df_summary.empty:
            rep_comp = df_summary.iloc[0]["公司"]
        else:
            rep_comp = list(group_data.keys())[0]

    return final_results, str(rep_comp)


def analyze_industry(industry: str, group_data: Dict[str, Dict[str, pd.DataFrame]],
                     cfg: Dict[str, Any],
                     log: Optional[Callable[[str], None]] = None,
                     group_by_tab: bool = False) -> Dict[str, pd.DataFrame]:
    """对单个行业跑 16 个维度，返回 {sheet名: DataFrame}。

    默认（group_by_tab=False，与桌面版一致）：
        {'0_综合得分排行', '维度1_总资产分析' ... '维度16_分红情况'} 共 17 个 sheet。
    group_by_tab=True：额外（可选，桌面版无此行为）按 FULL_CONFIG 的 tab 合并为
        '0_综合得分排行' + '维度 1-4' / '维度 5-8' / '维度 9-12' / '维度 13-16' 共 5 个 sheet。
    """
    _log = make_logger(log)
    sheets, _rep = _analyze_industry_full(industry, group_data, cfg, _log)
    if not sheets:
        return {}
    return group_sheets_by_tab(sheets, log=_log) if group_by_tab else sheets


def group_sheets_by_tab(sheets: Dict[str, pd.DataFrame],
                        log: Optional[Callable[[str], None]] = None) -> Dict[str, pd.DataFrame]:
    """可选附加能力：把 16 个维度按 '维度 1-4 / 5-8 / 9-12 / 13-16' 合并为 4 个 sheet。

    桌面版 16 维度结果文件是「每维度一个 sheet」；'维度 1-4' 等只是参数配置界面的分页名。
    本函数仅供云端需要四表视图时显式调用，默认流程不使用。
    """
    def _log(msg: str) -> None:  # 不再二次包装，避免调用方已包装的 logger 被重复打印
        if log is not None:
            try:
                log(msg)
            except Exception:
                pass

    if SUMMARY_SHEET in sheets:
        out: Dict[str, pd.DataFrame] = {SUMMARY_SHEET: sheets[SUMMARY_SHEET]}
    else:
        out = {}
    for lo, hi, tab in TAB_RANGES:
        parts: List[pd.DataFrame] = []
        for n in range(lo, hi + 1):
            name = next((sn for sn, _m in DIMENSION_ORDER if sn.startswith(f"维度{n}_")), None)
            if name and name in sheets:
                parts.append(sheets[name])
        if parts:
            out[tab] = pd.concat(parts, ignore_index=True) if len(parts) > 1 else parts[0].copy()
        else:
            _log(f"  ⚠️ 分组 [{tab}] 无可用维度数据")
    return out


# =====================================================================
# 三、出表：等价于桌面版 format_excel_sheet + pd.ExcelWriter(openpyxl)
# =====================================================================
def _format_excel_sheet(ws, is_summary: bool = False) -> None:
    """桌面版 format_excel_sheet（2689~2715 行）等价移植：列宽 / 淘汰行红字 / 居中换行。"""
    from openpyxl.styles import Alignment, Font, PatternFill

    red_font = Font(color="FF0000", bold=True)
    green_fill = PatternFill(start_color="00FF00", end_color="00FF00", fill_type="solid")
    if is_summary:
        ws.column_dimensions['A'].width = 25
        ws.column_dimensions['B'].width = 25
        ws.column_dimensions['C'].width = 20
        for row in ws.iter_rows(min_row=2, max_row=ws.max_row, min_col=2, max_col=3):
            for cell in row:
                cell.fill = green_fill
                cell.font = Font(bold=True)
                cell.alignment = Alignment(vertical='center', horizontal='center')
        return

    ws.column_dimensions['A'].width = 12
    ws.column_dimensions['B'].width = 25
    ws.column_dimensions['C'].width = 20
    ws.column_dimensions['D'].width = 15
    ws.column_dimensions['E'].width = 10
    ws.column_dimensions['F'].width = 25
    ws.column_dimensions['G'].width = 15
    ws.column_dimensions['H'].width = 20
    ws.column_dimensions['I'].width = 45

    for row in ws.iter_rows(min_row=2, max_row=ws.max_row, min_col=1, max_col=9):
        if len(row) > 6 and row[6].value and "淘汰" in str(row[6].value):
            for cell in row: cell.font = red_font
        for cell in row: cell.alignment = Alignment(vertical='center', wrap_text=True)


def _guess_rep_name(sheets: Dict[str, pd.DataFrame]) -> str:
    """未显式给代表公司时的兜底：综合得分排行第一名 -> 首个维度的首个公司。"""
    df = sheets.get(SUMMARY_SHEET)
    if df is not None and not df.empty and "公司" in df.columns:
        return str(df.iloc[0]["公司"])
    for name, d in sheets.items():
        if name == SUMMARY_SHEET:
            continue
        if d is not None and not d.empty and "公司" in d.columns:
            return str(d.iloc[0]["公司"])
    return "代表公司"


def export_industry_excel(industry: str, sheets: Dict[str, pd.DataFrame], output_dir: str,
                          rep_name: str = "", log: Optional[Callable[[str], None]] = None) -> str:
    """写出《{代表公司}_{行业}_16维度同行对比分析.xlsx》，返回文件路径。

    列名、sheet 名、sheet 顺序、列宽与红字样式与桌面版一致。
    """
    _log = make_logger(log)
    if not sheets:
        # 显式报错而非让 pd.ExcelWriter 抛 "At least one sheet must be visible"
        raise ValueError(f"行业 [{industry}] 没有任何可写出的 sheet（sheets 为空），未生成文件")
    os.makedirs(output_dir, exist_ok=True)

    rep = rep_name or _guess_rep_name(sheets)
    safe_rep = contracts.safe_filename(rep)
    safe_ind = contracts.safe_filename(industry)

    out_name = f"{safe_rep}_{safe_ind}_16维度同行对比分析.xlsx"
    out_file = os.path.join(output_dir, out_name)

    with pd.ExcelWriter(out_file, engine='openpyxl') as writer:
        for dim_name, df in sheets.items():
            df.to_excel(writer, sheet_name=dim_name, index=False)
            ws = writer.sheets[dim_name]
            _format_excel_sheet(ws, is_summary=(dim_name == SUMMARY_SHEET))

    _log(f"🎉 成功生成报告: {out_name}")
    return out_file


# =====================================================================
# 四、全流程：扫描 -> 按行业分析 -> 逐行业出表
# =====================================================================
def run_all(input_dir: str, output_dir: str, cfg: Dict[str, Any],
            log: Optional[Callable[[str], None]] = None,
            progress: Optional[Callable[[int, int, str], None]] = None) -> Dict[str, Any]:
    """全流程：扫描 -> 按行业分析 -> 逐行业出表，写到 output_dir/16维度对比结果/。

    progress(当前序号, 总数, 行业名) 在每个行业开始前回调一次（序号从 1 开始）。
    返回 {"ok": bool, "files": [路径...], "industries": [...], "error": str(可选)}，可 JSON 序列化。
    """
    _log = make_logger(log)
    files: List[str] = []
    industries: List[str] = []

    try:
        if not os.path.exists(input_dir):
            _log(f"⚠️ 输入源目录不存在，跳过分析: {input_dir}")
            return {"ok": False, "files": files, "industries": industries,
                    "error": f"输入源目录不存在: {input_dir}"}

        main_out = os.path.join(output_dir, contracts.DIR_16DIM)
        os.makedirs(main_out, exist_ok=True)

        all_data = scan_companies(input_dir, log=log)
        if not all_data:
            _log("⚠️ 未提取到有效数据，已跳过。")
            return {"ok": False, "files": files, "industries": industries,
                    "error": "未提取到有效数据（《报表提取完善》下没有可用的《统一整合输出》文件；"
                             "请先跑“年报数据搜索汇总”）"}

        rep_map = _scan_industry_rep_map(input_dir)
        merged_cfg = _merge_cfg(cfg)
        total = len(all_data)
        generated_count = 0

        # 明确提示“同行公司不足”：16 维度是**同行对比**，每个行业至少要有 2 家公司，
        # 否则逐维度对比没有对手，结果为空。旧版本只打印“未提取到有效数据”，
        # 用户无法判断是自己只跑了 1 家公司还是数据有问题。
        thin = {ind: len(cd) for ind, cd in all_data.items() if len(cd) < 2}
        if thin:
            need_total = sum(1 for cd in all_data.values() if len(cd) < 2)
            _log(f"   ⚠️ 有 {need_total} 个行业的同行公司不足 2 家，无法做同行对比："
                 + "、".join(f"{k}({v}家)" for k, v in list(thin.items())[:8]))
            _log("   ℹ️ 16 维度需要**同一行业内至少 2 家公司**。请在“年报数据”页多指定几家"
                 "同行业公司（例如乳品：伊利股份 600887、东鹏饮料 605499、养元饮品 603156）。")

        for idx, (industry, comp_dict) in enumerate(all_data.items(), start=1):
            if progress is not None:
                try:
                    progress(idx, total, industry)
                except Exception as e:  # noqa: BLE001
                    _log(f"  ⚠️ progress 回调异常: {type(e).__name__}: {str(e)[:120]}")
            if not comp_dict:
                continue

            try:
                sheets, rep = _analyze_industry_full(industry, comp_dict, merged_cfg, _log,
                                                     rep_hint=rep_map.get(industry, ""))
                if not sheets:
                    continue
                out_file = export_industry_excel(industry, sheets, main_out, rep_name=rep, log=log)
            except Exception as e:  # noqa: BLE001
                _log(f"❌ 行业 [{industry}] 生成报告失败: {e}")
                continue

            files.append(out_file)
            industries.append(industry)
            generated_count += 1

        if generated_count > 0:
            _log(f"🏁 完美收官！共生成 {generated_count} 份独立行业对比报告。")
        else:
            _log("⚠️ 所有行业的报告生成均失败。")

        return {"ok": generated_count > 0, "files": files, "industries": industries}

    except Exception as err:  # noqa: BLE001
        _log(f"❌ 运行发生严重异常: {err}")
        return {"ok": False, "files": files, "industries": industries, "error": str(err)}
