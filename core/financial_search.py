# -*- coding: utf-8 -*-
"""
年报数据搜索汇总（替代原“年报提取模块”）

原桌面版流程：下载年报 PDF → 解析 PDF 抽出三大表 → 联网补充员工/分红 → 产出
  ① 单年《{简称}（{代码}）{年份}年度年报提取表.xlsx》
  ② 《统一整合输出_{简称}({代码})_{起}-{止}.xlsx》
  ③ 《1_{代表公司}_公司属性表.xlsx》/《2_{代表公司}_同行排列表.xlsx》

云端版取消“下载年报”与“PDF 解析”，改为**直接用免费数据源检索汇总**，
但**产出的表结构与上述完全一致**（下游造假排雷/16维度/企业AI深度直接复用，无需改动）。
数据源：AkShare（新浪财报 / 东财数据中心 / 同花顺分红），全部免费、无需 token。
"""
from __future__ import annotations

import os
import re
import time
from typing import Any, Callable, Dict, Iterable, List, Optional, Tuple

import pandas as pd
from openpyxl.styles import Alignment, Font
from openpyxl.utils import get_column_letter

from .common import DiskCache, make_logger, norm_item, safe_float, to_yuan
from .contracts import (
    ATTR_COLUMNS,
    ATTR_FILENAME_TEMPLATE,
    DIV_COLUMNS,
    EMP_TOTAL_KEYWORDS,
    FIXED_ASSET_FALLBACK_ITEMS,
    PEER_COLUMNS,
    PEER_FILENAME_TEMPLATE,
    STATEMENT_ITEM_ALIASES,
    STATEMENT_KEY_COL,
    STATEMENT_YEAR_COL_FMT,
    UNIFIED_FILENAME_TEMPLATE,
    YEARLY_FILENAME_TEMPLATE,
    YEARLY_EXTRA_SHEETS,
    YEARLY_FINANCIAL_SHEETS,
    safe_filename,
)
from .universe import StockUniverse, get_industry, industry_label, remember_industry_member

# 信披/财务表里不参与“项目: 值”导出的元数据列
_META_COLS = {"数据源", "是否审计", "公告日期", "币种", "类型", "更新日期", "报告日",
              "SECUCODE", "SECURITY_CODE", "SECURITY_NAME_ABBR", "ORG_CODE", "ORG_TYPE"}


# ======================================================================
# 一、三大报表（AkShare 新浪源，中文科目名最贴近年报原文）
# ======================================================================
def _sina_symbol(code: str) -> str:
    code = str(code).zfill(6)
    if code[0] in ("6", "9"):
        return "sh" + code
    if code[0] in ("4", "8"):
        return "bj" + code
    return "sz" + code


def fetch_statement_raw(code: str, kind: str, cache: DiskCache,
                        log: Optional[Callable[[str], None]] = None) -> Optional[pd.DataFrame]:
    """kind: 资产负债表 / 利润表 / 现金流量表。

    返回**已转置**的表：**行 = 报告期(YYYYMMDD)**、**列 = 科目名**。
    说明：AkShare 新浪接口原始形态是“行=科目、列=报告期”，但列名并非可用的年份列
    （列只是中文科目名，报告期在 index），因此这里统一转置成“行=报告期”的形态，
    便于按 `df.loc[period, 科目]` 取数。
    """
    log = log or make_logger()
    key = f"sina_stmt_T_{kind}_{code}"
    cached = cache.get(key)
    if cached:
        try:
            df = pd.DataFrame(cached)
            if not df.empty:
                return df.set_index("报告期")
        except Exception:
            pass
    try:
        import akshare as ak
        raw = ak.stock_financial_report_sina(stock=_sina_symbol(code), symbol=kind)
        if raw is None or raw.empty:
            log(f"   ⚠️ {code} 的{kind}为空")
            return None
        raw = raw.copy()
        raw["报告日"] = raw["报告日"].astype(str).str.replace("-", "", regex=False).str[:8]
        raw = raw.drop_duplicates(subset=["报告日"], keep="first").set_index("报告日")
        # 原始：index=报告期，columns=科目  →  转置为 index=科目，columns=报告期
        t = raw.T
        t.index = [norm_item(i) for i in t.index]
        t.index.name = "报告期"
        # 与 Excel 一致的元数据列剔除（它们不是科目）
        col_drop = [c for c in t.columns if str(c) in _META_COLS]
        if col_drop:
            t = t.drop(columns=col_drop, errors="ignore")
        try:
            cache.set(key, t.reset_index().to_dict("records"))
        except Exception:
            pass
        return t
    except Exception as e:  # noqa: BLE001
        log(f"   ⚠️ 获取 {code} 的{kind}失败: {type(e).__name__}: {str(e)[:120]}")
        return None


def _period_cols(raw: pd.DataFrame, years: Iterable[int]) -> List[str]:
    """从“行=报告期”的表里挑出目标年份的年报列（优先 1231，缺失则同年其它期末兜底）。"""
    cols: List[str] = []
    idx = [str(i) for i in raw.columns]
    for y in sorted(set(int(x) for x in years)):
        prefer = f"{y}1231"
        if prefer in idx:
            cols.append(prefer)
            continue
        same_year = [c for c in idx if c.startswith(str(y))]
        if same_year:
            cols.append(sorted(same_year)[-1])
    return cols


def build_statement_sheet(raw: Optional[pd.DataFrame], years: Iterable[int],
                          log: Optional[Callable[[str], None]] = None) -> pd.DataFrame:
    """把“行=报告期”的表转成契约要求的结构：列 = ['项目', '20231231', '20241231', ...]。

    - 行 = 年报科目名；数值统一为“元”；
    - 科目名做轻量归一（去空格/全角括号/帐->账）并套用别名表（新浪“所支付的现金”→“支付的现金”）；
    - 重复科目名取第一次出现。
    """
    log = log or make_logger()
    if raw is None or raw.empty:
        return pd.DataFrame(columns=[STATEMENT_KEY_COL])
    cols = _period_cols(raw, years)
    if not cols:
        return pd.DataFrame(columns=[STATEMENT_KEY_COL])

    data: Dict[str, List[Optional[float]]] = {STATEMENT_KEY_COL: [norm_item(i) for i in raw.index]}
    for c in cols:
        try:
            series = raw.loc[:, c]
        except Exception:
            data[c] = [None] * len(raw.index)
            continue
        vals: List[Optional[float]] = []
        for v in series.tolist():
            vals.append(to_yuan(v) if isinstance(v, str) else safe_float(v))
        data[c] = vals

    out = pd.DataFrame(data)
    if out.empty:
        return pd.DataFrame(columns=[STATEMENT_KEY_COL])
    out[STATEMENT_KEY_COL] = out[STATEMENT_KEY_COL].replace(STATEMENT_ITEM_ALIASES)
    out = out[out[STATEMENT_KEY_COL].astype(str).str.len() > 0]
    out = out.drop_duplicates(subset=[STATEMENT_KEY_COL], keep="first").reset_index(drop=True)

    # 固定资产兜底：新浪没有“固定资产”总项，用净额/清理合计/净值依次回退
    if "固定资产" not in set(out[STATEMENT_KEY_COL]):
        for cand in FIXED_ASSET_FALLBACK_ITEMS:
            if cand in set(out[STATEMENT_KEY_COL]):
                row = out[out[STATEMENT_KEY_COL] == cand].copy()
                row[STATEMENT_KEY_COL] = "固定资产"
                out = pd.concat([out, row], ignore_index=True)
                break
    return out


# ======================================================================
# 二、分红情况（同花顺 + 东财双源）
# ======================================================================
def fetch_dividend_rows(code: str, cache: DiskCache,
                        log: Optional[Callable[[str], None]] = None,
                        ds_cfg: Optional[Dict[str, Any]] = None) -> List[Dict[str, Any]]:
    """取分红送转记录（同花顺为主、东财补缺、TuShare 兜底），统一按“除权年度”归年。

    返回 [{年份, 每10股派息(元), 方案进度, 除权除息日, 公告日期, 来源}]
    说明：年份取**除权除息日所在年**——即“该年度实际派发的现金”，与好价模块的 TTM 股息口径一致；
    若除权日缺失则回退公告日期/报告期年份。
    """
    log = log or make_logger()
    ds_cfg = ds_cfg or {}
    key = f"dividend_v3_{code}"
    cached = cache.get(key)
    if cached:
        return cached

    rows: List[Dict[str, Any]] = []

    # ① 同花顺分红（主源：派息 = 每10股派息(元)）
    try:
        import akshare as ak
        df = ak.stock_history_dividend_detail(symbol=str(code).zfill(6), indicator="分红")
        if df is not None and not df.empty:
            for _, r in df.iterrows():
                ex = str(r.get("除权除息日") or "")[:10]
                ann = str(r.get("公告日期") or "")[:10]
                ref = ex or ann
                year = int(ref[:4]) if re.match(r"^\d{4}", ref) else None
                per10 = safe_float(r.get("派息"))
                if per10 is None:
                    continue
                rows.append({"年份": year, "每10股派息(元)": float(per10),
                             "方案进度": str(r.get("进度") or ""), "除权除息日": ex,
                             "公告日期": ann, "来源": "同花顺"})
    except Exception as e:  # noqa: BLE001
        log(f"   ⚠️ 同花顺分红获取失败[{code}]: {type(e).__name__}: {str(e)[:100]}")

    # ② 东财分红（补缺：仅当同花顺没有同“除权日+派息额”的记录时才加入）
    try:
        import akshare as ak
        df2 = ak.stock_fhps_detail_em(symbol=str(code).zfill(6))
        if df2 is not None and not df2.empty:
            exist = {(r.get("除权除息日"), r.get("每10股派息(元)")) for r in rows}
            exist_ex = {r.get("除权除息日") for r in rows if r.get("除权除息日")}
            for _, r in df2.iterrows():
                ex = str(r.get("除权除息日") or "")[:10]
                rep = str(r.get("报告期") or "")[:10]
                per10 = safe_float(r.get("现金分红-现金分红比例"))
                if per10 is None:
                    continue
                if (ex, float(per10)) in exist:
                    continue
                # 同一除权日的记录视为同一笔，避免同花顺/东财重复计入
                if ex and ex in exist_ex:
                    continue
                year = int(rep[:4]) if re.match(r"^\d{4}", rep) else (
                    int(ex[:4]) if re.match(r"^\d{4}", ex) else None)
                rows.append({"年份": year, "每10股派息(元)": float(per10),
                             "方案进度": str(r.get("方案进度") or ""), "除权除息日": ex,
                             "公告日期": str(r.get("最新公告日期") or "")[:10], "来源": "东财"})
    except Exception as e:  # noqa: BLE001
        log(f"   ⚠️ 东财分红获取失败[{code}]: {type(e).__name__}: {str(e)[:100]}")

    # 再按 (年份, 派息额) 去重，保留有除权日/已实施的那条
    uniq: Dict[Tuple[Any, Any], Dict[str, Any]] = {}
    for r in rows:
        k = (r.get("年份"), round(float(r.get("每10股派息(元)") or 0), 4))
        old = uniq.get(k)
        if old is None:
            uniq[k] = r
        elif not old.get("除权除息日") and r.get("除权除息日"):
            uniq[k] = r
    out = sorted(uniq.values(), key=lambda x: (x.get("年份") or 0, x.get("除权除息日") or ""))
    # TuShare 兜底：仅补 AkShare 缺失的年份
    try:
        from .datasource import dividends_merged
        before = len(out)
        out = dividends_merged(code, out, ds_cfg=ds_cfg, log=log)
        if len(out) > before:
            log(f"   ℹ️ 已用 TuShare 补充 {len(out) - before} 条分红记录")
    except Exception:
        pass
    cache.set(key, out)
    log(f"   分红记录 {len(out)} 条（同花顺+东财+TuShare 合并去重）")
    return out


def build_dividend_sheet(div_rows: List[Dict[str, Any]], years: Iterable[int],
                         net_profit_by_year: Dict[int, Optional[float]],
                         total_shares: Optional[float] = None,
                         log: Optional[Callable[[str], None]] = None) -> pd.DataFrame:
    """产出《统一整合输出》里的【分红情况】表（**按年汇总，一年一行**）：
    列 = 股票代码 / 年份 / 同花顺现金分红总额(元) / 归属于母公司所有者的净利润(元) / 分红率

    - 同一年若有多笔分红（中期+年度），按“每10股派息”**累加**后换算成现金分红总额；
    - 分红率 = 现金分红总额 ÷ 归母净利润 ×100%，形如 '69.60%'（16维度模块按此解析）；
    - 总额需总股本（来自资产负债表“股本/实收资本”或缓存），缺失时该列留空但分红率仍尽力给出。
    """
    log = log or make_logger()
    want = sorted(set(int(y) for y in years))
    per_year: Dict[int, float] = {}
    for r in div_rows:
        y, per10 = r.get("年份"), r.get("每10股派息(元)")
        if y is None or per10 is None:
            continue
        try:
            y = int(y)
            per_year[y] = per_year.get(y, 0.0) + float(per10)
        except Exception:
            continue

    recs: List[Dict[str, Any]] = []
    for y in want:
        per10 = per_year.get(y)
        total_div = None
        if per10 is not None and total_shares:
            total_div = round(float(per10) / 10.0 * float(total_shares), 2)
        np_ = net_profit_by_year.get(y)
        rate = ""
        if total_div is not None and np_:
            try:
                rate = f"{total_div / float(np_) * 100:.2f}%"
            except Exception:
                rate = ""
        recs.append({
            "股票代码": "",
            "年份": y,
            "同花顺现金分红总额(元)": total_div,
            "归属于母公司所有者的净利润(元)": np_,
            "分红率": rate,
        })
    return pd.DataFrame(recs, columns=DIV_COLUMNS)


# ======================================================================
# 三、员工情况 / 董监高 / 母公司表
# ======================================================================
def fetch_employee_info(code: str, cache: DiskCache,
                        log: Optional[Callable[[str], None]] = None,
                        cash_flow_sheet: Optional[pd.DataFrame] = None,
                        ds_cfg: Optional[Dict[str, Any]] = None) -> pd.DataFrame:
    """员工情况：列 = 年份 / 项目 / 数值。

    原桌面版来自年报 PDF 的“员工情况”章节；免费数据源里**没有员工人数**，
    因此这里改为提供下游真正要用的“支付给职工以及为职工支付的现金”（取自现金流量表），
    并把公司概况里能拿到的员工/职工类字段一并带出（有多少给多少，不伪造）。
    """
    log = log or make_logger()
    key = f"employee_v2_{code}"
    recs: List[Dict[str, Any]] = []
    cached = cache.get(key)
    if cached:
        recs = list(cached)

    if not recs:
        # 公司概况（巨潮）：能拿到“注册资金/成立日期”等，但不含员工人数，尽力抓取相关字段
        try:
            import akshare as ak
            prof = ak.stock_profile_cninfo(symbol=str(code).zfill(6))
            if prof is not None and not prof.empty:
                for col in prof.columns:
                    if any(kw in str(col) for kw in ("员工", "职工")):
                        v = safe_float(prof.iloc[0][col])
                        if v:
                            recs.append({"年份": "", "项目": str(col), "数值": v})
        except Exception as e:  # noqa: BLE001
            log(f"   ⚠️ 公司概况获取失败[{code}]: {type(e).__name__}: {str(e)[:100]}")

    # 现金流量表里的“支付给职工以及为职工支付的现金”按年落到员工表（下游/报告都会用到）
    # 注意：必须按 (年份, 项目) 去重再追加 —— 否则命中缓存时会把同样的年度再写一遍，
    # 导致《员工情况(PDF提取)》出现完全重复的行（实测每行重复 2 次）。
    if cash_flow_sheet is not None and not cash_flow_sheet.empty:
        for key_name in ("支付给职工以及为职工支付的现金", "支付给职工以及为职工支付的现金(元)"):
            hit = cash_flow_sheet[cash_flow_sheet[STATEMENT_KEY_COL] == key_name]
            if hit.empty:
                continue
            row = hit.iloc[0]
            seen = {(str(r.get("年份")), str(r.get("项目"))) for r in recs}
            for col in cash_flow_sheet.columns:
                if col == STATEMENT_KEY_COL:
                    continue
                m = re.match(r"^(20\d{2})", str(col))
                v = safe_float(row[col])
                if m and v:
                    item = {"年份": int(m.group(1)),
                            "项目": "支付给职工以及为职工支付的现金", "数值": v}
                    if (str(item["年份"]), item["项目"]) not in seen:
                        recs.append(item)
                        seen.add((str(item["年份"]), item["项目"]))
            break

    # 最终再去重一次（防止历史缓存里已经存了重复行）
    if recs:
        _seen = set()
        _uniq = []
        for r in recs:
            k2 = (str(r.get("年份")), str(r.get("项目")))
            if k2 in _seen:
                continue
            _seen.add(k2)
            _uniq.append(r)
        recs = _uniq
        cache.set(key, recs)
    return pd.DataFrame(recs, columns=["年份", "项目", "数值"])


def fetch_executives(code: str, cache: DiskCache,
                     log: Optional[Callable[[str], None]] = None) -> pd.DataFrame:
    """董监高及报酬情况（来自同花顺高管变动 + 巨潮公司概况补充）。"""
    log = log or make_logger()
    key = f"executives_{code}"
    cached = cache.get(key)
    if cached:
        return pd.DataFrame(cached)
    recs: List[Dict[str, Any]] = []
    try:
        import akshare as ak
        df = ak.stock_management_change_ths(symbol=str(code).zfill(6))
        if df is not None and not df.empty:
            for _, r in df.head(200).iterrows():
                recs.append({str(k): ("" if pd.isna(v) else v) for k, v in r.items()})
    except Exception as e:  # noqa: BLE001
        log(f"   ⚠️ 高管信息获取失败[{code}]: {type(e).__name__}: {str(e)[:100]}")
    cache.set(key, recs)
    return pd.DataFrame(recs)


def _empty_parent_sheet() -> pd.DataFrame:
    """母公司报表：免费源不提供母公司口径，保留同名空表以维持契约（下游只读“合并”表）。"""
    return pd.DataFrame(columns=[STATEMENT_KEY_COL])


# ======================================================================
# 四、产出：单年《年报提取表》
# ======================================================================
def _fmt_sheet(ws, title: str, is_financial: bool, sheet_kind: str = "") -> None:
    ws.cell(1, 1, title).font = Font(bold=True, size=12)
    ws.cell(1, 1).alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws.row_dimensions[1].height = 40
    for row in ws.iter_rows(min_row=3, max_col=ws.max_column):
        item_name = str(row[0].value) if row[0].value else ""
        for cell in row[1:]:
            if isinstance(cell.value, (int, float)):
                if is_financial:
                    cell.number_format = "#,##0.0000" if "每股收益" in item_name else "#,##0.00"
                elif sheet_kind == "员工情况":
                    cell.number_format = "#,##0"
                else:
                    cell.number_format = "#,##0.00"
    for col in range(1, min(ws.max_column, 30) + 1):
        ws.column_dimensions[get_column_letter(col)].width = 38 if col == 1 else 16


def build_yearly_table(code: str, name: str, year: int, cache: DiskCache,
                       log: Optional[Callable[[str], None]] = None,
                       ds_cfg: Optional[Dict[str, Any]] = None) -> Dict[str, pd.DataFrame]:
    """组装单年《年报提取表》的 11 张 sheet（与原表同名同构）。"""
    log = log or make_logger()
    years = [year]
    sheets: Dict[str, pd.DataFrame] = {}

    bs = build_statement_sheet(fetch_statement_raw(code, "资产负债表", cache, log), years, log)
    is_ = build_statement_sheet(fetch_statement_raw(code, "利润表", cache, log), years, log)
    cf = build_statement_sheet(fetch_statement_raw(code, "现金流量表", cache, log), years, log)

    sheets["合并资产负债表"] = bs
    sheets["母公司资产负债表"] = _empty_parent_sheet()
    sheets["合并利润表"] = is_
    sheets["母公司利润表"] = _empty_parent_sheet()
    sheets["合并现金流量表"] = cf
    sheets["母公司现金流量表"] = _empty_parent_sheet()
    sheets["合并所有者权益变动表"] = _empty_parent_sheet()
    sheets["母公司所有者权益变动表"] = _empty_parent_sheet()

    div_rows = fetch_dividend_rows(code, cache, log, ds_cfg=ds_cfg)
    np_map = _net_profit_map(is_)
    shares = _shares_from_balance_sheet(bs) or _total_shares(cache, code)
    if shares:
        remember_total_shares(cache, code, shares)
    div_df = build_dividend_sheet(div_rows, years, np_map, shares, log)
    if not div_df.empty:
        div_df = div_df.copy()
        div_df["股票代码"] = str(code).zfill(6)
    sheets["分红情况"] = div_df
    sheets["员工情况"] = fetch_employee_info(code, cache, log, cash_flow_sheet=cf, ds_cfg=ds_cfg)
    sheets["董监高及报酬情况"] = fetch_executives(code, cache, log)
    return sheets


def _net_profit_map(income_sheet: pd.DataFrame) -> Dict[int, Optional[float]]:
    """从利润表里取“归属于母公司所有者的净利润”按年映射。"""
    out: Dict[int, Optional[float]] = {}
    if income_sheet is None or income_sheet.empty:
        return out
    keys = ["归属于母公司所有者的净利润", "归属于母公司股东的净利润", "净利润"]
    row = None
    for k in keys:
        hit = income_sheet[income_sheet[STATEMENT_KEY_COL] == k]
        if not hit.empty:
            row = hit.iloc[0]
            break
    if row is None:
        return out
    for col in income_sheet.columns:
        if col == STATEMENT_KEY_COL:
            continue
        m = re.match(r"^(20\d{2})", str(col))
        if m:
            out[int(m.group(1))] = safe_float(row[col])
    return out


def _shares_from_balance_sheet(bs: Optional[pd.DataFrame]) -> Optional[float]:
    """从合并资产负债表里取总股本（“股本 / 实收资本(或股本) / 实收资本”），取最近一期有值者。"""
    if bs is None or bs.empty:
        return None
    try:
        keys = ["股本", "实收资本(或股本)", "实收资本", "实收资本（或股本）"]
        hit = bs[bs[STATEMENT_KEY_COL].isin(keys)]
        if hit.empty:
            hit = bs[bs[STATEMENT_KEY_COL].astype(str).str.contains("股本|实收资本", na=False)]
        if hit.empty:
            return None
        row = hit.iloc[0]
        latest: Optional[float] = None
        for col in bs.columns:           # 列按年份升序，保留最后一个有值的
            if col == STATEMENT_KEY_COL:
                continue
            v = safe_float(row[col])
            if v and v > 0:
                latest = v
        return latest
    except Exception:
        return None


def _total_shares(cache: DiskCache, code: str) -> Optional[float]:
    """总股本（用于把“每10股派息”换算成现金分红总额）：优先缓存，其次腾讯行情字段[46]近似。"""
    key = f"shares_{code}"
    hit = cache.get(key)
    if hit:
        try:
            return float(hit)
        except Exception:
            pass
    return None


def remember_total_shares(cache: DiskCache, code: str, shares: Optional[float]) -> None:
    if shares:
        cache.set(f"shares_{code}", float(shares))


def write_yearly_excel(sheets: Dict[str, pd.DataFrame], out_dir: str, name: str, code: str, year: int,
                       log: Optional[Callable[[str], None]] = None) -> str:
    """写出《{简称}（{代码}）{年}年度年报提取表.xlsx》，版式与原程序一致。"""
    log = log or make_logger()
    os.makedirs(out_dir, exist_ok=True)
    path = os.path.join(out_dir, YEARLY_FILENAME_TEMPLATE.format(name=safe_filename(name), code=code, year=year))
    order = YEARLY_FINANCIAL_SHEETS + list(YEARLY_EXTRA_SHEETS)
    with pd.ExcelWriter(path, engine="openpyxl") as writer:
        for sheet_name in order:
            df = sheets.get(sheet_name)
            if df is None or df.empty:
                df = pd.DataFrame()
            df.to_excel(writer, sheet_name=sheet_name, startrow=2, index=False)
            ws = writer.sheets[sheet_name]
            is_financial = any(x in sheet_name for x in
                               ["资产负债表", "利润表", "现金流量表", "所有者权益变动表"])
            unit = "人民币元" if is_financial else "详见表内说明"
            _fmt_sheet(ws, f"{name}（{code}）{sheet_name}（{year}年度）\n单位：{unit}",
                       is_financial, sheet_name)
    log(f"   ✅ 已生成 {os.path.basename(path)}")
    return path


# ======================================================================
# 五、产出：统一整合输出 / 公司属性表 / 同行排列表
# ======================================================================
def build_unified_sheets(code: str, name: str, start_y: int, end_y: int, cache: DiskCache,
                         log: Optional[Callable[[str], None]] = None,
                         ds_cfg: Optional[Dict[str, Any]] = None) -> Dict[str, pd.DataFrame]:
    """组装《统一整合输出》的 5 张 sheet：员工情况(PDF提取) / 分红情况 / 三大表。"""
    log = log or make_logger()
    years = list(range(int(start_y), int(end_y) + 1))
    bs = build_statement_sheet(fetch_statement_raw(code, "资产负债表", cache, log), years, log)
    is_ = build_statement_sheet(fetch_statement_raw(code, "利润表", cache, log), years, log)
    cf = build_statement_sheet(fetch_statement_raw(code, "现金流量表", cache, log), years, log)

    div_rows = fetch_dividend_rows(code, cache, log, ds_cfg=ds_cfg)
    np_map = _net_profit_map(is_)
    shares = _shares_from_balance_sheet(bs) or _total_shares(cache, code)
    if shares:
        remember_total_shares(cache, code, shares)
    div = build_dividend_sheet(div_rows, years, np_map, shares, log)
    if not div.empty:
        div = div.copy()
        div["股票代码"] = str(code).zfill(6)

    emp = fetch_employee_info(code, cache, log, cash_flow_sheet=cf, ds_cfg=ds_cfg)
    if emp is None or emp.empty:
        emp = pd.DataFrame(columns=["年份", "项目", "数值"])

    return {
        "员工情况(PDF提取)": emp,
        "分红情况": div,
        "合并资产负债表": bs,
        "合并利润表": is_,
        "合并现金流量表": cf,
    }


def _merge_year_columns(sheets: Dict[str, pd.DataFrame]) -> Dict[str, pd.DataFrame]:
    """把三大表按“项目”合并去重（不同年份多次取数时避免重复行）。"""
    out: Dict[str, pd.DataFrame] = {}
    for k, df in sheets.items():
        if k in ("合并资产负债表", "合并利润表", "合并现金流量表") and df is not None and not df.empty:
            df = df.drop_duplicates(subset=[STATEMENT_KEY_COL], keep="first").reset_index(drop=True)
        out[k] = df
    return out


def write_unified_excel(sheets: Dict[str, pd.DataFrame], out_dir: str, name: str, code: str,
                        start_y: int, end_y: int, log: Optional[Callable[[str], None]] = None) -> str:
    """写出《统一整合输出_{简称}({代码})_{起}-{止}.xlsx》。

    **表头放在第 1 行**（与原桌面版“第1行合并标题 / 第2行表头”不同）：
    这样任何客户端直接 `pd.read_excel(path, sheet_name=...)` 就能拿到
    `['项目','20221231',...]`，无需再做表头探测；下游两个分析模块已同时兼容两种版式。
    说明文字改为放到首列上方的单独注释行之外——为保持零歧义，这里不再写合并标题。
    """
    log = log or make_logger()
    os.makedirs(out_dir, exist_ok=True)
    sheets = _merge_year_columns(sheets)
    path = os.path.join(out_dir, UNIFIED_FILENAME_TEMPLATE.format(
        name=safe_filename(name), code=code, start=int(start_y), end=int(end_y)))
    with pd.ExcelWriter(path, engine="openpyxl") as writer:
        for s_name, df in sheets.items():
            if df is None:
                df = pd.DataFrame()
            sheet_name = s_name[:31]
            df.to_excel(writer, sheet_name=sheet_name, index=False, header=True)
            ws = writer.sheets[sheet_name]
            max_col = max(1, df.shape[1])
            # 首行加粗置中（只是样式，不影响读取）
            for c in range(1, max_col + 1):
                cell = ws.cell(row=1, column=c)
                cell.font = Font(size=11, bold=True)
                cell.alignment = Alignment(horizontal="center", vertical="center")
            for col in range(1, min(max_col, 40) + 1):
                ws.column_dimensions[get_column_letter(col)].width = 40 if col == 1 else 16
    log(f"   ✅ 已生成 {os.path.basename(path)}")
    return path


def build_attribute_table(rep: Dict[str, Any], industry_info: Dict[str, str]) -> pd.DataFrame:
    """《1_{代表公司}_公司属性表.xlsx》——列名与契约一致（下游按“目标股票/二级行业”读取）。"""
    return pd.DataFrame([{
        "目标股票": rep.get("code", ""),
        "股票简称": rep.get("name", ""),
        "一级行业": industry_info.get("一级行业", ""),
        "二级行业": industry_info.get("二级行业", ""),
        "三级行业": industry_info.get("三级行业", ""),
        "完整行业路径": industry_info.get("完整行业路径", ""),
    }], columns=ATTR_COLUMNS)


def build_peer_table(rep_code: str, peers: List[Dict[str, Any]]) -> pd.DataFrame:
    """《2_{代表公司}_同行排列表.xlsx》——含代表公司自己，按市值排名（与原程序一致）。"""
    rows = []
    for i, p in enumerate(peers, start=1):
        rows.append({
            "原查询股票": rep_code,
            "排名": p.get("排名", i),
            "代码": p.get("代码", ""),
            "名称": p.get("名称", ""),
            "市值(亿)": p.get("市值(亿)"),
        })
    return pd.DataFrame(rows, columns=PEER_COLUMNS)


def write_attr_and_peer(attr_df: pd.DataFrame, peer_df: pd.DataFrame, ind_dir: str, rep_name: str,
                        log: Optional[Callable[[str], None]] = None) -> Tuple[str, str]:
    log = log or make_logger()
    os.makedirs(ind_dir, exist_ok=True)
    rep = safe_filename(rep_name)
    p1 = os.path.join(ind_dir, ATTR_FILENAME_TEMPLATE.format(rep=rep))
    p2 = os.path.join(ind_dir, PEER_FILENAME_TEMPLATE.format(rep=rep))
    attr_df.to_excel(p1, index=False)
    peer_df.to_excel(p2, index=False)
    log(f"   ✅ 已生成 {os.path.basename(p1)} / {os.path.basename(p2)}")
    return p1, p2


# ======================================================================
# 六、总入口：年报数据搜索汇总
# ======================================================================
def run_annual_report_search(targets: List[Dict[str, Any]], output_root: str, cache: DiskCache,
                             years_back: int = 5, start_year: Optional[int] = None,
                             end_year: Optional[int] = None,
                             log: Optional[Callable[[str], None]] = None,
                             progress: Optional[Callable[[int, int, str], None]] = None,
                             include_peers: bool = True,
                             ds_cfg: Optional[Dict[str, Any]] = None) -> Dict[str, Any]:
    """把“海选名单 / 手工输入的公司”汇总成原《报表提取完善》目录结构。

    参数
    ----
    targets: [{"code": "605499", "name": "东鹏饮料", "所属行业": "饮料乳品", "属性": "原查询公司"}, ...]
    output_root: 输出根目录（会创建 报表提取完善/ 子目录）
    """
    log = log or make_logger()
    ds_cfg = ds_cfg or {}
    out_root = os.path.join(output_root, "报表提取完善")
    os.makedirs(out_root, exist_ok=True)

    this_year = time.localtime().tm_year
    end_y = int(end_year or this_year)
    # 【年度夹取】当年年报通常要到次年 4 月底才披露完，请求“当年”只会得到空表
    # （用户反馈“提取表只有一个数字且没有表头”即由此产生）。
    # 因此把结束年度限制在“去年”，除非用户明确要求当年且当前已是年末。
    _now_month = time.localtime().tm_mon
    _latest_full = this_year - 1
    if end_y >= this_year and _now_month < 12:
        log(f"   ℹ️ 结束年度 {end_y} 的年度报告通常尚未披露（当年年报多在次年 4 月底前披露完），"
            f"已自动调整结束年度为 {_latest_full}。如需包含 {end_y} 年的中期数据请到“年报数据”页手工设置。")
        end_y = _latest_full
    start_y = int(start_year or (end_y - max(1, int(years_back)) + 1))
    if start_y > end_y:
        start_y = end_y
    log(f"🔄 年报数据搜索汇总启动：区间 {start_y}-{end_y}，共 {len(targets)} 家目标公司")

    generated: List[str] = []
    summary: List[Dict[str, Any]] = []

    # ① 逐家汇总（含同行）
    total = len(targets)
    for i, t in enumerate(targets, start=1):
        code, name = str(t.get("code", "")).zfill(6), str(t.get("name") or "")
        if not code:
            continue
        if progress:
            progress(i, total, f"{name}({code})")
        log(f"\n📁 [{i}/{total}] 正在汇总 {name}({code}) 的年报数据 ...")
        ind_info = get_industry(code, cache, log)
        ind = industry_label(ind_info) or str(t.get("所属行业") or "默认行业")
        remember_industry_member(cache, code, ind)
        rep_name = name or code
        ind_dir = os.path.join(out_root, f"{safe_filename(rep_name)}_行业_{safe_filename(ind)}")
        os.makedirs(ind_dir, exist_ok=True)

        # 公司属性表 + 同行排列表
        attr_df = build_attribute_table({"code": code, "name": rep_name}, ind_info)
        peers: List[Dict[str, Any]] = []
        if include_peers:
            try:
                from .universe import find_top_peers
                universe = StockUniverse(cache, log)
                peers = find_top_peers(code, ind, universe, cache, top_n=3, log=log)
            except Exception as e:  # noqa: BLE001
                log(f"   ⚠️ 同行检索失败: {type(e).__name__}: {str(e)[:100]}")
        self_row = {"排名": 1, "代码": code, "名称": rep_name,
                    "市值(亿)": _safe_cap(code, cache, log)}
        all_rows = [self_row] + [{"排名": i2 + 2, **p} for i2, p in enumerate(peers)]
        all_rows.sort(key=lambda x: -(float(x.get("市值(亿)") or 0)))
        for idx, r in enumerate(all_rows, start=1):
            r["排名"] = idx
        peer_df = build_peer_table(code, all_rows)
        write_attr_and_peer(attr_df, peer_df, ind_dir, rep_name, log)

        # 单年《年报提取表》（目标区间每一年一张）
        for y in range(start_y, end_y + 1):
            try:
                sheets = build_yearly_table(code, rep_name, y, cache, log, ds_cfg=ds_cfg)
                p = write_yearly_excel(sheets, ind_dir, rep_name, code, y, log)
                generated.append(p)
            except Exception as e:  # noqa: BLE001
                log(f"   ⚠️ {name}({code}) {y} 年报表汇总失败: {type(e).__name__}: {str(e)[:120]}")

        # 《统一整合输出》
        try:
            usheets = build_unified_sheets(code, rep_name, start_y, end_y, cache, log, ds_cfg=ds_cfg)
            p = write_unified_excel(usheets, ind_dir, rep_name, code, start_y, end_y, log)
            generated.append(p)
            _remember_shares_from_sheets(cache, code, usheets)
        except Exception as e:  # noqa: BLE001
            log(f"   ⚠️ {name}({code}) 统一整合输出失败: {type(e).__name__}: {str(e)[:120]}")

        summary.append({"code": code, "name": rep_name, "industry": ind,
                        "peers": [p.get("代码") for p in peers], "files": len(generated)})

    # ② 写目标清单（与原程序同名，便于回溯）
    try:
        from .contracts import TARGET_LIST_COLUMNS, TARGET_LIST_FILENAME
        list_df = pd.DataFrame(targets)
        list_df = list_df.rename(columns={"code": "代码", "name": "名称"})
        for col in TARGET_LIST_COLUMNS:
            if col not in list_df.columns:
                list_df[col] = ""
        list_path = os.path.join(output_root, TARGET_LIST_FILENAME)
        list_df[TARGET_LIST_COLUMNS].to_excel(list_path, index=False)
        log(f"\n✅ 已生成《{TARGET_LIST_FILENAME}》")
    except Exception as e:  # noqa: BLE001
        log(f"⚠️ 目标清单生成失败: {type(e).__name__}: {str(e)[:120]}")

    ok = len(generated) > 0
    log(f"\n{'✅' if ok else '❌'} 年报数据搜索汇总完成：共生成 {len(generated)} 个文件")
    return {"ok": ok, "files": generated, "companies": summary,
            "output_dir": out_root, "start_year": start_y, "end_year": end_y}


def _safe_cap(code: str, cache: DiskCache, log) -> Optional[float]:
    try:
        from .universe import get_market_cap
        return get_market_cap(code, cache, log)
    except Exception:
        return None


def _remember_shares_from_sheets(cache: DiskCache, code: str, sheets: Dict[str, pd.DataFrame]) -> None:
    """从资产负债表的“实收资本/股本”反推总股本（若可得），供分红总额换算。"""
    try:
        bs = sheets.get("合并资产负债表")
        if bs is None or bs.empty:
            return
        hit = bs[bs[STATEMENT_KEY_COL].isin(["实收资本(或股本)", "股本", "实收资本"])]
        if hit.empty:
            return
        row = hit.iloc[0]
        for col in bs.columns:
            if col == STATEMENT_KEY_COL:
                continue
            v = safe_float(row[col])
            if v and v > 0:
                remember_total_shares(cache, code, v)
                return
    except Exception:
        pass
