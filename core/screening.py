# -*- coding: utf-8 -*-
"""
海选公司（免费数据版）

原桌面版用 Selenium 打开同花顺问财抓多条件选股。云端无浏览器，改为：
1. **候选池**：AkShare 全市场清单 → 按市值降序取前 N 只（默认 300，可调），可选“排除北交所/金融股”；
2. **逐只取年度财务指标**：`ak.stock_financial_analysis_indicator`（新浪，按年/按报告期），
   得到 净资产收益率 / 加权净资产收益率 / 销售毛利率 / 经营现金净流量与净利润的比率 等；
3. **按条件过滤**：支持“连续 N 年”口径（最近 N 个完整会计年度全部满足）；
4. **产出与原程序一致**：
   - `海选情况/海选公司汇总表.xlsx`（列：代码 / 名称 / 筛选渠道）
   - `海选情况/问财_表1_多条件初始表_{ts}.xlsx`、`问财_表2_多条件整理表_{ts}.xlsx`、`问财_表3_代码及名称汇总.xlsx`
"""
from __future__ import annotations

import os
import re
import time
from concurrent.futures import ThreadPoolExecutor, as_completed
from typing import Any, Callable, Dict, List, Optional, Tuple

import pandas as pd

from .common import DiskCache, make_logger, safe_float
from .config import DEFAULT_WENCAI_QUERY
from .contracts import HAIXUAN_SUMMARY, HAIXUAN_SUMMARY_COLUMNS, DIR_HAIXUAN
from .universe import StockUniverse, get_industry, industry_label

# 判定“金融股”的行业关键词（与原程序“非金融股”口径对应）
_FINANCE_KEYS = ("银行", "保险", "证券", "多元金融", "信托", "期货", "金融服务", "资本市场")

# 问财原始字段 → 本项目指标列名
INDICATOR_ALIASES: Dict[str, List[str]] = {
    "roe_weighted": ["加权净资产收益率(%)", "净资产收益率(%)"],
    "roe": ["净资产收益率(%)", "加权净资产收益率(%)"],
    "gross_margin": ["销售毛利率(%)"],
    "cash_content": ["经营现金净流量与净利润的比率(%)"],
    "net_margin": ["销售净利率(%)"],
    "debt_ratio": ["资产负债率(%)"],
    "revenue_growth": ["主营业务收入增长率(%)"],
    "profit_growth": ["净利润增长率(%)"],
}


# ======================================================================
# 条件解析
# ======================================================================
def parse_query(query: str, default_years: int = 5) -> Dict[str, Any]:
    """把“连续5年加权roe>20，连续5年毛利率>40，上市时间>3年，剔除北交所，非金融股”这类
    文字条件解析成结构化条件。解析不到的条目会被忽略并记录在 unparsed。
    """
    q = str(query or "").replace("％", "%").replace("，", ",")
    conds: Dict[str, Any] = {"years": default_years, "exclude_bj": False, "exclude_finance": False,
                             "min_roe": None, "min_gross_margin": None, "min_cash_content": None,
                             "min_listed_years": None, "min_market_cap": None, "max_pe": None,
                             "unparsed": []}

    # 剔除类
    if "北交所" in q:
        conds["exclude_bj"] = True
    if "非金融" in q or "剔除金融" in q or "不含金融" in q:
        conds["exclude_finance"] = True

    # 连续N年
    m = re.search(r"连续\s*(\d+)\s*年", q)
    if m:
        conds["years"] = max(1, min(10, int(m.group(1))))

    def _num(pattern: str) -> Optional[float]:
        mm = re.search(pattern, q, re.IGNORECASE)
        if not mm:
            return None
        try:
            return float(mm.group(1))
        except Exception:
            return None

    conds["min_roe"] = _num(r"(?:加权)?\s*roe\s*[>＞=]+\s*([\d\.]+)")
    if conds["min_roe"] is None:
        conds["min_roe"] = _num(r"净资产收益率\s*[>＞=]+\s*([\d\.]+)")
    conds["min_gross_margin"] = _num(r"毛利率\s*[>＞=]+\s*([\d\.]+)")
    conds["min_cash_content"] = _num(r"净利润现金含量\s*[>＞=]+\s*([\d\.]+)")
    conds["min_listed_years"] = _num(r"上市时间\s*[>＞=]+\s*([\d\.]+)")
    conds["min_market_cap"] = _num(r"市值\s*[>＞=]+\s*([\d\.]+)")
    conds["max_pe"] = _num(r"(?:市盈率|pe)\s*[<＜=]+\s*([\d\.]+)")

    known = ["roe", "毛利率", "净利润现金含量", "上市时间", "市值", "市盈率", "北交所", "金融", "连续"]
    for part in [x for x in re.split(r"[,\n;；]", q) if x.strip()]:
        if not any(k in part.lower() for k in known):
            conds["unparsed"].append(part.strip())
    return conds


def conditions_to_text(c: Dict[str, Any]) -> str:
    parts = []
    if c.get("min_roe") is not None:
        parts.append(f"连续{c['years']}年加权ROE>{c['min_roe']}")
    if c.get("min_cash_content") is not None:
        parts.append(f"连续{c['years']}年净利润现金含量>{c['min_cash_content']}")
    if c.get("min_gross_margin") is not None:
        parts.append(f"连续{c['years']}年毛利率>{c['min_gross_margin']}")
    if c.get("min_listed_years") is not None:
        parts.append(f"上市时间>{c['min_listed_years']}年")
    if c.get("min_market_cap") is not None:
        parts.append(f"市值>{c['min_market_cap']}亿")
    if c.get("max_pe") is not None:
        parts.append(f"市盈率<{c['max_pe']}")
    if c.get("exclude_bj"):
        parts.append("剔除北交所")
    if c.get("exclude_finance"):
        parts.append("非金融股")
    return "，".join(parts)


# ======================================================================
# 财务指标
# ======================================================================
def fetch_indicators(code: str, cache: DiskCache, start_year: int,
                     log: Optional[Callable[[str], None]] = None) -> Optional[pd.DataFrame]:
    """取年度财务指标（按日期的年报记录过滤，保留 YYYY-12-31 或年度行）。"""
    log = log or make_logger()
    key = f"indicator_{code}_{start_year}"
    cached = cache.get(key)
    if cached:
        try:
            df = pd.DataFrame(cached)
            if not df.empty:
                return df
        except Exception:
            pass
    try:
        import akshare as ak
        df = ak.stock_financial_analysis_indicator(symbol=str(code).zfill(6), start_year=str(start_year))
        if df is None or df.empty:
            return None
        df = df.copy()
        df["日期"] = df["日期"].astype(str)
        cache.set(key, df.to_dict("records"))
        return df
    except Exception as e:  # noqa: BLE001
        log(f"      ⚠️ 财务指标获取失败[{code}]: {type(e).__name__}: {str(e)[:80]}")
        return None


def _annual_rows(df: pd.DataFrame, years: int) -> Dict[int, Dict[str, Any]]:
    """从指标表里挑出最近 N 个完整会计年度的年报行（日期形如 2024-12-31）。"""
    out: Dict[int, Dict[str, Any]] = {}
    if df is None or df.empty or "日期" not in df.columns:
        return out
    annual = df[df["日期"].astype(str).str.contains("-12-31", na=False)]
    if annual.empty:
        annual = df
    annual = annual.copy()
    annual["__y"] = annual["日期"].astype(str).str[:4]
    annual = annual.sort_values("__y")
    for _, r in annual.tail(years).iterrows():
        try:
            out[int(r["__y"])] = r.to_dict()
        except Exception:
            continue
    return out


def _pick(rec: Dict[str, Any], keys: List[str]) -> Optional[float]:
    for k in keys:
        if k in rec:
            v = safe_float(rec.get(k))
            if v is not None:
                return v
    return None


# 现金含量口径归一：新浪该字段虽是“(%)”但返回的是**倍数**（如 0.6056 表示 0.6056 倍）。
# 用户条件写的是“净利润现金含量>100”（百分数），因此需要统一乘 100。
_CASH_CONTENT_SCALE = 100.0


def _pick_cash_content(rec: Dict[str, Any], merged: Dict[int, Dict[str, Any]], year: int) -> Optional[float]:
    """取净利润现金含量并统一成百分数。"""
    v = _pick(rec, INDICATOR_ALIASES["cash_content"])
    if v is None:
        v = _pick(merged.get(int(year), {}) or {}, INDICATOR_ALIASES["cash_content"])
    if v is None:
        return None
    return round(v * _CASH_CONTENT_SCALE, 2)


def _pick_merged(rec: Dict[str, Any], keys: List[str], merged: Dict[int, Dict[str, Any]],
                 year: int) -> Optional[float]:
    """先从 AkShare 年度行取，取不到再从多源合并结果（Baostock/TuShare）取。"""
    v = _pick(rec, keys)
    if v is not None:
        return v
    return _pick(merged.get(int(year), {}) or {}, keys)


def _fix_code(v: Any) -> str:
    """代码归一：防止 pandas 把 '000568' 读成整数 568，必须补足 6 位。"""
    s = str(v).strip()
    if s.endswith(".0"):
        s = s[:-2]
    s = re.sub(r"\D", "", s)
    return s.zfill(6) if s else ""


def gross_margin_map(code: str, cache: DiskCache, years: List[int],
                     log: Optional[Callable[[str], None]] = None) -> Dict[int, float]:
    """毛利率兜底：AkShare 的“销售毛利率”字段在新浪源里常年为空，
    这里直接用财务报表口径计算 毛利率 =（营业收入 − 营业成本）/ 营业收入 × 100%。
    """
    out: Dict[int, float] = {}
    try:
        from .financial_search import build_statement_sheet, fetch_statement_raw
        is_ = build_statement_sheet(fetch_statement_raw(code, "利润表", cache, log=None), years, log=None)
        if is_ is None or is_.empty:
            return out
        key = "项目"

        def _row(names):
            for n in names:
                hit = is_[is_[key] == n]
                if not hit.empty:
                    return hit.iloc[0]
            return None

        rev = _row(["营业收入", "营业总收入"])
        cost = _row(["营业成本", "其中：营业成本"])
        if rev is None or cost is None:
            return out
        for col in is_.columns:
            if col == key:
                continue
            m = re.match(r"^(20\d{2})", str(col))
            if not m:
                continue
            rv, cv = safe_float(rev[col]), safe_float(cost[col])
            if rv and cv is not None and rv > 0:
                out[int(m.group(1))] = round((rv - cv) / rv * 100, 2)
    except Exception as e:  # noqa: BLE001
        (log or make_logger())(f"      ⚠️ 计算毛利率失败[{code}]: {type(e).__name__}: {str(e)[:80]}")
    return out


# ======================================================================
# 候选池
# ======================================================================
def build_candidate_pool(cond: Dict[str, Any], universe: StockUniverse, cache: DiskCache,
                         pool_size: int = 300,
                         log: Optional[Callable[[str], None]] = None) -> List[Dict[str, Any]]:
    """构造候选公司池：全市场按市值降序 → 排除北交所/金融股 → 取前 pool_size 只。"""
    log = log or make_logger()
    df = universe.load()
    if df is None or df.empty:
        return []

    pool: List[Dict[str, Any]] = []
    df = df.copy()
    df["code"] = df["code"].astype(str).str.zfill(6)
    if cond.get("exclude_bj"):
        df = df[~df["code"].str.startswith(("4", "8", "92"))]

    log(f"   候选池构建：全市场 {len(df)} 只，正在按市值排序…")
    # 【提速】先把已缓存的市值收集起来；未缓存的用**批量接口**一次补齐（腾讯行情支持多代码同查）
    from .universe import get_market_cap, get_market_caps_batch
    enriched: List[Tuple[str, str, float]] = []
    need: List[str] = []
    for _, row in df.iterrows():
        code = _fix_code(row["code"])
        name = str(row.get("name", ""))
        hit = cache.get(f"mktcap_{code}")
        if hit is not None:
            try:
                enriched.append((code, name, float(hit)))
                continue
            except Exception:
                pass
        need.append(code)
    enriched.sort(key=lambda x: -x[2])
    # 只给“还未拿到市值”的股票补数，最多补 pool_size*6 只，避免整市场空转
    need = need[: max(pool_size * 6, 200)]
    if need:
        log(f"   正在批量获取 {len(need)} 只候选股的市值…")
        caps = get_market_caps_batch(need, cache, log=None)
        for _, row in df.iterrows():
            code = _fix_code(row["code"])
            if code in caps:
                enriched.append((code, str(row.get("name", "")), float(caps[code])))
    enriched.sort(key=lambda x: -x[2])

    for code, name, cap in enriched:
        if cond.get("min_market_cap") is not None and cap < float(cond["min_market_cap"]):
            continue
        if cond.get("exclude_finance"):
            ind = industry_label(get_industry(code, cache, log=None))
            if any(k in ind for k in _FINANCE_KEYS):
                continue
        pool.append({"code": code, "name": name, "市值(亿)": cap})
        if len(pool) >= pool_size:
            break
    log(f"   候选池就绪：{len(pool)} 只（按市值降序）")
    return pool


# ======================================================================
# 主流程
# ======================================================================
def run_screening(query: str, output_root: str, cache: DiskCache,
                  pool_size: int = 300, workers: int = 6,
                  log: Optional[Callable[[str], None]] = None,
                  progress: Optional[Callable[[int, int, str], None]] = None,
                  extra_codes: Optional[List[Dict[str, Any]]] = None,
                  ds_cfg: Optional[Dict[str, Any]] = None) -> Dict[str, Any]:
    """执行海选，产出与原程序同名同结构的四张表。

    extra_codes: 额外强制纳入的公司（例如手工指定），[{"code","name"}]
    ds_cfg:      数据源开关（enable_baostock / enable_tushare / tushare_token 等）
    """
    log = log or make_logger()
    ds_cfg = ds_cfg or {}
    try:
        from .datasource import source_status
        log("   数据源可用性: " + str(source_status(ds_cfg.get("tushare_token", ""))))
    except Exception:
        pass
    cond = parse_query(query or DEFAULT_WENCAI_QUERY)
    log(f"🎯 海选条件：{conditions_to_text(cond) or '（未解析出任何阈值条件，仅按候选池输出）'}")
    if cond["unparsed"]:
        log(f"   ℹ️ 未能解析的条件片段（已忽略）：{cond['unparsed']}")

    out_dir = os.path.join(output_root, DIR_HAIXUAN)
    os.makedirs(out_dir, exist_ok=True)
    ts = time.strftime("%Y%m%d_%H%M%S")

    this_year = time.localtime().tm_year
    start_year = this_year - max(cond["years"] + 1, 3)

    universe = StockUniverse(cache, log)
    pool = build_candidate_pool(cond, universe, cache, pool_size=pool_size, log=log)
    if extra_codes:
        exist = {p["code"] for p in pool}
        for e in extra_codes:
            c = str(e.get("code", "")).zfill(6)
            if c and c not in exist:
                pool.insert(0, {"code": c, "name": e.get("name", ""), "市值(亿)": None})

    if not pool:
        log("❌ 候选池为空，无法海选（可能是网络受限或市值接口不可用）。")
        return {"ok": False, "stocks": [], "files": [], "conditions": cond}

    log(f"🔎 正在对 {len(pool)} 只候选股逐一取年度财务指标并过滤（并发 {workers}）…")

    results: List[Dict[str, Any]] = []
    detail_rows: List[Dict[str, Any]] = []
    done = 0

    def _work(item: Dict[str, Any]) -> Dict[str, Any]:
        code, name = item["code"], item.get("name", "")
        df = fetch_indicators(code, cache, start_year, log=None)
        annual = _annual_rows(df, cond["years"]) if df is not None else {}
        years_have = sorted(annual.keys())
        # 多源合并（AkShare → Baostock → TuShare），逐年逐字段补齐
        merged: Dict[int, Dict[str, Any]] = {}
        try:
            from .datasource import annual_indicators_merged
            if years_have:
                merged = annual_indicators_merged(code, years_have, annual, ds_cfg=ds_cfg, log=None)
        except Exception:
            merged = {}
        # 毛利率兜底（新浪指标的“销售毛利率”常为空，用利润表/多源口径计算）
        gm_fallback: Dict[int, float] = {}
        if cond["min_gross_margin"] is not None:
            try:
                from .datasource import gross_margin_from_sources
                gm_fallback = gross_margin_from_sources(
                    code, list(range(max(start_year, 2000), this_year + 1)),
                    ak_map=gross_margin_map(code, cache, list(range(max(start_year, 2000), this_year + 1))),
                    ds_cfg=ds_cfg, log=None)
            except Exception:
                gm_fallback = gross_margin_map(code, cache,
                                               list(range(max(start_year, 2000), this_year + 1)))
        rec: Dict[str, Any] = {
            "代码": code, "名称": name, "市值(亿)": item.get("市值(亿)"),
            "年度数": len(years_have), "年度": ",".join(str(y) for y in years_have),
        }
        # 每年指标
        roe_vals, gm_vals, cc_vals = [], [], []
        for y in years_have[-cond["years"]:]:
            r = annual[y]
            roe = _pick_merged(r, INDICATOR_ALIASES["roe_weighted"], merged, y)
            gm = _pick_merged(r, INDICATOR_ALIASES["gross_margin"], merged, y)
            if gm is None:
                gm = gm_fallback.get(int(y))
            cc = _pick_cash_content(r, merged, y)
            rec[f"ROE{y}"] = roe
            rec[f"毛利率{y}"] = gm
            rec[f"现金含量{y}"] = cc
            if roe is not None:
                roe_vals.append(roe)
            if gm is not None:
                gm_vals.append(gm)
            if cc is not None:
                cc_vals.append(cc)

        fails: List[str] = []
        n = cond["years"]
        if cond["min_roe"] is not None:
            if len(roe_vals) < n or min(roe_vals) < cond["min_roe"]:
                fails.append(f"ROE未连续{n}年≥{cond['min_roe']}")
        if cond["min_gross_margin"] is not None:
            if len(gm_vals) < n or min(gm_vals) < cond["min_gross_margin"]:
                fails.append(f"毛利率未连续{n}年≥{cond['min_gross_margin']}")
        if cond["min_cash_content"] is not None:
            if len(cc_vals) < n or min(cc_vals) < cond["min_cash_content"]:
                fails.append(f"现金含量未连续{n}年≥{cond['min_cash_content']}")

        rec["ROE最低"] = min(roe_vals) if roe_vals else None
        rec["毛利率最低"] = min(gm_vals) if gm_vals else None
        rec["现金含量最低"] = min(cc_vals) if cc_vals else None
        rec["是否入选"] = "是" if not fails else "否"
        rec["未通过原因"] = "；".join(fails)
        return rec

    with ThreadPoolExecutor(max_workers=max(1, workers)) as ex:
        futures = {ex.submit(_work, it): it for it in pool}
        for fut in as_completed(futures):
            done += 1
            it = futures[fut]
            if progress:
                progress(done, len(pool), f"{it.get('name','')}({it['code']})")
            try:
                rec = fut.result()
            except Exception as e:  # noqa: BLE001
                log(f"      ⚠️ {it['code']} 处理异常: {type(e).__name__}: {str(e)[:80]}")
                continue
            detail_rows.append(rec)
            if rec["是否入选"] == "是":
                results.append({"代码": rec["代码"], "名称": rec["名称"], "筛选渠道": "免费数据(多条件)"})

    # 排序：ROE 最低值降序，便于人工复核
    detail_rows.sort(key=lambda r: (-(r.get("ROE最低") or -999), r["代码"]))
    results.sort(key=lambda r: r["代码"])
    log(f"✅ 海选完成：候选 {len(pool)} 只 → 入选 {len(results)} 只")

    # 产出四张表（与原程序同名同义）
    files: List[str] = []
    summary_df = pd.DataFrame(results, columns=HAIXUAN_SUMMARY_COLUMNS)
    p_sum = os.path.join(out_dir, HAIXUAN_SUMMARY)
    summary_df.to_excel(p_sum, index=False)
    files.append(p_sum)

    detail_df = pd.DataFrame(detail_rows)
    p1 = os.path.join(out_dir, f"问财_表1_多条件初始表_{ts}.xlsx")
    detail_df.to_excel(p1, index=False)
    files.append(p1)

    keep_cols = [c for c in ["代码", "名称", "市值(亿)", "ROE最低", "毛利率最低", "现金含量最低",
                             "是否入选", "未通过原因", "年度"] if c in detail_df.columns]
    p2 = os.path.join(out_dir, f"问财_表2_多条件整理表_{ts}.xlsx")
    detail_df[keep_cols].to_excel(p2, index=False)
    files.append(p2)

    p3 = os.path.join(out_dir, "问财_表3_代码及名称汇总.xlsx")
    summary_df.to_excel(p3, index=False)
    files.append(p3)
    for f in files:
        log(f"   ✅ 已生成 {os.path.basename(f)}")

    return {"ok": True, "stocks": results, "files": files, "conditions": cond,
            "condition_text": conditions_to_text(cond), "candidates": len(pool),
            "detail": detail_rows, "output_dir": out_dir}
