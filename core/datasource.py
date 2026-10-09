# -*- coding: utf-8 -*-
"""
多数据源容错层：AkShare（主） → Baostock（备） → TuShare（备，需 token）

目标（对应需求“确保数据搜集齐全”）：
- 主源失败/字段缺失时，自动换源补齐，而不是直接留空；
- 三个源**全部可选**：未安装的包自动跳过，未配置 token 的 TuShare 自动跳过；
- 统一返回 pandas 结构，调用方无需关心底层差异。

各源能力对照
------------
| 数据            | AkShare（主）                      | Baostock（备）                    | TuShare（备）            |
|-----------------|-----------------------------------|----------------------------------|-------------------------|
| 三大报表        | stock_financial_report_sina ✅     | query_balance/profit/cashflow_data | 需积分，默认不用        |
| 年度财务指标    | stock_financial_analysis_indicator | query_profit_data（ROE/毛利等）   | fina_indicator          |
| 分红            | stock_history_dividend_detail      | ✗                                | dividend                |
| 个股行情        | 腾讯/新浪 HTTP ✅                  | query_history_k_data_plus        | daily                   |
"""
from __future__ import annotations

import os
import threading
from typing import Any, Callable, Dict, List, Optional

import pandas as pd

from .common import make_logger, safe_float

# Baostock 是全局单例式接口，必须串行 + 登录一次
_BS_LOCK = threading.Lock()
_BS_LOGGED = False
_TS_PRO = None


# ======================================================================
# 可用性探测
# ======================================================================
def baostock_available() -> bool:
    try:
        import baostock  # noqa: F401
        return True
    except Exception:
        return False


def tushare_available(token: str = "") -> bool:
    if not token:
        return False
    try:
        import tushare  # noqa: F401
        return True
    except Exception:
        return False


def source_status(token: str = "") -> Dict[str, bool]:
    """给前端展示的数据源可用性。"""
    return {"akshare": _ak_available(), "baostock": baostock_available(),
            "tushare": tushare_available(token)}


def _ak_available() -> bool:
    try:
        import akshare  # noqa: F401
        return True
    except Exception:
        return False


# ======================================================================
# Baostock
# ======================================================================
def bs_code(code: str) -> str:
    """6 位代码 → baostock 格式 sh.600887 / sz.000001。"""
    c = str(code).zfill(6)
    if c[0] == "6" or c.startswith("9"):
        return f"sh.{c}"
    if c[0] in ("0", "2", "3"):
        return f"sz.{c}"
    if c[0] in ("4", "8"):
        return f"bj.{c}"
    return f"sz.{c}"


def _bs_ensure_login(log: Optional[Callable[[str], None]] = None) -> bool:
    global _BS_LOGGED
    if not baostock_available():
        return False
    with _BS_LOCK:
        if _BS_LOGGED:
            return True
        try:
            import baostock as bs
            rs = bs.login()
            if getattr(rs, "error_code", "1") == "0":
                _BS_LOGGED = True
                return True
            (log or make_logger())(f"   ⚠️ Baostock 登录失败: {getattr(rs, 'error_msg', '')}")
        except Exception as e:  # noqa: BLE001
            (log or make_logger())(f"   ⚠️ Baostock 登录异常: {type(e).__name__}: {str(e)[:80]}")
    return False


def _bs_query_to_df(rs) -> Optional[pd.DataFrame]:
    """把 baostock 的 ResultData 转成 DataFrame。"""
    try:
        if getattr(rs, "error_code", "1") != "0":
            return None
        rows: List[List[Any]] = []
        while rs.next():
            rows.append(rs.get_row_data())
        if not rows:
            return None
        return pd.DataFrame(rows, columns=rs.fields)
    except Exception:
        return None


def bs_profit_indicators(code: str, years: List[int],
                         log: Optional[Callable[[str], None]] = None) -> Dict[int, Dict[str, Any]]:
    """Baostock 年度盈利能力：{年份: {roe, gross_margin, net_margin, eps, ...}}。

    query_profit_data 的字段含义（官方文档）：
      pubDate/statDate, roeAvg(净资产收益率均值), npMargin(销售净利率),
      gpMargin(销售毛利率), netProfit, epsTTM, MBRevenue, totalShare, liqaShare
    """
    out: Dict[int, Dict[str, Any]] = {}
    if not _bs_ensure_login(log):
        return out
    try:
        import baostock as bs
        with _BS_LOCK:
            for y in years:
                for q in (4, 3, 2, 1):     # 优先年报（第4季度）
                    try:
                        rs = bs.query_profit_data(code=bs_code(code), year=int(y), quarter=q)
                        df = _bs_query_to_df(rs)
                    except Exception:
                        df = None
                    if df is not None and not df.empty:
                        r = df.iloc[-1]
                        out[int(y)] = {
                            "净资产收益率(%)": safe_float(r.get("roeAvg")) * 100
                            if safe_float(r.get("roeAvg")) is not None and abs(safe_float(r.get("roeAvg"))) < 1.5
                            else safe_float(r.get("roeAvg")),
                            "销售毛利率(%)": safe_float(r.get("gpMargin")) * 100
                            if safe_float(r.get("gpMargin")) is not None and abs(safe_float(r.get("gpMargin"))) < 1.5
                            else safe_float(r.get("gpMargin")),
                            "销售净利率(%)": safe_float(r.get("npMargin")) * 100
                            if safe_float(r.get("npMargin")) is not None and abs(safe_float(r.get("npMargin"))) < 1.5
                            else safe_float(r.get("npMargin")),
                            "总资产(元)": safe_float(r.get("totalShare")),
                            "来源": "Baostock",
                        }
                        break
    except Exception as e:  # noqa: BLE001
        (log or make_logger())(f"   ⚠️ Baostock 指标获取失败[{code}]: {type(e).__name__}: {str(e)[:80]}")
    return out


def bs_income_annual(code: str, years: List[int],
                     log: Optional[Callable[[str], None]] = None) -> Dict[int, Dict[str, float]]:
    """Baostock 年度利润表关键项（用于毛利率兜底计算）。"""
    out: Dict[int, Dict[str, float]] = {}
    if not _bs_ensure_login(log):
        return out
    try:
        import baostock as bs
        with _BS_LOCK:
            for y in years:
                for q in (4, 3, 2, 1):
                    try:
                        rs = bs.query_profit_data(code=bs_code(code), year=int(y), quarter=q)
                        df = _bs_query_to_df(rs)
                    except Exception:
                        df = None
                    if df is not None and not df.empty:
                        r = df.iloc[-1]
                        rev = safe_float(r.get("MBRevenue"))
                        if rev:
                            out[int(y)] = {"营业收入": rev}
                            break
    except Exception:
        pass
    return out


# ======================================================================
# TuShare
# ======================================================================
def _ts_pro(token: str, log: Optional[Callable[[str], None]] = None):
    global _TS_PRO
    if _TS_PRO is not None:
        return _TS_PRO
    if not tushare_available(token):
        return None
    try:
        import tushare as ts
        ts.set_token(token)
        _TS_PRO = ts.pro_api()
        return _TS_PRO
    except Exception as e:  # noqa: BLE001
        (log or make_logger())(f"   ⚠️ TuShare 初始化失败: {type(e).__name__}: {str(e)[:80]}")
        return None


def ts_fina_indicator(code: str, years: List[int], token: str,
                      log: Optional[Callable[[str], None]] = None) -> Dict[int, Dict[str, Any]]:
    """TuShare 年度财务指标（需积分）。字段：roe/roe_waa/grossprofit_margin/netprofit_margin 等。"""
    out: Dict[int, Dict[str, Any]] = {}
    pro = _ts_pro(token, log)
    if pro is None:
        return out
    try:
        suffix = ".SH" if str(code).zfill(6)[0] in ("6", "9") else ".SZ"
        ts_code = str(code).zfill(6) + suffix
        start = f"{min(years)}0101"
        end = f"{max(years)}1231"
        df = pro.fina_indicator(ts_code=ts_code, start_date=start, end_date=end)
        if df is None or df.empty:
            return out
        df = df.copy()
        df["__y"] = df["end_date"].astype(str).str[:4]
        for _, r in df.iterrows():
            y = int(r["__y"])
            out[y] = {
                "净资产收益率(%)": safe_float(r.get("roe")),
                "加权净资产收益率(%)": safe_float(r.get("roe_waa")),
                "销售毛利率(%)": safe_float(r.get("grossprofit_margin")),
                "销售净利率(%)": safe_float(r.get("netprofit_margin")),
                "来源": "TuShare",
            }
    except Exception as e:  # noqa: BLE001
        (log or make_logger())(f"   ⚠️ TuShare 指标获取失败[{code}]: {type(e).__name__}: {str(e)[:80]}")
    return out


def ts_dividend(code: str, token: str,
                log: Optional[Callable[[str], None]] = None) -> List[Dict[str, Any]]:
    """TuShare 分红送转（cash_div_tax = 每股税前分红，元）。"""
    rows: List[Dict[str, Any]] = []
    pro = _ts_pro(token, log)
    if pro is None:
        return rows
    try:
        suffix = ".SH" if str(code).zfill(6)[0] in ("6", "9") else ".SZ"
        df = pro.dividend(ts_code=str(code).zfill(6) + suffix)
        if df is None or df.empty:
            return rows
        for _, r in df.iterrows():
            ex = str(r.get("ex_date") or "")
            year = int(ex[:4]) if len(ex) >= 4 and ex[:4].isdigit() else None
            cash = safe_float(r.get("cash_div_tax"))
            rows.append({
                "年份": year,
                "每10股派息(元)": (cash * 10) if cash is not None else None,
                "方案进度": str(r.get("div_proc") or ""),
                "除权除息日": ex,
                "公告日期": str(r.get("ann_date") or ""),
                "来源": "TuShare",
            })
    except Exception as e:  # noqa: BLE001
        (log or make_logger())(f"   ⚠️ TuShare 分红获取失败[{code}]: {type(e).__name__}: {str(e)[:80]}")
    return rows


# ======================================================================
# 统一入口：指标 / 毛利率
# ======================================================================
def annual_indicators_merged(code: str, years: List[int],
                             ak_annual: Optional[Dict[int, Dict[str, Any]]] = None,
                             ds_cfg: Optional[Dict[str, Any]] = None,
                             log: Optional[Callable[[str], None]] = None
                             ) -> Dict[int, Dict[str, Any]]:
    """把 AkShare 的年度指标与 Baostock/TuShare 合并，**逐字段补齐**（AkShare 优先）。"""
    log = log or make_logger()
    ds = ds_cfg or {}
    merged: Dict[int, Dict[str, Any]] = {int(k): dict(v) for k, v in (ak_annual or {}).items()}

    # 判断哪些年份/字段还缺
    def _missing() -> bool:
        for y in years:
            r = merged.get(int(y))
            if not r:
                return True
            if safe_float(r.get("销售毛利率(%)")) is None:
                return True
            if safe_float(r.get("净资产收益率(%)")) is None and safe_float(r.get("加权净资产收益率(%)")) is None:
                return True
        return False

    if not _missing():
        return merged

    if ds.get("enable_baostock", True) and baostock_available():
        bs_data = bs_profit_indicators(code, years, log)
        for y, r in bs_data.items():
            cur = merged.setdefault(int(y), {})
            for k, v in r.items():
                if v is None:
                    continue
                if safe_float(cur.get(k)) is None:
                    cur[k] = v
        log(f"      ℹ️ 已用 Baostock 补齐年度指标（{len(bs_data)} 年）")

    if _missing() and ds.get("enable_tushare") and ds.get("tushare_token"):
        ts_data = ts_fina_indicator(code, years, ds["tushare_token"], log)
        for y, r in ts_data.items():
            cur = merged.setdefault(int(y), {})
            for k, v in r.items():
                if v is None:
                    continue
                if safe_float(cur.get(k)) is None:
                    cur[k] = v
        log(f"      ℹ️ 已用 TuShare 补齐年度指标（{len(ts_data)} 年）")
    return merged


def gross_margin_from_sources(code: str, years: List[int],
                              ak_map: Optional[Dict[int, float]] = None,
                              ds_cfg: Optional[Dict[str, Any]] = None,
                              log: Optional[Callable[[str], None]] = None) -> Dict[int, float]:
    """毛利率多源兜底：AkShare 利润表 → Baostock gpMargin → TuShare grossprofit_margin。"""
    out: Dict[int, float] = {int(k): v for k, v in (ak_map or {}).items() if v is not None}
    ds = ds_cfg or {}
    need = [y for y in years if out.get(int(y)) is None]
    if not need:
        return out
    if ds.get("enable_baostock", True) and baostock_available():
        for y, r in bs_profit_indicators(code, need, log).items():
            v = safe_float(r.get("销售毛利率(%)"))
            if v is not None and out.get(int(y)) is None:
                out[int(y)] = round(v, 2)
    need = [y for y in years if out.get(int(y)) is None]
    if need and ds.get("enable_tushare") and ds.get("tushare_token"):
        for y, r in ts_fina_indicator(code, need, ds["tushare_token"], log).items():
            v = safe_float(r.get("销售毛利率(%)"))
            if v is not None and out.get(int(y)) is None:
                out[int(y)] = round(v, 2)
    return out


def dividends_merged(code: str, base_rows: List[Dict[str, Any]],
                     ds_cfg: Optional[Dict[str, Any]] = None,
                     log: Optional[Callable[[str], None]] = None) -> List[Dict[str, Any]]:
    """分红多源合并：AkShare(同花顺+东财) 为主，TuShare 仅在缺年份时补充。"""
    ds = ds_cfg or {}
    rows = list(base_rows or [])
    if not (ds.get("enable_tushare") and ds.get("tushare_token")):
        return rows
    have_years = {r.get("年份") for r in rows if r.get("年份")}
    try:
        ts_rows = ts_dividend(code, ds["tushare_token"], log)
    except Exception:
        ts_rows = []
    for r in ts_rows:
        if r.get("年份") not in have_years and r.get("每10股派息(元)"):
            rows.append(r)
            have_years.add(r.get("年份"))
    return rows
