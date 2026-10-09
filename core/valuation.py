# -*- coding: utf-8 -*-
"""
好价估值分析（移植自桌面版 main_logic_c / judge_stock / navigate_to_legulegu 体系）

与桌面版的差异（**只换数据来源，判断规则逐字保留**）
--------------------------------------------------
桌面版用 Selenium 打开乐咕乐股抓大盘指数 PE、打开雪球抓个股现价/PE/股息率。
云端没有浏览器，改为：
- 大盘/指数 PE 与分位：AkShare `stock_market_pe_lg`（深证A股/上证A股）与 `stock_a_ttm_lyr`（全A分位）；
- 中美国债收益率、美联储利率：`bond_zh_us_rate` / `macro_bank_usa_interest_rate`；
- 个股现价 / 市盈率(TTM)：腾讯、新浪行情 HTTP 接口（与桌面版修复版同一套逻辑）；
- 个股动态股息率：**严格 TTM**（近 12 个月“已实施/已除权”分红累加 ÷ 现价），数据来自东财分红接口；
- 判断规则 `judge_stock`、报告版式、颜色标注全部与桌面版一致。
"""
from __future__ import annotations

import os
import re
import time
from dataclasses import dataclass, field
from datetime import datetime, timedelta
from typing import Any, Callable, Dict, List, Optional, Tuple

import pandas as pd
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.shared import Inches, Pt, RGBColor

from .common import DiskCache, http_get, make_logger, safe_float
from .config import VALUATION_DEFAULTS

# ----------------------------------------------------------------------
# 行情（与桌面版修复版一致的腾讯/新浪双源）
# ----------------------------------------------------------------------
_QUOTE_HEADERS = {
    "User-Agent": ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
                   "(KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"),
}


def quote_symbol(user_input: str, cache: Optional[DiskCache] = None,
                 log: Optional[Callable[[str], None]] = None) -> Tuple[str, str]:
    """把“公司简称/代码”翻成行情符号。返回 (symbol, 市场)。

    解析顺序：纯代码 → 含代码文本 → 字母代码(美股) → 全市场清单反查中文简称 → 腾讯智能搜索。
    """
    raw = str(user_input or "").strip()
    if not raw:
        return "", "unknown"
    s = re.sub(r"\s+", "", raw)

    def _from_code(code: str) -> Tuple[str, str]:
        code = str(code).strip()
        if re.fullmatch(r"\d{6}", code):
            if code[0] in ("6", "9"):
                return "sh" + code, "A股"
            if code[0] in ("0", "2", "3"):
                return "sz" + code, "A股"
            if code[0] in ("4", "8"):
                return "bj" + code, "A股"
        if re.fullmatch(r"\d{4,5}", code):
            return "hk" + code.zfill(5), "港股"
        return "", "unknown"

    if re.fullmatch(r"\d{4,6}", s):
        sym, mkt = _from_code(s)
        if sym:
            return sym, mkt
    m = re.search(r"(\d{6})", s)
    if m:
        sym, mkt = _from_code(m.group(1))
        if sym:
            return sym, mkt
    if re.fullmatch(r"[A-Za-z][A-Za-z\.\-]{0,9}", s):
        return "us" + s.upper(), "美股"

    # 中文简称 → 先用全市场清单反查（零额外网络请求）
    if cache is not None:
        try:
            from .universe import StockUniverse
            code = StockUniverse(cache, log).code_of(s)
            if code:
                sym, mkt = _from_code(code)
                if sym:
                    return sym, mkt
        except Exception:
            pass
    # 再退到腾讯智能搜索（带重试）
    try:
        r = http_get("https://smartbox.gtimg.cn/s3/", params={"v": "2", "q": s, "t": "all"},
                     timeout=8, encoding="gbk", retries=3, label=f"代码解析[{s}]", log=log)
        if r is not None:
            mm = re.search(r'v_hint="([^"]*)"', r.text)
            if mm and mm.group(1):
                first = mm.group(1).split("^")[0].split("~")
                if len(first) >= 3:
                    mk, cd = first[0], first[1].split(".")[0]
                    code = cd if mk in ("sh", "sz", "bj") else (cd.zfill(5) if mk == "hk" else cd.upper())
                    sym, mkt2 = _from_code(code)
                    if sym:
                        return sym, mkt2
    except Exception:
        pass
    return "", "unknown"


def fetch_price_pe(symbol: str, log: Optional[Callable[[str], None]] = None) -> Dict[str, Any]:
    """腾讯→新浪 取现价与市盈率(TTM)。"""
    log = log or make_logger()
    # 腾讯
    for attempt in range(3):
        try:
            r = http_get(f"https://qt.gtimg.cn/q={symbol}", timeout=10, encoding="gbk", retries=2,
                         label=f"腾讯行情[{symbol}]", log=log)
            if r is not None:
                m = re.search(r'="([^"]+)"', r.text)
                if m:
                    f = m.group(1).split("~")
                    if len(f) > 6 and f[1]:
                        price = safe_float(f[3]) or safe_float(f[4])
                        pe = safe_float(f[39]) if len(f) > 39 else None
                        return {"来源": "腾讯行情", "名称": f[1], "现价": price, "PE_TTM": pe}
        except Exception:
            pass
        time.sleep(0.8 * (attempt + 1))
    # 新浪
    try:
        r = http_get(f"https://hq.sinajs.cn/list={symbol}",
                     headers={**_QUOTE_HEADERS, "Referer": "https://finance.sina.com.cn"},
                     timeout=10, encoding="gb18030", retries=3, label=f"新浪行情[{symbol}]", log=log)
        if r is not None:
            m = re.search(r'="([^"]*)"', r.text)
            if m and m.group(1).strip():
                f = m.group(1).split(",")
                if len(f) >= 4 and f[0]:
                    return {"来源": "新浪行情", "名称": f[0],
                            "现价": safe_float(f[3]) or safe_float(f[1]), "PE_TTM": None}
    except Exception:
        pass
    return {}


def fetch_ttm_dividend_yield(symbol: str, price: Optional[float],
                            log: Optional[Callable[[str], None]] = None) -> Tuple[Optional[float], str]:
    """严格 TTM 动态股息率 = 近12个月“已实施(已除权)”每股分红累加 ÷ 现价 ×100%。"""
    log = log or make_logger()
    code_m = re.search(r"(\d{6})", symbol)
    if not code_m or not price:
        return None, ""
    code = code_m.group(1)
    try:
        r = http_get("https://datacenter-web.eastmoney.com/api/data/v1/get", params={
            "reportName": "RPT_SHAREBONUS_DET",
            "columns": ("SECURITY_CODE,REPORT_DATE,EX_DIVIDEND_DATE,ASSIGN_PROGRESS,"
                        "IMPL_PLAN_PROFILE,PRETAX_BONUS_RMB"),
            "filter": f'(SECURITY_CODE="{code}")',
            "pageNumber": 1, "pageSize": 60, "sortColumns": "REPORT_DATE", "sortTypes": -1,
            "source": "WEB", "client": "WEB",
        }, timeout=12, retries=3, label=f"东财分红[{code}]", log=log)
        if r is None:
            return None, ""
        rows = (((r.json() or {}).get("result") or {}).get("data")) or []
    except Exception:
        return None, ""

    today = datetime.now().date()
    start = today - timedelta(days=365)
    total, hits = 0.0, []

    def _d(v):
        if not v:
            return None
        try:
            return datetime.strptime(str(v)[:10], "%Y-%m-%d").date()
        except Exception:
            return None

    for it in rows:
        ex, rd = _d(it.get("EX_DIVIDEND_DATE")), _d(it.get("REPORT_DATE"))
        prog = str(it.get("ASSIGN_PROGRESS") or "")
        used, note = None, ""
        if "实施" in prog:
            if ex is None or not (start <= ex <= today):
                continue
            used, note = ex, "除权"
        elif any(w in prog for w in ("董事会", "股东大会", "预案", "未实施")):
            continue
        elif ex is not None:
            if not (start <= ex <= today):
                continue
            used, note = ex, "除权"
        elif rd is not None and (today - rd).days >= 300:
            if not (start - timedelta(days=60) <= rd <= today):
                continue
            used, note = rd, "报告期"
        else:
            continue
        v = safe_float(it.get("PRETAX_BONUS_RMB"))
        if v is None:
            continue
        per_share = v / 10.0 if v > 1 else v
        if per_share <= 0:
            continue
        total += per_share
        hits.append(f"{used.isoformat()}{note} {per_share:.4g}元/股")
    if total > 0:
        return round(total / price * 100, 2), "近12个月已实施分红累加(" + "；".join(hits) + ")"
    return None, "近12个月无已实施分红"


# ----------------------------------------------------------------------
# 宏观与大盘
# ----------------------------------------------------------------------
def _latest_two(df: pd.DataFrame, col: str) -> Tuple[Optional[float], Optional[float]]:
    try:
        s = pd.to_numeric(df[col], errors="coerce").dropna()
        if s.empty:
            return None, None
        if len(s) == 1:
            return float(s.iloc[-1]), None
        return float(s.iloc[-1]), float(s.iloc[-2])
    except Exception:
        return None, None


def _fetch_a_market_pe(log: Optional[Callable[[str], None]] = None) -> Tuple[Optional[float], str]:
    """取 A 股整体市盈率（用于替代桌面版浏览器抓取的“深圳A股 PE”）。

    说明：桌面版用 Selenium 打开乐咕乐股取“深圳A股”PE。云端逐级降级取数：
    ① `stock_a_ttm_lyr` 的**滚动市盈率中位数** —— 官方口径的“A股整体估值”，最稳定；
    ② `stock_index_pe_lg('沪深300')` 的**等权滚动市盈率** —— 覆盖面接近全市场；
    ③ `stock_zh_index_value_csindex('000300')` 的 市盈率2。
    返回值第 2 项是口径说明，会写进日志和界面，方便你核对与桌面版数值的差异。
    """
    log = log or make_logger()
    # ① 全A 滚动市盈率中位数
    try:
        import akshare as ak
        df = ak.stock_a_ttm_lyr()
        if df is not None and not df.empty:
            for col in ("middlePETTM", "averagePETTM"):
                if col in df.columns:
                    val, _ = _latest_two(df, col)
                    if val:
                        return round(val, 2), f"A股整体{col}(stock_a_ttm_lyr)"
    except Exception as e:  # noqa: BLE001
        log(f"   ⚠️ stock_a_ttm_lyr 失败: {type(e).__name__}: {str(e)[:80]}")
    # ② 沪深300 等权滚动市盈率
    try:
        import akshare as ak
        df = ak.stock_index_pe_lg(symbol="沪深300")
        if df is not None and not df.empty:
            for col in ("等权滚动市盈率", "滚动市盈率中位数", "滚动市盈率"):
                if col in df.columns:
                    val, _ = _latest_two(df, col)
                    if val:
                        return round(val, 2), f"沪深300 {col}(stock_index_pe_lg)"
    except Exception as e:  # noqa: BLE001
        log(f"   ⚠️ stock_index_pe_lg 失败: {type(e).__name__}: {str(e)[:80]}")
    # ③ 中证指数官网
    try:
        import akshare as ak
        df = ak.stock_zh_index_value_csindex(symbol="000300")
        if df is not None and not df.empty and "市盈率2" in df.columns:
            val = safe_float(df.iloc[0]["市盈率2"])
            if val:
                return round(val, 2), "沪深300 市盈率2(中证指数)"
    except Exception as e:  # noqa: BLE001
        log(f"   ⚠️ csindex 失败: {type(e).__name__}: {str(e)[:80]}")
    return None, ""


def fetch_macro(cfg: Dict[str, Any], log: Optional[Callable[[str], None]] = None) -> Dict[str, Any]:
    """取宏观与大盘指标（全部免费接口，逐项容错）。"""
    log = log or make_logger()
    out: Dict[str, Any] = {
        "cn_bond_yield": None, "us_bond_yield": None, "fed_rate": None,
        "shenzhen_pe": None, "shanghai_pe": None, "all_a_pe": None,
        "sp500_pe": None, "hangseng_pe": None, "pe_source": "",
    }

    # 中美国债收益率
    if cfg.get("cn_10y", True) or cfg.get("us_10y", True):
        try:
            import akshare as ak
            df = ak.bond_zh_us_rate(start_date=(datetime.now() - timedelta(days=180)).strftime("%Y%m%d"))
            if df is not None and not df.empty:
                if cfg.get("cn_10y", True) and "中国国债收益率10年" in df.columns:
                    cn, _ = _latest_two(df, "中国国债收益率10年")
                    out["cn_bond_yield"] = cn
                if cfg.get("us_10y", True) and "美国国债收益率10年" in df.columns:
                    us, _ = _latest_two(df, "美国国债收益率10年")
                    out["us_bond_yield"] = us
        except Exception as e:  # noqa: BLE001
            log(f"   ⚠️ 中美国债收益率获取失败: {type(e).__name__}: {str(e)[:100]}")

    # 美联储利率
    if cfg.get("fed", True):
        try:
            import akshare as ak
            df = ak.macro_bank_usa_interest_rate()
            if df is not None and not df.empty and "今值" in df.columns:
                val, _ = _latest_two(df, "今值")
                out["fed_rate"] = val
        except Exception as e:  # noqa: BLE001
            log(f"   ⚠️ 美联储利率获取失败: {type(e).__name__}: {str(e)[:100]}")

    # A 股整体 PE（替代“深圳A股 PE”的浏览器抓取）
    if cfg.get("sz_pe", True):
        pe, src = _fetch_a_market_pe(log)
        out["shenzhen_pe"] = pe
        out["shanghai_pe"] = pe
        out["pe_source"] = src
        if pe:
            log(f"   A股整体PE = {pe}（口径：{src}）")

    log(f"   宏观: 中国10年国债={out['cn_bond_yield']} 美国10年={out['us_bond_yield']} "
        f"美联储={out['fed_rate']} A股PE={out['shenzhen_pe']}【{out['pe_source']}】")
    return out


# ----------------------------------------------------------------------
# 判断规则（逐字移植桌面版 judge_stock）
# ----------------------------------------------------------------------
def judge_stock(macro_data: Dict[str, Any], stock: Dict[str, Any]) -> Tuple[str, str, Optional[float]]:
    pe = stock.get("当前PE(TTM)")
    dividend = stock.get("动态股息率")
    price = stock.get("现价")
    category = stock.get("分类")
    if pe is None or dividend is None or price is None or not category:
        return "数据不足", "black", None

    judgment, color = "目前为观察区", "black"
    target_price = (15 * price / pe) if pe > 0 else None

    if category == "A股":
        sz_pe, cn_bond = macro_data.get("shenzhen_pe"), macro_data.get("cn_bond_yield")
        if sz_pe is None or cn_bond is None:
            return "宏观数据不足", "black", None
        if (sz_pe < 20) and (pe < 15) and (dividend > cn_bond):
            judgment, color = "目前为好价格", "red"
        elif (sz_pe < 40) and (pe < 30) and (dividend > cn_bond * 2 / 3):
            judgment, color = "目前为偏买区", "yellow"
        elif (40 < sz_pe < 60) and (30 < pe < 50) and (cn_bond / 3 < dividend < cn_bond * 2 / 3):
            judgment, color = "目前为偏卖区", "blue"
    elif category == "港股":
        hk_pe, cn_bond = macro_data.get("hangseng_pe"), macro_data.get("cn_bond_yield")
        if hk_pe is None or cn_bond is None:
            return "宏观数据不足", "black", None
        if (hk_pe < 10) and (pe < 15) and (dividend > cn_bond):
            judgment, color = "目前为好价格", "red"
        elif (hk_pe < 20) and (pe < 30) and (dividend > cn_bond * 2 / 3):
            judgment, color = "目前为偏买区", "yellow"
        elif (20 < hk_pe < 30) and (30 < pe < 50) and (cn_bond / 3 < dividend < cn_bond * 2 / 3):
            judgment, color = "目前为偏卖区", "blue"
    elif category == "美股":
        us_bond, fed_rate, sp_pe = (macro_data.get("us_bond_yield"), macro_data.get("fed_rate"),
                                    macro_data.get("sp500_pe"))
        if us_bond is None or fed_rate is None or sp_pe is None:
            return "宏观数据不足", "black", None
        sp_drop = 0
        if ((fed_rate < 4 and sp_pe < 15 and pe < 15 and dividend > us_bond)
                or (fed_rate > 4 and sp_pe < 10 and pe < 15 and dividend > us_bond)
                or (sp_drop > 50 and pe < 15 and dividend > us_bond)):
            judgment, color = "目前为好价格", "red"
        elif ((fed_rate < 4 and sp_pe < 30 and pe < 30 and dividend > us_bond * 2 / 3)
              or (fed_rate > 4 and sp_pe < 20 and pe < 30 and dividend > us_bond * 2 / 3)
              or (sp_drop > 30 and pe < 30 and dividend > us_bond * 2 / 3)):
            judgment, color = "目前为偏买区", "yellow"
        elif ((fed_rate < 4 and sp_pe < 45 and pe < 45 and dividend > us_bond / 3)
              or (fed_rate > 4 and sp_pe < 45 and pe < 45 and dividend > us_bond / 3)
              or (sp_drop > 20 and pe < 45 and dividend > us_bond / 3)):
            judgment, color = "目前为偏卖区", "blue"

    if target_price is not None:
        judgment += f"；股价小于{target_price:.2f}为好价格"
    return judgment, color, target_price


# ----------------------------------------------------------------------
# 报告
# ----------------------------------------------------------------------
def _set_cn_font(run, name: str = "微软雅黑") -> None:
    try:
        run.font.name = name
        run._element.rPr.rFonts.set(qn("w:eastAsia"), name)
    except Exception:
        pass


def create_word_report(output_path: str, records: List[Dict[str, Any]],
                       macro_data: Dict[str, Any], log: Optional[Callable[[str], None]] = None) -> str:
    """生成《深度资产估值分析报告》（版式与桌面版一致）。"""
    doc = Document()
    title = doc.add_paragraph()
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = title.add_run("资产深度估值分析报告")
    run.bold = True
    run.font.size = Pt(16)
    _set_cn_font(run)

    p = doc.add_paragraph()
    sz, cn, hk, sp, us, fed = (macro_data.get("shenzhen_pe"), macro_data.get("cn_bond_yield"),
                               macro_data.get("hangseng_pe"), macro_data.get("sp500_pe"),
                               macro_data.get("us_bond_yield"), macro_data.get("fed_rate"))
    p.add_run(f"深圳A股PE: {sz if sz else 'N/A'} | 中国10年国债: {cn if cn else 'N/A'}%\n")
    p.add_run(f"恒生指数PE: {hk if hk else 'N/A'}\n")
    p.add_run(f"标普500 PE: {sp if sp else 'N/A'} | 美国10年国债: {us if us else 'N/A'}% | "
              f"美联储利率: {fed if fed else 'N/A'}%")

    headers = ["分类", "名称/代码", "现价", "当前PE(TTM)", "动态股息率", "判断结果", "反推好价格上限", "来源文件"]
    table = doc.add_table(rows=1, cols=len(headers))
    table.style = "Table Grid"
    for i, h in enumerate(headers):
        cell = table.rows[0].cells[i]
        cell.text = h
        if cell.paragraphs[0].runs:
            cell.paragraphs[0].runs[0].bold = True
    for rec in records:
        cells = table.add_row().cells
        cells[0].text = str(rec.get("分类", ""))
        cells[1].text = str(rec.get("名称/代码", ""))
        cells[2].text = "" if rec.get("现价") is None else f"{rec.get('现价'):.2f}"
        cells[3].text = "" if rec.get("当前PE(TTM)") is None else f"{rec.get('当前PE(TTM)'):.2f}"
        cells[4].text = "" if rec.get("动态股息率") is None else f"{rec.get('动态股息率'):.2f}%"
        cells[5].text = str(rec.get("判断结果", ""))
        cells[6].text = "" if rec.get("反推好价格上限") is None else f"{rec.get('反推好价格上限'):.2f}"
        cells[7].text = os.path.basename(output_path)
        color = rec.get("颜色", "black")
        if color != "black":
            rgb = {"red": RGBColor(0xC0, 0x00, 0x00), "yellow": RGBColor(0xBF, 0x8F, 0x00),
                   "blue": RGBColor(0x1F, 0x4E, 0x79)}.get(color, RGBColor(0, 0, 0))
            for para in cells[5].paragraphs:
                for r in para.runs:
                    r.font.color.rgb = rgb
                    r.bold = True
    doc.save(output_path)
    (log or make_logger())(f"   ✅ 已生成 {os.path.basename(output_path)}")
    return output_path


# ----------------------------------------------------------------------
# 主流程
# ----------------------------------------------------------------------
def run_valuation(cfg: Dict[str, Any], output_dir: str,
                  log: Optional[Callable[[str], None]] = None,
                  progress: Optional[Callable[[int, int, str], None]] = None) -> Dict[str, Any]:
    """好价估值全流程：宏观/大盘 → 个股现价·PE·TTM股息率 → 判断 → 出报告。

    cfg 可含：a_stocks / hk_stocks / us_stocks（列表或逗号串）、sz_pe/hs_pe/sp_pe/cn_10y/us_10y/fed。
    也支持 `targets`: [{"code","name"}] 形式（由年报汇总结果直接带入）。
    """
    log = log or make_logger()
    os.makedirs(output_dir, exist_ok=True)
    merged = {**VALUATION_DEFAULTS, **(cfg or {})}

    def _as_list(v: Any) -> List[str]:
        if v is None:
            return []
        if isinstance(v, str):
            return [x.strip() for x in v.replace("，", ",").split(",") if x.strip()]
        return [str(x).strip() for x in v if str(x).strip()]

    a_list = _as_list(merged.get("a_stocks"))
    hk_list = _as_list(merged.get("hk_stocks"))
    us_list = _as_list(merged.get("us_stocks"))
    # 支持直接传入目标公司（来自年报汇总）：优先用代码，避免“名称/代码”混排
    for t in (merged.get("targets") or []):
        code = str(t.get("code") or "").strip()
        nm = str(t.get("name") or "").strip()
        key = code if re.fullmatch(r"\d{6}", code) else nm
        if key and key not in a_list:
            a_list.append(key)

    log("🔍 好价估值分析启动：先取宏观与大盘指标…")
    macro = fetch_macro(merged, log)

    groups = [("A股", a_list), ("港股", hk_list), ("美股", us_list)]
    total = sum(len(x) for _, x in groups) or 1
    done = 0
    records: List[Dict[str, Any]] = []
    rows: List[Dict[str, Any]] = []

    log("🔍 正在逐只获取个股现价 / PE(TTM) / TTM动态股息率…")
    for cat, items in groups:
        for raw in items:
            done += 1
            if progress:
                progress(done, total, raw)
            sym, mkt = quote_symbol(raw, cache=merged.get("_cache"), log=log)
            item: Dict[str, Any] = {"分类": cat, "名称/代码": raw, "现价": None,
                                    "当前PE(TTM)": None, "动态股息率": None,
                                    "行情来源": "", "股息率来源": ""}
            if not sym:
                log(f"   ⚠️ 无法识别标的：{raw}")
                rows.append(item)
                continue
            q = fetch_price_pe(sym, log)
            if q:
                item["现价"] = q.get("现价")
                item["当前PE(TTM)"] = q.get("PE_TTM")
                item["行情来源"] = q.get("来源", "")
                if q.get("名称"):
                    item["名称"] = q["名称"]
            div, tip = fetch_ttm_dividend_yield(sym, item.get("现价"), log)
            item["动态股息率"] = div
            item["股息率来源"] = tip
            if item["现价"] is not None:
                log(f"   ✅ {raw}: 现价={item['现价']} PE(TTM)={item['当前PE(TTM)']} "
                    f"股息率={item['动态股息率']}%（{tip[:40]}）")
            else:
                log(f"   ❌ {raw}: 未能取到行情（{sym}）")
            rows.append(item)

    # 判断
    for item in rows:
        judgment, color, target = judge_stock(macro, item)
        rec = {
            "分类": item["分类"], "名称/代码": item["名称/代码"],
            "现价": item["现价"], "当前PE(TTM)": item["当前PE(TTM)"],
            "动态股息率": item["动态股息率"], "判断结果": judgment,
            "颜色": color, "反推好价格上限": target,
            "行情来源": item.get("行情来源", ""), "股息率来源": item.get("股息率来源", ""),
        }
        records.append(rec)

    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    out_path = os.path.join(output_dir, f"深度资产估值分析报告_{ts}.docx")
    create_word_report(out_path, records, macro, log)

    # 可读性更好的明细表（供前端展示/下载）
    detail_path = os.path.join(output_dir, f"估值明细_{ts}.xlsx")
    try:
        pd.DataFrame([{
            "分类": r["分类"], "名称/代码": r["名称/代码"], "现价": r["现价"],
            "当前PE(TTM)": r["当前PE(TTM)"], "动态股息率(%)": r["动态股息率"],
            "判断结果": r["判断结果"], "反推好价格上限": r["反推好价格上限"],
            "行情来源": r["行情来源"], "股息率口径": r["股息率来源"],
        } for r in records]).to_excel(detail_path, index=False)
    except Exception as e:  # noqa: BLE001
        log(f"   ⚠️ 明细表导出失败: {e}")

    ok = any(r["现价"] is not None for r in records)
    log(f"{'✅' if ok else '❌'} 好价估值分析完成：{len(records)} 条记录")
    return {"ok": ok, "report": out_path, "detail": detail_path,
            "macro": macro, "records": records, "output_dir": output_dir}
