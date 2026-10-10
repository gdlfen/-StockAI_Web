# -*- coding: utf-8 -*-
"""
企业AI深度分析（云端版）

移植自桌面版 EnterpriseAIAnalyzer（原“AI 企业深度分析”标签页，
见 价值投资分析模型V1.4-3（扩展）1.0（修复版）.py 第 6835~7147 行）。

与桌面版的差异**只有数据来源**：
- 取消“自动联动下载年报”模式与 PDF 依赖，本地资料改为读取本项目产出的表格
  （《统一整合输出_*.xlsx》的 合并资产负债表 / 合并利润表 / 合并现金流量表 / 分红情况，
   以及单年《年报提取表》），统一压成紧凑文本；
- 目录下若仍有 PDF / DOCX，照桌面版一样尝试解析（try/except 容错），缺 PDF 不会失败；
- 所有 GUI / messagebox / tk / self.after 调用一律改为 log(...) 回调。

以下行为与桌面版**逐字保留**（不得改动）：
- 提示词模板与拼装顺序：template.format(company=..., competitors='、'.join(peers),
  web_info=web_info[:8000], doc_info=local_material[:15000])
- 系统提示词 "资深金融分析师"、temperature=0.3
- 联网检索的两个查询：f"{company} 最新深度研究报告" / f"{company} 最新公告 {peers} 对比"
- 报告命名：{clean_company}_深度报告_{HHMM}.docx，输出到 {输出根}/企业分析/
- 失败排错日志：{company}_生成失败排错日志（写 docx，不中断整批任务）
- 公司目录名 {代码}_{简称} 取简称：name.split('_')[-1]
"""
from __future__ import annotations

import os
import random
import re
import time
from datetime import datetime
from typing import Any, Callable, Dict, List, Optional, Tuple

# ----------------------------------------------------------------------
# 项目内依赖（只允许 core.* / 标准库 / 白名单第三方）
# ----------------------------------------------------------------------
try:  # 包内导入：from core.enterprise_ai import ...
    from . import common
    from . import contracts
    from .config import AI_DEFAULTS, DEFAULT_ENTERPRISE_PROMPT
except Exception:  # pragma: no cover - 把 core 目录直接加进 sys.path 时
    import common  # type: ignore
    import contracts  # type: ignore
    from config import AI_DEFAULTS, DEFAULT_ENTERPRISE_PROMPT  # type: ignore

# 桌面版提示词模板逐字保留（core/config.py::DEFAULT_ENTERPRISE_PROMPT）
DEFAULT_TEMPLATE: str = DEFAULT_ENTERPRISE_PROMPT

# 与桌面版 run_batch_analysis 完全一致的截断长度
WEB_INFO_LIMIT = 8000
DOC_INFO_LIMIT = 15000
# 采集阶段内部上限：只做内存保护，最终仍按桌面版 doc_info[:15000] 截断
MATERIAL_HARD_LIMIT = 20000
SYSTEM_PROMPT = "资深金融分析师"
DEFAULT_TEMPERATURE = 0.3
FAIL_LOG_SUFFIX = "生成失败排错日志"

# 采集哪些类型的本地资料
DOC_EXTENSIONS = (".pdf", ".docx")
EXCEL_EXTENSIONS = (".xlsx", ".xlsm")

# 优先读取的表格 sheet（顺序即文本顺序，与桌面版“合并三大表 + 分红”口径一致）
PRIMARY_SHEET_ORDER: Tuple[str, ...] = (
    contracts.UNIFIED_BS_SHEET,   # 合并资产负债表
    contracts.UNIFIED_IS_SHEET,   # 合并利润表
    contracts.UNIFIED_CF_SHEET,   # 合并现金流量表
    contracts.UNIFIED_DIV_SHEET,  # 分红情况
)
# 《统一整合输出》里除三大表与分红之外的附加 sheet（员工情况等），允许但优先级低
EXTRA_SHEET_ORDER: Tuple[str, ...] = (
    contracts.UNIFIED_EMP_SHEET,  # 员工情况(PDF提取)
)

# 分红情况表列名（契约：contracts.DIV_COLUMNS）
DIV_COLUMNS: Tuple[str, ...] = tuple(getattr(
    contracts, "DIV_COLUMNS",
    ["股票代码", "年份", "同花顺现金分红总额(元)", "归属于母公司所有者的净利润(元)", "分红率"]))
# 员工情况表列名（契约：contracts.EMP_COLUMNS = 年份 / 项目 / 数值）
EMP_COLUMNS: Tuple[str, ...] = tuple(getattr(contracts, "EMP_COLUMNS", ["年份", "项目", "数值"]))
# 逐行（逐年）呈现的 sheet：一行一个年份，不能按“项目: 期间=值”渲染
ROW_WISE_SHEETS: Tuple[str, ...] = (contracts.UNIFIED_DIV_SHEET, contracts.UNIFIED_EMP_SHEET)

# 单表/整份资料的字符预算（按优先级分配，保证核心三大表一定进得去）
SHEET_CHAR_BUDGET = 4200
EXTRA_SHEET_CHAR_BUDGET = 1200
YEARLY_FILE_CHAR_BUDGET = 3000

# 目录名里出现这些关键字说明不是公司目录（与桌面版跳过“新建文件夹”一致，并排除本项目产物目录）
SKIP_DIR_KEYS: Tuple[str, ...] = ("新建文件夹",) + tuple(contracts.SCAN_EXCLUDE_DIRS)

# 公司文件夹命名约定 {代码}_{简称}
COMPANY_DIR_RE = re.compile(r"^\s*\d{6}[_\-].+$")

# 表格里明显是表头/汇总说明的行，丢弃
_SKIP_ROW_KEYS = ("项目", "科目", "报表日期", "单位", "合计项", "注：", "注:")


# ======================================================================
# 日志小工具（禁止任何 GUI：统一走 log 回调 + stdout）
# ======================================================================
def _logger(log: Optional[Callable[[str], None]]) -> Callable[[str], None]:
    """统一日志出口：既打到 stdout（云端日志可见）又回调给上层（如任务队列）。

    **幂等包装**：``run_all`` 会把自己包装后的 logger 再传给 ``scan_company_tasks``、
    ``collect_local_material`` 等函数，而它们内部又调用 ``_logger`` —— 层层嵌套后
    同一条消息会被打印 2~6 次（实测日志里出现 ``[22:44:59] [22:44:59] [22:44:59] xxx``）。
    因此这里给包装函数打上标记，已是包装过的直接返回，保证整条链路只打印一次。

    另外保留同秒去重，防止上层 ``common.make_logger`` 再加一层前缀。
    """
    mk = getattr(common, "make_logger", None)
    if not callable(mk):
        return log or (lambda _m: None)
    # 已包装过 → 直接复用，避免嵌套重复输出
    if getattr(log, "_ea_logger_wrapped", False):
        return log
    inner = mk(log)
    state: Dict[str, Any] = {"last": None, "at": 0.0}

    def _once(msg: str) -> None:
        text = str(msg)
        now = time.time()
        if text == state["last"] and (now - state["at"]) < 1.0:
            return
        state["last"] = text
        state["at"] = now
        inner(text)

    try:
        _once._ea_logger_wrapped = True      # type: ignore[attr-defined]
    except Exception:  # noqa: BLE001
        pass
    return _once


def _to_float(x: Any) -> Optional[float]:
    fn = getattr(common, "safe_float", None)
    if callable(fn):
        try:
            return fn(x)
        except Exception:
            return None
    try:
        return float(str(x).replace(",", "").replace("%", "").strip())
    except Exception:
        return None


def _norm_text(x: Any) -> str:
    fn = getattr(common, "norm_text", None)
    if callable(fn):
        try:
            return fn(x)
        except Exception:
            pass
    return "" if x is None else str(x).strip()


def _ensure_dir(path: str) -> str:
    fn = getattr(common, "ensure_dir", None)
    if callable(fn):
        return fn(path)
    os.makedirs(path, exist_ok=True)
    return path


def _safe_filename(name: str) -> str:
    """与桌面版 re.sub(r'[\\\\/:*?\"<>|]', '', x) 一致。"""
    fn = getattr(contracts, "safe_filename", None)
    if callable(fn):
        return fn(name)
    return re.sub(r'[\\/:*?"<>|]', "", str(name or "")).strip()


def _strip_html(text: str) -> str:
    return re.sub(r"<[^>]+>", "", text or "").strip()


def _cell_text(v: Any) -> str:
    """单元格 -> 紧凑文本；空值返回空串。"""
    if v is None:
        return ""
    try:
        import pandas as pd  # 局部导入，避免模块级副作用
        if isinstance(v, float) and pd.isna(v):
            return ""
        if v is pd.NaT:
            return ""
    except Exception:
        pass
    if isinstance(v, float):
        if v != v:  # NaN
            return ""
        if v.is_integer():
            return str(int(v))
        return f"{v:.4f}".rstrip("0").rstrip(".")
    if isinstance(v, int):
        return str(v)
    s = str(v).strip().replace("\n", " ").replace("\r", " ")
    return "" if s.lower() in ("nan", "nat", "none", "null") else s


def _is_number_text(s: str, raw: Any) -> bool:
    if isinstance(raw, (int, float)) and not isinstance(raw, bool):
        return True
    if not s:
        return False
    return _to_float(s) is not None


# ======================================================================
# 1. 联网检索（保留桌面版 search_web 行为；优先 duckduckgo_search/ddgs，失败返回空串）
# ======================================================================
def _ddgs_results(query: str, max_res: int, timeout: int) -> List[Dict[str, Any]]:
    """优先用 duckduckgo_search / ddgs；不可用或抛错则返回空列表（不抛出）。"""
    DDGS = None
    try:
        from duckduckgo_search import DDGS as _D  # type: ignore
        DDGS = _D
    except Exception:
        try:
            from ddgs import DDGS as _D  # type: ignore
            DDGS = _D
        except Exception:
            DDGS = None
    if DDGS is None:
        return []
    for be in ("duckduckgo", "html", "lite"):
        try:
            kwargs: Dict[str, Any] = {"timeout": max(5, int(timeout))}
            with DDGS(**kwargs) as ddgs:
                hits = ddgs.text(query, region="cn-zh", safesearch="off",
                                 max_results=max_res, backend=be)
            hits = list(hits or [])
            if hits:
                return hits
        except Exception:
            continue
    return []


def _scrape_360(query: str, max_res: int, timeout: int) -> str:
    """桌面版首选通道：360 搜索（urllib + 正则，失败返回空串）。"""
    import urllib.parse
    import urllib.request

    url = f"https://www.so.com/s?q={urllib.parse.quote(query)}"
    req = urllib.request.Request(
        url, headers={"User-Agent": "Mozilla/5.0", "Accept": "*/*"})
    html = urllib.request.urlopen(req, timeout=timeout).read().decode("utf-8", "ignore")
    blocks = html.split('<li class="res-list')
    info = ""
    count = 0
    for block in blocks[1:max_res + 2]:
        title_match = re.search(r"<h3[^>]*>(.*?)</h3>", block, re.S)
        if not title_match:
            continue
        title = _strip_html(title_match.group(1))
        desc_match = (re.search(r'class="res-desc[^"]*"[^>]*>(.*?)</', block, re.S)
                      or re.search(r'class="res-rich[^"]*"[^>]*>(.*?)</', block, re.S))
        desc = _strip_html(desc_match.group(1)) if desc_match else "无摘要"
        if title:
            info += f"- [{title}]: {desc}\n"
            count += 1
            if count >= max_res:
                break
    return info


def _scrape_bing(query: str, max_res: int, timeout: int) -> str:
    """桌面版兜底通道：必应中文（urllib + 正则，失败返回空串）。"""
    import urllib.parse
    import urllib.request

    url = f"https://cn.bing.com/search?q={urllib.parse.quote(query)}"
    req = urllib.request.Request(url, headers={
        "User-Agent": "Mozilla/5.0",
        "Cookie": "SRCHHPGUSR=SRCHLANG=zh-Hans;",
    })
    html = urllib.request.urlopen(req, timeout=timeout).read().decode("utf-8", "ignore")
    blocks = re.findall(r"<h2[^>]*><a[^>]*>(.*?)</a></h2>(.*?)(?=<h2|$)", html, re.S)
    info = ""
    count = 0
    for title_raw, content_raw in blocks:
        title = _strip_html(title_raw)
        desc_match = re.search(r"<p[^>]*>(.*?)</p>", content_raw, re.S)
        desc = _strip_html(desc_match.group(1)) if desc_match else "无摘要"
        if title and count < max_res:
            info += f"- [{title}]: {desc}\n"
            count += 1
    return info


def search_web(query: str, max_res: int = 3, timeout: int = 15,
               log: Optional[Callable[[str], None]] = None) -> str:
    """联网检索，返回 "- [标题]: 摘要" 文本；全程容错，失败返回空串。

    与桌面版 EnterpriseAIAnalyzer.search_web 的行为一致（含 UA / 正则 / 摘要回退“无摘要”），
    差别只在云端优先走 duckduckgo_search/ddgs（桌面版走 360→必应，且带本地代理假设）。
    """
    log = _logger(log)
    text = ""
    query = str(query or "").strip()
    if not query:
        return ""
    sleep_fn = getattr(common, "sleep_between", None)  # 云端默认不睡；桌面版为反爬会 sleep(random 1.0~2.5)
    if callable(sleep_fn):
        try:
            sleep_fn()
        except Exception:
            pass
    else:
        try:
            time.sleep(random.uniform(0.2, 0.6))
        except Exception:
            pass

    # 第一优先：duckduckgo_search / ddgs
    for hit in _ddgs_results(query, max_res, timeout):
        title = _norm_text(hit.get("title") or hit.get("heading") or "")
        desc = _norm_text(hit.get("body") or hit.get("snippet") or hit.get("description") or "")
        if not title:
            continue
        text += f"- [{title}]: {desc or '无摘要'}\n"
    if text:
        return text

    # 第二优先：360 搜索（桌面版首选，保留其解析逻辑）
    try:
        text = _scrape_360(query, max_res, timeout)
        if text:
            log(f"检索命中（360 搜索）: {query}")
            return text
    except Exception as e:  # noqa: BLE001 - 桌面版此处为“记录并跳过”
        log(f"解析发生错误，已跳过该项: {str(e)}")

    # 第三优先：必应中文（桌面版兜底）
    try:
        text = _scrape_bing(query, max_res, timeout)
        if text:
            log(f"检索命中（必应）: {query}")
            return text
    except Exception as e:  # noqa: BLE001
        log(f"解析发生错误，已跳过该项: {str(e)}")

    log(f"⚠️ 联网检索无结果，已跳过: {query}")
    return ""


# ======================================================================
# 2. 本地资料采集：表格优先（替代桌面版的“PDF 年报”）
# ======================================================================
def _list_files(company_dir: str) -> List[str]:
    try:
        return sorted(f for f in os.listdir(company_dir)
                      if os.path.isfile(os.path.join(company_dir, f)))
    except Exception:
        return []


def _list_dirs(root: str) -> List[str]:
    try:
        return sorted(d for d in os.listdir(root) if os.path.isdir(os.path.join(root, d)))
    except Exception:
        return []


def _pick_unified_excel(files: List[str]) -> Optional[str]:
    """优先最新的《统一整合输出_*.xlsx》。"""
    prefix = getattr(contracts, "UNIFIED_FILENAME_PREFIX", "统一整合输出")
    hits = [f for f in files
            if f.lower().endswith(EXCEL_EXTENSIONS)
            and prefix in f
            and not f.startswith("~$")]
    if not hits:
        return None
    def _mtime(name: str) -> float:
        try:
            return os.path.getmtime(name)
        except Exception:
            return 0.0
    hits.sort(key=_mtime)
    return hits[-1]


def _year_from_name(filename: str) -> int:
    fn = getattr(contracts, "parse_year_from_yearly_filename", None)
    if callable(fn):
        try:
            y = fn(filename)
            if y:
                return int(y)
        except Exception:
            pass
    m = re.search(r"(\d{4})年度", filename or "")
    return int(m.group(1)) if m else 0


def _pick_yearly_excels(files: List[str]) -> List[str]:
    """兜底：按年份升序取单年《年报提取表》。"""
    key = getattr(contracts, "YEARLY_FILENAME_TEMPLATE", "{name}（{code}）{year}年度年报提取表.xlsx")
    suffix = str(key).split("}")[-1]           # "年度年报提取表.xlsx"
    stem = suffix.replace(".xlsx", "")
    hits = [f for f in files
            if f.lower().endswith(EXCEL_EXTENSIONS)
            and stem in f
            and not f.startswith("~$")]
    hits.sort(key=_year_from_name)
    return hits


def _read_sheet_df(path: str, sheet: str):
    """读取单个 sheet，失败返回 None（不抛出）。"""
    try:
        import pandas as pd
        with pd.ExcelFile(path, engine="openpyxl") as xls:
            names = [str(n) for n in xls.sheet_names]
            if sheet not in names:
                return None
            df = xls.parse(sheet)
    except Exception:
        return None
    if df is None or getattr(df, "empty", True):
        return None
    try:
        df = df.dropna(how="all").dropna(axis=1, how="all")
    except Exception:
        pass
    return None if getattr(df, "empty", True) else df


def _sheet_to_text(df, sheet_name: str, source: str, budget: int) -> str:
    """把一张财务表转成“项目: 期间=值; 期间=值”的紧凑文本。

    分红情况 / 员工情况是“一行一记录”的结构（年份在行里而非列里），
    另有专用渲染分支，否则会退化成没有列名的裸值。
    """
    if df is None or getattr(df, "empty", True):
        return ""
    try:
        cols = [str(c) for c in list(df.columns)]
    except Exception:
        return ""
    if not cols:
        return ""
    if sheet_name in ROW_WISE_SHEETS:
        return _row_wise_sheet_to_text(df, sheet_name, source, budget)

    key_col = cols[0]

    out_lines: List[str] = []
    used = 0
    truncated = False
    for _, row in df.iterrows():
        try:
            key = _cell_text(row.get(key_col, df.columns[0]))
        except Exception:
            key = ""
        if not key or key.startswith("Unnamed"):
            continue
        if len(key) > 40:
            key = key[:40]
        if any(key.startswith(k) for k in _SKIP_ROW_KEYS):
            continue

        pairs: List[str] = []
        for c, col in zip(cols[1:], list(df.columns)[1:]):
            raw = row.get(col)
            val = _cell_text(raw)
            if val == "":
                continue
            label = c if _is_number_text(c, col) else "值"
            pairs.append(f"{label}={val}")
        if not pairs:
            continue

        line = f"{key}: " + "; ".join(pairs)
        if used + len(line) + 4 > budget:
            truncated = True
            break
        out_lines.append(line)
        used += len(line) + 4

    if not out_lines:
        return ""
    return _finish_sheet(out_lines, truncated, source, sheet_name)


def _div_sheet_to_text(df, source: str, budget: int) -> str:
    """分红情况：一行一条记录 —— “2024年: 同花顺现金分红总额=…; 归母净利润=…; 分红率=…”。"""
    try:
        cols = list(df.columns)
        raw_names = [str(c) for c in cols]
    except Exception:
        return ""

    def _find(want: str) -> Any:
        if want in raw_names:
            return cols[raw_names.index(want)]
        for i, c in enumerate(raw_names):
            if want and (want in c or c in want):
                return cols[i]
        return None

    c_year, c_cash = _find(DIV_COLUMNS[1]), _find(DIV_COLUMNS[2])
    c_profit, c_ratio = _find(DIV_COLUMNS[3]), _find(DIV_COLUMNS[4])
    if c_year is None:
        return ""

    out_lines: List[str] = []
    used = 0
    truncated = False
    for _, row in df.iterrows():
        year = _cell_text(row.get(c_year))
        if not year:
            continue
        pairs: List[str] = []
        for label, col in ((DIV_COLUMNS[2], c_cash), (DIV_COLUMNS[3], c_profit),
                           (DIV_COLUMNS[4], c_ratio)):
            if col is None:
                continue
            val = _cell_text(row.get(col))
            if val != "":
                pairs.append(f"{label}={val}")
        if not pairs:
            continue
        line = f"{year}年: " + "; ".join(pairs)
        if used + len(line) + 2 > budget:
            truncated = True
            break
        out_lines.append(line)
        used += len(line) + 2
    return _finish_sheet(out_lines, truncated, source, "分红情况")


def _emp_sheet_to_text(df, source: str, budget: int) -> str:
    """员工情况(PDF提取)：列 = 年份 / 项目 / 数值。"""
    try:
        cols = list(df.columns)
        raw_names = [str(c) for c in cols]
    except Exception:
        return ""

    def _idx(want: str, fallback: int) -> Any:
        if want in raw_names:
            return cols[raw_names.index(want)]
        return cols[fallback] if fallback < len(cols) else None

    c_year = _idx(EMP_COLUMNS[0], 0)
    c_item = _idx(EMP_COLUMNS[1], 1)
    c_val = _idx(EMP_COLUMNS[2], 2)
    if c_item is None or c_val is None:
        return ""

    # 年份升序，便于“按年”阅读
    try:
        order = sorted(range(len(df)),
                       key=lambda i: _to_float(df.iloc[i].get(c_year)) or 0)
    except Exception:
        order = list(range(len(df)))

    by_year: Dict[str, List[str]] = {}
    seq: List[str] = []
    for i in order:
        row = df.iloc[i]
        year = _cell_text(row.get(c_year)) or "未知年份"
        item = _cell_text(row.get(c_item))
        val = _cell_text(row.get(c_val))
        if not item or val == "":
            continue
        if year not in by_year:
            by_year[year] = []
            seq.append(year)
        by_year[year].append(f"{item}={val}")

    out_lines: List[str] = []
    used = 0
    truncated = False
    for year in seq:
        line = f"{year}年: " + "; ".join(by_year[year])
        if used + len(line) + 2 > budget:
            truncated = True
            break
        out_lines.append(line)
        used += len(line) + 2
    return _finish_sheet(out_lines, truncated, source, "员工情况(PDF提取)")


def _row_wise_sheet_to_text(df, sheet_name: str, source: str, budget: int) -> str:
    if sheet_name == contracts.UNIFIED_DIV_SHEET:
        return _div_sheet_to_text(df, source, budget)
    if sheet_name == contracts.UNIFIED_EMP_SHEET:
        return _emp_sheet_to_text(df, source, budget)
    return ""


def _finish_sheet(out_lines: List[str], truncated: bool, source: str, sheet_name: str) -> str:
    if not out_lines:
        return ""
    body = "\n".join(out_lines)
    if truncated:
        body += "\n（该表内容较多，已截断）"
    return f"【{source}｜{sheet_name}】\n{body}"


def _append_section(sections: List[str], chunk: str, used: int, hard_limit: int) -> int:
    if not chunk:
        return used
    remain = hard_limit - used
    if remain <= 0:
        return used
    if len(chunk) + 4 > remain:
        chunk = chunk[:max(0, remain - 20)] + "\n（后续资料已按长度上限截断）"
    sections.append(chunk)
    return used + len(chunk) + 4


def collect_local_material(company_dir: str,
                           log: Optional[Callable[[str], None]] = None,
                           max_chars: int = MATERIAL_HARD_LIMIT) -> str:
    """把公司目录下的资料转成紧凑文本（替代桌面版的“读 PDF 年报”）。

    数据来源优先级：
      1) 《统一整合输出_*.xlsx》的 合并资产负债表 / 合并利润表 / 合并现金流量表 / 分红情况
      2) 单年《年报提取表》的同名 sheet（兜底）
      3) 目录下残留的 PDF / DOCX（照桌面版行为尝试解析，缺了不算失败）
    绝不因为“没有 PDF”而失败；任何单个文件解析失败都只记日志后跳过。
    """
    log = _logger(log)
    company_dir = str(company_dir or "")
    if not company_dir or not os.path.isdir(company_dir):
        log(f"⚠️ 本地资料目录不存在，按无本地资料继续: {company_dir}")
        return ""
    try:
        hard_limit = max(2000, int(max_chars))
    except Exception:
        hard_limit = MATERIAL_HARD_LIMIT

    files = _list_files(company_dir)
    sections: List[str] = []
    used = 0
    read_sheets: set = set()

    unified = _pick_unified_excel(files)
    if unified:
        path = os.path.join(company_dir, unified)
        log(f"📊 读取整合表格: {unified}")
        for sheet in PRIMARY_SHEET_ORDER:
            if used >= hard_limit:
                break
            df = _read_sheet_df(path, sheet)
            if df is None:
                continue
            chunk = _sheet_to_text(df, sheet, unified, SHEET_CHAR_BUDGET)
            if chunk:
                used = _append_section(sections, chunk, used, hard_limit)
                read_sheets.add(sheet)
        # 附加 sheet（员工情况等）仅在还有余量时读
        for sheet in EXTRA_SHEET_ORDER:
            if used >= hard_limit:
                break
            if sheet in read_sheets:
                continue
            df = _read_sheet_df(path, sheet)
            if df is None:
                continue
            chunk = _sheet_to_text(df, sheet, unified, EXTRA_SHEET_CHAR_BUDGET)
            if chunk:
                used = _append_section(sections, chunk, used, hard_limit)
    else:
        yearly = _pick_yearly_excels(files)
        if yearly:
            log(f"📊 未发现《统一整合输出》，改用单年《年报提取表》{len(yearly)} 份")
        for name in yearly:
            if used >= hard_limit:
                break
            path = os.path.join(company_dir, name)
            year = _year_from_name(name)
            # 三大表：逐年追加（这是单年《年报提取表》的价值所在）
            for sheet in PRIMARY_SHEET_ORDER:
                if used >= hard_limit:
                    break
                if sheet == contracts.UNIFIED_DIV_SHEET:
                    continue      # 分红情况只读取一次，见下方
                df = _read_sheet_df(path, sheet)
                if df is None:
                    continue
                chunk = _sheet_to_text(df, sheet, f"{year or ''}年 {name}",
                                       YEARLY_FILE_CHAR_BUDGET)
                if chunk:
                    used = _append_section(sections, chunk, used, hard_limit)
            # 分红情况：只取最新一年（年份已升序）
            if used < hard_limit and contracts.UNIFIED_DIV_SHEET not in read_sheets:
                df = _read_sheet_df(path, contracts.UNIFIED_DIV_SHEET)
                if df is not None:
                    chunk = _sheet_to_text(df, contracts.UNIFIED_DIV_SHEET, name,
                                           YEARLY_FILE_CHAR_BUDGET)
                    if chunk:
                        used = _append_section(sections, chunk, used, hard_limit)
                        read_sheets.add(contracts.UNIFIED_DIV_SHEET)

    if not sections:
        log(f"⚠️ 目录内未找到可用表格资料: {os.path.basename(company_dir)}")

    # 残留 PDF / DOCX：保留桌面版能力，但完全容错
    doc_files = [f for f in files if f.lower().endswith(DOC_EXTENSIONS)]
    for f in doc_files:
        if used >= hard_limit:
            break
        path = os.path.join(company_dir, f)
        try:
            text = _read_doc_file(path)
        except Exception as e:  # noqa: BLE001 - 与桌面版一致：记录并跳过
            log(f"解析发生错误，已跳过该项: {str(e)}")
            continue
        if not text:
            continue
        used = _append_section(sections, f"【本地文档｜{f}】\n{text}", used, hard_limit)

    return "\n\n".join(sections)


def _read_doc_file(path: str) -> str:
    """PDF / DOCX 文本抽取（容错；两个解析器都缺就返回空串）。"""
    low = path.lower()
    if low.endswith(".pdf"):
        try:
            import pdfplumber  # type: ignore
        except Exception:
            return ""
        text = ""
        with pdfplumber.open(path) as pdf:
            for page in pdf.pages[:35]:   # 与桌面版一致，只取前 35 页
                text += (page.extract_text() or "") + "\n"
        return text.strip()
    if low.endswith(".docx"):
        try:
            import docx  # type: ignore
        except Exception:
            return ""
        d = docx.Document(path)
        return "\n".join(p.text for p in d.paragraphs if p.text).strip()
    return ""


# ======================================================================
# 3. 单家公司分析：联网检索 + 本地表格资料 -> LLM
# ======================================================================
def _build_web_info(company: str, peers_text: str, ai_cfg: Dict[str, Any],
                    log: Callable[[str], None]) -> str:
    """与桌面版 run_batch_analysis 的两路检索逐字一致。"""
    if not bool(ai_cfg.get("web_search", AI_DEFAULTS.get("web_search", True))):
        return ""
    try:
        max_res = int(ai_cfg.get("web_results", AI_DEFAULTS.get("web_results", 3)) or 3)
    except Exception:
        max_res = 3
    try:
        timeout = int(ai_cfg.get("timeout", ai_cfg.get("request_timeout", 15)) or 15)
    except Exception:
        timeout = 15
    web_info = ""
    web_info += search_web(f"{company} 最新深度研究报告", 2, timeout, log=log)
    web_info += search_web(f"{company} 最新公告 {peers_text} 对比", 2, timeout, log=log)
    return web_info


def _extract_content(resp: Any) -> str:
    """兼容 openai SDK 的两种返回形态（对象 / dict）。"""
    try:
        return resp.choices[0].message.content or ""
    except Exception:
        pass
    try:
        return resp["choices"][0]["message"]["content"] or ""
    except Exception:
        return ""


def render_template(template: str, *, company: str, competitors: str,
                    web_info: str, doc_info: str, log=None) -> str:
    """安全渲染提示词模板。

    支持的占位符：``{company}`` / ``{competitors}`` / ``{web_info}`` / ``{doc_info}``
    （与桌面版一致；本地财报资料对应 ``{doc_info}``）。

    原来直接用 ``str.format``：模板里只要出现**任何**未知占位符（例如写成
    ``{local_material}`` 或误留 ``{`` ），就会抛 ``KeyError`` 让整份报告生成失败。
    现在改为：已知占位符逐个替换，未知占位符原样保留并逐条告警，
    保证「模板写错」只影响那一处文字，而不是整份报告作废。
    """
    log = _logger(log)
    text = str(template or "")
    mapping = {"company": str(company or ""), "competitors": str(competitors or ""),
               "web_info": str(web_info or ""), "doc_info": str(doc_info or "")}

    # 找出模板里所有 {name} 形式的占位符
    found = set(re.findall(r"\{([A-Za-z_][A-Za-z0-9_]*)\}", text))
    unknown = sorted(found - set(mapping))
    for key, val in mapping.items():
        text = text.replace("{" + key + "}", val)
    if unknown:
        log(f"   ⚠️ 提示词模板含未知占位符 {unknown}（已原样保留，不会被替换）；"
            f"可用占位符：company / competitors / web_info / doc_info")
    return text


def analyze_company(company: str, peers: List[str], local_material: str,
                    template: str, ai_cfg: Dict[str, Any],
                    log: Optional[Callable[[str], None]] = None) -> str:
    """调 LLM 生成深度报告文本（prompt 组装与桌面版完全一致）。

    - peers 传入的应是“已取简称”的同行名；在这里再按 '、'.join 组装；
    - web_info 截断 8000、doc_info 截断 15000、system="资深金融分析师"、temperature=0.3；
    - ai_cfg 支持 {"base_url","api_key","model","temperature","web_search","web_results"}。
    """
    log = _logger(log)
    ai_cfg = dict(ai_cfg or {})
    company = str(company or "").strip()
    peer_names = [str(p or "").strip() for p in (peers or []) if str(p or "").strip()]
    competitor_text = "、".join(peer_names) or "暂无"

    web_info = _build_web_info(company, competitor_text, ai_cfg, log)

    tpl = str(template or "").strip() or DEFAULT_TEMPLATE
    prompt = render_template(tpl, company=company, competitors=competitor_text,
                             web_info=web_info[:WEB_INFO_LIMIT],
                             doc_info=str(local_material or "")[:DOC_INFO_LIMIT],
                             log=log)

    try:
        from openai import OpenAI  # type: ignore
    except Exception as e:  # pragma: no cover
        raise RuntimeError(f"未安装 openai SDK: {e}") from e

    base_url = ai_cfg.get("base_url") or AI_DEFAULTS["base_url"]
    api_key = ai_cfg.get("api_key") or "sk-local"
    model = ai_cfg.get("model") or AI_DEFAULTS["model"]
    temperature = ai_cfg.get("temperature", DEFAULT_TEMPERATURE)
    try:
        temperature = float(temperature)
    except Exception:
        temperature = DEFAULT_TEMPERATURE

    client = OpenAI(api_key=api_key, base_url=base_url)
    log(f"🤖 调用模型 {model}（temperature={temperature}，prompt {len(prompt)} 字）")
    resp = client.chat.completions.create(
        model=model,
        messages=[{"role": "system", "content": SYSTEM_PROMPT},
                  {"role": "user", "content": prompt}],
        temperature=temperature,
    )
    content = _extract_content(resp)
    if not str(content).strip():
        raise RuntimeError("模型返回内容为空")
    return str(content)


# ======================================================================
# 4. 报告落盘（docx）
# ======================================================================
def _write_docx(content: str, save_path: str, title_line: str, subtitle: str) -> str:
    """生成 docx：主标题 + 副标题 + 正文（与桌面版 save_to_word 的 Markdown 渲染口径一致）。"""
    try:
        from docx import Document
        from docx.shared import Pt
    except Exception:
        # python-docx 缺失时降级为同名 txt，仍然返回“已产出”的路径
        fallback = os.path.splitext(save_path)[0] + ".txt"
        with open(fallback, "w", encoding="utf-8") as f:
            f.write(f"{title_line}\n{subtitle}\n" + "=" * 50 + "\n" + str(content or ""))
        return fallback

    doc = Document()
    doc.add_heading(title_line, 0)
    doc.add_paragraph(subtitle)
    doc.add_paragraph("=" * 50)
    for line in str(content or "").split("\n"):
        if not line.strip():
            continue
        stripped = line.strip()
        if stripped.startswith("#"):
            doc.add_heading(stripped.replace("#", "").strip(),
                            level=min(stripped.count("#"), 3))
        elif stripped.startswith("**"):
            doc.add_paragraph().add_run(stripped.replace("**", "")).bold = True
        else:
            doc.add_paragraph(line).style.font.size = Pt(11)
    doc.save(save_path)
    return save_path


def save_report(content: str, company: str, output_dir: str) -> str:
    """保存为 docx，命名与桌面版一致：``{company}_深度报告_{HHMM}.docx``。

    桌面版直接用 ``%H%M``，**同一分钟内跑两家公司就会互相覆盖**（后写的把先写的盖掉，
    而且日志都显示"已生成"，用户不易察觉）。这里保持同样的命名风格，
    但一旦重名就依次退避到 ``%H%M%S`` → ``_2`` → ``_3``…，确保绝不静默覆盖。
    """
    _ensure_dir(output_dir)
    clean = _safe_filename(company) or "未命名公司"
    base = f"{clean}_深度报告_{datetime.now().strftime('%H%M')}"
    save_path = os.path.join(output_dir, f"{base}.docx")
    if os.path.exists(save_path):
        save_path = os.path.join(output_dir, f"{clean}_深度报告_{datetime.now().strftime('%H%M%S')}.docx")
    n = 2
    while os.path.exists(save_path):
        save_path = os.path.join(output_dir, f"{base}_{n}.docx")
        n += 1
    subtitle = f"对标: 暂无 | 时间: {datetime.now().strftime('%Y-%m-%d')}"
    return _write_docx(content, save_path, f"{clean} 深度分析报告", subtitle)


def _save_failure_log(company: str, err: Exception, output_dir: str) -> str:
    """AI 调用失败时写 {company}_生成失败排错日志（桌面版行为），不中断整批任务。"""
    _ensure_dir(output_dir)
    clean = _safe_filename(company) or "未命名公司"
    body = (
        f"生成报告失败。API 调用出现异常：\n{type(err).__name__}: {str(err)}\n\n"
        "常见原因：1.模型上下文 Token 超出限制；2.当前填写的 API Key 余额不足或网络受阻；"
        "3.本地表格资料为空或模板占位符不匹配。\n"
    )
    save_path = os.path.join(output_dir, f"{clean}_{FAIL_LOG_SUFFIX}.docx")
    subtitle = f"时间: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}"
    try:
        return _write_docx(body, save_path, f"{clean} {FAIL_LOG_SUFFIX}", subtitle)
    except Exception:
        txt = os.path.join(output_dir, f"{clean}_{FAIL_LOG_SUFFIX}.txt")
        with open(txt, "w", encoding="utf-8") as f:
            f.write(body)
        return txt


# ======================================================================
# 5. 批处理：扫描 {代码}_{简称} 目录 -> 采集 -> 逐家分析 -> 出报告
# ======================================================================
def _clean_name(folder_name: str) -> str:
    """与桌面版一致：{代码}_{简称} 取最后一段。"""
    raw = str(folder_name or "").strip()
    return raw.split("_")[-1].strip() if "_" in raw else raw


def _is_company_like(folder_name: str) -> bool:
    if any(k and k in folder_name for k in SKIP_DIR_KEYS):
        return False
    if COMPANY_DIR_RE.match(folder_name):      # 首选：{代码}_{简称}
        return True
    # 兼容：目录名不带 6 位代码时，只要目录里有项目产出的表格也算
    return False


# 公司文件名形如：东鹏饮料（605499）2023年度年报提取表.xlsx / 统一整合输出_东鹏饮料(605499)_2021-2025.xlsx
_COMPANY_FILE_RES = (
    re.compile(r"[（(](\d{6})[)）]"),
    re.compile(r"^\s*(\d{6})[_\-]"),
)


def _company_from_filename(fname: str) -> Optional[Tuple[str, str]]:
    """从文件名解析 ``(代码, 简称)``；解析不到返回 None。"""
    base = os.path.splitext(str(fname))[0]
    for rx in _COMPANY_FILE_RES:
        m = rx.search(base)
        if not m:
            continue
        code = m.group(1)
        # 简称：取代码之前、去掉“统一整合输出_”等前缀的那一段
        head = base[:m.start()].strip(" _-")
        head = re.sub(r"^(统一整合输出|年报提取表|合并|提取)\s*[_\-]*", "", head).strip(" _-")
        return code, (head or code)
    return None


def scan_flat_tasks(input_dir: str, log=None) -> List[Dict[str, Any]]:
    """兼容本项目**实际产出结构**的任务扫描。

    本项目《报表提取完善》下的布局是「按行业一个文件夹、公司文件平铺在里面」：
        {输入根}/{简称}_行业_{行业}/{简称}（{代码}）{年}年度年报提取表.xlsx
        {输入根}/{简称}_行业_{行业}/统一整合输出_{简称}({代码})_{起}-{止}.xlsx
    并没有 ``{代码}_{简称}`` 子目录 —— 所以 ``scan_company_tasks`` 会一条都扫不到。
    这里按**文件名**识别公司，并以同行业文件夹内的其它公司作为同行。
    """
    log = _logger(log)
    tasks: List[Dict[str, Any]] = []
    if not input_dir or not os.path.isdir(input_dir):
        return tasks

    def _scan_dir(dir_path: str, group: str) -> None:
        try:
            names = sorted(os.listdir(dir_path))
        except Exception:  # noqa: BLE001
            return
        # 每个代码 -> 该公司的文件路径列表
        per: Dict[str, Dict[str, Any]] = {}
        for f in names:
            fp = os.path.join(dir_path, f)
            if not os.path.isfile(fp):
                continue
            if any(k and k in f for k in SKIP_DIR_KEYS):
                continue
            if not f.lower().endswith((".xlsx", ".xls", ".csv", ".docx", ".doc", ".pdf")):
                continue
            got = _company_from_filename(f)
            if not got:
                continue
            code, name = got
            rec = per.setdefault(code, {"code": code, "name": name, "files": []})
            rec["files"].append(fp)
            # 简称优先取“统一整合输出”/“年报提取表”文件名里的那一段（更可靠）
            if name and (rec["name"] == code or len(name) > len(rec["name"])):
                rec["name"] = name

        if not per:
            return
        # 只保留有可用资料的（统一整合输出或年报提取表）
        usable = {}
        for code, rec in per.items():
            if any(("统一整合输出" in os.path.basename(p)) or ("年报提取表" in os.path.basename(p))
                   for p in rec["files"]):
                usable[code] = rec
        if not usable:
            return
        for code, rec in usable.items():
            peers = [f"{c}_{r['name']}" for c, r in usable.items() if c != code]
            tasks.append({
                "path": dir_path,                 # 文件平铺在行业目录里
                "name": f"{code}_{rec['name']}",  # 与结构 A/B 一致的命名
                "code": code,
                "display_name": rec["name"],
                "group": group,
                "files": rec["files"],
                "peers": peers,
                "peers_text": "、".join(_clean_name(p) for p in peers) or "暂无",
            })

    top_dirs = _list_dirs(input_dir)
    # 顶层若直接就是公司文件（无行业目录），也支持
    _scan_dir(input_dir, "")
    for group in top_dirs:
        if any(k and k in group for k in SKIP_DIR_KEYS):
            continue
        _scan_dir(os.path.join(input_dir, group), group)

    if tasks:
        log(f"   ℹ️ 按「行业目录 + 公司文件平铺」结构识别到 {len(tasks)} 家公司"
            f"（行业组 {len({t['group'] for t in tasks})} 个）")
    return tasks


def _peer_names_from_table(company_dir: str, display_name: str) -> List[str]:
    """从《2_{简称}_同行排列表.xlsx》读取同行名称（本项目产出里就有）。"""
    try:
        for f in _list_files(company_dir):
            if "同行排列表" not in f:
                continue
            df = pd.read_excel(os.path.join(company_dir, f))
            if "名称" not in df.columns:
                continue
            out: List[str] = []
            for _, r in df.iterrows():
                nm = str(r.get("名称") or "").strip()
                cd = str(r.get("代码") or "").strip()
                if not nm:
                    continue
                if nm == display_name or cd in f:
                    continue          # 排除代表公司自己
                out.append(f"{cd}_{nm}" if cd else nm)
            if out:
                return out
    except Exception:  # noqa: BLE001
        pass
    return []


def scan_company_tasks(input_dir: str, log: Optional[Callable[[str], None]] = None
                       ) -> List[Dict[str, Any]]:
    """扫描输入根目录，返回 [{"path","name","peers","peers_text","group"}]。

    兼容三种结构：
      A) {输入根}/{行业文件夹}/{代码}_{简称}/      （本项目标准结构，同行=同行业其余公司）
      B) {输入根}/{代码}_{简称}/                   （扁平结构，同行=其余公司）
      C) {输入根}/{行业文件夹}/ 公司文件平铺/**    （本项目**实际产出**结构，按文件名识别公司）
    """
    log = _logger(log)
    tasks: List[Dict[str, Any]] = []
    if not input_dir or not os.path.isdir(input_dir):
        log(f"⚠️ 输入目录不存在: {input_dir}")
        return tasks

    def _collect(group: str, folders: List[str]) -> None:
        for comp in folders:
            path = os.path.join(input_dir, group, comp) if group else os.path.join(input_dir, comp)
            if not os.path.isdir(path):
                continue
            peers = [c for c in folders if c != comp]
            tasks.append({
                "path": path,
                "name": comp,
                "group": group,
                "peers": peers,
                "peers_text": "、".join(_clean_name(p) for p in peers) or "暂无",
            })

    top_dirs = _list_dirs(input_dir)
    top_files = _list_files(input_dir)
    company_like = [d for d in top_dirs if _is_company_like(d)]

    if len(company_like) >= 2:
        # 结构 B：扁平
        _collect("", company_like)
    elif len(company_like) <= 1 and (top_files or top_dirs):
        # 结构 A：深入行业文件夹
        for group in top_dirs:
            if any(k and k in group for k in SKIP_DIR_KEYS):
                continue
            group_path = os.path.join(input_dir, group)
            folders = [d for d in _list_dirs(group_path) if _is_company_like(d)]
            if folders:
                _collect(group, folders)
    else:
        for group in top_dirs:
            if any(k and k in group for k in SKIP_DIR_KEYS):
                continue
            group_path = os.path.join(input_dir, group)
            folders = [d for d in _list_dirs(group_path) if _is_company_like(d)]
            if folders:
                _collect(group, folders)

    # 结构 C 兜底：以上都没扫到 → 按文件名识别（本项目实际产出结构）
    if not tasks:
        tasks = scan_flat_tasks(input_dir, log)
        # 用《同行排列表》里的真实同行覆盖“同目录其余公司”
        for t in tasks:
            real = _peer_names_from_table(t.get("path", ""), t.get("display_name", ""))
            if real:
                t["peers"] = real
                t["peers_text"] = "、".join(_clean_name(p) for p in real)

    if not tasks:
        log("⚠️ 未扫描到任何公司资料（可尝试 company_filter 或检查目录结构）")
    return tasks


def _match_filter(task: Dict[str, Any], wanted: List[str]) -> bool:
    name = str(task.get("name") or "")
    clean = _clean_name(name)
    for w in wanted:
        w = str(w or "").strip()
        if not w:
            continue
        if w == name or w == clean or w in name or clean in w:
            return True
    return False


def run_all(input_dir: str, output_dir: str, template: str, ai_cfg: Dict[str, Any],
            log: Optional[Callable[[str], None]] = None,
            progress: Optional[Callable[[int, int, str], None]] = None,
            company_filter: Optional[List[str]] = None) -> Dict[str, Any]:
    """全流程：扫描 -> 采集本地表格资料 -> 逐家调 LLM -> 出 docx 到 output_dir/企业分析/。

    与桌面版 run_batch_analysis 一致：单家失败绝不中断整批，失败信息落到
    {company}_生成失败排错日志 并记入返回值的 errors。
    返回 {"ok": True, "reports": [...], "companies": [...], "errors": [...]}（另带 counts / output_dir）。
    """
    log = _logger(log)
    ai_cfg = dict(ai_cfg or {})
    result: Dict[str, Any] = {
        "ok": True, "reports": [], "companies": [], "errors": [],
        "output_dir": "", "total": 0, "failed": 0,
    }

    def _emit(done: int, total: int, name: str) -> None:
        if progress is None:
            return
        try:
            progress(int(done), int(total), str(name))
        except Exception:
            pass

    if not input_dir or not os.path.isdir(input_dir):
        msg = f"输入目录不存在: {input_dir}"
        log(f"⚠️ {msg}")
        result["ok"] = False
        result["errors"].append(msg)
        return result
    if not output_dir:
        msg = "未指定输出目录"
        log(f"⚠️ {msg}")
        result["ok"] = False
        result["errors"].append(msg)
        return result

    output_root = os.path.join(output_dir, contracts.DIR_ENTERPRISE)
    _ensure_dir(output_root)
    result["output_dir"] = output_root
    log(f"📁 报告输出目录: {output_root}")

    tasks = scan_company_tasks(input_dir, log=log)
    if company_filter:
        wanted = [str(x) for x in company_filter if str(x or "").strip()]
        tasks = [t for t in tasks if _match_filter(t, wanted)]
    total = len(tasks)
    result["total"] = total
    if total == 0:
        log("⚠️ 没有待分析的公司任务，任务结束")
        _emit(0, 0, "")
        return result

    log(f"🔎 共 {total} 家公司待分析")
    tpl = str(template or "").strip() or DEFAULT_TEMPLATE
    started = time.time()

    for idx, task in enumerate(tasks, 1):
        folder = str(task.get("name") or "")
        clean_company = _clean_name(folder)
        peers_text = str(task.get("peers_text") or "暂无")
        log(f"\n开始分析: {folder}")
        _emit(idx - 1, total, clean_company)
        try:
            company_dir = str(task.get("path") or "")
            local_material = collect_local_material(company_dir, log=log)
            log(f"📚 本地资料 {len(local_material)} 字（送入模型时按 {DOC_INFO_LIMIT} 字截断）")

            content = analyze_company(clean_company, _peer_clean_list(task),
                                      local_material, tpl, ai_cfg, log=log)
            save_path = save_report_with_peers(content, clean_company, output_root, peers_text)
            result["reports"].append(save_path)
            result["companies"].append(clean_company)
            log(f"✅ 报告已生成: {os.path.basename(save_path)}")
        except Exception as e:  # noqa: BLE001 - 单家失败不影响其余公司
            log(f"❌ {clean_company} 生成失败: {type(e).__name__}: {str(e)}")
            result["failed"] += 1
            result["errors"].append(f"{clean_company}: {type(e).__name__}: {str(e)}")
            try:
                bad = _save_failure_log(clean_company, e, output_root)
                log(f"📝 已写入排错日志: {os.path.basename(bad)}")
            except Exception as e2:  # noqa: BLE001
                log(f"⚠️ 排错日志写入失败: {type(e2).__name__}: {str(e2)}")
        finally:
            _emit(idx, total, clean_company)

    result["ok"] = result["failed"] == 0 or bool(result["reports"])
    log(f"🎉 全部完成！成功 {len(result['reports'])} 家，失败 {result['failed']} 家，"
        f"耗时 {time.time() - started:.1f}s")
    return result


def _peer_clean_list(task: Dict[str, Any]) -> List[str]:
    """同行名取简称（与桌面版 p.split('_')[-1] 一致）。"""
    return [_clean_name(p) for p in (task.get("peers") or [])]


def save_report_with_peers(content: str, company: str, output_dir: str, peers_text: str) -> str:
    """与 save_report 同规则，仅额外把“对标同行”写入副标题（保持桌面版副标题信息量）。"""
    path = save_report(content, company, output_dir)
    try:
        if not path.lower().endswith(".docx"):
            return path
        from docx import Document
        doc = Document(path)
        if len(doc.paragraphs) >= 2:
            doc.paragraphs[1].text = (f"对标: {peers_text} | "
                                      f"时间: {datetime.now().strftime('%Y-%m-%d')}")
            doc.save(path)
    except Exception:
        pass
    return path
