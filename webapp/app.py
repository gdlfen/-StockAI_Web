# -*- coding: utf-8 -*-
"""
Streamlit 前端（手机可直接访问）

两种运行模式（侧边栏可切换）：
- **内置模式（推荐，Streamlit 云端一键部署）**：直接调用 core 包在同一进程内跑流水线；
- **API 模式**：填写后端地址（server/main.py 起的 FastAPI），由后端执行、前端只做展示与下载。

部署（Streamlit Community Cloud）：
    Main file path 填 `app.py`（仓库根的那个入口）或 `webapp/app.py`，两者都可以；
    入口会通过 bootstrap.py 自动定位 `core/` 所在目录，因此仓库怎么摆都能跑。

本地启动： streamlit run app.py   或   streamlit run webapp/app.py
"""
from __future__ import annotations

import io
import json
import os
import shutil
import sys
import tempfile
import time
import zipfile
from datetime import datetime
from typing import Any, Dict, List, Optional

import pandas as pd
import streamlit as st

# ----------------------------------------------------------------------
# 路径引导：向上查找含 core/pipeline.py 的项目根（兼容仓库根 / webapp / 中文目录 等摆放）
# ----------------------------------------------------------------------
_HERE = os.path.dirname(os.path.abspath(__file__))
_BOOT_CANDIDATES = [_HERE, os.path.dirname(_HERE), os.getcwd(), os.path.dirname(os.getcwd()), "/mount/src"]
for _b in list(_BOOT_CANDIDATES):
    _cur = _b
    for _ in range(6):
        _p = os.path.dirname(_cur)
        if not _p or _p == _cur:
            break
        _BOOT_CANDIDATES.append(_p)
        _cur = _p

_PROJECT_ROOT = None
for _b in _BOOT_CANDIDATES:
    try:
        if _b and os.path.isfile(os.path.join(_b, "core", "pipeline.py")):
            _PROJECT_ROOT = _b
            break
    except Exception:
        continue
if _PROJECT_ROOT:
    if _PROJECT_ROOT not in sys.path:
        sys.path.insert(0, _PROJECT_ROOT)
elif _HERE not in sys.path:
    sys.path.insert(0, _HERE)

try:
    from core import config as cfg_mod          # noqa: E402
    from core import contracts as C             # noqa: E402
    from core import pipeline as pipe           # noqa: E402
    from core.common import DiskCache, make_logger  # noqa: E402
except ModuleNotFoundError as _imp_err:         # noqa: F841
    st.set_page_config(page_title="启动失败", page_icon="⚠️", layout="centered")
    st.error("### 未找到项目代码目录 `core/`，无法启动")
    st.markdown(
        "请确认 GitHub 仓库里与入口文件 **同级** 存在 `core/` 文件夹"
        "（`core/pipeline.py`、`core/screening.py` 等都在）。正确结构：\n\n"
        "```\n"
        "<仓库根>/app.py            ← Streamlit 的 Main file path\n"
        "<仓库根>/core/…            ← 业务代码\n"
        "<仓库根>/server/…  <仓库根>/webapp/…\n"
        "<仓库根>/requirements.txt  <仓库根>/packages.txt\n"
        "```"
    )
    with st.expander("诊断信息（排查用）"):
        st.write({"入口文件": os.path.abspath(__file__), "当前工作目录": os.getcwd()})
        _seen = []
        for _b in _BOOT_CANDIDATES[:10]:
            try:
                _seen.append(f"{_b} -> {sorted(os.listdir(_b))[:20]}")
            except Exception as _e:
                _seen.append(f"{_b} -> <无法读取: {_e}>")
        st.code("\n".join(_seen))
        st.code(f"{type(_imp_err).__name__}: {_imp_err}")
    # 关键：显式终止，不依赖 st.stop() 的“停止执行”语义（桩件/测试环境下它不会中断）
    st.stop()
    raise SystemExit(0)

st.set_page_config(page_title="价值投资智能分析（云端版）", page_icon="📈",
                   layout="wide", initial_sidebar_state="collapsed")

# ======================================================================
# 会话级数据目录（云端文件系统不保证持久，故每次会话独立，并支持一键打包下载）
# ======================================================================
def _session_dir() -> str:
    if "data_dir" not in st.session_state:
        base = os.environ.get("VIM_DATA_DIR") or tempfile.mkdtemp(prefix="vim_")
        st.session_state.data_dir = base
        os.makedirs(base, exist_ok=True)
    return st.session_state.data_dir


def _cache_dir() -> str:
    d = os.path.join(_session_dir(), "_cache")
    os.makedirs(d, exist_ok=True)
    return d


def _cfg() -> Dict[str, Any]:
    if "user_config" not in st.session_state:
        st.session_state.user_config = cfg_mod.load_user_config(_session_dir())
    return st.session_state.user_config


# ======================================================================
# 侧边栏
# ======================================================================
with st.sidebar:
    st.markdown("## 📈 价值投资分析")
    st.caption("云端版 · 手机可用")
    mode = st.radio("运行模式", ["内置引擎（推荐）", "远程 API"], index=0, horizontal=False)
    api_base = ""
    if mode == "远程 API":
        api_base = st.text_input("后端地址", value=st.session_state.get("api_base", "http://127.0.0.1:8000"))
        st.session_state.api_base = api_base
        st.caption("后端启动：`uvicorn server.main:app --host 0.0.0.0 --port 8000`")

    st.divider()
    # 数据源可用性：让“某个可选依赖没装上”这种情况一眼可见，
    # 而不是等到跑某个模块才发现数据为空。
    try:
        from core.datasource import source_status
        _ds = source_status((st.session_state.get("user_config") or {}).get("data_source", {}).get("tushare_token", ""))
        _marks = {"akshare": "AkShare(主)", "baostock": "Baostock(备)", "tushare": "TuShare(备)"}
        _txt = " · ".join(f"{'✅' if ok else '⬜'} {_marks.get(k, k)}" for k, ok in _ds.items())
        st.caption("数据源：" + _txt)
        if not _ds.get("akshare"):
            st.warning("未检测到 AkShare（主数据源）。年报汇总与海选将无法工作，"
                       "请确认 requirements.txt 已包含 `akshare` 并重新部署。")
    except Exception:
        pass

    st.metric("数据目录", os.path.basename(_session_dir()))
    if st.button("🧹 清空本次会话数据", use_container_width=True):
        shutil.rmtree(_session_dir(), ignore_errors=True)
        for k in ("data_dir", "job", "last_result", "user_config"):
            st.session_state.pop(k, None)
        st.rerun()

MODE_LOCAL = mode.startswith("内置")

# ======================================================================
# HTTP 客户端（仅在远程 API 模式下使用）
# ======================================================================
import requests  # noqa: E402


def api(method: str, path: str, **kw) -> Any:
    url = st.session_state.get("api_base", api_base).rstrip("/") + path
    r = requests.request(method, url, timeout=kw.pop("timeout", 60), **kw)
    r.raise_for_status()
    return r.json()


# ======================================================================
# 内置引擎：线程安全的"内存 Job"（与 server.jobs 同结构，供前端直接轮询）
# ======================================================================
class LocalJob:
    def __init__(self, stage_keys: List[str], params: Dict[str, Any]):
        import threading
        from collections import deque
        self.id = "local-" + datetime.now().strftime("%H%M%S")
        self.stage_keys = stage_keys
        self.params = params
        self.status = "pending"
        self.stage_total = len(stage_keys)
        self.stage_index = 0
        self.stage_name = ""
        self.step_done = 0
        self.step_total = 0
        self.step_label = ""
        self.logs = deque(maxlen=3000)
        self.result: Optional[Dict[str, Any]] = None
        self.error = ""
        self.started_at = self.finished_at = None
        self._thread = threading.Thread(target=self._run, daemon=True)
        self._thread.start()

    # ---- 供 pipeline 回调 ----
    def log(self, msg: str) -> None:
        self.logs.append(str(msg))

    def progress(self, done: int, total: int, label: str = "") -> None:
        self.step_done, self.step_total, self.step_label = done, total, label

    def snapshot(self) -> Dict[str, Any]:
        return {
            "id": self.id, "status": self.status,
            "stage_index": self.stage_index, "stage_total": self.stage_total,
            "stage_name": self.stage_name,
            "step_done": self.step_done, "step_total": self.step_total, "step_label": self.step_label,
            "logs": list(self.logs), "result": self.result, "error": self.error,
            "elapsed": (self.finished_at or time.time()) - (self.started_at or time.time()),
        }

    # ---- 执行 ----
    def _run(self) -> None:
        self.status = "running"
        self.started_at = time.time()
        try:
            cfg = _cfg()
            results: Dict[str, Any] = {}
            for idx, key in enumerate(self.stage_keys):
                self.stage_index = idx + 1
                self.stage_name = next((s["name"] for s in pipe.STAGES if s["key"] == key), key)
                self.step_done = self.step_total = 0
                self.log("=" * 60)
                self.log(f"▶️ {self.stage_name}")
                ctx = pipe.make_context(_session_dir(), _cache_dir(), user_config=cfg,
                                        log=make_logger(self.log), progress=self.progress)
                sp = self.params.get(key)
                sp = dict(sp) if isinstance(sp, dict) else {}
                fn = pipe._STAGE_FUNCS.get(key)
                if fn is None:
                    continue
                try:
                    res = fn(ctx, sp)
                except Exception as e:  # noqa: BLE001
                    import traceback
                    self.log(f"❌ {self.stage_name} 失败：{type(e).__name__}: {e}")
                    self.log(traceback.format_exc()[:1500])
                    res = {"ok": False, "error": f"{type(e).__name__}: {e}"}
                results[key] = res
                self.log(f"{'✅' if (res or {}).get('ok') else '⚠️'} {self.stage_name} 结束")
                if key == "report" and not (res or {}).get("ok"):
                    self.log("⛔ 年报汇总未成功，中止后续阶段")
                    break
            self.result = {"results": pipe._jsonable(results)}
            self.status = "done"
        except Exception as e:  # noqa: BLE001
            self.error = f"{type(e).__name__}: {e}"
            self.status = "failed"
        finally:
            self.finished_at = time.time()


# ======================================================================
# 通用：启动任务 + 进度展示
# ======================================================================
def start_job(stage_keys: List[str], params: Dict[str, Any]) -> None:
    if MODE_LOCAL:
        st.session_state.job = LocalJob(stage_keys, params)
    else:
        payload = {"stages": stage_keys, "params": params, "user_config": _cfg()}
        try:
            resp = api("POST", "/api/jobs", json=payload)
            st.session_state.job_id = resp["job_id"]
            st.session_state.job = None
        except Exception as e:  # noqa: BLE001
            st.error(f"调用后端失败：{e}")
            st.stop()


class RemoteJob:
    """远程后端任务的轻量代理（每次 snapshot 都去后端拉最新状态）。"""

    def __init__(self, job_id: str, base: str):
        self.job_id = job_id
        self.base = base

    def snapshot(self) -> Dict[str, Any]:
        url = self.base.rstrip("/") + f"/api/jobs/{self.job_id}"
        r = requests.get(url, timeout=30)
        r.raise_for_status()
        return r.json()


def _current_job():
    """取出当前任务对象：内置模式用 LocalJob，远程模式用 RemoteJob。"""
    if MODE_LOCAL:
        return st.session_state.get("job")
    jid = st.session_state.get("job_id")
    if not jid:
        return None
    return RemoteJob(jid, st.session_state.get("api_base", ""))


def render_job_progress() -> Optional[Dict[str, Any]]:
    """展示运行进度；完成时返回结果快照。"""
    job = _current_job()
    if job is None:
        return None
    try:
        snap = job.snapshot()
    except Exception as e:  # noqa: BLE001
        st.error(f"读取任务状态失败：{e}")
        return None
    status = snap.get("status")
    label_map = {"pending": "排队中", "running": "运行中", "done": "已完成",
                 "failed": "失败", "cancelled": "已取消"}
    st.info(f"任务 {snap.get('id')} · 状态：**{label_map.get(status, status)}** · "
            f"阶段 {snap.get('stage_index')}/{snap.get('stage_total')} {snap.get('stage_name')} · "
            f"用时 {snap.get('elapsed', 0):.0f}s")
    if snap.get("step_total"):
        st.progress(min(1.0, snap["step_done"] / max(1, snap["step_total"])),
                    text=f"{snap['step_label']} ({snap['step_done']}/{snap['step_total']})")
    with st.expander("📜 运行日志", expanded=(status == "running")):
        st.code("\n".join(snap.get("logs", [])[-400:]) or "（暂无日志）")
    if status in ("done", "failed", "cancelled"):
        return snap
    # 仍在运行：稍等后自动刷新（无需手动点按钮）
    time.sleep(1.5)
    st.rerun()
    return None


# ======================================================================
# 结果与文件
# ======================================================================
def list_output_files() -> List[Dict[str, Any]]:
    root = _session_dir()
    out: List[Dict[str, Any]] = []
    for r, dirs, files in os.walk(root):
        dirs[:] = [d for d in dirs if d != "_cache"]
        for f in files:
            if f.startswith("~$"):
                continue
            p = os.path.join(r, f)
            out.append({"name": f, "rel": os.path.relpath(p, root), "path": p,
                        "size": os.path.getsize(p), "mtime": os.path.getmtime(p)})
    out.sort(key=lambda x: -x["mtime"])
    return out


def zip_all(root: str) -> bytes:
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as z:
        for r, dirs, files in os.walk(root):
            dirs[:] = [d for d in dirs if d != "_cache"]
            for f in files:
                if f.startswith("~$"):
                    continue
                p = os.path.join(r, f)
                z.write(p, os.path.relpath(p, root))
    return buf.getvalue()


def render_files() -> None:
    files = list_output_files()
    st.subheader("📦 产物文件")
    if not files:
        st.caption("还没有产物。先在上面的页面跑一个模块。")
        return
    st.caption(f"共 {len(files)} 个文件（云端文件不持久，请及时下载）")
    if st.button("⬇️ 打包下载全部产物（zip）", use_container_width=True, type="primary"):
        st.download_button("点此保存 zip",
                           data=zip_all(_session_dir()),
                           file_name=f"价值投资分析产物_{datetime.now():%Y%m%d_%H%M}.zip",
                           mime="application/zip", use_container_width=True)
    for f in files[:120]:
        col1, col2 = st.columns([3, 1])
        with col1:
            st.markdown(f"**{f['name']}**  \n<span style='color:#888;font-size:0.8em'>"
                        f"{f['rel']} · {f['size']/1024:.0f} KB</span>", unsafe_allow_html=True)
        with col2:
            with open(f["path"], "rb") as fh:
                st.download_button("下载", data=fh.read(), file_name=f["name"],
                                   key=f"dl_{f['rel']}", use_container_width=True)


# ======================================================================
# 页面
# ======================================================================
st.title("📈 价值投资智能分析模型 · 云端版")
tabs = st.tabs(["🏠 总览", "🎯 海选", "📚 年报数据", "🕵️ 造假排雷",
                "📊 16维度", "🤖 企业AI", "💰 好价估值", "⚙️ 设置", "📦 产物"])

# ---------------- 总览 ----------------
with tabs[0]:
    st.markdown("#### 一键跑完整流水线")
    st.caption("顺序：海选公司 → 年报数据搜索汇总 → 造假排雷 → 16维度 → 企业AI深度 → 好价估值")
    presets = {
        "全部（推荐）": ["haixuan", "report", "fraud", "dim16", "enterprise", "valuation"],
        "跳过海选（用已有名单）": ["report", "fraud", "dim16", "enterprise", "valuation"],
        "只要数据与财务分析": ["haixuan", "report", "fraud", "dim16"],
        "只要好价估值": ["valuation"],
    }
    pick = st.selectbox("运行范围", list(presets.keys()), index=0)
    with st.expander("常用参数（可留默认）"):
        c1, c2 = st.columns(2)
        with c1:
            pool_size = st.number_input("海选候选池大小（按市值降序取前 N 只）", 20, 1000, 200, 20,
                                        help="越大越全，但更慢。手机端建议 100~300。")
            years_back = st.number_input("年报回溯年数", 2, 10, 5, 1)
        with c2:
            manual = st.text_area("手工指定公司（每行一个，可填代码或简称；留空则用海选结果）",
                                  height=90, placeholder="600887\n605499\n重庆啤酒")
    if st.button("🚀 开始运行", type="primary", use_container_width=True):
        targets = []
        for line in (manual or "").replace("，", ",").splitlines():
            s = line.strip()
            if s:
                targets.append({"code": s if s.isdigit() else "", "name": s})
        params = {
            "haixuan": {"pool_size": int(pool_size)},
            "report": {"years_back": int(years_back),
                       "targets": [t for t in targets if t.get("code")] or None},
            "enterprise": {},
        }
        start_job(presets[pick], params)
    snap = render_job_progress()
    if snap and snap.get("result"):
        st.success("任务结束，见下方各模块页面与「产物」页。")
        st.json(snap["result"], expanded=False)

# ---------------- 海选 ----------------
with tabs[1]:
    st.markdown("#### 🎯 海选公司（免费数据源）")
    st.caption("用 AkShare 全市场数据按条件筛选，替代原桌面版的问财浏览器抓取。")
    cfg = _cfg()
    preset = st.selectbox("条件模板", list(cfg_mod.HAIXUAN_PRESETS.keys()), index=0)
    query = st.text_area("筛选条件（可自由编辑）", value=cfg_mod.HAIXUAN_PRESETS[preset], height=90)
    pool = st.slider("候选池大小", 20, 1000, 200, 20)
    if st.button("开始海选", type="primary", use_container_width=True):
        start_job(["haixuan"], {"haixuan": {"query": query, "pool_size": int(pool)}})
    snap = render_job_progress()
    if snap and snap.get("result"):
        res = (snap["result"].get("results") or {}).get("haixuan") or {}
        stocks = res.get("stocks") or []
        st.success(f"入选 {len(stocks)} 只（候选 {res.get('candidates')} 只）")
        if stocks:
            df = pd.DataFrame(stocks)
            st.dataframe(df, use_container_width=True, hide_index=True)
            st.download_button("下载《海选公司汇总表》CSV",
                               df.to_csv(index=False).encode("utf-8-sig"),
                               file_name="海选公司汇总表.csv", mime="text/csv")
        detail = res.get("detail") or []
        if detail:
            with st.expander("查看全部候选明细（含未通过原因）"):
                st.dataframe(pd.DataFrame(detail), use_container_width=True, hide_index=True)

# ---------------- 年报数据 ----------------
with tabs[2]:
    st.markdown("#### 📚 年报数据搜索汇总")
    st.caption("用 AkShare（新浪财报 / 东财分红 / 同花顺）汇总三大表、分红、员工数据，"
               "产出与原《年报提取表》《统一整合输出》完全一致的表结构。")
    manual = st.text_area("目标公司（每行一个：代码或简称；留空则读海选结果）", height=110,
                          placeholder="600887\n605499\n603156")
    c1, c2 = st.columns(2)
    with c1:
        sy = st.number_input("起始年份", 2000, 2030, datetime.now().year - 4)
    with c2:
        ey = st.number_input("结束年份", 2000, 2030, datetime.now().year)
    with st.expander("高级：当年报汇总遇到同名多文件时，行业分组与同行"):
        inc_peers = st.checkbox("同时生成同行排列表（同行前三）", value=True)
    if st.button("开始汇总", type="primary", use_container_width=True):
        tks = []
        for line in (manual or "").replace("，", ",").splitlines():
            s = line.strip()
            if s:
                tks.append({"code": s if s.isdigit() else "", "name": s})
        start_job(["report"], {"report": {"targets": tks or None, "start_year": int(sy),
                                          "end_year": int(ey), "include_peers": inc_peers}})
    snap = render_job_progress()
    if snap and snap.get("result"):
        res = (snap["result"].get("results") or {}).get("report") or {}
        st.success(f"生成 {len(res.get('files') or [])} 个文件")
        comps = res.get("companies") or []
        if comps:
            st.dataframe(pd.DataFrame(comps), use_container_width=True, hide_index=True)

# ---------------- 造假排雷 ----------------
with tabs[3]:
    st.markdown("#### 🕵️ 财务造假排雷（18 项）")
    st.caption("读取《统一整合输出》逐公司逐年度跑 18 条判断，产出《深度排雷报告.xlsx》。")
    cfg = _cfg()
    st.caption("阈值可在「设置」页调整；这里点一下就用当前阈值运行。")
    if st.button("开始排雷", type="primary", use_container_width=True):
        start_job(["fraud"], {"fraud": {}})
    snap = render_job_progress()
    if snap and snap.get("result"):
        res = (snap["result"].get("results") or {}).get("fraud") or {}
        comps = res.get("companies") or []
        if comps:
            st.dataframe(pd.DataFrame(comps), use_container_width=True, hide_index=True)

# ---------------- 16维度 ----------------
with tabs[4]:
    st.markdown("#### 📊 16 维度同行对比分析")
    st.caption("按行业把同行公司放在一起，逐维度打分并给出投资建议。")
    if st.button("开始分析", type="primary", use_container_width=True):
        start_job(["dim16"], {"dim16": {}})
    snap = render_job_progress()
    if snap and snap.get("result"):
        res = (snap["result"].get("results") or {}).get("dim16") or {}
        st.success(f"生成 {len(res.get('files') or [])} 个行业报告")
        if res.get("industries"):
            st.write("行业：" + "、".join(map(str, res["industries"])))

# ---------------- 企业AI ----------------
with tabs[5]:
    st.markdown("#### 🤖 企业 AI 深度分析")
    st.caption("把本项目汇总的财务表格 + 联网检索情报喂给 LLM，逐家生成企业深度报告。"
               "需要在「设置」里填 API Key。")
    cfg = _cfg()
    ai = cfg.get("ai") or {}
    st.caption(f"当前模型：`{ai.get('model','-')}` · 接口：`{ai.get('base_url','-')}` · "
               f"Key：{'已配置' if (ai.get('api_key') or '').strip() else '未配置'}")
    tmpl = st.text_area("提示词模板（{company} / {competitors} / {web_info} / {doc_info} 占位符）",
                        value=ai.get("template") or cfg_mod.DEFAULT_ENTERPRISE_PROMPT, height=160)
    if st.button("开始生成", type="primary", use_container_width=True):
        start_job(["enterprise"], {"enterprise": {"template": tmpl, "ai": ai}})
    snap = render_job_progress()
    if snap and snap.get("result"):
        res = (snap["result"].get("results") or {}).get("enterprise") or {}
        if res.get("skipped"):
            st.warning(res.get("reason"))
        else:
            st.success(f"生成 {len(res.get('reports') or [])} 份报告")
        if res.get("errors"):
            with st.expander("错误明细"):
                st.write(res["errors"])

# ---------------- 好价估值 ----------------
with tabs[6]:
    st.markdown("#### 💰 好价估值分析")
    st.caption("现价 / 市盈率(TTM) / **严格 TTM 动态股息率** → 好价格判断（规则与原程序一致）。")
    cfg = _cfg()
    val = cfg.get("valuation") or cfg_mod.VALUATION_DEFAULTS
    manual = st.text_area("估值标的（每行一个：代码或简称；留空则读年报汇总出的名单）",
                          height=100, placeholder="600887\n605499")
    c1, c2, c3 = st.columns(3)
    with c1:
        sz = st.checkbox("A股大盘PE", value=bool(val.get("sz_pe", True)))
        cn = st.checkbox("中国10Y国债", value=bool(val.get("cn_10y", True)))
    with c2:
        hk = st.checkbox("恒生指数PE", value=bool(val.get("hs_pe", True)))
        us = st.checkbox("美国10Y国债", value=bool(val.get("us_10y", True)))
    with c3:
        sp = st.checkbox("标普500PE", value=bool(val.get("sp_pe", True)))
        fed = st.checkbox("美联储利率", value=bool(val.get("fed", True)))
    if st.button("开始估值", type="primary", use_container_width=True):
        tks = [{"code": s.strip() if s.strip().isdigit() else "", "name": s.strip()}
               for s in (manual or "").splitlines() if s.strip()]
        vcfg = {"sz_pe": sz, "hs_pe": hk, "sp_pe": sp, "cn_10y": cn, "us_10y": us, "fed": fed,
                "targets": [t for t in tks if t.get("code")] or None}
        start_job(["valuation"], {"valuation": {"valuation": vcfg}})
    snap = render_job_progress()
    if snap and snap.get("result"):
        res = (snap["result"].get("results") or {}).get("valuation") or {}
        recs = res.get("records") or []
        if recs:
            df = pd.DataFrame(recs)
            keep = [c for c in ["分类", "名称/代码", "现价", "当前PE(TTM)", "动态股息率",
                                "判断结果", "反推好价格上限"] if c in df.columns]
            st.dataframe(df[keep], use_container_width=True, hide_index=True)
        macro = res.get("macro") or {}
        if macro:
            with st.expander("宏观与大盘指标"):
                st.json(macro)

# ---------------- 设置 ----------------
with tabs[7]:
    st.markdown("#### ⚙️ 设置")
    cfg = _cfg()
    sub = st.tabs(["AI 接口", "造假排雷阈值", "16维度参数", "数据源"])

    with sub[0]:
        ai = dict(cfg.get("ai") or {})
        ai["base_url"] = st.text_input("API 地址", value=ai.get("base_url", cfg_mod.AI_DEFAULTS["base_url"]))
        ai["api_key"] = st.text_input("API Key", value=ai.get("api_key", ""), type="password")
        ai["model"] = st.text_input("模型", value=ai.get("model", cfg_mod.AI_DEFAULTS["model"]))
        ai["temperature"] = st.slider("temperature", 0.0, 1.0, float(ai.get("temperature", 0.3)), 0.05)
        if st.button("保存 AI 设置", use_container_width=True):
            cfg["ai"] = ai
            cfg_mod.save_user_config(_session_dir(), cfg)
            st.session_state.user_config = cfg
            st.success("已保存")

    with sub[1]:
        fp = dict(cfg.get("fraud_params") or cfg_mod.default_fraud_params())
        meta = {k: v["name"] for k, v in cfg_mod.FRAUD_PARAM_DEFS.items()}
        cols = st.columns(2)
        for i, (k, v) in enumerate(fp.items()):
            with cols[i % 2]:
                if isinstance(v, str):
                    fp[k] = st.text_input(meta.get(k, k), value=v, key=f"fp_{k}")
                elif isinstance(v, int):
                    fp[k] = st.number_input(meta.get(k, k), value=int(v), step=1, key=f"fp_{k}")
                else:
                    fp[k] = st.number_input(meta.get(k, k), value=float(v), step=0.1,
                                            format="%.4f", key=f"fp_{k}")
        if st.button("保存排雷阈值", use_container_width=True):
            cfg["fraud_params"] = fp
            cfg_mod.save_user_config(_session_dir(), cfg)
            st.session_state.user_config = cfg
            st.success("已保存")

    with sub[2]:
        d16 = dict(cfg.get("config_16dim") or cfg_mod.default_16dim_config())
        meta16 = {k: v for k, v in cfg_mod.FULL_CONFIG.items()}
        cols = st.columns(2)
        for i, (k, v) in enumerate(d16.items()):
            desc = meta16.get(k, {}).get("desc", k)
            with cols[i % 2]:
                if isinstance(v, str):
                    d16[k] = st.text_input(desc, value=v, key=f"d16_{k}")
                else:
                    d16[k] = st.number_input(desc, value=float(v), step=0.01,
                                             format="%.4f", key=f"d16_{k}")
        if st.button("保存 16 维度参数", use_container_width=True):
            cfg["config_16dim"] = d16
            cfg_mod.save_user_config(_session_dir(), cfg)
            st.session_state.user_config = cfg
            st.success("已保存")

    with sub[3]:
        ds = dict(cfg.get("data_source") or cfg_mod.DATA_SOURCE_DEFAULTS)
        ds["years_back"] = st.number_input("默认回溯年数", 2, 10, int(ds.get("years_back", 5)))
        ds["max_peers"] = st.number_input("同行取前 N 名", 1, 10, int(ds.get("max_peers", 3)))
        ds["request_timeout"] = st.number_input("请求超时(秒)", 5, 120, int(ds.get("request_timeout", 20)))
        ds["sleep_between"] = st.number_input("请求间隔(秒)", 0.0, 3.0, float(ds.get("sleep_between", 0.35)), 0.05)
        st.info("本版本全部使用免费数据源：AkShare（新浪/东财/同花顺/巨潮）、腾讯行情、新浪行情。"
                "Tushare 为可选（需 token）。")
        if st.button("保存数据源设置", use_container_width=True):
            cfg["data_source"] = ds
            cfg_mod.save_user_config(_session_dir(), cfg)
            st.session_state.user_config = cfg
            st.success("已保存")

# ---------------- 产物 ----------------
with tabs[8]:
    render_files()
    st.divider()
    st.caption("提示：Streamlit Community Cloud 的磁盘在重启后会清空，重要结果请及时下载。")
