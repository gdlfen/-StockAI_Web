# -*- coding: utf-8 -*-
"""
FastAPI 后端：把流水线能力暴露成 HTTP 接口，供 Streamlit 前端 / 手机浏览器 / 第三方调用。

启动：
    uvicorn server.main:app --host 0.0.0.0 --port 8000
接口文档：
    http://localhost:8000/docs
"""
from __future__ import annotations

import os
import sys
from typing import Any, Dict, List, Optional

from fastapi import FastAPI, HTTPException, Query
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse, JSONResponse
from pydantic import BaseModel, Field

# 允许从项目根目录直接 `uvicorn server.main:app` 启动
ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if ROOT not in sys.path:
    sys.path.insert(0, ROOT)

from core import config as cfg_mod           # noqa: E402
from core import contracts as C              # noqa: E402
from core import pipeline as pipe            # noqa: E402
from server.jobs import JOB_MANAGER          # noqa: E402

# ----------------------------------------------------------------------
# 数据根目录：云端用环境变量覆盖（例如挂载卷），默认项目下 runtime/
# ----------------------------------------------------------------------
DATA_ROOT = os.environ.get("VIM_DATA_DIR") or os.path.join(ROOT, "runtime")
CACHE_ROOT = os.environ.get("VIM_CACHE_DIR") or os.path.join(DATA_ROOT, "_cache")
os.makedirs(DATA_ROOT, exist_ok=True)
os.makedirs(CACHE_ROOT, exist_ok=True)

app = FastAPI(
    title="价值投资智能分析模型（云端版）",
    description="海选 → 年报数据搜索汇总 → 造假排雷 → 16维度 → 企业AI深度 → 好价估值",
    version="2.0.0",
)
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"], allow_credentials=True,
    allow_methods=["*"], allow_headers=["*"],
)


# ======================================================================
# 模型
# ======================================================================
class JobRequest(BaseModel):
    stages: List[str] = Field(default_factory=lambda: list(pipe.STAGE_KEYS),
                              description="要运行的阶段 key，顺序不限（会按标准顺序执行）")
    params: Dict[str, Any] = Field(default_factory=dict, description="各阶段参数，key 与阶段同名")
    user_config: Optional[Dict[str, Any]] = Field(default=None, description="一次性带入的用户配置")


class ConfigRequest(BaseModel):
    config: Dict[str, Any]


# ======================================================================
# 基础接口
# ======================================================================
@app.get("/api/health")
def health() -> Dict[str, Any]:
    return {"ok": True, "app": "value-invest-cloud", "version": "2.0.0",
            "stages": pipe.STAGES, "data_root": DATA_ROOT}


@app.get("/api/stages")
def stages() -> Dict[str, Any]:
    return {"stages": pipe.STAGES}


@app.get("/api/config")
def get_config() -> Dict[str, Any]:
    cfg = cfg_mod.load_user_config(DATA_ROOT)
    # api_key 脱敏返回
    ai = dict(cfg.get("ai") or {})
    if ai.get("api_key"):
        k = str(ai["api_key"])
        ai["api_key_masked"] = (k[:4] + "****" + k[-2:]) if len(k) > 8 else "****"
        ai["api_key"] = ""
    cfg["ai"] = ai
    cfg["_ai_defaults"] = cfg_mod.AI_DEFAULTS
    cfg["_fraud_defaults"] = cfg_mod.default_fraud_params()
    cfg["_fraud_meta"] = {k: v["name"] for k, v in cfg_mod.FRAUD_PARAM_DEFS.items()}
    cfg["_dim16_defaults"] = cfg_mod.default_16dim_config()
    cfg["_dim16_meta"] = {k: {"desc": v["desc"], "tab": v["tab"], "is_pct": v["is_pct"]}
                          for k, v in cfg_mod.FULL_CONFIG.items()}
    cfg["_valuation_defaults"] = cfg_mod.VALUATION_DEFAULTS
    cfg["_haixuan_presets"] = cfg_mod.HAIXUAN_PRESETS
    return cfg


@app.put("/api/config")
def put_config(req: ConfigRequest) -> Dict[str, Any]:
    path = cfg_mod.save_user_config(DATA_ROOT, req.config or {})
    return {"ok": True, "path": path}


# ======================================================================
# 任务接口
# ======================================================================
@app.post("/api/jobs")
def create_job(req: JobRequest) -> Dict[str, Any]:
    keys = [k for k in req.stages if k in pipe.STAGE_KEYS]
    if not keys:
        raise HTTPException(status_code=400, detail="没有有效的阶段，请检查 stages 参数")
    # 按标准顺序执行（依赖关系：海选 → 年报汇总 → 后续）
    ordered = [s["key"] for s in pipe.STAGES if s["key"] in keys]
    job = JOB_MANAGER.submit(ordered, req.params or {}, DATA_ROOT, CACHE_ROOT)
    return {"ok": True, "job_id": job.id, "stages": ordered, "data_root": DATA_ROOT}


@app.get("/api/jobs")
def list_jobs(limit: int = 20) -> Dict[str, Any]:
    return {"jobs": JOB_MANAGER.list(limit=limit)}


@app.get("/api/jobs/{job_id}")
def job_status(job_id: str, log_tail: int = 200) -> Dict[str, Any]:
    job = JOB_MANAGER.get(job_id)
    if not job:
        raise HTTPException(status_code=404, detail="任务不存在")
    return job.snapshot(log_tail=log_tail)


@app.post("/api/jobs/{job_id}/cancel")
def cancel_job(job_id: str) -> Dict[str, Any]:
    ok = JOB_MANAGER.cancel(job_id)
    return {"ok": ok}


@app.post("/api/jobs/{job_id}/pause")
def pause_job(job_id: str, paused: bool = True) -> Dict[str, Any]:
    """暂停 / 继续任务（协作式：在阶段内的循环检查点生效）。"""
    ok = JOB_MANAGER.pause(job_id, paused=paused)
    return {"ok": ok, "paused": paused}


# ======================================================================
# 文件接口
# ======================================================================
def _inside(path: str) -> bool:
    try:
        return os.path.commonpath([os.path.abspath(path), os.path.abspath(DATA_ROOT)]) == os.path.abspath(DATA_ROOT)
    except Exception:
        return False


@app.get("/api/files")
def list_files(subdir: str = "", exts: str = ".xlsx,.docx,.csv,.json") -> Dict[str, Any]:
    base = os.path.join(DATA_ROOT, subdir) if subdir else DATA_ROOT
    if not _inside(base) or not os.path.isdir(base):
        raise HTTPException(status_code=404, detail="目录不存在")
    allow = tuple(x.strip().lower() for x in exts.split(",") if x.strip())
    items: List[Dict[str, Any]] = []
    for r, dirs, files in os.walk(base):
        dirs[:] = [d for d in dirs if not d.startswith("_") and d != "_cache"]
        for f in files:
            if f.startswith("~$"):
                continue
            if allow and not f.lower().endswith(allow):
                continue
            p = os.path.join(r, f)
            try:
                st = os.stat(p)
            except Exception:
                continue
            items.append({
                "name": f,
                "rel": os.path.relpath(p, DATA_ROOT).replace("\\", "/"),
                "dir": os.path.relpath(r, DATA_ROOT).replace("\\", "/"),
                "size": st.st_size,
                "mtime": st.st_mtime,
            })
    items.sort(key=lambda x: -x["mtime"])
    return {"ok": True, "root": DATA_ROOT, "files": items[:500]}


@app.get("/api/download")
def download(rel: str = Query(..., description="相对 data_root 的文件路径")):
    p = os.path.join(DATA_ROOT, rel)
    if not _inside(p) or not os.path.isfile(p):
        raise HTTPException(status_code=404, detail="文件不存在")
    return FileResponse(p, filename=os.path.basename(p))


@app.get("/api/summary")
def summary() -> JSONResponse:
    """一眼看清各阶段产物是否齐全（前端首页用）。"""
    out: Dict[str, Any] = {}
    for d in (C.DIR_HAIXUAN, C.DIR_EXTRACT, C.DIR_FRAUD, C.DIR_16DIM,
              C.DIR_ENTERPRISE, C.DIR_VALUATION):
        p = os.path.join(DATA_ROOT, d)
        cnt = 0
        if os.path.isdir(p):
            for r, _dirs, files in os.walk(p):
                cnt += sum(1 for f in files if not f.startswith("~$"))
        out[d] = {"exists": os.path.isdir(p), "files": cnt}
    return JSONResponse({"ok": True, "dirs": out, "data_root": DATA_ROOT})


# ======================================================================
# 直接运行
# ======================================================================
if __name__ == "__main__":  # pragma: no cover
    import uvicorn
    uvicorn.run("server.main:app", host="0.0.0.0", port=int(os.environ.get("PORT", "8000")), reload=False)
