# -*- coding: utf-8 -*-
"""
任务调度层：单个进程内的后台线程任务管理（适配 Streamlit 云端单实例场景）。

- 不依赖 Celery/Redis，进程重启即清空（云端本来也不需要持久队列）；
- 每个任务保存：状态、阶段进度、环形日志、结构化结果、产物文件列表；
- 线程安全：所有写操作加锁；日志用 deque 限长，避免内存膨胀。
"""
from __future__ import annotations

import os
import threading
import time
import uuid
from collections import deque
from dataclasses import dataclass, field
from typing import Any, Callable, Deque, Dict, List, Optional

from core import config as cfg_mod
from core import pipeline as pipe
from core.common import make_logger

MAX_LOG_LINES = 4000


@dataclass
class Job:
    id: str
    stage_keys: List[str]
    params: Dict[str, Any]
    data_dir: str
    cache_dir: str
    status: str = "pending"           # pending / running / paused / done / failed / cancelled
    created_at: float = field(default_factory=time.time)
    started_at: Optional[float] = None
    finished_at: Optional[float] = None
    stage_index: int = 0
    stage_total: int = 0
    stage_name: str = ""
    step_done: int = 0
    step_total: int = 0
    step_label: str = ""
    logs: Deque[str] = field(default_factory=lambda: deque(maxlen=MAX_LOG_LINES))
    result: Optional[Dict[str, Any]] = None
    error: str = ""
    # ---- 协作式暂停 / 停止 ----
    pause_flag: bool = False
    stop_flag: bool = False
    _lock: threading.Lock = field(default_factory=threading.Lock, repr=False)

    # ---------- 控制 ----------
    def checkpoint(self) -> None:
        """在阶段内的循环里调用：暂停时阻塞等待，停止时抛出中断。

        之所以做成“检查点”而不是真挂起线程：所有耗时都发生在同步网络请求里，
        无法安全地打断；在循环边界检查是最稳、且不会破坏数据一致性的做法。
        """
        import time as _t
        while self.pause_flag and not self.stop_flag:
            _t.sleep(0.3)
        if self.stop_flag:
            raise KeyboardInterrupt("用户停止任务")

    def set_paused(self, paused: bool) -> None:
        self.pause_flag = bool(paused)
        if paused and self.status == "running":
            self.status = "paused"
        elif not paused and self.status == "paused":
            self.status = "running"

    # ---------- 读取 ----------
    def snapshot(self, log_tail: int = 200) -> Dict[str, Any]:
        with self._lock:
            return {
                "id": self.id,
                "status": self.status,
                "paused": self.pause_flag,
                "stopping": self.stop_flag,
                "stage_keys": list(self.stage_keys),
                "stage_index": self.stage_index,
                "stage_total": self.stage_total,
                "stage_name": self.stage_name,
                "step_done": self.step_done,
                "step_total": self.step_total,
                "step_label": self.step_label,
                "created_at": self.created_at,
                "started_at": self.started_at,
                "finished_at": self.finished_at,
                "elapsed": (self.finished_at or time.time()) - (self.started_at or self.created_at),
                "error": self.error,
                "logs": list(self.logs)[-log_tail:],
                "result": self.result,
            }

    # ---------- 写入 ----------
    def log(self, msg: str) -> None:
        with self._lock:
            self.logs.append(str(msg))

    def set_step(self, done: int, total: int, label: str = "") -> None:
        with self._lock:
            self.step_done, self.step_total, self.step_label = done, total, label


class JobManager:
    def __init__(self) -> None:
        self._jobs: Dict[str, Job] = {}
        self._lock = threading.Lock()

    # ------------------------------------------------------------------
    def submit(self, stage_keys: List[str], params: Dict[str, Any],
               data_dir: str, cache_dir: str) -> Job:
        job = Job(id=uuid.uuid4().hex[:12], stage_keys=list(stage_keys), params=dict(params or {}),
                  data_dir=data_dir, cache_dir=cache_dir,
                  stage_total=len(stage_keys))
        with self._lock:
            self._jobs[job.id] = job
        t = threading.Thread(target=self._run, args=(job,), daemon=True, name=f"job-{job.id}")
        t.start()
        return job

    def get(self, job_id: str) -> Optional[Job]:
        with self._lock:
            return self._jobs.get(job_id)

    def list(self, limit: int = 20) -> List[Dict[str, Any]]:
        with self._lock:
            jobs = sorted(self._jobs.values(), key=lambda j: -j.created_at)[:limit]
        return [j.snapshot(log_tail=0) for j in jobs]

    def cancel(self, job_id: str) -> bool:
        """停止任务：置停止标志；正在跑的阶段会在下一个检查点退出。"""
        job = self.get(job_id)
        if not job or job.status in ("done", "failed", "cancelled"):
            return False
        job.stop_flag = True
        job.pause_flag = False          # 取消暂停以便线程能走到检查点
        job.status = "cancelled"
        job.log("⏹️ 已请求停止任务，将在当前这一步完成后中断…")
        return True

    def pause(self, job_id: str, paused: bool = True) -> bool:
        """暂停 / 继续：协作式，在阶段内的循环检查点生效。"""
        job = self.get(job_id)
        if not job or job.status in ("done", "failed", "cancelled"):
            return False
        job.set_paused(paused)
        job.log("⏸️ 已暂停（当前这一步跑完后挂起）" if paused else "▶️ 已继续运行")
        return True

    # ------------------------------------------------------------------
    def _run(self, job: Job) -> None:
        job.status = "running"
        job.started_at = time.time()
        try:
            cfg = cfg_mod.load_user_config(job.data_dir)
            if isinstance(job.params.get("user_config"), dict):
                # 前端可一次性把设置带过来（避免云端多实例下读不到磁盘配置）
                for k, v in job.params["user_config"].items():
                    if isinstance(v, dict) and isinstance(cfg.get(k), dict):
                        cfg[k].update(v)
                    else:
                        cfg[k] = v
            job.log(f"🚀 任务开始：{len(job.stage_keys)} 个阶段 | 数据目录 {job.data_dir}")

            for idx, key in enumerate(job.stage_keys):
                # 阶段之间也响应停止；暂停则在此等待
                try:
                    job.checkpoint()
                except KeyboardInterrupt:
                    job.log("⏹️ 任务已停止。")
                    job.finished_at = time.time()
                    return
                if job.stop_flag:
                    job.log("⏹️ 任务已停止，放弃后续阶段。")
                    job.finished_at = time.time()
                    return
                job.stage_index = idx + 1
                job.stage_name = next((s["name"] for s in pipe.STAGES if s["key"] == key), key)
                job.set_step(0, 0, "")

                # 进度回调同时充当“协作式检查点”：阶段内的循环只要有进度就会响应暂停/停止
                def _progress(done: int, total: int, label: str = "", _job=job) -> None:
                    _job.set_step(done, total, label)
                    _job.checkpoint()

                ctx = pipe.make_context(job.data_dir, job.cache_dir, user_config=cfg,
                                        log=make_logger(job.log), progress=_progress)
                fn = pipe._STAGE_FUNCS.get(key)
                if fn is None:
                    job.log(f"⚠️ 未知阶段 {key}，跳过")
                    continue
                stage_params = job.params.get(key) if isinstance(job.params.get(key), dict) else {}
                stage_params = dict(stage_params or {})
                stage_params.setdefault("_job", job.id)
                try:
                    res = fn(ctx, stage_params)
                except KeyboardInterrupt:
                    # 用户在阶段内的检查点按了“停止”
                    job.log("⏹️ 任务已被用户停止（当前阶段中断，已生成的产物会保留）。")
                    job.status = "cancelled"
                    job.finished_at = time.time()
                    return
                except Exception as e:  # noqa: BLE001
                    import traceback
                    job.log(f"❌ 阶段【{job.stage_name}】失败：{type(e).__name__}: {e}")
                    job.log(traceback.format_exc()[:2000])
                    res = {"ok": False, "error": f"{type(e).__name__}: {e}"}
                job.result = {"stage": key, "stage_name": job.stage_name, "result": res}
                job.log(f"{'✅' if (res or {}).get('ok') else '⚠️'} 阶段【{job.stage_name}】结束")

                if key == "report" and not (res or {}).get("ok"):
                    job.log("⛔ 年报数据汇总未成功，后续依赖阶段已中止。")
                    break

            job.status = "done"
            job.finished_at = time.time()
            job.log(f"🎉 任务完成，用时 {job.finished_at - job.started_at:.0f} 秒")
        except Exception as e:  # noqa: BLE001
            import traceback
            job.status = "failed"
            job.error = f"{type(e).__name__}: {e}"
            job.finished_at = time.time()
            job.log("❌ 任务异常终止：" + job.error)
            job.log(traceback.format_exc()[:2000])


# 单例（Streamlit / FastAPI 同进程共用一个）
JOB_MANAGER = JobManager()
