# -*- coding: utf-8 -*-
"""
启动引导：让 `core` / `server` 无论仓库怎么摆放都能被导入。

背景
----
Streamlit Community Cloud 把仓库挂载到 `/mount/src/<repo>`，并把「Main file path」
指定的那个文件当作脚本执行。`app.py` 的位置可能是下面任意一种：

    <repo>/app.py               ← 直接放在仓库根（Streamlit 默认就是 app.py）
    <repo>/webapp/app.py        ← 按本项目 README 推荐的位置
    <repo>/价值投资网络程序/app.py        ← 整目录上传（含中文项目文件夹）
    <repo>/价值投资网络程序/webapp/app.py

原来 `webapp/app.py` 里写的是 `sys.path.insert(0, 上一级目录)`，只有最后一种摆放才对；
放在仓库根时会去 `/mount/src` 找 `core`，于是报 ModuleNotFoundError。

本模块的做法：**从当前文件与工作目录出发，向上逐级查找含 `core/pipeline.py` 的目录**，
找到就把项目根加入 `sys.path`；找不到时给出明确的中文指引，而不是抛裸 ImportError。
"""
from __future__ import annotations

import os
import sys
from typing import List, Optional

# 判定“项目根”的标志文件
_MARKER = os.path.join("core", "pipeline.py")
# 兜底搜索起点（Streamlit Cloud 的挂载点是 /mount/src）
_EXTRA_ROOTS = ("/mount/src",)


def _candidates() -> List[str]:
    cands: List[str] = []

    def _add(p: Optional[str]) -> None:
        if not p:
            return
        p = os.path.abspath(p)
        if p and p not in cands:
            cands.append(p)

    here = os.path.dirname(os.path.abspath(__file__))
    _add(here)
    _add(os.getcwd())
    for r in _EXTRA_ROOTS:
        _add(r)

    # 从每个起点向上逐级（最多 6 层）
    for start in list(cands):
        cur = start
        for _ in range(6):
            parent = os.path.dirname(cur)
            if not parent or parent == cur:
                break
            _add(parent)
            cur = parent
    return cands


def find_project_root() -> Optional[str]:
    """返回含 core/pipeline.py 的目录；找不到返回 None。"""
    for c in _candidates():
        try:
            if os.path.isfile(os.path.join(c, _MARKER)):
                return c
        except Exception:
            continue
    return None


def setup_path(verbose: bool = False) -> Optional[str]:
    """把项目根加入 sys.path（幂等）。返回项目根或 None。"""
    root = find_project_root()
    if root and root not in sys.path:
        sys.path.insert(0, root)
    if verbose:
        print(f"[bootstrap] project_root={root}  sys.path[:3]={sys.path[:3]}", flush=True)
    return root


def require_core() -> str:
    """确保 core 可导入；失败时抛出带中文指引的 RuntimeError（便于在页面上显示）。"""
    root = setup_path()
    try:
        import importlib
        importlib.import_module("core.pipeline")
    except Exception as exc:  # noqa: BLE001
        listing = []
        for base in _candidates()[:8]:
            try:
                listing.append(f"  {base}: {sorted(os.listdir(base))[:20]}")
            except Exception:
                listing.append(f"  {base}: <无法读取>")
        raise RuntimeError(
            "未找到项目代码目录 `core/`，无法启动。\n\n"
            "请确认 GitHub 仓库里与 `app.py` **同级** 存在 `core/` 文件夹"
            "（即 `core/pipeline.py`、`core/screening.py` 等都在）。\n"
            "本项目的正确结构是：\n"
            "    <仓库根>/app.py            ← Streamlit 的 Main file path\n"
            "    <仓库根>/core/…            ← 业务代码\n"
            "    <仓库根>/server/…  <仓库根>/webapp/…\n"
            "    <仓库根>/requirements.txt  <仓库根>/packages.txt\n\n"
            f"已查找过的位置：\n" + "\n".join(listing) + f"\n\n原始错误：{type(exc).__name__}: {exc}"
        ) from exc
    return root or ""


# 导入本模块即自动生效（放在 app.py 最前面 import 即可）
PROJECT_ROOT = setup_path(verbose=os.environ.get("VIM_DEBUG_BOOTSTRAP") == "1")
