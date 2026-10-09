# -*- coding: utf-8 -*-
"""
Streamlit 入口（仓库根版本）

Streamlit Community Cloud 的 Main file path 默认就是 `app.py`。
本文件只做两件事：
  1. 把项目根（含 `core/` 的目录）加入 sys.path；
  2. 执行真正的前端 `webapp/app.py`（用真实路径设置 __file__，保证它内部定位正确）。

因此下面两种 Main file path **都能正常工作**：
    app.py             ← 本文件（仓库根）
    webapp/app.py      ← 前端本体
"""
from __future__ import annotations

import os
import sys

_BOOT_CANDIDATES = [os.path.dirname(os.path.abspath(__file__)), os.getcwd(), "/mount/src"]
for _b in list(_BOOT_CANDIDATES):
    _cur = _b
    for _ in range(6):
        _p = os.path.dirname(_cur)
        if not _p or _p == _cur:
            break
        _BOOT_CANDIDATES.append(_p)
        _cur = _p

_ROOT = None
for _b in _BOOT_CANDIDATES:
    try:
        if _b and os.path.isfile(os.path.join(_b, "core", "pipeline.py")):
            _ROOT = _b
            break
    except Exception:
        continue

if _ROOT is None:
    import streamlit as st
    st.set_page_config(page_title="启动失败", page_icon="⚠️", layout="centered")
    st.error("### 未找到 `core/` 目录")
    st.markdown(
        "请确认仓库里 `core/`（含 `core/pipeline.py`）与 `app.py` **同级**。\n\n"
        "正确结构：\n```\n"
        "<仓库根>/app.py\n<仓库根>/core/…\n<仓库根>/webapp/app.py\n"
        "<仓库根>/requirements.txt\n<仓库根>/packages.txt\n```"
    )
    with st.expander("诊断信息"):
        st.write({"入口": os.path.abspath(__file__), "cwd": os.getcwd()})
        _seen = []
        for _b in _BOOT_CANDIDATES[:10]:
            try:
                _seen.append(f"{_b} -> {sorted(os.listdir(_b))[:20]}")
            except Exception as _e:
                _seen.append(f"{_b} -> <无法读取: {_e}>")
        st.code("\n".join(_seen))
    # 关键：显式终止，不依赖 st.stop() 的“停止执行”语义
    st.stop()
    raise SystemExit(0)

if _ROOT not in sys.path:
    sys.path.insert(0, _ROOT)

_TARGET = os.path.join(_ROOT, "webapp", "app.py")
if not os.path.isfile(_TARGET):
    import streamlit as st
    st.set_page_config(page_title="启动失败", page_icon="⚠️", layout="centered")
    st.error("### 未找到前端文件 `webapp/app.py`")
    st.code(f"项目根：{_ROOT}\n期望文件：{_TARGET}")
    st.stop()
    raise SystemExit(0)

# 用真实路径与 __name__ 执行前端，等价于把它当成入口脚本
with open(_TARGET, "r", encoding="utf-8") as _f:
    _code = compile(_f.read(), _TARGET, "exec")
exec(_code, {"__name__": "__main__", "__file__": _TARGET})
