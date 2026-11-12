#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
ttkbootstrap 1.x / 2.x 兼容层
===========================================================================
实测现场（麒麟 V10 aarch64，2026-11-12）：

    ttkbootstrap 版本: 2.2.7
    AttributeError: module 'ttkbootstrap' has no attribute 'PanedWindow'
      File "gui_app.py", line 763, in create_widgets
        self.main_paned = ttk.PanedWindow(self.root, orient=tk.HORIZONTAL)

原因：
    ttkbootstrap 2.0 起把分割面板类统一改名为 ``Panedwindow``（与
    ``tkinter.ttk`` 完全一致的拼写），1.x 里的驼峰别名 ``PanedWindow``
    没有保留，于是老代码 ``ttk.PanedWindow(...)`` 在 2.x 上直接抛
    AttributeError，主窗口建到一半就崩，表现为「启动失败」弹窗。

本模块在 ``import`` 时做两件事：

    1. **补齐别名**：把 1.x 的驼峰写法映射到 2.x 的真实类上，并反向补齐
       2.x 新增拼写；两个版本的写法都能用，业务代码不需要改调用方式。
    2. **启动自检**：``verify_widgets()`` 一次性检查本项目真正用到的全部
       控件是否可用，缺失时直接给出可执行的修复命令，而不是让界面建到
       三分之二才炸。

用法（把原来的 ``import ttkbootstrap as ttk`` 换成下面一行即可）::

    from ttk_compat import ttk

    # 其余代码一个字都不用改：
    self.main_paned = ttk.PanedWindow(self.root, orient=tk.HORIZONTAL)
"""

import sys

import tkinter.ttk as _tk_ttk

try:
    import ttkbootstrap as _tb
except Exception as _exc:  # pragma: no cover - 依赖缺失时给中文提示
    raise ImportError(
        "无法导入 ttkbootstrap，请先安装图形界面依赖：\n"
        "    python3 -m pip install 'ttkbootstrap>=1.10.1,<3'"
    ) from _exc


# ---------------------------------------------------------------------------
# 版本号：1.x 不导出 __version__，2.x 才有，统一走 importlib.metadata
# ---------------------------------------------------------------------------

def version():
    """返回 ttkbootstrap 发行版本号字符串（取不到时返回 'unknown'）。"""
    v = getattr(_tb, "__version__", None)
    if v:
        return str(v)
    try:
        from importlib.metadata import version as _dist_version
        return _dist_version("ttkbootstrap")
    except Exception:
        return "unknown"


# ---------------------------------------------------------------------------
# 别名表：1.x 驼峰名 <-> 2.x / tkinter.ttk 标准拼写（双向）
# ---------------------------------------------------------------------------
# 左边是历史代码里写的名字，右边是 ttk 标准拼写（2.x 与 tkinter.ttk 都用它）。
# 只列真正需要互通的别名对；两个拼写指向同一个控件时才补，否则不猜。
_ALIAS_PAIRS = (
    ("PanedWindow", "Panedwindow"),   # 2.0 起只剩 Panedwindow
    ("LabelFrame", "Labelframe"),
)

_repaired = []


def _resolve(name):
    """按优先级解析一个控件类：ttkbootstrap 模块 -> ttkbootstrap.widgets -> tkinter.ttk。"""
    for holder in (_tb, getattr(_tb, "widgets", None), getattr(_tb, "ttk", None), _tk_ttk):
        if holder is None:
            continue
        obj = getattr(holder, name, None)
        if obj is not None:
            return obj
    return None


def _repair_aliases():
    """把缺失的拼写补到 ttkbootstrap 模块上，返回本次实际补上的名字列表。"""
    added = []
    for camel, standard in _ALIAS_PAIRS:
        for missing, source in ((camel, standard), (standard, camel)):
            if hasattr(_tb, missing):
                continue
            obj = _resolve(source)
            if obj is None:
                continue
            try:
                setattr(_tb, missing, obj)
            except Exception:
                continue
            added.append(missing)
    return added


_repaired = _repair_aliases()


# ---------------------------------------------------------------------------
# 启动自检：本项目用到的全部 ttkbootstrap 控件（见 gui_app.py / rule_management_ui.py）
# ---------------------------------------------------------------------------
REQUIRED_WIDGETS = (
    "Window",        # ttkbootstrap 的根窗口（App 的永久别名）
    "Frame",
    "Label",
    "Button",
    "Entry",
    "Combobox",
    "Checkbutton",
    "Radiobutton",
    "Scrollbar",
    "Separator",
    "Spinbox",
    "Treeview",
    "Notebook",
    "LabelFrame",
    "PanedWindow",   # 1.x 驼峰名，2.x 由本模块补齐
)


def verify_widgets():
    """
    检查依赖控件是否齐全。

    返回 (版本号, 缺失控件名列表)。缺失列表为空表示可以安全启动。
    """
    missing = [name for name in REQUIRED_WIDGETS if not hasattr(_tb, name)]
    return version(), missing


def describe():
    """一行诊断信息，写日志用。"""
    v, missing = verify_widgets()
    state = "正常" if not missing else "缺少 " + ",".join(missing)
    fixed = ("，已自动补齐 " + ",".join(_repaired)) if _repaired else ""
    return f"ttkbootstrap {v}（Python {sys.version.split()[0]}，控件别名{state}{fixed}）"


# 对外暴露的模块对象：业务代码 `from ttk_compat import ttk` 后，ttk.XXX 与官方用法一致
ttk = _tb
