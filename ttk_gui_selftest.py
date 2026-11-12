#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
无头 GUI 冒烟测试（CI 用，需要 Xvfb 提供虚拟 X11）

为什么需要它：
    CI 原来只做 `import` 和 `py_compile`，这两步都**不会**建窗口。
    2026-11-12 麒麟现场的崩溃发生在 `create_widgets()` —— import 全部通过、
    窗口建到一半才炸。只有真正把界面搭起来才能抓到这类问题。

本脚本做四件事：
    1. 控件自检（ttk_compat.verify_widgets）
    2. 建主窗口 + 规则管理窗口 + 全部工具对话框，逐个销毁
    3. 结果同时打印到标准输出和 AppData/ttk_selftest.log
       （PyInstaller console=False 打包后 stdout 为 None，只能靠这个文件取证）
    4. 用退出码表达结论：0 全通过；1 有失败项

用法：
    xvfb-run -a python3 ttk_gui_selftest.py
    ./投诉批办单管理系统 --selftest          # 打包产物自检
    python3 gui_app.py --selftest            # 源码自检
"""

import os
import sys
import traceback

_HERE = os.path.dirname(os.path.abspath(__file__))
if _HERE not in sys.path:
    sys.path.insert(0, _HERE)

_LOG_PATH = None
_LINES = []


def _report_path():
    """报告文件位置：优先 AppData/，退化到临时目录（AppData 不可写时）。"""
    global _LOG_PATH
    if _LOG_PATH is not None:
        return _LOG_PATH
    candidates = []
    try:
        import platform_compat as pc
        candidates.append(os.path.join(pc.get_appdata_path(), "ttk_selftest.log"))
    except Exception:
        pass
    candidates.append(os.path.join(_HERE, "AppData", "ttk_selftest.log"))
    candidates.append(os.path.join("/tmp", "ttk_selftest.log"))
    for p in candidates:
        try:
            d = os.path.dirname(p)
            if d and not os.path.isdir(d):
                os.makedirs(d, exist_ok=True)
            with open(p, "w", encoding="utf-8"):
                pass
            _LOG_PATH = p
            break
        except Exception:
            continue
    return _LOG_PATH


def out(msg=""):
    """同时写标准输出与报告文件。console=False 下 stdout 为 None，print 本身安全。"""
    _LINES.append(msg)
    try:
        print(msg)
    except Exception:
        pass


def flush_report():
    path = _report_path()
    if not path:
        return
    try:
        with open(path, "a", encoding="utf-8") as f:
            f.write("\n".join(_LINES) + "\n")
    except Exception:
        pass


FAILS = []


def _fail(label, exc):
    FAILS.append(f"{label}: {type(exc).__name__}: {exc}")
    out(f"  FAIL {label}")
    try:
        traceback.print_exc(limit=4)
    except Exception:
        pass
    # traceback 在 console=False 下进不了文件，手动记一份
    _LINES.append("".join(traceback.format_exc(limit=4)))


def main():
    out("=" * 62)
    try:
        from importlib.metadata import version
        out("ttkbootstrap: " + version("ttkbootstrap"))
    except Exception:
        out("ttkbootstrap: <未安装或元数据缺失>")
    out("python      : " + sys.version.split()[0])
    out("frozen      : " + str(bool(getattr(sys, "frozen", False))))
    out("DISPLAY     : " + os.environ.get("DISPLAY", "<空>"))
    out("=" * 62)

    # 1) 控件自检
    try:
        import ttk_compat
        _ver, missing = ttk_compat.verify_widgets()
        out("控件自检    : " + ttk_compat.describe())
        if missing:
            FAILS.append("缺少控件: " + ",".join(missing))
            out("  FAIL 缺少控件: " + str(missing))
    except Exception as e:
        _fail("ttk_compat 自检", e)
        out("RESULT: FAILED")
        flush_report()
        return 1

    try:
        import tkinter as tk
        import tkinter.messagebox as mb
    except Exception as e:
        _fail("import tkinter", e)
        out("RESULT: FAILED")
        flush_report()
        return 1

    # 对话框会阻塞，自检时全部打成空操作
    for name in ("showinfo", "showwarning", "showerror", "askyesno",
                 "askokcancel", "askquestion", "askstring", "askinteger"):
        setattr(mb, name, lambda *a, **k: (False if name.startswith("ask") else None))

    _orig_toplevel = tk.Toplevel.__init__

    def _patched_toplevel(self, *a, **k):
        _orig_toplevel(self, *a, **k)
        self.grab_set = lambda *x, **y: None
        self.wait_window = lambda *x, **y: None

    tk.Toplevel.__init__ = _patched_toplevel

    try:
        import ttkbootstrap as ttk
    except Exception as e:
        _fail("import ttkbootstrap", e)
        out("RESULT: FAILED")
        flush_report()
        return 1

    # 2) 主窗口（历史上炸在这一步）
    try:
        root = ttk.Window(title="selftest", themename="litera", size=(1400, 900))
        root.withdraw()
        out("  OK   主窗口创建")
    except Exception as e:
        _fail("主窗口创建", e)
        out("RESULT: FAILED")
        flush_report()
        return 1

    try:
        import gui_app
        app = gui_app.ComplaintManagementGUI(root)
        root.update_idletasks()
        out("  OK   三栏布局; main_paned = %s | sash0 = %s" % (
            type(app.main_paned).__name__, app.main_paned.sashpos(0)))
    except Exception as e:
        _fail("GUI 初始化（create_widgets）", e)
        try:
            root.destroy()
        except Exception:
            pass
        out("RESULT: FAILED")
        flush_report()
        return 1

    # 3) 规则管理窗口（另一处 PanedWindow 使用者）
    try:
        import rule_management_ui
        rule_management_ui.RuleManagementWindow(root)
        root.update_idletasks()
        out("  OK   规则管理窗口")
    except Exception as e:
        _fail("规则管理窗口", e)
    for w in list(root.winfo_children()):
        if isinstance(w, tk.Toplevel):
            try:
                w.destroy()
            except Exception:
                pass

    # 4) 工具菜单里的全部对话框
    for label, meth in [
        ("主题切换", "change_theme"),
        ("快捷填充设置", "open_quick_fill_settings"),
        ("路径设置", "manage_paths"),
        ("编号管理", "manage_counter"),
        ("二维码设置", "manage_qrcode"),
        ("帮助", "show_help"),
        ("关于", "show_about"),
        ("数据库统计", "show_database_stats"),
        ("统计报表", "show_statistics"),
        ("保存窗口配置", "save_window_config"),
    ]:
        fn = getattr(app, meth, None)
        if fn is None:
            continue
        before = set(root.winfo_children())
        try:
            fn()
            root.update_idletasks()
            out("  OK   " + label)
        except Exception as e:
            _fail(label, e)
        for w in set(root.winfo_children()) - before:
            try:
                w.destroy()
            except Exception:
                pass

    try:
        root.destroy()
    except Exception:
        pass

    out("=" * 62)
    if FAILS:
        out("RESULT: FAILED")
        for f in FAILS:
            out("  - " + f)
        flush_report()
        return 1
    out("RESULT: ALL PASS")
    flush_report()
    return 0


if __name__ == "__main__":
    sys.exit(main())
