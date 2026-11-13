#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
界面布局辅助 —— 解决麒麟 / 统信小屏、高 DPI 下「窗口底部按钮看不见」的问题。

现场反馈（麒麟 V10）：
  1. 二维码设置窗口固定 500x580 且 resizable(False, False)，屏幕稍矮或缩放后
     最下面「保存 / 取消」两个按钮被挤出可视区，且不能拖大窗口。
  2. 批办单编辑页的表单放在 Canvas 里，但没绑鼠标滚轮，内容比可视区高时
     用户滚不动，最下面一排按钮永远露不出来。

这里提供三个与业务无关的小工具：
  - fit_dialog_to_screen()：按屏幕可用尺寸算初始窗口大小（带最小尺寸下限），
    允许拖动调整，并居中。
  - bind_mousewheel()：给 Canvas / Text / Listbox 等可滚动控件绑定滚轮
    （Linux 用 Button-4/5，Windows/macOS 用 MouseWheel），支持嵌套容器。
  - make_scrollable_frame()：一次性创建「Canvas + 竖向滚动条 + 内容 Frame」，
    自动处理 scrollregion、滚轮和宽度自适应。
"""

import tkinter as tk


def screen_workarea(widget):
    """返回 (可用宽, 可用高)。优先用 winfo_screen*，异常时给保守默认值。"""
    try:
        w = widget.winfo_screenwidth()
        h = widget.winfo_screenheight()
    except Exception:
        return 1280, 720
    # 给系统任务栏 / 标题栏留余量；国产桌面常见的顶栏 30~60px
    return max(640, w), max(480, h - 90)


def fit_dialog_to_screen(dialog, parent=None, width=500, height=580,
                         min_width=420, min_height=420, max_ratio=0.90,
                         resizable=True, center=True):
    """
    设置一个 Toplevel 的初始几何尺寸，使其在当前屏幕上完整可见。

    参数：
        width/height  期望尺寸
        min_*         最小尺寸（低于它内容会挤压，靠滚动条兜底）
        max_ratio     最多占屏幕可用宽/高的比例
        resizable     是否允许拖动调整（默认允许）
        center        是否居中到 parent 或屏幕

    返回实际使用的 (w, h)。
    """
    ref = parent if parent is not None else dialog
    sw, sh = screen_workarea(ref)

    w = min(int(width), int(sw * max_ratio))
    h = min(int(height), int(sh * max_ratio))
    w = max(min_width, w)
    h = max(min_height, h)
    # 下限也不能超过屏幕，否则又会被裁掉
    w = min(w, sw)
    h = min(h, sh)

    try:
        dialog.resizable(resizable, resizable)
        if resizable:
            dialog.minsize(min(min_width, w), min(min_height, h))
    except Exception:
        pass

    if center:
        try:
            dialog.update_idletasks()
            if parent is not None:
                px, py = parent.winfo_rootx(), parent.winfo_rooty()
                pw, ph = parent.winfo_width(), parent.winfo_height()
            else:
                px, py, pw, ph = 0, 0, sw, sh
            x = px + max(0, (pw - w) // 2)
            y = py + max(0, (ph - h) // 2)
            x = min(x, max(0, sw - w))
            y = min(y, max(0, sh - h))
            dialog.geometry(f"{w}x{h}+{x}+{y}")
        except Exception:
            dialog.geometry(f"{w}x{h}")
    else:
        dialog.geometry(f"{w}x{h}")

    return w, h


def bind_mousewheel(widget, target=None, recursive=True):
    """
    给可滚动控件绑定滚轮。target 是真正要 yview 的控件（默认 widget 本身）。

    Linux/X11 上 Tk 发的是 Button-4 / Button-5，Windows/macOS 发的是 MouseWheel；
    两者都绑。表单控件（Entry / Text / Frame）常是之后才创建的，且 <Enter>
    不会向父级冒泡，所以除直接绑定外，还在所属 Toplevel 上用 bind_all 兜底：

      * 每个滚动区（以 target 的归属根为准）只注册一个 handler，避免一次滚轮
        滚动多格；
      * handler 先判断事件来源是否落在本滚动区内，只处理自己这块，不干扰
        同窗口的文件列表、日志框等其它滚动区域。
    """
    if target is None:
        target = widget

    # 归属根：优先用 canvas 本身（content Frame 是它的子窗口）
    scope = target

    def _is_ours(w):
        try:
            while w is not None:
                if w == scope or w == widget:
                    return True
                w = w.master
        except Exception:
            pass
        return False

    def _on_wheel(event):
        if not _is_ours(getattr(event, 'widget', None)):
            return None
        try:
            if getattr(event, 'num', None) == 5 or getattr(event, 'delta', 0) < 0:
                target.yview_scroll(3, 'units')
            elif getattr(event, 'num', None) == 4 or getattr(event, 'delta', 0) > 0:
                target.yview_scroll(-3, 'units')
        except Exception:
            pass
        return 'break'

    def _bind_one(w):
        try:
            w.bind('<MouseWheel>', _on_wheel)
            w.bind('<Button-4>', _on_wheel)
            w.bind('<Button-5>', _on_wheel)
        except Exception:
            pass

    def _walk(w):
        for c in w.winfo_children():
            _bind_one(c)
            _walk(c)

    _bind_one(widget)
    if recursive:
        try:
            _walk(widget)
        except Exception:
            pass

    # 动态创建的子控件：Toplevel 级 bind_all 兜底，按 scope 去重
    try:
        top = target.winfo_toplevel()
        key = '_ui_layout_wheel_scopes'
        scopes = getattr(top, key, None)
        if scopes is None:
            scopes = set()
            setattr(top, key, scopes)
        scope_id = str(target)
        if scope_id not in scopes:
            scopes.add(scope_id)
            for seq in ('<MouseWheel>', '<Button-4>', '<Button-5>'):
                top.bind_all(seq, _on_wheel, add='+')
    except Exception:
        pass

    return _on_wheel


def make_scrollable_frame(parent, ttk_module, bind_wheel=True):
    """
    创建 Canvas + 竖向滚动条 + 内容 Frame，返回 (canvas, content_frame, outer_frame)。

    outer_frame 已经 pack(fill=BOTH, expand=True)，调用方只需往 content_frame
    里放控件。Canvas 宽度变化时内容 Frame 自动跟随，scrollregion 自动更新。
    """
    outer = ttk_module.Frame(parent)
    outer.pack(fill=tk.BOTH, expand=True)

    canvas = tk.Canvas(outer, highlightthickness=0, borderwidth=0)
    scrollbar = ttk_module.Scrollbar(outer, orient='vertical', command=canvas.yview)
    content = ttk_module.Frame(canvas)

    win_id = canvas.create_window((0, 0), window=content, anchor='nw')
    canvas.configure(yscrollcommand=scrollbar.set)

    scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
    canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

    def _on_configure(_event=None):
        try:
            canvas.configure(scrollregion=canvas.bbox('all'))
        except Exception:
            pass

    def _on_canvas_resize(event):
        try:
            canvas.itemconfigure(win_id, width=max(event.width, 1))
        except Exception:
            pass

    content.bind('<Configure>', _on_configure)
    canvas.bind('<Configure>', _on_canvas_resize)

    if bind_wheel:
        bind_mousewheel(canvas, canvas)
        bind_mousewheel(content, canvas)

    return canvas, content, outer


def ensure_visible_bottom(parent, ttk_module, button_frame_builder):
    """
    在 parent 里创建一个「永远贴在底部、不会被内容挤走」的按钮区。

    用法：先调用本函数拿到 button_frame，再往里面放按钮；
    之后才 pack 主体内容。这样即使内容超高，按钮仍留在可视区底部。
    """
    bar = ttk_module.Frame(parent)
    bar.pack(fill=tk.X, side=tk.BOTTOM, pady=(6, 0))
    button_frame_builder(bar)
    return bar
