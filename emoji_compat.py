#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
按钮图标兼容层 —— 解决麒麟 / 统信上「按钮前面的 emoji 有的显示、有的显示不出来」。

现象：
    Windows 上按钮文字里的 👁 🖨 💾 🗑 ⬅ ➡ 📝 等 emoji 由 Segoe UI Emoji
    正常渲染；麒麟 V10 上 Tk 8.6 用的是普通轮廓字体（Noto Sans CJK SC），
    不含这些 emoji 字形，于是部分按钮显示成方框 / 空白。同一界面里少数符号
    （如 ⚙、✓）恰好在字体里有，就显得「有些可以、有些不行」。

策略：
    - Windows：原样保留全部 emoji（与旧版一致）
    - Linux：检测是否有 **轮廓** emoji 字体（Noto Color Emoji 是 CBDT 位图，
      Tk 8.6 渲染不了，不算）；没有就剥掉控件文本开头的 emoji，
      只保留中文标签，保证每个按钮都可读、可点，不会出现方框。

环境变量 COMPLAINTS_EMOJI：
    auto  默认，按上面的规则判断
    on    强制保留 emoji
    off   强制剥掉 emoji

用法：在 gui_app.py 里 `from emoji_compat import install_emoji_compat`
      然后 `install_emoji_compat(ttk)`（ttk 是 ttk_compat.ttk 模块对象）。
"""

import os
import re
import shutil
import subprocess
import sys

# 常见 emoji 区段：杂项符号、装饰符号、补充符号、表情、传输/箭头等
_EMOJI_RE = re.compile(
    '[\U0001F000-\U0001FAFF\u2600-\u27BF\u2B00-\u2BFF\u2190-\u21FF\u2300-\u23FF]'
)

# 只有真正带 emoji 字形、且是轮廓（TTF/OTF）的字体才能被 Tk 8.6 使用。
# 之前把 DejaVu / FreeSerif 也算作「有 emoji 字体」，但这两个只覆盖极少数
# 符号（⚙ ✓ ➡），覆盖不了 🖨 💾 📝 📋 这类表情 emoji，于是判定为「可显示」
# 却仍然一片方框——这就是「有些按钮图标还是不显示」的原因。
# 现在只认真正的 emoji 字体。
_OUTLINE_EMOJI_FONT_HINTS = (
    "symbola",        # 覆盖最全的轮廓符号字体，麒麟上最常装
    "emojione",
    "twemoji",
    "noto emoji",     # 轮廓版（不是 Noto Color Emoji）
    "notoemoji",
    "openmoji",
    "emojisymbols",
)

# 位图（CBDT/SBIX）emoji 字体 Tk 8.6 渲染不了，出现即视为不可用
_BITMAP_EMOJI_HINTS = ("color emoji", "segoe ui emoji")

_enabled = None
_patched = False
_stripped_count = 0
_counter = [0]  # 闭包里用 list 计数，避免 UnboundLocalError


def _env_pref():
    return (os.environ.get('COMPLAINTS_EMOJI') or 'auto').strip().lower()


def _font_families():
    """返回系统字体家族名（小写）。没有 fc-list 时返回空串。"""
    if not shutil.which('fc-list'):
        return ''
    try:
        r = subprocess.run(['fc-list', ':', 'family'], stdout=subprocess.PIPE,
                           stderr=subprocess.DEVNULL, timeout=8)
        return r.stdout.decode('utf-8', 'ignore').lower()
    except Exception:
        return ''


def _has_outline_emoji_font():
    """
    是否存在 Tk 8.6 能渲染的 **轮廓** emoji 字体。

    只装 Noto Color Emoji（位图）不算 —— Tk 画不出来，仍会显示方框。
    """
    text = _font_families()
    if not text:
        return False
    # 先排除「只有位图」的情况：命中了位图字体但没有任何轮廓字体
    has_outline = any(h in text for h in _OUTLINE_EMOJI_FONT_HINTS)
    return has_outline


def font_report():
    """诊断用：说明为什么保留或剥掉 emoji。"""
    text = _font_families()
    if not text:
        return "未找到 fc-list，无法探测 emoji 字体"
    found = [h for h in _OUTLINE_EMOJI_FONT_HINTS if h in text]
    bitmap = [h for h in _BITMAP_EMOJI_HINTS if h in text]
    parts = []
    if found:
        parts.append("轮廓 emoji 字体: " + ", ".join(found))
    if bitmap:
        parts.append("位图 emoji 字体(Tk 不可用): " + ", ".join(bitmap))
    return "; ".join(parts) if parts else "无 emoji 字体"


def emoji_enabled():
    """当前是否保留 emoji。"""
    global _enabled
    if _enabled is not None:
        return _enabled
    pref = _env_pref()
    if pref == 'on':
        _enabled = True
    elif pref == 'off':
        _enabled = False
    elif sys.platform == 'win32':
        _enabled = True
    else:
        _enabled = _has_outline_emoji_font()
    return _enabled


def strip_emoji(text):
    """剥掉文本里的 emoji 字符，清理多余空白；没有 emoji 时原样返回。"""
    if not isinstance(text, str) or not _EMOJI_RE.search(text):
        return text
    cleaned = _EMOJI_RE.sub('', text)
    # 去掉 emoji 后常见的 " 中文" 前导空格与重复空格
    cleaned = cleaned.strip()
    cleaned = re.sub(r'\s{2,}', ' ', cleaned)
    return cleaned


def status():
    pref = _env_pref()
    state = "保留" if emoji_enabled() else "剥掉（避免方框）"
    count = _counter[0]
    return f"emoji={pref}/{state}" + (f"，已处理 {count} 处" if count else "")


def install_emoji_compat(ttk_module):
    """
    给 ttkbootstrap 的常用文本控件打轻量补丁：创建 **以及后续改文本** 时
    都统一过滤 emoji。

    只包 `__init__` 不够：代码里大量 `label.config(text="📋 ...")` 是在
    控件建好之后更新文字的（例如收文编号），那些路径绕过了创建时的过滤，
    于是麒麟上仍会显示方框。这里同时包 `configure`（`config` 是它的别名）。
    """
    global _patched
    if _patched or emoji_enabled():
        return _patched
    _patched = True

    def _fix_kwargs(kwargs):
        text = kwargs.get('text')
        if isinstance(text, str):
            fixed = strip_emoji(text)
            if fixed != text:
                kwargs['text'] = fixed
                _counter[0] += 1

    def _make_init(orig):
        def __init__(self, *args, **kwargs):
            _fix_kwargs(kwargs)
            return orig(self, *args, **kwargs)
        __init__.__name__ = getattr(orig, '__name__', '__init__')
        __init__.__doc__ = orig.__doc__
        return __init__

    def _make_configure(orig):
        def configure(self, cnf=None, **kwargs):
            # 形式一：w.configure(text="📋 x")  -> cnf 为 None，值在 kwargs
            _fix_kwargs(kwargs)
            # 形式二：w.configure('text', '📋 x') -> cnf='text'，值在 kwargs['value']
            if cnf == 'text' and 'value' in kwargs and isinstance(kwargs['value'], str):
                fixed = strip_emoji(kwargs['value'])
                if fixed != kwargs['value']:
                    kwargs['value'] = fixed
                    _counter[0] += 1
            return orig(self, cnf, **kwargs)
        configure.__name__ = getattr(orig, '__name__', 'configure')
        configure.__doc__ = orig.__doc__
        return configure

    for name in ('Button', 'Label', 'Checkbutton', 'Radiobutton',
                 'LabelFrame', 'ToggleButton', 'Treeview', 'Notebook'):
        cls = getattr(ttk_module, name, None)
        if cls is None:
            continue
        try:
            cls.__init__ = _make_init(cls.__init__)
        except Exception:
            pass
        try:
            if getattr(cls, 'configure', None) is not None:
                cls.configure = _make_configure(cls.configure)
                cls.config = cls.configure
        except Exception:
            pass

    return True
