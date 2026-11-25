#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
外部程序运行环境清理层 —— 解决打包版「调不动系统程序」的一类问题。

现场故障（麒麟 V10 aarch64，2026-11-25）
---------------------------------------------------------------------------
症状：点打印时 LibreOffice 转换失败，弹窗报

    /usr/lib/libreoffice/program/soffice.bin: symbol lookup error:
    /lib/aarch64-linux-gnu/libcairo.so.2: undefined symbol: FT_Done_MM_Var

根因 **不是** LibreOffice 坏了，而是 **库版本冲突**：

    PyInstaller 打包后，程序启动时会把自己 `_internal/` 目录塞进
    LD_LIBRARY_PATH（Bootloader 的机制）。我们的 _internal 里带着
    Pillow 等依赖打包进来的 `libfreetype.so.6`、`libfontconfig.so.1`，
    它们是**较旧**的版本。

    当我们 subprocess 启动 soffice 时，子进程**继承**了 LD_LIBRARY_PATH，
    于是系统的 libcairo 去找系统 freetype 的 FT_Done_MM_Var 符号时，
    实际加载的却是 _internal 里的旧 libfreetype —— 没有这个符号 —— 直接报错。

同样的污染还会影响：
    · lp / cups（打印）
    · xdg-open（打开文件、打开目录）
    · fc-list / fc-cache（字体探测与缓存，导致中文版式判断错误）
    · WPS（它自己的 Qt 依赖也可能被顶掉）
    · ImageMagick / Ghostscript 这类打印链上的转换工具

本模块做什么
---------------------------------------------------------------------------
提供一个 `clean_env()`：返回**剔除应用自身库路径**的环境变量字典，
所有「调用系统外部程序」的地方都用它。这样无论打包版装了多少自带库，
子进程看到的都是**系统自己的库**，行为与直接在终端里敲命令一致。

另外提供 `sanitize_library_path()` 在进程启动早期调用，
把本进程的 LD_LIBRARY_PATH 也清干净 —— 有些库在 dlopen 时才解析，
早清理能避免更隐蔽的问题。
"""

import os
import sys

# 这些环境变量会改变动态库的解析顺序，必须从子进程里剔除/修正
_LIB_PATH_VARS = (
    'LD_LIBRARY_PATH',
    'LD_PRELOAD',
    'LD_AUDIT',
    'DYLD_LIBRARY_PATH',      # macOS
    'DYLD_INSERT_LIBRARIES',  # macOS
)

# PyInstaller 运行时的特征变量，也应该从子进程剔除
_PYINSTALLER_VARS = (
    'LD_RUN_PATH',
    'PYTHONPATH',
    'PYTHONHOME',
    'PYTHONSTARTUP',
)


def _app_dirs():
    """
    返回「应用自身目录」集合：打包版是 _internal 与其父目录，
    源码运行时是脚本所在目录与可能的 .venv。
    """
    dirs = set()
    try:
        # PyInstaller onedir: sys._MEIPASS 指向 _internal
        meipass = getattr(sys, '_MEIPASS', None)
        if meipass:
            dirs.add(os.path.abspath(meipass))
            dirs.add(os.path.abspath(os.path.dirname(meipass)))
        if getattr(sys, 'frozen', False):
            exe_dir = os.path.dirname(os.path.abspath(sys.executable))
            dirs.add(exe_dir)
        # 源码运行
        here = os.path.dirname(os.path.abspath(__file__))
        dirs.add(here)
        venv = os.path.join(here, '.venv')
        if os.path.isdir(venv):
            dirs.add(venv)
    except Exception:
        pass
    return {d for d in dirs if d}


def _looks_like_app_path(entry, app_dirs):
    """判断某个路径条目是否属于「应用自带」，需要从子进程环境里剔除。"""
    try:
        ap = os.path.abspath(entry)
    except Exception:
        return False
    for d in app_dirs:
        if ap == d or ap.startswith(d + os.sep):
            return True
    # _internal 特征（PyInstaller 固定命名）
    if os.sep + '_internal' in ap:
        return True
    return False


def _clean_path_list(value, app_dirs):
    """
    清理一个冒号分隔的路径列表：去掉应用自带目录，保留系统目录。
    返回 (新值, 是否发生变化)
    """
    if not value:
        return value, False
    kept = []
    changed = False
    for part in value.split(os.pathsep):
        if not part:
            continue
        if _looks_like_app_path(part, app_dirs):
            changed = True
            continue
        kept.append(part)
    return os.pathsep.join(kept), changed


def clean_env(base=None):
    """
    返回一个「干净的」环境变量字典，用于启动**系统外部程序**。

    做的事：
      1. 剔除 LD_LIBRARY_PATH 中属于应用自身的条目（保留用户自己设的系统路径）
      2. LD_LIBRARY_PATH 清空后直接删除该变量，避免空值仍有语义
      3. 删除 LD_PRELOAD / LD_AUDIT（这些会强制注入库，风险最高）
      4. 删除 PyInstaller 相关变量（PYTHONPATH/PYTHONHOME 等会干扰
         Python 版 WPS 包装脚本或 unoconv）
      5. 设置 PYTHONNOUSERSITE=1，避免用户 site-packages 干扰子进程

    用法：
        subprocess.run(cmd, env=clean_env())
    """
    env = dict(os.environ if base is None else base)
    app_dirs = _app_dirs()

    for var in _LIB_PATH_VARS:
        if var not in env:
            continue
        if var in ('LD_PRELOAD', 'LD_AUDIT',
                   'DYLD_INSERT_LIBRARIES', 'DYLD_LIBRARY_PATH'):
            # 强制注入类：直接删除，不保留任何条目
            env.pop(var, None)
            continue
        newv, changed = _clean_path_list(env.get(var), app_dirs)
        if changed:
            if newv:
                env[var] = newv
            else:
                env.pop(var, None)

    for var in _PYINSTALLER_VARS:
        env.pop(var, None)

    env['PYTHONNOUSERSITE'] = '1'
    return env


def sanitize_library_path():
    """
    清理**当前进程**的 LD_LIBRARY_PATH，移除应用自带目录。

    在程序启动早期（main 入口）调用。PyInstaller 会用它来找自带库，
    但 PyInstaller 自己是通过 sys._MEIPASS 与 rpath 找库的，
    清掉 LD_LIBRARY_PATH 不会影响本进程加载自带扩展模块 ——
    而且能防止后续 dlopen（如 Pango/Cairo）误用旧库。

    返回 (原值, 新值) 便于日志。
    """
    old = os.environ.get('LD_LIBRARY_PATH')
    if not old:
        return old, old
    app_dirs = _app_dirs()
    newv, changed = _clean_path_list(old, app_dirs)
    if not changed:
        return old, old
    if newv:
        os.environ['LD_LIBRARY_PATH'] = newv
    else:
        os.environ.pop('LD_LIBRARY_PATH', None)
    return old, newv


def describe():
    """一行状态，供启动日志与诊断脚本使用。"""
    app_dirs = _app_dirs()
    raw = os.environ.get('LD_LIBRARY_PATH') or ''
    dirty = []
    for part in raw.split(os.pathsep):
        if part and _looks_like_app_path(part, app_dirs):
            dirty.append(part)
    if dirty:
        return ("LD_LIBRARY_PATH 含应用自带路径 %d 条（调用外部程序时会自动剔除）"
                % len(dirty))
    return "LD_LIBRARY_PATH 干净（不含应用自带路径）"


def bundled_lib_report(extra_dirs=None):
    """
    检查 _internal 里是否带了会与系统冲突的库，返回列表。

    用于打包后自检：如果发现带了 libfreetype / libfontconfig 等，
    说明打包配置需要加排除规则（见 complaints-linux.spec 的
    _filter_binaries / _SYSTEM_LIB_PREFIXES）。

    extra_dirs：额外要扫描的目录（诊断脚本、单元测试用）。
    """
    risky_prefixes = (
        'libfreetype', 'libfontconfig', 'libcairo', 'libpango',
        'libharfbuzz', 'libpng', 'libjpeg', 'libz.',
    )
    scan_dirs = set(_app_dirs())
    for d in (extra_dirs or ()):
        if d:
            try:
                scan_dirs.add(os.path.abspath(d))
            except Exception:
                pass
    found = []
    for d in scan_dirs:
        if not os.path.isdir(d):
            continue
        try:
            names = os.listdir(d)
        except OSError:
            continue
        for name in names:
            low = name.lower()
            if not low.endswith(('.so', '.so.6', '.so.1')) and '.so.' not in low:
                continue
            for p in risky_prefixes:
                if low.startswith(p):
                    found.append(os.path.join(d, name))
                    break
    return sorted(set(found))
