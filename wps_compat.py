#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
WPS Office 调用层（统信 UOS / 麒麟 Kylin V10 专用）
========================================================================
国产桌面基本都预装 WPS。老二进制 .doc（OLE2）纯 Python 读不了，
WordML / OOXML 我们已经能直接解析，但遇到真正的老 .doc 时必须借助
WPS（或 LibreOffice）转换成 .docx 才能提取表格。

本模块解决现有实现的三处硬伤：
  1. 原代码把 WPS 当 LibreOffice 用，只试 `wps --headless --convert-to pdf`；
     WPS Linux 多数版本不吃这套参数，且可能弹授权/首页窗口。
  2. `doc_to_docx()` 只走 LibreOffice，WPS 机器上没有 LO 就直接失败。
  3. 没有路径探测、没有超时、没有 HOME 隔离、没有 DISPLAY 处理，
     批量处理时一个文件卡住，整批就停。

调用策略（每个候选都是「启动 -> 等产物 -> 超时杀掉」）：
  - 结果驱动：不靠返回码，只看目标文件是否真的生成且非空
  - 多参数候选：不同 WPS 版本 / 打包方式参数不同，逐一试
  - 环境隔离：临时 HOME + 临时 XDG 目录，避免污染用户 WPS 配置
  - 允许用环境变量 REA_WPS_COMMAND / WPS_COMMAND 指定可执行文件

对外主函数：
  find_wps()                  -> WPS 可执行文件路径或 None
  wps_available()             -> bool
  convert_with_wps(src, fmt)  -> 转换后的文件路径或 None
  extract_table_with_wps(p)   -> 字典（转换后读表格）或 None
  diagnose()                  -> 一行诊断文本
"""

import os
import shutil
import subprocess
import sys
import tempfile
import glob

# ---------------------------------------------------------------------------
# 可执行文件探测
# ---------------------------------------------------------------------------

_WPS_NAMES = ("wps", "et", "wpp", "wpspdf", "kingsoft-wps")

_WPS_PATHS = (
    # 官方 Linux 版
    "/opt/kingsoft/wps-office/office6/wps",
    "/opt/kingsoft/wps-office/office6/et",
    "/opt/kingsoft/wps-office/office6/wpp",
    "/usr/bin/wps",
    "/usr/bin/et",
    "/usr/bin/wpp",
    "/usr/local/bin/wps",
    # 统信/麒麟商店版、平铺目录
    "/opt/apps/com.kingsoft.wps-office/files/bin/wps",
    "/opt/apps/cn.kingsoft.wps-office/files/bin/wps",
    "/opt/wps-office/office6/wps",
    "/usr/lib/wps-office/office6/wps",
    "/usr/lib/wps/office6/wps",
    # 龙芯 / 申威 等移植版
    "/opt/kingsoft/wps-office/office6/wps_loongarch64",
    "/opt/kingsoft/wps-office/office6/wps_mips64el",
)

_CACHED_WPS = None
_PROBED = False


def _candidate_commands():
    """按优先级产出候选 WPS 命令路径。"""
    seen = set()

    def add(p):
        if p and p not in seen:
            seen.add(p)
            return True
        return False

    # 1) 显式指定（现场运维可以直接写绝对路径，也支持 .py 包装脚本）
    for var in ("REA_WPS_COMMAND", "WPS_COMMAND", "WPS_PATH"):
        val = os.environ.get(var)
        if val and os.path.isfile(val) and (os.access(val, os.X_OK) or val.lower().endswith('.py')) and add(val):
            yield val

    # 2) PATH
    for name in _WPS_NAMES:
        p = shutil.which(name)
        if p and add(p):
            yield p

    # 3) 常见安装路径
    for p in _WPS_PATHS:
        if os.path.isfile(p) and os.access(p, os.X_OK) and add(p):
            yield p

    # 4) 通配兜底（一些发行版把版本号写进目录）
    for pattern in ("/opt/kingsoft/wps-office*/office6/wps",
                    "/opt/apps/*wps*/files/bin/wps",
                    "/usr/lib/wps*/office6/wps"):
        for p in sorted(glob.glob(pattern)):
            if os.path.isfile(p) and os.access(p, os.X_OK) and add(p):
                yield p


def find_wps():
    """返回 WPS 可执行文件路径；未安装返回 None（结果缓存，避免每文件探测）。"""
    global _CACHED_WPS, _PROBED
    if _PROBED:
        return _CACHED_WPS
    _PROBED = True
    for p in _candidate_commands():
        _CACHED_WPS = p
        return p
    return None


def wps_available():
    return find_wps() is not None


def reset_cache():
    """测试或用户新装 WPS 后调用，重新探测。"""
    global _CACHED_WPS, _PROBED
    _CACHED_WPS = None
    _PROBED = False


# ---------------------------------------------------------------------------
# 转换
# ---------------------------------------------------------------------------

# 不同 WPS 版本接受的命令行差异很大，按「成功率从高到低」排列，逐个试。
# {exe} / {src} / {out} / {fmt} 为占位符。
_ARG_TEMPLATES = (
    # 新版 WPS Linux（兼容 LibreOffice 风格）
    ["{exe}", "--headless", "--norestore", "--invisible",
     "--convert-to", "{fmt}", "--outdir", "{out}", "{src}"],
    # 带显式输出文件
    ["{exe}", "--headless", "--convert-to", "{fmt}", "-o", "{outfile}", "{src}"],
    # WPS 传统 -f 打开后转换
    ["{exe}", "-f", "{src}", "--convert-to", "{fmt}", "--outdir", "{out}"],
    # 部分政企定制版用 -convert-to（单横线）
    ["{exe}", "--headless", "-convert-to", "{fmt}", "--outdir", "{out}", "{src}"],
    # 只带 --convert-to 的极简形式
    ["{exe}", "--convert-to", "{fmt}", "--outdir", "{out}", "{src}"],
)

# WPS 对格式名的写法也不统一
_FMT_ALIASES = {
    "docx": ("docx:MS Word 2007 XML", "docx"),
    "pdf": ("pdf:writer_pdf_Export", "pdf"),
    "html": ("html", "htm"),
    "txt": ("txt:Text (encoded):UTF8", "txt"),
}


def _env_for_wps(home_dir):
    """
    构造隔离环境：WPS 首次启动会写 ~/.config/Kingsoft、弹用户协议/首页，
    用临时 HOME 可以避免污染用户配置并绕开部分首次运行弹窗。
    """
    env = os.environ.copy()
    env["HOME"] = home_dir
    env["XDG_CONFIG_HOME"] = os.path.join(home_dir, ".config")
    env["XDG_CACHE_HOME"] = os.path.join(home_dir, ".cache")
    env["XDG_DATA_HOME"] = os.path.join(home_dir, ".local", "share")
    env["WPS_NO_QT_SESSION_MANAGER"] = "1"
    env["QT_QPA_PLATFORM"] = "offscreen"       # 无 X 显示也能启动 Qt
    env["SAL_USE_VCLPLUGIN"] = "gen"
    env.pop("DISPLAY", None)                   # 避免连到真实桌面弹窗
    env.pop("WAYLAND_DISPLAY", None)
    for k in ("WPS_AUTO_SHOW", "WPS_SHOW_FIRST_PAGE"):
        env[k] = "0"
    return env


def _run(cmd, env, timeout):
    """启动子进程，超时杀掉；返回 CompletedProcess 或 None。"""
    # 允许把 Python 脚本当作转换器（现场自定义 WPS 包装 / 测试替身），
    # 否则 Windows 上 .py 不能直接 CreateProcess。
    try:
        if cmd and str(cmd[0]).lower().endswith('.py'):
            cmd = [sys.executable] + list(cmd)
    except Exception:
        pass
    try:
        return subprocess.run(
            cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE,
            timeout=timeout, env=env,
            stdin=subprocess.DEVNULL,
            start_new_session=True,
        )
    except subprocess.TimeoutExpired:
        return None
    except Exception:
        return None


def convert_with_wps(src, fmt="docx", out_dir=None, timeout=45):
    """
    用 WPS 把文档转换成 fmt（docx / pdf / html / txt）。

    成功返回转换产物路径，失败返回 None。结果驱动：不看返回码，
    只看目标文件是否真的出现且非空。自动清理临时目录（保留产物）。
    """
    src = os.path.abspath(src)
    if not os.path.isfile(src):
        return None
    exe = find_wps()
    if not exe:
        return None

    fmt = fmt.lower()
    if fmt not in _FMT_ALIASES:
        return None

    keep_dir = out_dir is not None
    work_dir = os.path.abspath(out_dir) if keep_dir else tempfile.mkdtemp(prefix="wpsconv_")
    home_dir = tempfile.mkdtemp(prefix="wpshome_")
    os.makedirs(work_dir, exist_ok=True)

    base = os.path.splitext(os.path.basename(src))[0]
    result_path = None

    try:
        env = _env_for_wps(home_dir)
        for fmt_name in _FMT_ALIASES[fmt]:
            outfile = os.path.join(work_dir, base + "." + fmt.split(":")[0])
            for tpl in _ARG_TEMPLATES:
                # 关键：清掉 work_dir 里同名扩展名的所有旧产物。
                # 调用方常复用同一个临时目录，若不清，上一轮成功的文件会被
                # 本轮（超时/失败）误判为「转换成功」，导致读到过期数据。
                ext = fmt.split(":")[0].lower()
                for stale in glob.glob(os.path.join(work_dir, "*." + ext)):
                    if os.path.abspath(stale) == os.path.abspath(src):
                        continue
                    try:
                        os.remove(stale)
                    except OSError:
                        pass

                cmd = [p.format(exe=exe, src=src, out=work_dir,
                                outfile=outfile, fmt=fmt_name) for p in tpl]
                _run(cmd, env, timeout)

                # WPS 有时输出成 .DOCX / 带其它后缀，扫描目录兜底
                produced = None
                if os.path.isfile(outfile) and os.path.getsize(outfile) > 0:
                    produced = outfile
                else:
                    ext = fmt.split(":")[0].lower()
                    for p in glob.glob(os.path.join(work_dir, "*")):
                        if (p.lower().endswith("." + ext)
                                and os.path.isfile(p) and os.path.getsize(p) > 0
                                and os.path.abspath(p) != src):
                            produced = p
                            break
                if produced:
                    result_path = produced
                    return result_path
    finally:
        shutil.rmtree(home_dir, ignore_errors=True)
        if not keep_dir:
            # 产物要留下：移到系统临时目录下的稳定位置
            if result_path and os.path.isfile(result_path):
                stable = os.path.join(tempfile.gettempdir(),
                                      "wps_" + os.path.basename(result_path))
                try:
                    shutil.copy2(result_path, stable)
                    result_path = stable
                except OSError:
                    pass
            shutil.rmtree(work_dir, ignore_errors=True)

    return None


def doc_to_docx_with_wps(doc_path, out_dir=None):
    """老 .doc -> .docx 的 WPS 专用入口。"""
    return convert_with_wps(doc_path, "docx", out_dir=out_dir)


def word_to_pdf_with_wps(word_path, out_dir=None):
    """Word -> PDF 的 WPS 专用入口。"""
    return convert_with_wps(word_path, "pdf", out_dir=out_dir)


# ---------------------------------------------------------------------------
# 用 WPS 转换后再提取表格（这才是「调用 WPS 提取数据」的完整答案）
# ---------------------------------------------------------------------------

def extract_table_with_wps(path, reader=None, out_dir=None):
    """
    用 WPS 把任意 Word（尤其老二进制 .doc）转成 .docx，再交给 reader 读表格。

    reader: 可调用对象，接收 docx 路径并返回 dict；默认用 python-docx
            读取第一张表的全部单元格文本（调用方再按标签映射字段）。
    返回 dict 或 None。
    """
    converted = doc_to_docx_with_wps(path, out_dir=out_dir)
    if not converted:
        return None
    try:
        if reader is not None:
            return reader(converted)

        from docx import Document
        doc = Document(converted)
        if not doc.tables:
            return None
        rows = []
        for row in doc.tables[0].rows:
            rows.append([c.text.replace('\r', '').replace('\x07', '').strip()
                         for c in row.cells])
        return {'_rows': rows}
    except Exception:
        return None
    finally:
        # 转换产物放在临时目录时用完即删
        if out_dir is None and converted and os.path.isfile(converted):
            try:
                os.remove(converted)
            except OSError:
                pass


def diagnose():
    """给 诊断.sh / 启动日志用的一行说明。"""
    exe = find_wps()
    if not exe:
        return "WPS: 未检测到（老 .doc 将改用 LibreOffice；纯 WordML/OOXML 不依赖 WPS）"
    return "WPS: " + exe
