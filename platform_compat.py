#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
跨平台兼容层 —— 让「投诉批办单管理系统」在 Windows / 统信 UOS / 麒麟 Kylin V10 上
保持同一套 GUI 界面与功能。

Windows 分支保留原有行为（Word COM、ShellExecute 打印、docx2pdf、微软雅黑/仿宋），
Linux 分支提供等价实现（LibreOffice/WPS 转换、CUPS lp 打印、xdg-open、中文字体映射）。
本模块不依赖 tkinter，可在命令行模式下安全导入。
"""

import os
import sys
import shutil
import subprocess
import tempfile

IS_WINDOWS = sys.platform == "win32"
IS_MACOS = sys.platform == "darwin"
IS_LINUX = sys.platform.startswith("linux")


# ---------------------------------------------------------------------------
# 路径：区分「只读资源」与「可写用户数据」
# ---------------------------------------------------------------------------

def get_app_dir():
    """可执行文件所在目录（打包后）/ 源码目录（开发环境）。"""
    if getattr(sys, "frozen", False):
        return os.path.dirname(os.path.abspath(sys.executable))
    return os.path.dirname(os.path.abspath(__file__))


def get_resource_path(relative_path=""):
    """
    只读资源路径。PyInstaller 打包后指向 sys._MEIPASS（onedir 的 _internal 目录），
    开发环境指向源码目录。模板 .docx、默认规则 JSON 等均从这里读取。
    """
    if getattr(sys, "frozen", False):
        base = getattr(sys, "_MEIPASS", get_app_dir())
    else:
        base = get_app_dir()
    return os.path.join(base, relative_path) if relative_path else base


def get_user_data_path(relative_path=""):
    """可写用户数据路径，位于可执行文件同级目录。"""
    base = get_app_dir()
    return os.path.join(base, relative_path) if relative_path else base


def get_appdata_path(filename=""):
    """可写的 AppData 目录（自动创建）。所有配置、数据库、日志都落在这里。"""
    d = os.path.join(get_app_dir(), "AppData")
    os.makedirs(d, exist_ok=True)
    return os.path.join(d, filename) if filename else d


def ensure_app_cwd():
    """
    把进程工作目录切到应用目录。这一步让源码里全部相对路径
    （AppData、待打印、已打印、已处理、模板、配置）在双击启动时依旧指向正确位置，
    与 Windows 快捷方式「起始位置」设为程序目录的行为完全一致。
    """
    d = get_app_dir()
    try:
        os.chdir(d)
    except OSError:
        pass
    return d


# 打包时这些文件被放进 _internal/AppData，首次运行复制到可写的 exe 同级 AppData
_SEED_FILES = (
    "filling_rules_default.json",
    "filling_rules.json",
    "path_config.json",
    "quick_buttons_config.json",
    "qrcode_config.json",
    "input_history.json",
    "counter.json",
    "config.json",
    "window_config.json",
    "last_database.json",
)

_SEED_FROM_ROOT = ("print_delay_config.json",)

# 中文路径键 -> 打包时放在 _internal 下的目录名
_DATA_DIRS = ("已处理", "待打印", "已打印")


def seed_user_data():
    """首次运行时把默认配置/数据库从只读资源目录复制到可写用户目录（不覆盖已有数据）。"""
    src_dir = get_resource_path("AppData")
    dst_dir = get_appdata_path()
    same = os.path.abspath(src_dir) == os.path.abspath(dst_dir)

    if not same and os.path.isdir(src_dir):
        for name in _SEED_FILES:
            src = os.path.join(src_dir, name)
            dst = os.path.join(dst_dir, name)
            if os.path.isfile(src) and not os.path.exists(dst):
                try:
                    shutil.copy2(src, dst)
                except OSError:
                    pass
        # 种子数据库：与 Windows 版交付内容一致，首次运行即带历史库
        src_db = os.path.join(src_dir, "complaints.db")
        dst_db = os.path.join(dst_dir, "complaints.db")
        if os.path.isfile(src_db) and not os.path.exists(dst_db):
            try:
                shutil.copy2(src_db, dst_db)
            except OSError:
                pass

    # 根目录级配置（打印延迟等）
    for name in _SEED_FROM_ROOT:
        src = get_resource_path(name)
        dst = get_user_data_path(name)
        if os.path.abspath(src) != os.path.abspath(dst) and os.path.isfile(src) and not os.path.exists(dst):
            try:
                shutil.copy2(src, dst)
            except OSError:
                pass

    # 保证三个工作目录在首次启动就存在，行为与 Windows 版一致
    for d in _DATA_DIRS:
        try:
            os.makedirs(get_user_data_path(d), exist_ok=True)
        except OSError:
            pass


# ---------------------------------------------------------------------------
# 字体：Windows 字体名 -> 国产 Linux 发行版预装的中文字体
# ---------------------------------------------------------------------------

_FONT_CANDIDATES = {
    "微软雅黑": ["Noto Sans CJK SC", "Source Han Sans SC", "思源黑体",
                 "WenQuanYi Micro Hei", "文泉驿微米黑", "DejaVu Sans"],
    "Microsoft YaHei": ["Noto Sans CJK SC", "Source Han Sans SC", "思源黑体",
                        "WenQuanYi Micro Hei", "DejaVu Sans"],
    "仿宋_GB2312": ["FangSong", "仿宋", "Noto Serif CJK SC", "AR PL UMing CN",
                    "DejaVu Serif"],
    "仿宋": ["FangSong", "仿宋", "Noto Serif CJK SC", "DejaVu Serif"],
    "宋体": ["SimSun", "宋体", "AR PL UMing CN", "AR PL SungtiL GB",
             "WenQuanYi Zen Hei", "Noto Serif CJK SC", "DejaVu Serif"],
    "黑体": ["SimHei", "黑体", "WenQuanYi Zen Hei", "WenQuanYi Micro Hei",
             "Noto Sans CJK SC", "DejaVu Sans"],
    "Consolas": ["DejaVu Sans Mono", "Liberation Mono", "Noto Sans Mono CJK SC"],
}

_FONT_CACHE = {}


def _installed_families():
    """
    一次性读取 fontconfig 的字体家族列表；失败时返回空串（交给 Tk 自行回退）。

    fc-list 也必须用干净环境：打包版若让 fc-list 加载到 _internal 里的旧
    libfontconfig，会解析不出系统字体，导致中文字体判断错误、版式偏移。
    """
    try:
        import print_env
        env = print_env.clean_env()
    except Exception:
        env = None
    try:
        r = subprocess.run(["fc-list", ":", "family"], env=env,
                           stdout=subprocess.PIPE, stderr=subprocess.DEVNULL, timeout=5)
        return (r.stdout or b"").decode("utf-8", "ignore")
    except Exception:
        return ""


def ui_font(family):
    """
    返回当前平台可用的字体家族名。Windows 原样返回，界面与旧版完全一致；
    统信/麒麟上映射到系统已装的中文字体，未装时 Tk 仍按默认字体回退。
    """
    if IS_WINDOWS or family in _FONT_CACHE:
        return family if IS_WINDOWS else _FONT_CACHE[family]
    fams = ""
    try:
        fams = _installed_families()
    except Exception:
        pass
    candidates = _FONT_CANDIDATES.get(family, [])
    if fams:
        fam_lines = [x.strip() for x in fams.splitlines() if x.strip()]
        for c in candidates:
            # Windows 专有字体名在 Linux 上不可能装到，跳过精确匹配
            if c in ("SimSun", "宋体", "SimHei", "黑体", "FangSong", "仿宋",
                     "Microsoft YaHei", "微软雅黑"):
                continue
            if any(c.lower() == line.lower() for line in fam_lines):
                _FONT_CACHE[family] = c
                return c
        # 宽松匹配：fontconfig 常输出 "Noto Sans CJK SC:style=Regular"
        for c in candidates:
            if c in ("SimSun", "宋体", "SimHei", "黑体", "FangSong", "仿宋",
                     "Microsoft YaHei", "微软雅黑"):
                continue
            if any(c.lower() in line.lower() for line in fam_lines):
                _FONT_CACHE[family] = c
                return c
    # fc-list 不可用或没有候选字体时，选一个 Linux 上真实存在的通用家族
    if family == "Consolas":
        chosen = "DejaVu Sans Mono"
    elif family in ("微软雅黑", "Microsoft YaHei", "黑体"):
        chosen = "Noto Sans CJK SC"
    else:
        chosen = "Noto Serif CJK SC"
    _FONT_CACHE[family] = chosen
    return chosen


# ---------------------------------------------------------------------------
# 打开 / 打印 / 文档转换
# ---------------------------------------------------------------------------

def open_path(path):
    """
    用系统默认程序打开文件或目录，等价 Windows 的 os.startfile。

    xdg-open 同样要用干净环境：它最终调用桌面环境的程序（WPS、文档查看器），
    若继承 _internal 的库路径，那些程序也可能起不来或字体异常。
    """
    path = os.path.abspath(path)
    if IS_WINDOWS:
        os.startfile(path)  # noqa: F821  (仅 Windows 存在)
        return True
    try:
        import print_env
        env = print_env.clean_env()
    except Exception:
        env = None
    if IS_MACOS:
        subprocess.Popen(["open", path], stdout=subprocess.DEVNULL,
                         stderr=subprocess.DEVNULL, start_new_session=True, env=env)
        return True
    for cmd in ("xdg-open", "gio", "gnome-open", "kde-open"):
        exe = shutil.which(cmd)
        if not exe:
            continue
        argv = [exe, "open", path] if cmd == "gio" else [exe, path]
        try:
            subprocess.Popen(argv, stdout=subprocess.DEVNULL,
                             stderr=subprocess.DEVNULL, start_new_session=True,
                             env=env)
            return True
        except OSError:
            continue
    return False


def _find_soffice():
    """定位 LibreOffice / WPS 自带的可执行程序。"""
    for name in ("soffice", "libreoffice"):
        p = shutil.which(name)
        if p:
            return p
    for p in (
        "/usr/bin/soffice", "/usr/bin/libreoffice",
        "/usr/lib/libreoffice/program/soffice",
        "/opt/libreoffice/program/soffice",
        "/opt/libreoffice24.8/program/soffice",
        "/opt/libreoffice7/program/soffice",
    ):
        if os.path.isfile(p) and os.access(p, os.X_OK):
            return p
    return None


def _find_wps():
    """定位 WPS；探测逻辑在 wps_compat 里统一维护。"""
    try:
        import wps_compat
        return wps_compat.find_wps()
    except Exception:
        return shutil.which("wps")


def _run_office(cmd, timeout=120):
    """
    运行外部程序（soffice / lp / unoconv 等）。

    关键：必须用 print_env.clean_env() 清除 LD_LIBRARY_PATH 里属于本应用
    _internal/ 的路径。否则打包版的子进程会加载到我们自带的旧
    libfreetype / libfontconfig，导致外部程序启动失败 —— 麒麟现场那次

        soffice.bin: symbol lookup error:
        /lib/aarch64-linux-gnu/libcairo.so.2: undefined symbol: FT_Done_MM_Var

    就是这个原因（应用自带的旧 freetype 顶掉了系统的），
    而不是 LibreOffice 本身损坏。
    """
    env = None
    try:
        import print_env
        env = print_env.clean_env()
    except Exception:
        env = None
    try:
        return subprocess.run(cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE,
                              timeout=timeout, env=env)
    except Exception:
        return None


def _safe_copy_for_conversion(src, work_dir, ext=None):
    """
    把源文件复制到 work_dir 并改成纯 ASCII 安全名，返回新路径。

    为什么必须这么做：麒麟生成的批办单文件名是「投诉批办单#C-26-113.docx」。
    LibreOffice 内部按 file:// URI 处理路径，`#` 是 URI 片段分隔符，
    `&` 是参数分隔符，空格和中文在部分版本/语言环境下也会出问题，
    结果就是「转换命令返回成功但产物名字对不上」，上层判定为失败。
    复制成 safe_<n>.docx 后这些坑全部消失。
    """
    ext = ext or os.path.splitext(src)[1].lower() or '.docx'
    safe = os.path.join(work_dir, 'src_convert' + ext)
    shutil.copy2(src, safe)
    return safe


def _pick_produced(work_dir, before, ext):
    """在 work_dir 里找出本次新产生的、扩展名匹配的非空文件。"""
    found = []
    try:
        for name in os.listdir(work_dir):
            fp = os.path.join(work_dir, name)
            if not os.path.isfile(fp):
                continue
            if os.path.abspath(fp) in before:
                continue
            if name.lower().endswith('.' + ext) and os.path.getsize(fp) > 0:
                found.append(fp)
    except OSError:
        pass
    found.sort(key=lambda x: os.path.getsize(x), reverse=True)
    return found[0] if found else None


_LAST_CONVERT_ERROR = ['']
_LAST_PAGE_WARNING = ['']
# LibreOffice 本机是否「自身损坏」（缺符号/缺库），一旦确认就不再重试
_LO_BROKEN = ['']

# LibreOffice 启动即失败的典型特征：动态库符号/版本不匹配。
# 麒麟现场实测：
#   /usr/lib/libreoffice/program/soffice.bin: symbol lookup error:
#   /lib/aarch64-linux-gnu/libcairo.so.2: undefined symbol: FT_Done_MM_Var
# 这是 freetype 与 cairo 版本不配套的系统级问题，不是本程序能修的，
# 但必须让用户看到可执行的解决方案，而不是反复失败。
_LO_BROKEN_SIGNS = (
    'symbol lookup error',
    'undefined symbol',
    'FT_Done_MM_Var',
    'error while loading shared libraries',
    'cannot open shared object file',
)


def lo_broken_reason():
    """LibreOffice 若被判定为自身损坏，返回原因；否则空串。"""
    return _LO_BROKEN[0]


def _looks_like_broken_lo(text):
    low = (text or '').lower()
    return any(s.lower() in low for s in _LO_BROKEN_SIGNS)


def last_convert_error():
    """最近一次转换失败的原因（给 GUI 提示用）。"""
    return _LAST_CONVERT_ERROR[0]


def page_warning():
    """最近一次转换的页数提示（可能为空）；用于提醒「一页变两页」。"""
    return _LAST_PAGE_WARNING[0]


def _try_soffice(src, work_dir, target_fmt, timeout=180):
    """
    用 LibreOffice 转换一次。返回产物路径或 None，失败原因写入 _LAST_CONVERT_ERROR。
    会在临时目录里做「安全名副本」，避免 # & 空格 中文 等字符导致的静默失败。
    """
    soffice = _find_soffice()
    if not soffice:
        _LAST_CONVERT_ERROR[0] = '未找到 LibreOffice（soffice）'
        return None

    # 已确认本机 LibreOffice 自身损坏：直接跳过，避免每次都等一次超时
    if _LO_BROKEN[0]:
        _LAST_CONVERT_ERROR[0] = (
            '本机 LibreOffice 自身损坏（缺少动态库符号），无法转换：\n'
            + _LO_BROKEN[0].splitlines()[0]
            + '\n修复：sudo apt install --reinstall libreoffice-writer '
              'libreoffice-core libfreetype6 libcairo2')
        return None

    ext = os.path.splitext(src)[1].lower().lstrip('.') or 'docx'
    stage = tempfile.mkdtemp(prefix='lo_stage_')
    out_dir = tempfile.mkdtemp(prefix='lo_out_')
    try:
        safe_src = _safe_copy_for_conversion(src, stage, '.' + ext)

        # 转 PDF 前去掉 docGrid（行网格）。
        # LibreOffice 把网格当「最小行高」叠加在段落行距上，比 Word 高出
        # 零点几磅；模板页边距紧，累积后表格后的「备注」被挤到第 2 页。
        # 新生成的文档已在 main.strip_docgrid 里处理过，这里再兜一次，
        # 是为了让「待打印」里遗留的旧文件（带 docGrid）也能一页打出来。
        # 只作用于临时副本，不改用户的原文件。
        if target_fmt == 'pdf' and ext == 'docx':
            try:
                import main as _main_mod
                _main_mod.strip_docgrid(safe_src)
            except Exception:
                pass

        before = set()
        # 转换串，pdf 与 docx 的 filter 不同
        if target_fmt == 'pdf':
            conv = 'pdf:writer_pdf_Export'
        elif target_fmt == 'docx':
            conv = 'docx:MS Word 2007 XML'
        else:
            conv = target_fmt

        profile = 'file://' + os.path.join(stage, 'lo_profile')

        # 字体度量兼容：模板用「仿宋_GB2312」等 Windows 字体，麒麟上没有，
        # LibreOffice 会替换成度量不同的字体，导致行高变化、一页变两页。
        # 转换前确保用户级 fontconfig 别名已就位（映射到度量最接近的本地字体）。
        try:
            import font_compat
            font_compat.ensure_font_aliases()
        except Exception:
            pass

        cmd = [soffice, '-env:UserInstallation=' + profile,
               '--headless', '--norestore', '--invisible', '--nodefault',
               '--nolockcheck', '--nofirststartwizard',
               '--convert-to', conv, '--outdir', out_dir, safe_src]
        r = _run_office(cmd, timeout=timeout)
        if r is None:
            _LAST_CONVERT_ERROR[0] = 'LibreOffice 调用超时或被中断（首次启动较慢，可重试）'
            return None

        err = (r.stderr or b'').decode('utf-8', 'ignore').strip()
        out = (r.stdout or b'').decode('utf-8', 'ignore').strip()

        produced = _pick_produced(out_dir, before, target_fmt)
        if produced:
            return produced

        # 没有产物：把真实原因带回去，别再静默
        detail = err or out or f'返回码 {r.returncode}'
        if _looks_like_broken_lo(err) or _looks_like_broken_lo(out):
            _LO_BROKEN[0] = (err or out).strip()[:400]
            _LAST_CONVERT_ERROR[0] = (
                '本机 LibreOffice 自身损坏（缺少动态库符号），无法转换：\n'
                + _LO_BROKEN[0].splitlines()[0]
                + '\n修复：sudo apt install --reinstall libreoffice-writer '
                  'libreoffice-core libfreetype6 libcairo2')
        else:
            _LAST_CONVERT_ERROR[0] = 'LibreOffice 转换未产生文件：' + detail[:300]
        return None
    finally:
        shutil.rmtree(stage, ignore_errors=True)


def _verify_page_count(word_path, pdf_path):
    """
    比对 Word 与 PDF 的页数，返回提示文本（可能为空）。

    现场反馈「Word 一页，转 PDF 变两页」：模板页边距很紧，字体替换后
    行高微增就把最后一行挤到第二页。这里做一次校验，把差异摆到台面上，
    而不是让用户打完才发现多一张纸。
    """
    try:
        import font_compat
        expect = font_compat.docx_page_estimate(word_path)
        actual = font_compat.pdf_page_count(pdf_path)
        if expect and actual and actual != expect:
            miss = font_compat.docx_missing_fonts(word_path)
            tip = ('缺少字体：' + '、'.join(miss)) if miss else '字体已替换'
            return ('页数不一致：Word %d 页 -> PDF %d 页（%s）。\n'
                    '如需完全一致，请安装模板字体（仿宋_GB2312 等），'
                    '或用 WPS 直接打开打印。' % (expect, actual, tip))
    except Exception:
        pass
    return ''


def word_to_pdf(word_path, out_pdf_path=None):
    """
    Word -> PDF。
      Windows: docx2pdf（原逻辑）
      Linux  : LibreOffice 优先（唯一可靠），WPS 兜底，unoconv 最后
    成功返回 PDF 路径，失败返回 None；失败原因见 last_convert_error()。
    页数变化会写入 _LAST_PAGE_WARNING（见 page_warning()）。
    """
    word_path = os.path.abspath(word_path)
    base = os.path.splitext(os.path.basename(word_path))[0]
    if out_pdf_path is None:
        out_pdf_path = os.path.join(os.path.dirname(word_path), base + '.pdf')
    out_pdf_path = os.path.abspath(out_pdf_path)
    os.makedirs(os.path.dirname(out_pdf_path), exist_ok=True)
    _LAST_CONVERT_ERROR[0] = ''
    _LAST_PAGE_WARNING[0] = ''

    if IS_WINDOWS:
        try:
            from docx2pdf import convert
            convert(word_path, out_pdf_path)
            if os.path.isfile(out_pdf_path) and os.path.getsize(out_pdf_path) > 0:
                return out_pdf_path
            _LAST_CONVERT_ERROR[0] = 'docx2pdf 未生成 PDF（请确认已安装 Word）'
        except Exception as e:
            _LAST_CONVERT_ERROR[0] = 'docx2pdf 失败: %s' % e

    # 1) LibreOffice（安全名副本 + 产物扫描）
    pdf = _try_soffice(word_path, None, 'pdf')
    if pdf:
        if os.path.exists(out_pdf_path):
            try:
                os.remove(out_pdf_path)
            except OSError:
                pass
        try:
            shutil.move(pdf, out_pdf_path)
            _LAST_PAGE_WARNING[0] = _verify_page_count(word_path, out_pdf_path)
            return out_pdf_path
        except OSError:
            pass

    # 2) WPS 兜底（内部有试转探测，不支持时立即返回 None，不弹窗）
    if _find_wps():
        try:
            import wps_compat
            tmp_dir = tempfile.mkdtemp(prefix='wpspdf_')
            try:
                wps_out = wps_compat.word_to_pdf_with_wps(word_path, out_dir=tmp_dir)
                if wps_out and os.path.isfile(wps_out) and os.path.getsize(wps_out) > 0:
                    if os.path.exists(out_pdf_path):
                        try:
                            os.remove(out_pdf_path)
                        except OSError:
                            pass
                    shutil.move(wps_out, out_pdf_path)
                    _LAST_PAGE_WARNING[0] = _verify_page_count(word_path, out_pdf_path)
                    return out_pdf_path
            finally:
                shutil.rmtree(tmp_dir, ignore_errors=True)
        except Exception as e:
            _LAST_CONVERT_ERROR[0] = (_LAST_CONVERT_ERROR[0] + ' / WPS: ' + str(e)).strip(' /')

    # 3) unoconv 最后
    uno = shutil.which('unoconv')
    if uno:
        tmp_dir = tempfile.mkdtemp(prefix='unoconv_')
        try:
            _run_office([uno, '-f', 'pdf', '-o',
                         os.path.join(tmp_dir, base + '.pdf'), word_path],
                        timeout=180)
            produced = os.path.join(tmp_dir, base + '.pdf')
            if os.path.isfile(produced) and os.path.getsize(produced) > 0:
                if os.path.exists(out_pdf_path):
                    try:
                        os.remove(out_pdf_path)
                    except OSError:
                        pass
                shutil.move(produced, out_pdf_path)
                _LAST_PAGE_WARNING[0] = _verify_page_count(word_path, out_pdf_path)
                return out_pdf_path
        finally:
            shutil.rmtree(tmp_dir, ignore_errors=True)

    return None


def doc_to_docx(doc_path, out_dir):
    """
    旧版 .doc -> .docx，用于提取表格。
    Windows 用 Word COM；Linux 用 LibreOffice 优先、WPS 兜底。
    成功返回新文件路径，失败返回 None。
    """
    doc_path = os.path.abspath(doc_path)
    base = os.path.splitext(os.path.basename(doc_path))[0]
    os.makedirs(out_dir, exist_ok=True)
    out_path = os.path.join(out_dir, base + '.docx')

    if IS_WINDOWS:
        try:
            import win32com.client
            word = win32com.client.Dispatch('Word.Application')
            word.Visible = False
            word.DisplayAlerts = 0
            try:
                d = word.Documents.Open(doc_path)
                d.SaveAs(out_path, FileFormat=16)  # wdFormatXMLDocument
                d.Close(False)
            finally:
                word.Quit()
            if os.path.isfile(out_path) and os.path.getsize(out_path) > 0:
                return out_path
        except Exception:
            pass

    # LibreOffice 优先（安全名副本 + 产物扫描）
    produced = _try_soffice(doc_path, None, 'docx')
    if produced:
        try:
            if os.path.exists(out_path):
                os.remove(out_path)
            shutil.move(produced, out_path)
            return out_path
        except OSError:
            pass

    # WPS 兜底
    if os.environ.get('COMPLAINTS_WPS_PREFERENCE', '').strip().lower() != 'off' and _find_wps():
        try:
            import wps_compat
            converted = wps_compat.doc_to_docx_with_wps(doc_path, out_dir=out_dir)
            if converted and os.path.isfile(converted) and os.path.getsize(converted) > 0:
                return converted
        except Exception:
            pass
    return None


def print_file_direct_word(path, printer=None):
    """
    直接打印 Word（不生成 PDF），供「不想因转换而变版式」的场景使用。

    为什么能直接打：CUPS 的 lp 只认 PDF/PS/图片，但 LibreOffice 自带
    `--print` 参数，会自己调用系统打印子系统把 docx 送进队列，中间不落
    PDF 文件 —— 排版由 LibreOffice 直接输出到打印机驱动，不会因为
    「转 PDF 再打」多一道字体替换。

    注意：LibreOffice 打印时仍然用本机的字体替换，所以「字体缺失导致
    行高变化」的问题依然存在；它解决的是「多一道 PDF 转换」这层风险，
    而不是字体本身。要彻底一致仍需安装模板字体。

    返回 (是否成功, 说明文字)。
    """
    path = os.path.abspath(path)
    soffice = _find_soffice()
    if not soffice:
        return False, "未找到 LibreOffice，无法直接打印 Word"

    # LibreOffice 自身损坏（如 FT_Done_MM_Var 缺符号）时它连启动都做不到
    if _LO_BROKEN[0]:
        return False, ('本机 LibreOffice 自身损坏，无法直接打印：\n'
                       + _LO_BROKEN[0].splitlines()[0])

    tmp = tempfile.mkdtemp(prefix='loprint_')
    try:
        # 同样先复制成安全名，规避 # & 空格 中文 的 URI 问题
        ext = os.path.splitext(path)[1].lower() or '.docx'
        safe = _safe_copy_for_conversion(path, tmp, ext)
        profile = 'file://' + os.path.join(tmp, 'lo_profile')
        cmd = [soffice, '-env:UserInstallation=' + profile,
               '--headless', '--norestore', '--invisible', '--nodefault',
               '--nolockcheck', '--nofirststartwizard']
        if printer:
            cmd += ['--print-to-printer', printer]
        cmd += ['--print', safe]
        r = _run_office(cmd, timeout=180)
        if r is None:
            return False, "LibreOffice 直接打印超时或被中断"
        err = (r.stderr or b'').decode('utf-8', 'ignore').strip()
        if r.returncode != 0:
            return False, err or ('LibreOffice 打印返回码 %d' % r.returncode)
        return True, "已通过 LibreOffice 直接发送到打印队列"
    except Exception as e:
        return False, "直接打印失败：%s" % e
    finally:
        shutil.rmtree(tmp, ignore_errors=True)


def print_file(path):
    """
    打印单个文档。

    Windows: ShellExecute('print') —— 与旧版一致。
    Linux  : 两种通路，可用 COMPLAINTS_PRINT_MODE 选择：
        pdf（默认）  先转 PDF 再 lp。兼容性最好（任何打印机驱动都认 PDF），
                     但会因字体替换出现「一页变两页」。
        direct       用 LibreOffice 的 --print 直接送打印队列，不落 PDF，
                     少一道转换；排版仍受本机字体影响。
        word         把 .docx 交给 WPS 的图形界面打开（人工确认后打印）。

    PDF 与 direct 都失败时不再「静默把 docx 喂给 lp」
    （CUPS 会报「不支持的 document-format」）。

    返回 (是否成功, 说明文字)。
    """
    path = os.path.abspath(path)
    if IS_WINDOWS:
        try:
            import win32api
            win32api.ShellExecute(0, "print", path, None, ".", 0)
            return True, "已发送到 Windows 打印队列"
        except Exception as e:
            return False, str(e)

    mode = (os.environ.get('COMPLAINTS_PRINT_MODE') or 'pdf').strip().lower()
    is_word = path.lower().endswith((".doc", ".docx", ".wps", ".rtf", ".odt"))

    # --- direct 模式：直接打印 Word，不生成 PDF ---
    if is_word and mode in ('direct', 'word'):
        if mode == 'word':
            # 交给 WPS 图形界面，由人确认后打印（最保真，但需人工点一下）
            try:
                import wps_compat
                exe = wps_compat.find_wps()
                if exe:
                    try:
                        import print_env
                        _env = print_env.clean_env()
                    except Exception:
                        _env = None
                    subprocess.Popen([exe, path], env=_env,
                                     stdout=subprocess.DEVNULL,
                                     stderr=subprocess.DEVNULL)
                    return True, "已在 WPS 中打开，请用 WPS 的打印按钮打印"
            except Exception:
                pass
        ok, info = print_file_direct_word(path)
        if ok:
            return True, info
        # 直接打印失败，继续走 PDF 通路（不直接放弃）

    tmp_dir = None
    target = path
    try:
        if is_word:
            tmp_dir = tempfile.mkdtemp(prefix="printconv_")
            base = os.path.splitext(os.path.basename(path))[0]
            pdf = word_to_pdf(path, os.path.join(tmp_dir, base + ".pdf"))
            if not pdf or not os.path.isfile(pdf) or os.path.getsize(pdf) <= 0:
                # 关键：转换失败必须直接报错，不能拿 docx 去喂 lp。
                # 把 LibreOffice 的真实失败原因带出来，避免「装了 LO 却不知道为啥失败」。
                reason = last_convert_error() or "未得到具体原因"
                soffice = _find_soffice()
                broken = lo_broken_reason()
                if broken:
                    # LibreOffice 自身坏了：给出可直接执行的修复命令，
                    # 并明确告知「现在唯一能保证版式的是 WPS」。
                    return (False,
                            "打印已中止：本机 LibreOffice 自身损坏，无法转换。\n\n"
                            f"报错：{broken.splitlines()[0]}\n\n"
                            "修复 LibreOffice（选一个，需管理员）：\n"
                            "  sudo apt install --reinstall libreoffice-writer "
                            "libreoffice-core libfreetype6 libcairo2\n"
                            "  # 若源不可用，直接卸载后重装：\n"
                            "  sudo apt remove --purge libreoffice-* && "
                            "sudo apt install libreoffice-writer\n\n"
                            "不改系统也能打印（版式与 Word 一致）：\n"
                            "  · 本程序「打印」菜单 → 「在 WPS 中打开当前文件并手动打印」\n"
                            "  · 或直接双击用 WPS 打开，点 WPS 的打印按钮")
                hint = ("检测到 LibreOffice: " + soffice) if soffice else \
                       ("未检测到 LibreOffice，请安装：\n"
                        "  sudo apt install -y libreoffice-writer")
                return (False,
                        "无法把 Word 转成 PDF，打印已中止。\n\n"
                        f"原因：{reason}\n"
                        f"{hint}\n\n"
                        "可选办法：\n"
                        "  · 用 WPS 打开该文件后手动打印（版式最保真）\n"
                        "  · COMPLAINTS_PRINT_MODE=direct ./启动GUI.sh 直接打印\n"
                        "  · 先用「批量转PDF」转换后打印 PDF")
            target = pdf

        # 页数变化提醒（转 PDF 后一页变两页）
        warn = page_warning()

        lp = shutil.which("lp")
        if not lp:
            return False, "未找到 CUPS 打印命令 lp，请检查打印机配置"
        r = _run_office([lp, target], timeout=90)
        if r is None:
            return False, "调用 lp 超时或被中断"
        out = (r.stdout or b"").decode("utf-8", "ignore").strip()
        err = (r.stderr or b"").decode("utf-8", "ignore").strip()
        if r.returncode != 0:
            return False, err or out or "lp 返回非零状态"
        msg = out or "已发送到打印队列"
        if warn:
            msg += "\n\n注意：" + warn
        return True, msg
    finally:
        if tmp_dir:
            shutil.rmtree(tmp_dir, ignore_errors=True)
