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
    """一次性读取 fontconfig 的字体家族列表；失败时返回空串（交给 Tk 自行回退）。"""
    try:
        r = subprocess.run(["fc-list", ":", "family"],
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
    """用系统默认程序打开文件或目录，等价 Windows 的 os.startfile。"""
    path = os.path.abspath(path)
    if IS_WINDOWS:
        os.startfile(path)  # noqa: F821  (仅 Windows 存在)
        return True
    if IS_MACOS:
        subprocess.Popen(["open", path], stdout=subprocess.DEVNULL,
                         stderr=subprocess.DEVNULL, start_new_session=True)
        return True
    for cmd in ("xdg-open", "gio", "gnome-open", "kde-open"):
        exe = shutil.which(cmd)
        if not exe:
            continue
        argv = [exe, "open", path] if cmd == "gio" else [exe, path]
        try:
            subprocess.Popen(argv, stdout=subprocess.DEVNULL,
                             stderr=subprocess.DEVNULL, start_new_session=True)
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
    try:
        return subprocess.run(cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE,
                              timeout=timeout)
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


def last_convert_error():
    """最近一次转换失败的原因（给 GUI 提示用）。"""
    return _LAST_CONVERT_ERROR[0]


def _try_soffice(src, work_dir, target_fmt, timeout=180):
    """
    用 LibreOffice 转换一次。返回产物路径或 None，失败原因写入 _LAST_CONVERT_ERROR。
    会在临时目录里做「安全名副本」，避免 # & 空格 中文 等字符导致的静默失败。
    """
    soffice = _find_soffice()
    if not soffice:
        _LAST_CONVERT_ERROR[0] = '未找到 LibreOffice（soffice）'
        return None

    ext = os.path.splitext(src)[1].lower().lstrip('.') or 'docx'
    stage = tempfile.mkdtemp(prefix='lo_stage_')
    out_dir = tempfile.mkdtemp(prefix='lo_out_')
    try:
        safe_src = _safe_copy_for_conversion(src, stage, '.' + ext)
        before = set()
        # 转换串，pdf 与 docx 的 filter 不同
        if target_fmt == 'pdf':
            conv = 'pdf:writer_pdf_Export'
        elif target_fmt == 'docx':
            conv = 'docx:MS Word 2007 XML'
        else:
            conv = target_fmt

        profile = 'file://' + os.path.join(stage, 'lo_profile')
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
        _LAST_CONVERT_ERROR[0] = 'LibreOffice 转换未产生文件：' + detail[:300]
        return None
    finally:
        shutil.rmtree(stage, ignore_errors=True)


def word_to_pdf(word_path, out_pdf_path=None):
    """
    Word -> PDF。
      Windows: docx2pdf（原逻辑）
      Linux  : LibreOffice 优先（唯一可靠），WPS 兜底，unoconv 最后
    成功返回 PDF 路径，失败返回 None；失败原因见 last_convert_error()。
    """
    word_path = os.path.abspath(word_path)
    base = os.path.splitext(os.path.basename(word_path))[0]
    if out_pdf_path is None:
        out_pdf_path = os.path.join(os.path.dirname(word_path), base + '.pdf')
    out_pdf_path = os.path.abspath(out_pdf_path)
    os.makedirs(os.path.dirname(out_pdf_path), exist_ok=True)
    _LAST_CONVERT_ERROR[0] = ''

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


def print_file(path):
    """
    打印单个文档，等价 Windows 的 win32api.ShellExecute(0,"print",...)。

    Linux 下必须先转成 PDF 再交给 CUPS 的 lp：CUPS 不认识 docx 的
    MIME 类型，直接递 .docx 会报
        /bin/lp: 不支持的 document-format
        "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    这正是麒麟现场的报错。所以转换失败时**直接返回失败**，
    绝不把原 Word 文件塞给 lp。

    PDF 直接 lp。
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

    tmp_dir = None
    target = path
    try:
        if path.lower().endswith((".doc", ".docx", ".wps", ".rtf", ".odt")):
            tmp_dir = tempfile.mkdtemp(prefix="printconv_")
            base = os.path.splitext(os.path.basename(path))[0]
            pdf = word_to_pdf(path, os.path.join(tmp_dir, base + ".pdf"))
            if not pdf or not os.path.isfile(pdf) or os.path.getsize(pdf) <= 0:
                # 关键：转换失败必须直接报错，不能拿 docx 去喂 lp。
                # 把 LibreOffice 的真实失败原因带出来，避免「装了 LO 却不知道为啥失败」。
                reason = last_convert_error() or "未得到具体原因"
                soffice = _find_soffice()
                hint = ("检测到 LibreOffice: " + soffice) if soffice else \
                       ("未检测到 LibreOffice，请安装：\n"
                        "  sudo apt install -y libreoffice-writer")
                return (False,
                        "无法把 Word 转成 PDF，打印已中止。\n\n"
                        f"原因：{reason}\n"
                        f"{hint}\n\n"
                        "可先试：先用「批量转PDF」手动转换，成功后再打印。")
            target = pdf

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
        return True, out or "已发送到打印队列"
    finally:
        if tmp_dir:
            shutil.rmtree(tmp_dir, ignore_errors=True)
