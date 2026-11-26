# -*- mode: python ; coding: utf-8 -*-
"""
投诉批办单管理系统 —— 统信 UOS / 麒麟 Kylin V10 打包配置
在 x86_64 或 aarch64 的国产 Linux 上用 PyInstaller --onedir 生成可执行文件。
用法（在本文件所在目录执行）：
    pyinstaller complaints-linux.spec --noconfirm
产物：dist/投诉批办单管理系统/投诉批办单管理系统 （ELF 可执行文件，chmod +x 后双击/命令行运行）
"""
import os
from PyInstaller.utils.hooks import collect_all

# 资源一律从 spec 文件所在目录（仓库根目录）取；PyInstaller 会把它们放进 _internal
ROOT = globals().get('SPECPATH', os.path.abspath(os.getcwd()))


def res(rel, dst):
    """生成绝对路径的 (源, 目标) 数据项。"""
    return (os.path.join(ROOT, rel), dst)

datas = []
binaries = []
hiddenimports = []

for pkg in ('ttkbootstrap', 'charset_normalizer', 'PyPDF2',
            'pypinyin', 'qrcode', 'openpyxl', 'PIL', 'docx', 'pdfplumber'):
    try:
        d, b, h = collect_all(pkg)
        datas += d
        binaries += b
        hiddenimports += h
    except Exception:
        pass

hiddenimports += [
    'tkinter', 'tkinter.ttk', 'tkinter.filedialog', 'tkinter.messagebox',
    'tkinter.scrolledtext', 'tkinter.simpledialog',
    '_tkinter', 'sqlite3', 'platform_compat', 'ttk_compat',
    'ttk_gui_selftest', 'ui_layout', 'wps_compat', 'emoji_compat', 'font_compat',
    'print_env',
    'asyncio', 'asyncio.unix_events', 'asyncio.selector_events',
]

# 只读资源：模板、图标、默认配置（仅收录仓库中实际存在的文件）
datas += [
    res('投诉批办单示例#1.docx', '.'),
    res('1.png', '.'),
    res('1.ico', '.'),
    res('AppData/filling_rules_default.json', 'AppData'),
    res('AppData/filling_rules.json', 'AppData'),
    res('AppData/path_config.json', 'AppData'),
    res('AppData/quick_buttons_config.json', 'AppData'),
    res('AppData/qrcode_config.json', 'AppData'),
    res('AppData/input_history.json', 'AppData'),
    res('AppData/counter.json', 'AppData'),
    res('AppData/config.json', 'AppData'),
    res('AppData/window_config.json', 'AppData'),
    res('AppData/last_database.json', 'AppData'),
    res('AppData/complaints.db', 'AppData'),
    res('print_delay_config.json', '.'),
]
datas = [item for item in datas if os.path.exists(item[0])]


# ---------------------------------------------------------------------------
# 剔除不应打包的系统库
# ---------------------------------------------------------------------------
# 背景（麒麟现场 2026-11-25）：
#   Pillow 等依赖会把自带的 libfreetype.so.6 / libfontconfig.so.1 收集进
#   _internal/。PyInstaller 运行时会用 LD_LIBRARY_PATH 指向该目录，
#   于是我们启动的 soffice 子进程继承了它，加载到旧版 freetype
#   （缺 FT_Done_MM_Var），系统的 libcairo 找不到该符号 —— LibreOffice
#   直接报 symbol lookup error，打印/转 PDF 全部失败。
#
# 修法有两层：
#   1) 这里的 filter —— 根本不要把这类系统库打进来（最彻底）
#   2) 运行期 print_env.clean_env() —— 即使漏打进来，子进程也用系统库
#
# 注意：只剔除「系统必然自带、且 ABI 稳定」的基础库。Python 扩展自己
# 需要的库（如 libjpeg 之于 Pillow）不进这个列表，否则 Pillow 会坏。
_SYSTEM_LIB_PREFIXES = (
    'libfreetype',
    'libfontconfig',
    'libcairo',
    'libpango',
    'libharfbuzz',
    'libglib-2.0',
    'libgobject-2.0',
    'libgio-2.0',
    'libX11',
    'libXext',
    'libXrender',
    'libxcb',
    'libexpat',
    'libuuid',
)
# 说明：libz.so 不在此列。zlib 被多个扩展依赖，且冲突极少；
# 排除它反而可能让某些模块找不到符号。只排「确认冲突且系统必有」的库。


def _is_conflicting_system_lib(path):
    """
    判断某个路径的文件名是否属于会与系统冲突的基础库。

    注意：判断依据是**文件名**，必须传源文件路径（src），
    不能传打包目标目录（dest）。PyInstaller 的 binaries 项是
    (src, dest)，其中 dest 常常就是 "."，拿它取 basename 永远是 "."
    —— 之前就是这么写的，导致过滤函数一个都没排除掉。
    """
    name = os.path.basename(str(path)).lower()
    for p in _SYSTEM_LIB_PREFIXES:
        if name.startswith(p.lower()):
            return True
    return False


def _filter_binaries(items):
    """
    从 PyInstaller 的 binaries 列表里去掉冲突的系统库。

    binaries 元素形态：(源路径, 打包内目标目录)。
    目标目录为空表示放在 _internal 根目录 —— 正是这些会通过
    LD_LIBRARY_PATH 顶掉系统库的位置。
    """
    kept = []
    removed = []
    for item in items:
        try:
            src, dest = item[0], item[1]
        except Exception:
            kept.append(item)
            continue
        # 只处理「放在根」或「放在 .libs」这类参与运行时库搜索的项。
        # 判定必须用 src 的文件名（dest 往往就是 "."）。
        if _is_conflicting_system_lib(src):
            removed.append(os.path.basename(str(src)))
            continue
        kept.append(item)
    if removed:
        print('[spec] 已排除 %d 个可能与系统冲突的库: %s'
              % (len(removed), ', '.join(sorted(set(removed))[:8])))
    return kept


binaries = _filter_binaries(binaries)

a = Analysis(
    ['gui_app.py'],
    pathex=[ROOT],
    binaries=binaries,
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=['matplotlib', 'numpy', 'scipy', 'pandas',
              'win32com', 'win32api', 'win32print', 'pythoncom',
              'docx2pdf', 'pywin32'],
    noarchive=False,
)

# ---------------------------------------------------------------------------
# Analysis 之后再过滤一次 —— 这一步才是真正生效的
# ---------------------------------------------------------------------------
# 上面传入 binaries= 的过滤只作用于「我们手动收集」的那部分；
# PyInstaller 分析依赖图时，会把 Pillow 依赖的 libfreetype /
# libfontconfig 等**重新加进** a.binaries（它们出现在 hook 与
# binary dependency 扫描里），于是仍然被打进 _internal。
# 实测：只在传参前过滤时，产物里 libfreetype.so.6 与
# libfontconfig.so.1 依然存在。
# 因此必须在 Analysis 结束后对 a.binaries 再过滤一次。
_before_cnt = len(a.binaries)
a.binaries = _filter_binaries(a.binaries)
print('[spec] Analysis 后过滤 binaries: %d -> %d'
      % (_before_cnt, len(a.binaries)))

pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name='投诉批办单管理系统',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)

coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=False,
    upx_exclude=[],
    name='投诉批办单管理系统',
)
