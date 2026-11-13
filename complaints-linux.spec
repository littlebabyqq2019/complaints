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
    'ttk_gui_selftest', 'ui_layout', 'wps_compat',
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
