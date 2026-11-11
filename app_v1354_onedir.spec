# -*- mode: python ; coding: utf-8 -*-
# 投诉批办单管理系统 v1.38.2 文件夹模式（最稳定）
# 使用--onedir模式，避免PKG Archive兼容性问题
# 这是最可靠的打包方式

from PyInstaller.utils.hooks import collect_all

# 收集所有资源
datas = []
binaries = []
hiddenimports = []

# 收集ttkbootstrap的所有资源（主题和图标）
tmp_ret = collect_all('ttkbootstrap')
datas += tmp_ret[0]; binaries += tmp_ret[1]; hiddenimports += tmp_ret[2]

# 收集charset_normalizer的所有资源（包括C扩展）
tmp_ret = collect_all('charset_normalizer')
datas += tmp_ret[0]; binaries += tmp_ret[1]; hiddenimports += tmp_ret[2]

# 收集PyPDF2
tmp_ret = collect_all('PyPDF2')
datas += tmp_ret[0]; binaries += tmp_ret[1]; hiddenimports += tmp_ret[2]

# 收集docx2pdf
tmp_ret = collect_all('docx2pdf')
datas += tmp_ret[0]; binaries += tmp_ret[1]; hiddenimports += tmp_ret[2]

# 收集tqdm（docx2pdf依赖）
tmp_ret = collect_all('tqdm')
datas += tmp_ret[0]; binaries += tmp_ret[1]; hiddenimports += tmp_ret[2]

# 收集pypinyin（拼音首字母匹配功能）
tmp_ret = collect_all('pypinyin')
datas += tmp_ret[0]; binaries += tmp_ret[1]; hiddenimports += tmp_ret[2]

# 收集qrcode（二维码生成功能）
try:
    tmp_ret = collect_all('qrcode')
    datas += tmp_ret[0]; binaries += tmp_ret[1]; hiddenimports += tmp_ret[2]
    print("✓ qrcode模块已收集")
except:
    print("警告：qrcode未安装")

# 收集openpyxl（Excel导出功能）
try:
    tmp_ret = collect_all('openpyxl')
    datas += tmp_ret[0]; binaries += tmp_ret[1]; hiddenimports += tmp_ret[2]
    print("✓ openpyxl模块已收集")
except:
    print("警告：openpyxl未安装")

# 添加asyncio相关的所有隐藏导入（完整收集）
hiddenimports += [
    # asyncio核心模块
    'asyncio',
    'asyncio.base_events',
    'asyncio.selector_events',
    'asyncio.windows_events',
    'asyncio.proactor_events',
    'asyncio.futures',
    'asyncio.tasks',
    'asyncio.streams',
    'asyncio.subprocess',
    'asyncio.windows_utils',
    'asyncio.coroutines',
    'asyncio.events',
    'asyncio.exceptions',
    'asyncio.locks',
    'asyncio.protocols',
    'asyncio.queues',
    'asyncio.runners',
    'asyncio.base_subprocess',
    'asyncio.constants',
    'asyncio.format_helpers',
    'asyncio.log',
    'asyncio.sslproto',
    'asyncio.staggered',
    'asyncio.taskgroups',
    'asyncio.threads',
    'asyncio.transports',
    'asyncio.trsock',
    'asyncio.unix_events',
    # tqdm.asyncio依赖
    'tqdm.asyncio',
    'tqdm.auto',
    'tqdm.std',
]

# 添加模板文件
datas += [('投诉批办单示例#1.docx', '.')]

print(f"收集了 {len(datas)} 个数据文件")
print(f"收集了 {len(hiddenimports)} 个隐藏导入")

a = Analysis(
    ['gui_app.py'],
    pathex=[],
    binaries=binaries,
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=['matplotlib', 'numpy', 'scipy', 'pandas'],  # 排除不需要的大型库
    noarchive=False,
)

pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    [],  # ← 关键：空列表，不打包binaries和datas到exe中
    exclude_binaries=True,  # ← 关键：不将依赖打包到exe中
    name='投诉批办单管理系统v1.38.2',
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
    icon='1.ico',
)

# ← 关键：创建COLLECT，生成文件夹结构
coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=False,
    upx_exclude=[],
    name='投诉批办单管理系统v1.38.2',
)
