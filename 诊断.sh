#!/usr/bin/env bash
# 现场诊断：在麒麟 / 统信机器上确认「启动失败」的真实原因，一条命令出结论。
# 用法：./诊断.sh            （发布包目录）
#       bash 诊断.sh        （源码目录，同样可用）
# 只读，不改动任何数据。
set -u
cd "$(dirname "$0")" 2>/dev/null || true

echo "=================================================="
echo " 投诉批办单管理系统 · 现场诊断"
echo "=================================================="
echo "时间    : $(date)"
echo "系统    : $(uname -a)"
echo "架构    : $(uname -m)"
echo "glibc   : $(ldd --version 2>&1 | head -1)"
echo "显示    : DISPLAY=${DISPLAY:-<空>} WAYLAND_DISPLAY=${WAYLAND_DISPLAY:-<空>}"
echo

# 1) 找 Python：优先源码方式的 .venv，其次打包产物自带的运行时，最后系统 python3
PY=""
if [ -x "./.venv/bin/python" ]; then
    PY="./.venv/bin/python"
    echo "Python  : 源码虚拟环境 ./.venv/bin/python"
elif [ -x "./投诉批办单管理系统" ] && [ -d "./_internal" ]; then
    echo "Python  : 发布包自带运行时（_internal/），跳过下面基于 python3 的检查"
    echo "          可执行文件: $(file 投诉批办单管理系统 2>/dev/null)"
    echo "          缺失动态库: $(ldd ./投诉批办单管理系统 2>&1 | grep -i 'not found' || echo 无)"
elif command -v python3 >/dev/null 2>&1; then
    PY="python3"
    echo "Python  : 系统 python3"
fi

if [ -n "$PY" ]; then
    "$PY" - <<'PYEOF'
import sys
print(f"          版本: {sys.version.split()[0]}  可执行: {sys.executable}")

# 2) Tk / ttkbootstrap 依赖
try:
    import tkinter
    print(f"  tkinter : OK (Tk {tkinter.TkVersion})")
except Exception as e:
    print(f"  tkinter : 不可用 -> {e}")
    print("            修复: sudo apt install -y python3-tk")
    raise SystemExit(0)

try:
    import ttk_compat
    v, missing = ttk_compat.verify_widgets()
    print(f"  ttk     : {ttk_compat.describe()}")
    if missing:
        print(f"            缺少控件: {', '.join(missing)}")
        print("            修复: python3 -m pip install 'ttkbootstrap>=1.10.1,<3'")
    else:
        print("            控件自检: 通过（可以启动）")
except ModuleNotFoundError:
    print("  ttk     : 找不到 ttk_compat.py —— 这是旧版本源码，"
          "请用最新源码包（含 ttk_compat.py）")
except Exception as e:
    print(f"  ttk     : 检查失败 -> {e}")

# WPS / LibreOffice：只影响老二进制 .doc，WordML/OOXML 不需要
try:
    import main as _m
    print(f"  office  : {_m.wps_status()}")
    print(f"            prefer_wps={_m.prefer_wps()}")
except Exception as e:
    print(f"  office  : 检查失败 -> {e}")
try:
    import platform_compat as _pc
    print(f"  LO      : {_pc._find_soffice() or '未检测到 LibreOffice'}")
except Exception as e:
    print(f"  LO      : 检查失败 -> {e}")
PYEOF
fi

echo
# 3) 最近一次的启动日志（打包版与源码版共用 AppData/app_startup.log）
if [ -f "AppData/app_startup.log" ]; then
    echo "---- AppData/app_startup.log（最近 20 行）----"
    tail -n 20 "AppData/app_startup.log"
else
    echo "未找到 AppData/app_startup.log（还没启动过，或启动器在更早一步就退出了）"
    [ -f "AppData/启动诊断.log" ] && { echo "---- AppData/启动诊断.log ----"; tail -n 20 "AppData/启动诊断.log"; }
fi

echo
# 4) 深度自检：真把界面搭一遍（有 X 显示时才有意义）
if [ -n "${DISPLAY:-}${WAYLAND_DISPLAY:-}" ]; then
    if [ -x "./投诉批办单管理系统" ] && [ -d "./_internal" ]; then
        echo "---- 深度自检：./投诉批办单管理系统 --selftest ----"
        ./投诉批办单管理系统 --selftest >/dev/null 2>&1
        echo "退出码: $?  （0 = 界面能完整建起来）"
        [ -f "AppData/ttk_selftest.log" ] && tail -n 4 "AppData/ttk_selftest.log"
    elif [ -n "$PY" ]; then
        echo "---- 深度自检：$PY ttk_gui_selftest.py ----"
        "$PY" ttk_gui_selftest.py >/dev/null 2>&1
        echo "退出码: $?  （0 = 界面能完整建起来）"
        [ -f "AppData/ttk_selftest.log" ] && tail -n 4 "AppData/ttk_selftest.log"
    fi
else
    echo "跳过深度自检（没有图形显示环境）"
fi

echo
echo "=================================================="
echo " 结论提示"
echo "   出现 PanedWindow 相关报错 / 控件自检不通过 -> 依赖版本问题，"
echo "     python3 -m pip install 'ttkbootstrap>=1.10.1,<3' 后重试"
echo "   tkinter 不可用 -> sudo apt install -y python3-tk"
echo "   架构不匹配 -> 下载与 uname -m 对应的包（见 使用说明.txt 第五节）"
echo "=================================================="
