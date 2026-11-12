#!/usr/bin/env bash
# 命令行批处理模式（等价于 Windows 版 main.py）
# 注意：必须写 "./启动命令行.sh" 或 "bash 启动命令行.sh"，
#       Linux 默认不从当前目录查找命令，直接写 "启动命令行.sh" 会提示「未找到命令」。
set -u
cd "$(dirname "$0")" || exit 1

APP="./投诉批办单管理系统"
mkdir -p ./AppData 2>/dev/null

# 优先直接使用打包后的可执行文件（自带 Python，不需要系统装 Python）
if [ -f "$APP" ] && [ -d "./_internal" ]; then
    [ -x "$APP" ] || chmod +x "$APP" 2>/dev/null || true
    if [ -z "${DISPLAY:-}" ] && [ -z "${WAYLAND_DISPLAY:-}" ] && command -v ldd >/dev/null 2>&1; then
        MISSING=$(ldd "$APP" 2>&1 | grep -i "not found" || true)
        if [ -n "$MISSING" ]; then
            echo "缺少系统动态库，请先安装对应软件包：" >&2
            echo "$MISSING" >&2
            echo "本机 glibc: $(ldd --version 2>&1 | head -1)" >&2
            exit 1
        fi
    fi
    exec "$APP" --cli "$@"
fi

# 源码运行兜底
PY=$(command -v python3 || command -v python || true)
if [ -z "$PY" ]; then
    echo "未找到 python3。打包版不需要 Python；源码方式请执行：" >&2
    echo "  sudo apt install python3 python3-tk python3-pip" >&2
    echo "  python3 -m pip install -r requirements-linux.txt" >&2
    exit 1
fi
exec "$PY" main.py "$@"
