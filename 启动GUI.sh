#!/usr/bin/env bash
# 统信 UOS / 麒麟 Kylin V10 启动脚本
# 用法：在文件管理器中双击，或终端执行 ./启动GUI.sh
set -e
cd "$(dirname "$0")"

# 情况一：与 PyInstaller 打包出的可执行文件同级（解压发布包后的目录）
if [ -f "./投诉批办单管理系统" ] && [ -x "./投诉批办单管理系统" ]; then
    exec "./投诉批办单管理系统" "$@"
fi

# 情况二：dist/ 目录结构（exe 在子目录中）
if [ -x "./投诉批办单管理系统/投诉批办单管理系统" ]; then
    exec "./投诉批办单管理系统/投诉批办单管理系统" "$@"
fi

# 情况三：源码运行
PY=$(command -v python3 || command -v python)
if [ -z "$PY" ]; then
    echo "未找到 python3，请先安装：sudo apt install python3 python3-tk" >&2
    exit 1
fi
exec "$PY" gui_app.py "$@"
