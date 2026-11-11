#!/usr/bin/env bash
# 命令行批处理模式（等价于 Windows 版 main.py）
cd "$(dirname "$0")"
PY=$(command -v python3 || command -v python)
exec "$PY" main.py "$@"
