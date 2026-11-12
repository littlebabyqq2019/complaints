#!/usr/bin/env bash
# 投诉批办单管理系统 —— 统信 UOS / 麒麟 Kylin V10 启动器
# 用法：终端执行 ./启动GUI.sh ，或在文件管理器中双击本文件（选「在终端中运行」）
# 注意：Linux 不会从当前目录查找命令，必须写 "./启动GUI.sh"，只写 "启动GUI.sh" 会提示未找到命令。
set -u
cd "$(dirname "$0")" || exit 1
APP="./投诉批办单管理系统"
LOGDIR="./AppData"
mkdir -p "$LOGDIR" 2>/dev/null
LOG="$LOGDIR/启动诊断.log"

# 源码方式（源码方式安装.sh 创建的 .venv）：优先用虚拟环境启动，不依赖打包产物
if [ -x "./.venv/bin/python" ] && [ -f "./gui_app.py" ]; then
    exec ./.venv/bin/python gui_app.py "$@"
fi

fail() {
    echo "==================================================" >&2
    echo "  投诉批办单管理系统 启动失败" >&2
    echo "==================================================" >&2
    echo "$*" >&2
    echo "详细信息已写入: $LOG" >&2
    echo "==================================================" >&2
    {
        echo "时间: $(date)"
        echo "系统: $(uname -a)"
        echo "架构: $(uname -m)"
        echo "glibc: $(ldd --version 2>&1 | head -1)"
        echo "原因: $*"
        echo
    } > "$LOG" 2>&1
    # 有图形环境时弹窗提示，避免"双击没反应"
    if [ -n "${DISPLAY:-}${WAYLAND_DISPLAY:-}" ]; then
        if command -v zenity >/dev/null 2>&1; then
            zenity --error --title="投诉批办单管理系统 启动失败" --text="$*" 2>/dev/null || true
        elif command -v kdialog >/dev/null 2>&1; then
            kdialog --error "$*" 2>/dev/null || true
        elif command -v xmessage >/dev/null 2>&1; then
            xmessage -center "投诉批办单管理系统 启动失败: $*" 2>/dev/null || true
        fi
    fi
    # 双击场景下保留窗口，让用户能看到错误
    if [ -t 0 ]; then
        read -r -p "按回车键关闭..." _ || true
    else
        sleep 12
    fi
    exit 1
}

# 1) 可执行文件存在性 / 权限 / 架构
[ -f "$APP" ] || fail "未找到可执行文件 $APP，请确认解压完整（应包含 _internal 目录）。"
[ -x "$APP" ] || chmod +x "$APP" 2>/dev/null || true
[ -x "$APP" ] || fail "$APP 没有可执行权限，请执行: chmod +x \"$APP\""
[ -d "./_internal" ] || fail "缺少 _internal 依赖目录，请重新完整解压发布包。"
case "$(uname -m)" in
    aarch64|arm64) WANT="ARM aarch64" ;;
    x86_64|amd64)  WANT="x86-64" ;;
    *) fail "未知 CPU 架构 $(uname -m)，本包只提供 x86_64 / aarch64 两个版本。" ;;
esac
if command -v file >/dev/null 2>&1; then
    file "$APP" | grep -q "$WANT" || fail "架构不匹配：本机为 $(uname -m)，但此包是另一种架构。请下载 complaints-linux-$(uname -m | sed 's/amd64/x86_64/;s/arm64/aarch64/').tar.gz"
fi

# 2) glibc 版本预检：给出可操作的提示，而不是静默失败
if command -v ldd >/dev/null 2>&1; then
    GLIBC_VER=$(ldd --version 2>&1 | head -1 | grep -oE '[0-9]+\.[0-9]+$' || true)
    if [ -n "$GLIBC_VER" ]; then
        GMAJ=${GLIBC_VER%%.*}; GMIN=${GLIBC_VER##*.}
        if [ "$GMAJ" -lt 2 ] || { [ "$GMAJ" -eq 2 ] && [ "$GMIN" -lt 28 ]; }; then
            fail "系统 glibc $GLIBC_VER 过低（需 >= 2.28）。请升级系统或用源码方式运行（见使用说明.txt 第四节）。"
        fi
    fi
fi

# 3) 依赖库检查：把缺失库名直接打印出来
if command -v ldd >/dev/null 2>&1; then
    MISSING=$(ldd "$APP" 2>&1 | grep -i "not found" || true)
    if [ -n "$MISSING" ]; then
        fail "缺少系统动态库，请安装对应软件包：\n$MISSING"
    fi
fi

# 4) 图形环境检查
if [ -z "${DISPLAY:-}" ] && [ -z "${WAYLAND_DISPLAY:-}" ]; then
    fail "没有检测到图形显示环境（DISPLAY / WAYLAND_DISPLAY 均为空）。请在桌面会话中运行，或配置 X11 转发。"
fi

# 5) 启动，stdout/stderr 全部留痕；失败时展示诊断而不是静默退出
echo "启动投诉批办单管理系统... (日志: $LOG)"
if "$APP" "$@" >>"$LOG" 2>&1; then
    exit 0
fi
RC=$?
echo "程序退出码: $RC" >> "$LOG"
TAIL=$(tail -n 30 "$LOG" 2>/dev/null)
fail "程序异常退出（退出码 $RC）。最后日志：\n$TAIL"
