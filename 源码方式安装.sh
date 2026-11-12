#!/usr/bin/env bash
# 源码方式一键安装（适用于 glibc < 2.28 的老版麒麟/统信、龙芯 LoongArch、申威等任意架构）
# 源码是纯 Python，不依赖打包时的 glibc 版本，也不需要本仓库的可执行文件。
# 用法：在源码目录执行 ./源码方式安装.sh
set -euo pipefail
cd "$(dirname "$0")"

echo "=================================================="
echo " 投诉批办单管理系统 · 源码方式安装"
echo "=================================================="
echo "本机架构: $(uname -m)"
echo "glibc   : $(ldd --version 2>&1 | head -1)"
echo

# 1. 系统依赖
PKGS="python3 python3-tk python3-pip python3-venv fonts-noto-cjk fonts-wqy-microhei libreoffice-writer cups-client xdg-utils"
MISSING=""
for p in $PKGS; do
    if ! dpkg -s "$p" >/dev/null 2>&1; then MISSING="$MISSING $p"; fi
done
if [ -n "$MISSING" ]; then
    echo "需要安装以下系统软件包：$MISSING"
    echo "请在执行本脚本的终端里输入密码（或提前手动执行下面这行命令）："
    echo "  sudo apt install -y$MISSING"
    if sudo -n true 2>/dev/null; then
        sudo apt-get update -qq
        sudo apt-get install -y $MISSING
    else
        read -r -p "现在自动安装？[Y/n] " yn
        if [ "${yn:-Y}" = "Y" ] || [ "${yn:-Y}" = "y" ]; then
            sudo apt-get update -qq
            sudo apt-get install -y $MISSING
        else
            echo "已取消。装好依赖后重新运行本脚本。"
            exit 1
        fi
    fi
fi

# 2. 独立虚拟环境
if [ ! -d .venv ]; then
    echo "创建虚拟环境 .venv ..."
    python3 -m venv --system-site-packages .venv
fi
.venv/bin/python -m pip install --upgrade pip -q
echo "安装 Python 依赖..."
.venv/bin/python -m pip install -r requirements-linux.txt

# 3. 校验
.venv/bin/python -m py_compile gui_app.py main.py database.py rule_engine.py \
    rule_management_ui.py platform_compat.py ttk_compat.py
.venv/bin/python -c "import platform_compat, main, database, rule_engine, rule_management_ui, tkinter; print('模块与 Tk 校验通过')"
# ttkbootstrap 1.x/2.x 控件别名自检（2.x 去掉了 PanedWindow 驼峰别名）
.venv/bin/python -c "import ttk_compat; v,m=ttk_compat.verify_widgets(); assert not m, '缺少控件: '+','.join(m); print('ttk 控件校验通过:', ttk_compat.describe())"

# 4. 生成桌面入口（绝对路径，避免双击无效）
DESK="$HOME/.local/share/applications/投诉批办单管理系统.desktop"
mkdir -p "$(dirname "$DESK")" "$HOME/.local/share/icons"
cp 1.png "$HOME/.local/share/icons/投诉批办单管理系统.png" 2>/dev/null || true
cat > "$DESK" <<EOF
[Desktop Entry]
Type=Application
Version=1.0
Name=投诉批办单管理系统
Comment=投诉批办单自动生成与打印管理（源码方式）
Exec=/bin/bash -c 'cd "$PWD" && .venv/bin/python gui_app.py'
Path=$PWD
Icon=$HOME/.local/share/icons/投诉批办单管理系统.png
Terminal=false
Categories=Office;Utility;
EOF
chmod +x "$DESK"

echo
echo "=================================================="
echo " 安装完成"
echo "=================================================="
echo "启动 GUI："
echo "   ./启动GUI.sh           # 会优先用 .venv 运行"
echo "   .venv/bin/python gui_app.py"
echo "命令行批处理："
echo "   .venv/bin/python main.py"
echo
echo "桌面入口已写入：$DESK"
echo "（部分国产桌面需要在应用列表里刷新后才显示）"
echo "=================================================="
