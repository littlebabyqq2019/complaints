#!/usr/bin/env bash
# 在 glibc 2.28 基线容器（ubuntu:18.04 / debian:10 同级）内编译 Python 3.10 并打包。
# 目的：让产物能在麒麟 V10 SP1、统信 UOS 20 等 glibc 2.28~2.31 的国产系统上运行。
# 由 GitHub Actions 的 docker run 调用，工作目录挂载为 /src。
set -euo pipefail

ARCH="${1:-x86_64}"
PYVER="3.10.14"
PREFIX="/opt/py310"
SRC="/src"

echo "=== 目标架构: ${ARCH} / 容器 glibc: $(ldd --version 2>&1 | head -1) ==="
uname -m

# ---------- 1. 系统依赖：编译 Python + Tk/Tcl + 中文字体 + Office + 打印 ----------
export DEBIAN_FRONTEND=noninteractive
# Ubuntu 18.04 已进入 ESM，部分镜像的 archive.ubuntu.com 源失效，切换到 old-releases
if ! apt-get update -qq 2>/tmp/apt.err; then
    echo "apt 源不可用，切换到 old-releases.ubuntu.com"
    sed -i 's|http://archive.ubuntu.com/ubuntu|http://old-releases.ubuntu.com/ubuntu|g; s|http://security.ubuntu.com/ubuntu|http://old-releases.ubuntu.com/ubuntu|g' /etc/apt/sources.list || true
    apt-get update -qq || { cat /tmp/apt.err; exit 1; }
fi
apt-get install -y -qq --no-install-recommends \
    build-essential wget ca-certificates xz-utils zlib1g-dev \
    libssl-dev libffi-dev libbz2-dev liblzma-dev libsqlite3-dev \
    libreadline-dev libncursesw5-dev libgdbm-dev libuuid1 uuid-dev \
    tcl-dev tk-dev libx11-dev libxext-dev libxft-dev libxrender-dev \
    libxss-dev libxrandr-dev libxi-dev libxcursor-dev \
    libfreetype6-dev fontconfig libfontconfig1-dev \
    fonts-noto-cjk fonts-wqy-microhei \
    libreoffice-writer cups-client xdg-utils file patchelf >/dev/null

# ---------- 2. 编译 Python 3.10（静态 libpython，避免运行时找系统 Python） ----------
if [ ! -x "${PREFIX}/bin/python3.10" ]; then
    cd /tmp
    wget -q "https://www.python.org/ftp/python/${PYVER}/Python-${PYVER}.tar.xz"
    tar -xf "Python-${PYVER}.tar.xz"
    cd "Python-${PYVER}"
    ./configure --prefix="${PREFIX}" --enable-shared \
        --with-ensurepip=install \
        LDFLAGS="-Wl,-rpath,${PREFIX}/lib" >/dev/null
    make -j"$(nproc)" >/dev/null
    make install >/dev/null
    cd / && rm -rf "/tmp/Python-${PYVER}" "/tmp/Python-${PYVER}.tar.xz"
fi
export PATH="${PREFIX}/bin:${PATH}"
export LD_LIBRARY_PATH="${PREFIX}/lib${LD_LIBRARY_PATH:+:$LD_LIBRARY_PATH}"
python3.10 -c "import sys, _tkinter, tkinter, sqlite3, ssl; print('Python', sys.version.split()[0], '| Tk', tkinter.TkVersion, '| OK')"

# ---------- 3. 安装 Python 依赖 ----------
python3.10 -m pip install --no-cache-dir --upgrade pip setuptools wheel >/dev/null
python3.10 -m pip install --no-cache-dir \
    "pdfplumber==0.11.0" "python-docx==1.1.0" \
    "openpyxl>=3.1.0" "Pillow>=10.0.0,<11" "ttkbootstrap>=1.10.1" \
    "PyPDF2>=3.0.0" "pypinyin>=0.50.0" "qrcode>=7.4.2" \
    "pyinstaller>=6.5.0"

# ---------- 4. 校验源码 ----------
cd "${SRC}"
python3.10 -m py_compile gui_app.py main.py database.py rule_engine.py \
    rule_management_ui.py platform_compat.py
python3.10 -c "import platform_compat, main, database, rule_engine, rule_management_ui; print('模块导入通过')"
python3.10 -c "from PyPDF2 import PdfWriter; print('PdfWriter OK')"

# ---------- 5. PyInstaller 打包 ----------
python3.10 -m PyInstaller complaints-linux.spec --noconfirm --distpath "${SRC}/dist" --workpath /tmp/pyibuild
chmod +x "${SRC}/dist/投诉批办单管理系统/投诉批办单管理系统"

# ---------- 6. 校验产物架构与最高 GLIBC 符号版本 ----------
EXE="${SRC}/dist/投诉批办单管理系统/投诉批办单管理系统"
echo "=== 产物架构 ==="
file "$EXE"
echo "=== GLIBC 最高符号版本（必须 <= 2.28） ==="
MAXGLIBC=$(python3.10 - <<'PY'
import re, glob, os
root = "/src/dist/投诉批办单管理系统"
best = (0, 0)
for p in glob.glob(root + "/**/*", recursive=True):
    if not os.path.isfile(p):
        continue
    try:
        with open(p, "rb") as f:
            head = f.read(4)
            if head != b"\x7fELF":      # 只检查 ELF 二进制，跳过 Python 源码里的字符串
                continue
            f.seek(0)
            data = f.read()
    except OSError:
        continue
    for a, b in re.findall(rb"GLIBC_(\d+)\.(\d+)", data):
        v = (int(a), int(b))
        if v > best:
            best = v
print("%d.%d" % best)
PY
)
echo "MAX_GLIBC=${MAXGLIBC}"
MAJ=${MAXGLIBC%%.*}; MIN=${MAXGLIBC##*.}
if [ "$MAJ" -gt 2 ] || { [ "$MAJ" -eq 2 ] && [ "$MIN" -gt 28 ]; }; then
    echo "!! 产物依赖 GLIBC_${MAXGLIBC}，高于基线 2.28，麒麟 V10 SP1 / UOS 20 无法运行" >&2
    exit 1
fi
echo "=== glibc 基线校验通过（<= 2.28） ==="

# ---------- 7. 补齐发布包内容 ----------
OUT="${SRC}/dist/投诉批办单管理系统"
cd "${SRC}"
cp 启动GUI.sh 启动命令行.sh 投诉批办单管理系统.desktop 使用说明.txt "$OUT/"
chmod +x "$OUT/启动GUI.sh" "$OUT/启动命令行.sh" "$OUT/投诉批办单管理系统"
cp 1.png "$OUT/1.png"
echo "${ARCH}" > "$OUT/构建架构.txt"

# ---------- 8. tar.gz ----------
cd "${SRC}/dist"
tar -czf "complaints-linux-${ARCH}.tar.gz" "投诉批办单管理系统"
sha256sum "complaints-linux-${ARCH}.tar.gz" > "complaints-linux-${ARCH}.tar.gz.sha256"
ls -lh "complaints-linux-${ARCH}.tar.gz"
echo "=== BUILD_DONE ${ARCH} ==="
