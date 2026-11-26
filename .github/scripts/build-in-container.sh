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
    xvfb xauth x11-utils \
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
    "openpyxl>=3.1.0" "Pillow>=10.0.0,<11" "ttkbootstrap>=1.10.1,<3" \
    "PyPDF2>=3.0.0" "pypinyin>=0.50.0" "qrcode>=7.4.2" \
    "pyinstaller>=6.5.0"

# ---------- 4. 校验源码 ----------
cd "${SRC}"
python3.10 -m py_compile gui_app.py main.py database.py rule_engine.py \
    rule_management_ui.py platform_compat.py ttk_compat.py ttk_gui_selftest.py ui_layout.py wps_compat.py emoji_compat.py font_compat.py print_env.py
python3.10 -c "import platform_compat, main, database, rule_engine, rule_management_ui, ttk_compat, ttk_gui_selftest, ui_layout, wps_compat, emoji_compat, font_compat, print_env; print('模块导入通过')"
# ttkbootstrap 1.x/2.x 的控件别名校验（不需要图形环境）。2.x 去掉了 PanedWindow
# 驼峰别名，若这里失败说明打包环境的 ttkbootstrap 版本越界，产物运行会启动失败。
python3.10 -c "import ttk_compat; v,m = ttk_compat.verify_widgets(); assert not m, '缺少控件: ' + ','.join(m); print('ttk 控件校验通过:', ttk_compat.describe())"
# WPS 调用层自检：容器里没有 WPS，验证「未安装时能优雅降级」，不阻断构建
python3.10 -c "import wps_compat, main; assert wps_compat.convert_with_wps('/nonexistent.doc','docx') is None; print('WPS 调用层:', main.wps_status())"
python3.10 -c "from PyPDF2 import PdfWriter; print('PdfWriter OK')"

# ---------- 4b. 无头 GUI 冒烟测试（源码） ----------
# import 通过 ≠ 窗口能建起来。2026-11-12 麒麟现场就崩在 create_widgets()，
# 只有真的把三栏布局、规则窗口、全部对话框建一遍才能提前发现。
echo "=== 源码 GUI 冒烟测试（Xvfb） ==="
export DISPLAY=:99
Xvfb :99 -screen 0 1600x1000x24 >/tmp/xvfb.log 2>&1 &
XVFB_PID=$!
trap 'kill ${XVFB_PID} 2>/dev/null || true' EXIT
for _ in $(seq 1 50); do
    if xdpyinfo -display :99 >/dev/null 2>&1; then break; fi
    sleep 0.2
done
if ! xdpyinfo -display :99 >/dev/null 2>&1; then
    echo "!! Xvfb 未就绪，GUI 冒烟测试无法进行" >&2
    cat /tmp/xvfb.log >&2 || true
    exit 1
fi
python3.10 ttk_gui_selftest.py || { echo "!! 源码 GUI 冒烟测试失败" >&2; exit 1; }

# ---------- 5. PyInstaller 打包 ----------
python3.10 -m PyInstaller complaints-linux.spec --noconfirm --distpath "${SRC}/dist" --workpath /tmp/pyibuild
chmod +x "${SRC}/dist/投诉批办单管理系统/投诉批办单管理系统"

# ---------- 5b. 打包产物自检（真正跑一遍 Release 里那个 ELF） ----------
# 源码能跑不代表打包后的产物能跑：hiddenimports 漏了、资源没进去，
# 都只会在产物里暴露。--selftest 会把界面完整搭一遍再退出。
echo "=== 打包产物自检（--selftest） ==="
DIST_DIR="${SRC}/dist/投诉批办单管理系统"
# 注意：脚本顶部是 set -e，直接调用失败会立刻退出，这里显式捕获退出码
set +e
( cd "${DIST_DIR}" && ./投诉批办单管理系统 --selftest )
SELFTEST_RC=$?
set -e
echo "--- 产物自检报告 ---"
cat "${DIST_DIR}/AppData/ttk_selftest.log" 2>/dev/null || echo "（无报告文件）"
if [ "${SELFTEST_RC}" -ne 0 ]; then
    echo "!! 打包产物自检失败（exit ${SELFTEST_RC}），产物不可发布" >&2
    exit 1
fi
echo "=== 产物自检通过 ==="

# ---------- 5c. 用系统库替换 _internal 里的旧 freetype / fontconfig ----------
# 背景：Pillow 等依赖把旧版 libfreetype.so.6 / libfontconfig.so.1 打进
# _internal。PyInstaller 用 LD_LIBRARY_PATH 指向该目录，我们启动的 soffice
# 子进程继承后加载到旧 freetype，而系统 libcairo 需要 FT_Done_MM_Var
# （旧 freetype 没有）—— 报 symbol lookup error，打印/转 PDF 全挂。
#
# 为什么不用 spec 的排除规则：PyInstaller 的依赖库是成套改名的
# （libxcb-<hash>.so.x），删掉会打断依赖链导致产物起不来（已实测崩过）。
# 而且 Tk 的字体栈本身就要 freetype/fontconfig。
#
# 所以采用「替换」而不是「删除」：保留文件名与 hashed 名不变，只把内容
# 换成构建机上的系统版本（glibc 基线容器里的 freetype 含所需符号）。
#
# 若构建机没有同名系统库，则跳过并给出提示，不影响构建。
echo "=== 用系统库替换 _internal 内的 freetype / fontconfig ==="
REPLACED=0
for f in "${DIST_DIR}/_internal"/libfreetype*.so* "${DIST_DIR}/_internal"/libfontconfig*.so*; do
    [ -e "$f" ] || continue
    base=$(basename "$f")
    # 去掉 PyInstaller 加的 hash 段：libfreetype-e3eca772.so.6.20.1 -> libfreetype.so.6
    plain=$(echo "$base" | sed -E 's/^lib(freetype|fontconfig)-[0-9a-f]+\.so/lib\1.so/')
    # 库族名（libfreetype / libfontconfig），用于最后的模糊兜底
    fam=$(echo "$base" | sed -E 's/^(-?lib(freetype|fontconfig)).*/\1/')
    syslib=""
    for cand in "/lib/$(uname -m)-linux-gnu/$plain" \
                "/usr/lib/$(uname -m)-linux-gnu/$plain" \
                "/lib/$plain" "/usr/lib/$plain"; do
        if [ -f "$cand" ]; then syslib="$cand"; break; fi
    done
    # 再按 soname 模糊找（版本号可能略有不同）
    if [ -z "$syslib" ]; then
        syslib=$(ls /lib/*-linux-gnu/${plain%.so*}*.so.* 2>/dev/null | head -1 || true)
    fi
    # 最后按库族名兜底：系统里任意 libfreetype.so.* 都算
    if [ -z "$syslib" ] && [ -n "$fam" ]; then
        syslib=$(ls /lib/*-linux-gnu/${fam}.so* /usr/lib/*-linux-gnu/${fam}.so* \
                       /lib/${fam}.so* /usr/lib/${fam}.so* 2>/dev/null | head -1 || true)
    fi
    if [ -n "$syslib" ] && [ -f "$syslib" ]; then
        cp -f "$syslib" "$f"
        echo "    替换 $base  <-  $syslib"
        REPLACED=$((REPLACED + 1))
    else
        echo "    (跳过) $base 未找到对应系统库 $plain"
    fi
done
if [ "${REPLACED}" -gt 0 ]; then
    echo "OK   已用系统库替换 ${REPLACED} 个文件"
else
    echo "提示 _internal 内没有需要替换的 freetype/fontconfig"
fi
# 校验：替换后 _internal 里的 freetype 必须含 FT_Done_MM_Var
FT_CHECK=$(find "${DIST_DIR}/_internal" -maxdepth 1 -name 'libfreetype*.so*' | head -1)
if [ -n "${FT_CHECK}" ]; then
    if nm -D "${FT_CHECK}" 2>/dev/null | grep -q FT_Done_MM_Var \
       || strings -a "${FT_CHECK}" 2>/dev/null | grep -q FT_Done_MM_Var; then
        echo "OK   _internal 内的 freetype 含 FT_Done_MM_Var"
    else
        echo "!! _internal 内的 freetype 仍缺 FT_Done_MM_Var，soffice 会启动失败" >&2
        echo "   文件: ${FT_CHECK}" >&2
        exit 1
    fi
fi

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
cp 源码方式安装.sh 诊断.sh "$OUT/"
chmod +x "$OUT/启动GUI.sh" "$OUT/启动命令行.sh" "$OUT/源码方式安装.sh" "$OUT/诊断.sh" "$OUT/投诉批办单管理系统"
cp 1.png "$OUT/1.png"
cp requirements-linux.txt "$OUT/requirements-linux.txt"
echo "${ARCH}" > "$OUT/构建架构.txt"

# ---------- 8. tar.gz ----------
cd "${SRC}/dist"
tar -czf "complaints-linux-${ARCH}.tar.gz" "投诉批办单管理系统"
sha256sum "complaints-linux-${ARCH}.tar.gz" > "complaints-linux-${ARCH}.tar.gz.sha256"
ls -lh "complaints-linux-${ARCH}.tar.gz"
echo "=== BUILD_DONE ${ARCH} ==="
