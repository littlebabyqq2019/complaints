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

# ---------- 5c. freetype / fontconfig 的说明与校验 ----------
# 背景：Pillow 等依赖把 libfreetype.so.6 / libfontconfig.so.1 打进 _internal，
# PyInstaller 用 LD_LIBRARY_PATH 指向该目录，soffice 子进程继承后可能加载到
# 这里的 freetype。若该 freetype 缺 FT_Done_MM_Var，系统 libcairo 就报
# symbol lookup error（麒麟现场打印失败的根因）。
#
# 尝试过的两种构建期方案都不行：
#   a) spec 里排除 -> PyInstaller 依赖库成套改名（libxcb-<hash>.so.x），
#      删掉会断链，产物启动报 ImportError（已实测崩过）
#   b) 用构建机的系统库替换 -> 本容器是 ubuntu:18.04，其 freetype 是 2.8.1，
#      **比打包进来的还旧**，同样没有 FT_Done_MM_Var（实测校验失败）
#
# 因此采取两条真正有效的措施：
#   1) 运行期 print_env.clean_env()（本项目的核心修复）：
#      所有外部子进程（soffice/wps/lp/fc-list…）启动前剔除 _internal 路径，
#      强制使用**目标机自己的**系统库 —— 麒麟装的是 2.10.1，含所需符号。
#      这与构建机的字体版本无关，是根治手段。
#   2) 构建期只做「体检」，不替换：
#      记录容器 freetype 版本供排查；若容器 freetype 恰好含该符号则更好，
#      但不强求（不强求才不会误杀构建）。
echo "=== freetype / fontconfig 体检（仅记录，不替换）==="
CONTAINER_FT=$(ls /usr/lib/*-linux-gnu/libfreetype.so.6 /lib/*-linux-gnu/libfreetype.so.6 2>/dev/null | head -1 || true)
if [ -n "${CONTAINER_FT}" ]; then
    echo "    构建机 freetype: ${CONTAINER_FT}"
    if nm -D "${CONTAINER_FT}" 2>/dev/null | grep -q FT_Done_MM_Var \
       || strings -a "${CONTAINER_FT}" 2>/dev/null | grep -q FT_Done_MM_Var; then
        echo "    含 FT_Done_MM_Var"
    else
        # 这是预期情况（ubuntu 18.04 的 freetype 2.8.1 没有该符号），
        # 不影响发布：运行期会剔除 _internal 让子进程用目标机系统库。
        echo "    不含 FT_Done_MM_Var（构建机较旧，属正常）"
        echo "    -> 由运行期 print_env.clean_env() 保证外部程序使用目标机系统库"
    fi
else
    echo "    未找到构建机 freetype（不阻断构建）"
fi
# 确认打包产物里确实带了 freetype（否则说明 Tk 字体栈可能不完整）
if find "${DIST_DIR}/_internal" -maxdepth 1 -name 'libfreetype*.so*' | grep -q .; then
    echo "OK   _internal 内含 freetype（Tk/Pillow 依赖完整）"
else
    echo "提示 _internal 内没有 freetype（可能由系统提供）"
fi

# ---------- 5d. 验证运行期库隔离真的生效 ----------
# 这是本项目修复 FT_Done_MM_Var 的核心机制，必须在 CI 里真实跑一遍：
# 模拟 PyInstaller 的 LD_LIBRARY_PATH 污染，然后确认 print_env.clean_env()
# 产生的环境里不含 _internal，且用它启动外部程序能正常工作。
echo "=== 验证 print_env 库隔离（模拟打包版污染）==="
python3.10 - <<'PY'
import os, sys, subprocess
sys.path.insert(0, '/src')
import print_env

fake_internal = '/src/dist/投诉批办单管理系统/_internal'
saved = dict(os.environ)
try:
    os.environ['LD_LIBRARY_PATH'] = fake_internal
    sys._MEIPASS = fake_internal
    sys.frozen = True
    env = print_env.clean_env()
    lp = env.get('LD_LIBRARY_PATH', '')
    assert '_internal' not in lp, 'clean_env 未剔除 _internal: %r' % lp
    print('  OK    clean_env 已剔除 _internal（LD_LIBRARY_PATH=%r）' % lp)

    # 用清理后的环境启动一个真实系统程序，确认能跑
    r = subprocess.run(['fc-list', ':', 'family'], env=env,
                       stdout=subprocess.PIPE, stderr=subprocess.PIPE, timeout=30)
    assert r.returncode == 0, 'fc-list 在清理环境下失败: %s' % r.stderr[:200]
    fams = r.stdout.decode('utf-8', 'ignore')
    print('  OK    fc-list 正常（读到 %d 行字体）' % len(fams.splitlines()))
finally:
    os.environ.clear(); os.environ.update(saved)
    if hasattr(sys, '_MEIPASS'): del sys._MEIPASS
    if hasattr(sys, 'frozen'): del sys.frozen
print('  OK    运行期库隔离验证通过')
PY
if [ $? -ne 0 ]; then
    echo "!! print_env 库隔离验证失败，产物不可发布" >&2
    exit 1
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
cp 源码方式安装.sh 诊断.sh 修复打印.sh 校准版式.sh "$OUT/"
chmod +x "$OUT/启动GUI.sh" "$OUT/启动命令行.sh" "$OUT/源码方式安装.sh" "$OUT/诊断.sh" \
         "$OUT/修复打印.sh" "$OUT/校准版式.sh" "$OUT/投诉批办单管理系统"
cp 1.png "$OUT/1.png"
cp requirements-linux.txt "$OUT/requirements-linux.txt"
echo "${ARCH}" > "$OUT/构建架构.txt"

# ---------- 8. tar.gz ----------
cd "${SRC}/dist"
tar -czf "complaints-linux-${ARCH}.tar.gz" "投诉批办单管理系统"
sha256sum "complaints-linux-${ARCH}.tar.gz" > "complaints-linux-${ARCH}.tar.gz.sha256"
ls -lh "complaints-linux-${ARCH}.tar.gz"
echo "=== BUILD_DONE ${ARCH} ==="
