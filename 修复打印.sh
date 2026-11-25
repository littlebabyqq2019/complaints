#!/usr/bin/env bash
# =============================================================================
#  修复打印：定位并解决 LibreOffice「FT_Done_MM_Var 缺符号」问题
#
#  在麒麟 / 统信机器上运行：
#      bash 修复打印.sh          # 只诊断，不改任何东西
#      bash 修复打印.sh --fix    # 诊断并尝试自动修复（使用临时 LD_PRELOAD）
#
#  背景：
#     /lib/aarch64-linux-gnu/libcairo.so.2: undefined symbol: FT_Done_MM_Var
#
#     这个符号是 freetype 2.9+ 才提供的。麒麟 V10 自带 freetype 2.10.1，
#     按理说**有这个符号**。所以报错通常不是「系统库太旧」，而是
#     cairo 实际加载到了**另一个更旧的 libfreetype**（常见于某个软件
#     自带的 .so 被塞进了 LD_LIBRARY_PATH，或者 /usr/local/lib 下有一份）。
#
#     这种情况不需要联网、不需要重装，把正确的库优先加载即可。
# =============================================================================

set +u
set +e

RED=''; GRN=''; YEL=''; RST=''
if [ -t 1 ]; then
    RED=$'\033[31m'; GRN=$'\033[32m'; YEL=$'\033[33m'; RST=$'\033[0m'
fi

ok()   { echo "  ${GRN}OK${RST}   $*"; }
bad()  { echo "  ${RED}问题${RST} $*"; }
warn() { echo "  ${YEL}注意${RST} $*"; }
hr()   { echo "------------------------------------------------------------"; }

DO_FIX=0
[ "${1:-}" = "--fix" ] && DO_FIX=1

echo "============================================================"
echo " 修复打印 · LibreOffice 动态库诊断"
echo "============================================================"
echo "时间: $(date)"
echo "架构: $(uname -m)"
echo

SOFFICE=""
for c in /usr/bin/soffice /usr/lib/libreoffice/program/soffice \
         /opt/libreoffice/program/soffice "$(command -v soffice 2>/dev/null)"; do
    if [ -n "$c" ] && [ -x "$c" ]; then SOFFICE="$c"; break; fi
done

hr
echo "[1] LibreOffice 可执行文件"
if [ -n "$SOFFICE" ]; then
    ok "$SOFFICE"
    echo "      $(file "$SOFFICE" 2>/dev/null | cut -c1-100)"
else
    bad "未找到 soffice（LibreOffice 可能未安装或已损坏）"
fi

hr
echo "[2] 复现报错（这一步能直接看到真实错误）"
if [ -n "$SOFFICE" ]; then
    OUT=$("$SOFFICE" --version 2>&1 | head -5)
    if echo "$OUT" | grep -qi "symbol lookup error\|undefined symbol"; then
        bad "LibreOffice 启动失败："
        echo "$OUT" | sed 's/^/        /'
    else
        ok "可以启动：$(echo "$OUT" | head -1)"
    fi
else
    warn "跳过（没有 soffice）"
fi

hr
echo "[3] cairo 实际加载的 freetype 是哪一个"
CAIRO=""
for c in /lib/aarch64-linux-gnu/libcairo.so.2 \
         /usr/lib/aarch64-linux-gnu/libcairo.so.2 \
         /lib/x86_64-linux-gnu/libcairo.so.2 \
         /usr/lib/x86_64-linux-gnu/libcairo.so.2; do
    [ -f "$c" ] && CAIRO="$c" && break
done

if [ -n "$CAIRO" ]; then
    echo "      cairo: $CAIRO"
    FREETYPE_LOADED=$(ldd "$CAIRO" 2>/dev/null | grep -i freetype | awk '{print $3}')
    if [ -n "$FREETYPE_LOADED" ]; then
        echo "      -> 加载的 freetype: $FREETYPE_LOADED"
    else
        warn "ldd 未显示 freetype 依赖（可能静态链接或 ldd 不可用）"
    fi
else
    warn "未找到 libcairo.so.2"
fi

hr
echo "[4] 系统 freetype 是否真的含有 FT_Done_MM_Var"
SYS_FT=""
for c in /lib/aarch64-linux-gnu/libfreetype.so.6 \
         /usr/lib/aarch64-linux-gnu/libfreetype.so.6 \
         /lib/x86_64-linux-gnu/libfreetype.so.6 \
         /usr/lib/x86_64-linux-gnu/libfreetype.so.6; do
    [ -f "$c" ] && SYS_FT="$c" && break
done

HAS_SYMBOL=0
if [ -n "$SYS_FT" ]; then
    echo "      系统 freetype: $SYS_FT"
    echo "      版本: $(strings -a "$SYS_FT" 2>/dev/null | grep -m1 -E '^2\.[0-9]+\.[0-9]+$' || echo 未知)"
    if nm -D "$SYS_FT" 2>/dev/null | grep -q FT_Done_MM_Var; then
        ok "含有 FT_Done_MM_Var —— 系统库是好的，问题是加载错了库"
        HAS_SYMBOL=1
    elif strings -a "$SYS_FT" 2>/dev/null | grep -q FT_Done_MM_Var; then
        ok "含有 FT_Done_MM_Var（符号表不可读但字符串存在）"
        HAS_SYMBOL=1
    else
        bad "不含该符号 —— 系统 freetype 版本确实偏旧"
    fi
else
    bad "未找到系统 libfreetype.so.6"
fi

hr
echo "[5] 系统里是否有另一份更旧的 libfreetype（冲突源）"
echo "      搜索中，可能要十几秒..."
DUP=$(find /usr/local /opt /home /root /usr/lib -name "libfreetype.so*" 2>/dev/null \
      | grep -v "^/usr/lib/.*-linux-gnu/libfreetype" | head -20)
if [ -n "$DUP" ]; then
    bad "发现其它位置的 freetype（很可能就是它盖住了系统库）："
    echo "$DUP" | sed 's/^/        /'
else
    ok "没有发现重复的 freetype"
fi

hr
echo "[6] LD_LIBRARY_PATH / LD_PRELOAD 是否被污染"
echo "      LD_LIBRARY_PATH=${LD_LIBRARY_PATH:-<空>}"
echo "      LD_PRELOAD=${LD_PRELOAD:-<空>}"
if [ -n "${LD_LIBRARY_PATH:-}" ]; then
    warn "LD_LIBRARY_PATH 非空，可能导致加载到旧库"
    echo "$LD_LIBRARY_PATH" | tr ':' '\n' | sed 's/^/        /'
fi

hr
echo "[7] apt 缓存里是否有可离线安装的 deb"
CACHED=$(ls /var/cache/apt/archives/ 2>/dev/null | grep -Ei \
         "libreoffice-(writer|core)|libfreetype6|libcairo2" | head -10)
if [ -n "$CACHED" ]; then
    ok "找到已下载的 deb，可离线重装："
    echo "$CACHED" | sed 's/^/        /'
    echo "      命令： sudo dpkg -i /var/cache/apt/archives/<上面文件名>"
else
    warn "apt 缓存里没有相关 deb（需要联网下载）"
fi

hr
echo "[8] 网络（apt 失败的原因）"
DNS_OK=0
if getent hosts archive.kylinos.cn >/dev/null 2>&1; then
    DNS_OK=1
    ok "能解析 archive.kylinos.cn"
else
    bad "无法解析 archive.kylinos.cn（DNS 不通，apt 必然失败）"
    echo "        当前 DNS: $(grep -m3 '^nameserver' /etc/resolv.conf 2>/dev/null | tr '\n' ' ')"
fi
if [ "$DNS_OK" = "1" ]; then
    if curl -s -m 5 -o /dev/null -w "%{http_code}" https://archive.kylinos.cn/ 2>/dev/null | grep -qE '^[23]'; then
        ok "能访问源"
    else
        warn "能解析但访问不通（可能需代理或源暂时不可用）"
    fi
fi

# ---------------------------------------------------------------------------
# 结论与修复
# ---------------------------------------------------------------------------
hr
echo "结论"
echo

if [ "$HAS_SYMBOL" = "1" ] && [ -n "$SOFFICE" ]; then
    echo "  系统 freetype 是好的（含 FT_Done_MM_Var），"
    echo "  所以问题几乎肯定是『cairo 加载到了另一份旧 freetype』。"
    echo
    echo "  ${GRN}不需要联网、不需要重装，两种修法：${RST}"
    echo
    echo "  方式 A（临时，验证用）—— 用正确库启动 LibreOffice："
    echo "      LD_PRELOAD=$SYS_FT soffice --version"
    echo "      若这条能出版本号，说明修复方向正确。"
    echo
    echo "  方式 B（永久）—— 清掉多余的那份 freetype："
    echo "      1) 把上面 [5] 列出的旧 libfreetype 改名备份："
    echo "         sudo mv <那个路径> <那个路径>.bak"
    echo "      2) 刷新库缓存：sudo ldconfig"
    echo "      3) 验证：soffice --version"
    echo
    if [ "$DO_FIX" = "1" ] && [ -n "$FREETYPE_LOADED" ]; then
        echo "  ---- 已执行 --fix：用系统 freetype 生成启动包装脚本 ----"
        WRAP="/usr/local/bin/soffice-fixed"
        cat > /tmp/soffice-fixed <<EOF
#!/bin/sh
# 由「修复打印.sh --fix」生成：强制使用系统 freetype，绕开旧库
export LD_PRELOAD=$SYS_FT
exec $SOFFICE "\$@"
EOF
        chmod +x /tmp/soffice-fixed
        echo "      已生成：/tmp/soffice-fixed"
        echo "      测试： /tmp/soffice-fixed --version"
        echo
        echo "      若测试通过，安装为系统命令："
        echo "        sudo cp /tmp/soffice-fixed $WRAP"
        echo "      然后重新打开本程序即可正常转 PDF / 打印。"
    fi
elif [ "$HAS_SYMBOL" = "0" ]; then
    echo "  系统 freetype 确实缺 FT_Done_MM_Var，需要更新它。"
    echo "  但你的源不通，可行的替代："
    echo
    echo "  1) 先修 DNS（apt 报的『暂时不能解析域名』）："
    echo "       sudo sh -c 'echo \"nameserver 223.5.5.5\" > /etc/resolv.conf'"
    echo "     或检查是否被 NetworkManager 覆盖："
    echo "       nmcli dev show | grep DNS"
    echo "     然后： sudo apt update && sudo apt install --reinstall libfreetype6 libcairo2"
    echo
    echo "  2) 源地址可能已变更，换成国内可达镜像后再试。"
    echo
    echo "  3) ${GRN}不改系统也能打印：${RST}用 WPS 打开批办单，点 WPS 的打印按钮。"
    echo "     WPS 排版与 Word 一致，不需要转 PDF，也就不会变两页。"
else
    echo "  未能定位到 soffice，先确认 LibreOffice 是否真的装了："
    echo "      dpkg -l | grep -i libreoffice"
fi

if [ "$DNS_OK" = "0" ]; then
    echo
    echo "  ${YEL}另：apt 装不了是因为 DNS 解析失败（archive.kylinos.cn 不通），${RST}"
    echo "  ${YEL}不是软件包不存在。先修网络再谈重装。${RST}"
fi

echo
echo "============================================================"
echo " 如果以上都嫌麻烦，最省事的路（版式最保真）："
echo "   用 WPS 打开批办单 -> 点 WPS 的「打印」"
echo "   或在程序里：「打印」菜单 -> 「在 WPS 中打开当前文件并手动打印」"
echo "============================================================"
