#!/usr/bin/env bash
# =============================================================================
#  版式校准：把「打印 PDF 比 Word 偏了多少毫米」换算成需设置的参数
#
#  背景：
#    程序删掉 docGrid 后要补偿三处间距（标题段前距、标题与表格之间空行、
#    表格行高）。这些值是在 Windows 上用 Word COM 基准标定的，
#    麒麟的 LibreOffice 6.4.7 渲染略有差异，现场会整体偏下。
#
#  用法：
#    bash 校准版式.sh            # 用默认现场值（上沿+1mm / 下沿+2mm / 备注+5mm）
#    bash 校准版式.sh 1 2 5      # 自定义三个偏移（毫米，正数=偏下）
#    bash 校准版式.sh 0.5 1 3
#
#  它会打印应写入 AppData/config.json 的 layout_* 值，并可选直接写入。
# =============================================================================

set +u
set +e

TOP_MM="${1:-1}"      # 表格上沿偏移
BOT_MM="${2:-2}"      # 表格下沿偏移
NOTE_MM="${3:-5}"     # 备注文字偏移

SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
cd "$SCRIPT_DIR" 2>/dev/null

PY=""
# 依次尝试，并且**验证解释器真的能执行**（
# 某些系统上 python3 只是占位程序，能查到路径但执行不了）
for cand in "./.venv/bin/python" "$(command -v python3 2>/dev/null)" \
            "$(command -v python 2>/dev/null)"; do
    if [ -n "$cand" ] && [ -x "$cand" ]; then
        if "$cand" -c "print(1)" >/dev/null 2>&1; then
            PY="$cand"
            break
        fi
    fi
done

if [ -z "$PY" ]; then
    echo "未找到可用的 Python，无法自动计算。" >&2
    echo "可手工按此公式换算（1mm = 56.69 twips）：" >&2
    echo "  layout_spacer_line = 430 - 表格上沿偏移(mm)*56.69" >&2
    echo "  layout_row_height  = 640 - (下沿-上沿)(mm)*56.69/10" >&2
    echo "  layout_note_gap    = 0   - (备注-下沿)(mm)*56.69" >&2
else
    "$PY" - "$TOP_MM" "$BOT_MM" "$NOTE_MM" <<'PYEOF'
import json
import os
import sys

MM2TW = (72.0 / 25.4) * 20.0   # 1mm = 56.69 twips

try:
    top_mm = float(sys.argv[1]); bot_mm = float(sys.argv[2]); note_mm = float(sys.argv[3])
except Exception:
    top_mm, bot_mm, note_mm = 1.0, 2.0, 5.0

# 当前默认值（与 main.py 的 _DEFAULT_* 一致）
CUR_TITLE = 100      # 标题段前距
CUR_SPACER = 430     # 标题与表格之间的空行行高
CUR_ROW = 640        # 表格行高
CUR_NOTE = 0         # 表格后备注段前距

top_tw = top_mm * MM2TW
bot_tw = bot_mm * MM2TW
note_tw = note_mm * MM2TW

# 推导：
#   表格上沿偏下 top  -> 标题与表格之间的间距要减少 top
#   表格下沿额外偏移 = bot - top  -> 表格本身高了，行高要减少 (bot-top)/10 行
#   备注额外偏移 = note - bot      -> 表格与备注之间要减少 (note-bot)
table_taller = bot_tw - top_tw
new_spacer = int(round(CUR_SPACER - top_tw))
new_row = int(round(CUR_ROW - table_taller / 10.0))
new_note = int(round(CUR_NOTE - (note_tw - bot_tw)))
new_title = CUR_TITLE  # 顶部留白未反馈偏移，保持

print("=" * 58)
print(" 版式校准计算")
print("=" * 58)
print(f" 实测偏移：表格上沿 +{top_mm}mm / 下沿 +{bot_mm}mm / 备注 +{note_mm}mm")
print(f" 换算系数：1mm = {MM2TW:.1f} twips")
print()
print(" 各段偏移（twips）：")
print(f"   表格上沿      {top_tw:7.1f}")
print(f"   表格本身增高  {table_taller:7.1f}  （10 行分摊 {table_taller/10:.1f}/行）")
print(f"   表格→备注     {note_tw - bot_tw:7.1f}")
print()
print(" 建议写入 AppData/config.json 的值：")
print(f'   "layout_title_before": {new_title},')
print(f'   "layout_spacer_line" : {new_spacer},')
print(f'   "layout_row_height"  : {new_row},')
print(f'   "layout_note_gap"    : {new_note}')
print()
if new_spacer < 0 or new_row < 200:
    print(" ⚠ 计算结果偏小，可能偏移量填得过大，请复核实测值。")
print("=" * 58)

# 直接写入（存在旧配置则合并）
targets = []
if os.path.isdir("AppData"):
    targets.append(os.path.join("AppData", "config.json"))
cfgp = (targets[0] if targets else None)
if cfgp:
    cfg = {}
    if os.path.isfile(cfgp):
        try:
            with open(cfgp, "r", encoding="utf-8-sig") as f:
                cfg = json.load(f)
        except Exception:
            cfg = {}
    cfg["layout_title_before"] = new_title
    cfg["layout_spacer_line"] = new_spacer
    cfg["layout_row_height"] = new_row
    cfg["layout_note_gap"] = new_note
    try:
        with open(cfgp, "w", encoding="utf-8") as f:
            json.dump(cfg, f, ensure_ascii=False, indent=2)
        print(f" 已写入 {cfgp}（重启程序生效）")
    except Exception as e:
        print(f" 写入失败：{e}；请手工把上面的键值加到 {cfgp}")
else:
    print(" 未找到 AppData 目录，请手工把上面的键值加到 AppData/config.json")
PYEOF
fi

echo
echo "提示：这些参数只影响「删网格后的间距补偿」，不改变纸张与页边距。"
echo "      调好后重新生成一份批办单并打印，再目测一次是否对齐。"
