#!/usr/bin/env bash
# =============================================================================
#  版式校准：按「这次打印与 Word 还差多少毫米」增量微调
#
#  背景：
#    程序删掉 docGrid 后要补偿几处间距。默认值是在 Windows 上用 Word COM
#    基准标定的，麒麟的 LibreOffice 渲染略有差异，需要现场微调。
#
#  用法：
#    bash 校准版式.sh                  # 用默认推测值（上沿+1 / 下沿+2 / 备注+5）
#    bash 校准版式.sh 0 -1 -2          # 自定义三个偏差（毫米）
#    bash 校准版式.sh 0 0 2            # 只调备注
#
#  三个参数（**正数 = PDF 比 Word 更靠下；负数 = PDF 比 Word 更靠上**）：
#    1) 表格上沿偏差
#    2) 表格下沿偏差
#    3) 备注文字偏差
#
#  ⚠ 它是**增量**的：会读取 AppData/config.json 里当前的 layout_* 值，
#    在此基础上按你这次的测量结果继续修正。所以：
#      · 每次打印后重新量一次，把「还差多少」填进去再跑
#      · 不要重复填同一组数（那会叠加两次）
#
#  例：打印后发现表格上沿与 Word 齐平、下沿比 Word 高 1mm、备注比 Word 高 2mm
#      -> 上沿 0、下沿 -1、备注 -2
#      bash 校准版式.sh 0 -1 -2
# =============================================================================

set +u
set +e

TOP_MM="${1:-1}"
BOT_MM="${2:-2}"
NOTE_MM="${3:-5}"

SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
cd "$SCRIPT_DIR" 2>/dev/null

PY=""
# 依次尝试，并验证解释器真的能执行（某些系统 python3 只是占位程序）
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
    echo "可手工按此公式换算（1mm = 56.69 twips；正数=PDF 偏下）：" >&2
    echo "  layout_spacer_line = 当前值 - 上沿偏差*56.69" >&2
    echo "  layout_row_height  = 当前值 - (下沿-上沿)*56.69/10" >&2
    echo "  layout_note_gap    = 当前值 - (备注-下沿)*56.69" >&2
    exit 1
fi

"$PY" - "$TOP_MM" "$BOT_MM" "$NOTE_MM" <<'PYEOF'
import json
import os
import sys

MM2TW = (72.0 / 25.4) * 20.0   # 1mm = 56.69 twips

try:
    top_mm = float(sys.argv[1]); bot_mm = float(sys.argv[2]); note_mm = float(sys.argv[3])
except Exception:
    top_mm, bot_mm, note_mm = 1.0, 2.0, 5.0

# 默认值（与 main.py 的 _DEFAULT_* 一致）
DEF = {
    'layout_title_before': 100,
    'layout_spacer_line': 430,
    'layout_row_height': 640,
    'layout_note_gap': 0,
}

cfgp = os.path.join("AppData", "config.json") if os.path.isdir("AppData") else None
cfg = {}
if cfgp and os.path.isfile(cfgp):
    try:
        with open(cfgp, "r", encoding="utf-8-sig") as f:
            cfg = json.load(f)
    except Exception:
        cfg = {}

def cur(key):
    try:
        v = cfg.get(key)
        return int(float(v)) if v is not None else DEF[key]
    except Exception:
        return DEF[key]

# 当前生效值（增量基准）
cur_title = cur('layout_title_before')
cur_spacer = cur('layout_spacer_line')
cur_row = cur('layout_row_height')
cur_note = max(0, cur('layout_note_gap'))

top_tw = top_mm * MM2TW
bot_tw = bot_mm * MM2TW
note_tw = note_mm * MM2TW
table_delta = bot_tw - top_tw          # 表格本身的高度偏差
note_delta = note_tw - bot_tw          # 表格底到备注的间距偏差

# 增量公式：新值 = 当前值 - 偏差
new_title = cur_title
new_spacer = int(round(cur_spacer - top_tw))
new_row = int(round(cur_row - table_delta / 10.0))
new_note = int(round(cur_note - note_delta))

# 段前距不能为负：需要把备注往上提超过现有段前距时，
# 改为压缩行高来带动备注上移（同时保证 new_note 不小于 0）
if new_note < 0:
    extra = -new_note
    new_row -= int(round(extra / 10.0))
    new_note = 0

# 安全下限，避免算飞
new_spacer = max(80, new_spacer)
new_row = max(300, new_row)

print("=" * 64)
print(" 版式校准（增量）")
print("=" * 64)
print(f" 本次测量偏差（正数=PDF 比 Word 靠下）：")
print(f"   表格上沿 {top_mm:+.1f}mm   下沿 {bot_mm:+.1f}mm   备注 {note_mm:+.1f}mm")
print(f" 1mm = {MM2TW:.1f} twips")
print()
print(f" 当前配置： spacer={cur_spacer}  row={cur_row}  note={cur_note}")
print(f" 本次修正：")
print(f"   表格上方间距 {top_tw:+8.1f} twips")
print(f"   表格高度     {table_delta:+8.1f} twips（10 行分摊 {table_delta/10:+.1f}/行）")
print(f"   表格→备注     {note_delta:+8.1f} twips")
print()
print(" 写入 AppData/config.json：")
print(f'   "layout_title_before": {new_title},')
print(f'   "layout_spacer_line" : {new_spacer},')
print(f'   "layout_row_height"  : {new_row},')
print(f'   "layout_note_gap"    : {new_note}')
print()

if cfgp:
    cfg["layout_title_before"] = new_title
    cfg["layout_spacer_line"] = new_spacer
    cfg["layout_row_height"] = new_row
    cfg["layout_note_gap"] = new_note
    try:
        with open(cfgp, "w", encoding="utf-8") as f:
            json.dump(cfg, f, ensure_ascii=False, indent=2)
        print(f" 已写入 {cfgp}，重启程序生效")
    except Exception as e:
        print(f" 写入失败：{e}；请手工把上面的键值加到 {cfgp}")
else:
    print(" 未找到 AppData 目录，请手工把上面的键值加到 AppData/config.json")

print("=" * 64)
print(" 提示：三个参数填的是「这次还差多少」，不是累计总量。")
print("       再打印一次、重新量，把新的差值填进来再跑即可收敛。")
PYEOF
