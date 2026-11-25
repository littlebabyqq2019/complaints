#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
字体度量兼容层 —— 解决「Word 转 PDF 后一页变两页」。

根因
----
批办单模板用的是 Windows 字体：仿宋_GB2312、方正小标宋简体。
麒麟 / 统信上没有这些字体，LibreOffice 会做字体替换：
   仿宋_GB2312  ->  Noto Serif CJK SC（或别的）
替换字体的**度量不同**（字面高度、行距、字符宽度），而模板页边距很紧
（A4 上下 1440 twips、左右 1800 twips），行高或字宽一变大，最后一行就
被挤到第 2 页 —— 这就是「原来一页，转 PDF 变两页」。

做法
----
给 fontconfig 写一份**用户级**别名，把 Windows 字体名映射到度量最接近的
本地字体，并priority 靠前，让 LibreOffice 优先选它：

   仿宋_GB2312 / 仿宋  ->  FangSong / AR PL UMing CN
   宋体 / SimSun       ->  AR PL UMing CN / Noto Serif CJK SC
   方正小标宋简体      ->  Noto Serif CJK SC
   黑体 / 微软雅黑     ->  Noto Sans CJK SC

同时提供「页数校验」：转换前后比对页数，不一致就把原因报出来，
不再让用户拿到错的打印结果还不知道。

注意：这是「尽量逼近」，不是「完全等同」。要做到像素级一致，只有两条路：
  1. 系统装上真正的仿宋_GB2312（有版权的商业字体，需自行购买授权）
  2. 用 Word/WPS 的 GUI 直接打印（不在本模块范围）
"""

import os
import shutil
import subprocess

# Windows 字体名 -> 本地候选（按度量接近程度排序，第一个优先）
_FONT_ALIASES = {
    "仿宋_GB2312": ["FangSong", "AR PL UMing CN", "Noto Serif CJK SC"],
    "仿宋": ["FangSong", "AR PL UMing CN", "Noto Serif CJK SC"],
    "FangSong_GB2312": ["FangSong", "AR PL UMing CN"],
    "宋体": ["AR PL UMing CN", "Noto Serif CJK SC", "AR PL SungtiL GB"],
    "SimSun": ["AR PL UMing CN", "Noto Serif CJK SC"],
    "方正小标宋简体": ["Noto Serif CJK SC", "AR PL UMing CN"],
    "黑体": ["Noto Sans CJK SC", "WenQuanYi Micro Hei"],
    "SimHei": ["Noto Sans CJK SC", "WenQuanYi Micro Hei"],
    "微软雅黑": ["Noto Sans CJK SC", "WenQuanYi Micro Hei"],
    "Microsoft YaHei": ["Noto Sans CJK SC", "WenQuanYi Micro Hei"],
}

_CONF_NAME = "99-complaints-font-alias.conf"


def _clean_env():
    """外部字体命令用的干净环境（剔除应用自带库路径）。"""
    try:
        import print_env
        return print_env.clean_env()
    except Exception:
        return None


def _installed_families():
    """
    读取 fontconfig 的已安装字体家族（小写比较用）。

    必须用干净环境：打包版若让 fc-list 加载到 _internal 里的旧
    libfontconfig，会解析不出系统字体（或解析结果异常），
    从而误判「缺字体」并把版式调偏。
    """
    if not shutil.which("fc-list"):
        return ""
    try:
        r = subprocess.run(["fc-list", ":", "family"], env=_clean_env(),
                           stdout=subprocess.PIPE, stderr=subprocess.DEVNULL,
                           timeout=8)
        return (r.stdout or b"").decode("utf-8", "ignore")
    except Exception:
        return ""


def _font_dirs():
    """fontconfig 用户配置目录（按 XDG 规则）。"""
    home = os.path.expanduser("~")
    dirs = []
    xdg = os.environ.get("XDG_CONFIG_HOME")
    if xdg:
        dirs.append(os.path.join(xdg, "fontconfig", "conf.d"))
    dirs.append(os.path.join(home, ".config", "fontconfig", "conf.d"))
    return dirs


def missing_fonts(families_text=None):
    """返回文档会用到、但本机没有的 Windows 字体名列表。"""
    text = families_text if families_text is not None else _installed_families()
    if not text:
        return []
    low = text.lower()
    missing = []
    for name in _FONT_ALIASES:
        if name.lower() not in low:
            missing.append(name)
    return missing


def _pick_local(candidates, families_text):
    low = (families_text or "").lower()
    for c in candidates:
        if c and c.lower() in low:
            return c
    return None


def build_alias_conf():
    """
    生成 fontconfig 别名配置文本。只映射「本机确实有的」本地字体，
    避免又指向一个不存在的字体。
    """
    text = _installed_families()
    if not text:
        return None

    lines = ['<?xml version="1.0"?>',
             '<!DOCTYPE fontconfig SYSTEM "fonts.dtd">',
             '<fontconfig>',
             '  <!-- 由「投诉批办单管理系统」自动生成：',
             '       Windows 字体 -> 本地度量接近字体，减少 Word->PDF 排版漂移 -->']
    mapped = 0
    for win_name, cands in _FONT_ALIASES.items():
        local = _pick_local(cands, text)
        if not local:
            continue
        # prefer：让 LibO 优先用 local
        lines += ['  <alias>',
                  '    <family>%s</family>' % win_name,
                  '    <prefer><family>%s</family></prefer>' % local,
                  '  </alias>']
        mapped += 1
    lines.append('</fontconfig>')
    if not mapped:
        return None
    return "\n".join(lines) + "\n", mapped


def ensure_font_aliases(force=False):
    """
    确保用户级 fontconfig 别名已就位。返回 (是否已写入, 说明)。

    只写用户目录（~/.config/fontconfig/conf.d），不需要 root，
    不污染系统；同一内容只写一次。
    """
    if not shutil.which("fc-list"):
        return False, "未找到 fc-list（fontconfig），跳过字体别名"

    built = build_alias_conf()
    if not built:
        return False, "本机找不到可用的中文字体，无法建立别名"
    content, mapped = built

    for d in _font_dirs():
        target = os.path.join(d, _CONF_NAME)
        try:
            if not os.path.isdir(d):
                os.makedirs(d, exist_ok=True)
            old = None
            if os.path.isfile(target):
                with open(target, "r", encoding="utf-8") as f:
                    old = f.read()
            if old == content and not force:
                return True, "字体别名已存在（%d 条）" % mapped
            with open(target, "w", encoding="utf-8") as f:
                f.write(content)
            # 刷新缓存，让 LibreOffice 立刻生效
            if shutil.which("fc-cache"):
                subprocess.run(["fc-cache", "-f", d], env=_clean_env(),
                               stdout=subprocess.DEVNULL,
                               stderr=subprocess.DEVNULL, timeout=30)
            return True, "已写入字体别名 %d 条：%s" % (mapped, target)
        except Exception as e:
            last = e
            continue
    return False, "写入字体别名失败：%s" % (locals().get("last"),)


def docx_fonts(docx_path):
    """从 docx 里读出声明用到的字体名（用于判断是否需要别名字体）。"""
    import re
    import zipfile
    fonts = set()
    try:
        with zipfile.ZipFile(docx_path) as z:
            for part in ("word/document.xml", "word/styles.xml"):
                try:
                    xml = z.read(part).decode("utf-8", "ignore")
                except KeyError:
                    continue
                for m in re.finditer(r'w:(?:eastAsia|ascii|hAnsi)="([^"]+)"', xml):
                    fonts.add(m.group(1))
    except Exception:
        pass
    # 过滤非字体名的值（语言标记等）
    return {f for f in fonts if f and not f.startswith("zh-") and len(f) > 1}


def docx_missing_fonts(docx_path):
    """该 docx 用到、但本机缺的字体。"""
    need = docx_fonts(docx_path)
    if not need:
        return []
    text = _installed_families()
    low = (text or "").lower()
    return sorted(f for f in need if f.lower() not in low)


def pdf_page_count(pdf_path):
    """读 PDF 页数；失败返回 None。"""
    try:
        import re
        with open(pdf_path, "rb") as f:
            data = f.read()
        # /Type /Pages 的 /Count 最可靠
        counts = [int(x) for x in re.findall(rb"/Count\s+(\d+)", data)]
        if counts:
            return max(counts)
        # 兜底：数 /Type /Page
        return len(re.findall(rb"/Type\s*/Page[^s]", data)) or None
    except Exception:
        return None


def docx_page_estimate(docx_path):
    """
    估算 docx 的页数。docx 本身不存页数（由排版决定），
    只能看 app.xml 里的 <Pages>（Word 保存时写入）。取不到返回 None。
    """
    import re
    import zipfile
    try:
        with zipfile.ZipFile(docx_path) as z:
            xml = z.read("docProps/app.xml").decode("utf-8", "ignore")
        m = re.search(r"<Pages>(\d+)</Pages>", xml)
        if m:
            return int(m.group(1))
    except Exception:
        pass
    return None


def diagnose():
    """一行状态，给日志/诊断脚本用。"""
    text = _installed_families()
    if not text:
        return "字体: 无法读取 fontconfig"
    miss = missing_fonts(text)
    if not miss:
        return "字体: Windows 字体齐全（无需别名）"
    return "字体: 缺少 %s（将使用度量接近的本地字体替代）" % "、".join(miss[:4])
