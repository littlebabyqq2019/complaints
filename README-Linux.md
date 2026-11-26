# 投诉批办单管理系统 · 统信 UOS / 麒麟 Kylin V10 版

本目录是 Windows 版「投诉批办单管理系统」的国产 Linux 移植版。GUI 布局、菜单、按钮、
字段、Word 模板、编号规则、数据库结构与全部业务流程保持原样，仅把 Windows 专属调用
替换成国产系统上的等价实现。

## 一、运行方式

### 1. 直接运行 GitHub Actions 打包出的可执行文件（推荐）

在 [Releases](https://github.com/littlebabyqq2019/complaints/releases) 下载**与 CPU 架构对应的包**。
先在终端执行 `uname -m` 判断架构：

| `uname -m` 输出 | 下载文件 | 适用 CPU |
| --- | --- | --- |
| `x86_64` | `complaints-linux-x86_64.tar.gz` | Intel / AMD、兆芯、海光 |
| `aarch64` | `complaints-linux-aarch64.tar.gz` | 飞腾、鲲鹏、麒麟芯、瑞芯微等 arm64 国产 CPU |

两个包**不能混用**：包内 ELF 可执行文件和 `_internal` 下的 `.so` 是对应架构的原生机器码。
GitHub Actions 用 `ubuntu-22.04` 与 `ubuntu-22.04-arm` 两个原生 runner 分别构建。

```bash
tar -xzf complaints-linux-aarch64.tar.gz
cd 投诉批办单管理系统
./启动GUI.sh
```

或直接双击 `投诉批办单管理系统` / `启动GUI.sh`。命令行批处理模式：`./启动命令行.sh`。

### 2. 源码运行（任意架构，含龙芯 LoongArch / 申威）

源码为纯 Python，与 CPU 架构无关：

```bash
sudo apt install python3 python3-tk python3-pip python3-venv \
                 fonts-noto-cjk fonts-wqy-microhei \
                 libreoffice-writer cups-client xdg-utils
python3 -m pip install -r requirements-linux.txt
./启动GUI.sh
```

如需在目标架构上生成原生可执行文件（构建基线 glibc 2.35+）：

```bash
python3 -m pip install -r requirements-linux.txt
python3 -m PyInstaller complaints-linux.spec --noconfirm
chmod +x "dist/投诉批办单管理系统/投诉批办单管理系统"
```

## 二、系统要求

| 项目 | 要求 |
| --- | --- |
| 操作系统 | 统信 UOS 20+ / 麒麟 Kylin V10 SP1+，x86_64 或 aarch64 |
| glibc | ≥ 2.35（Ubuntu 22.04 构建基线）；更老系统请用源码方式 |
| 架构 | 每个架构单独原生打包，见上表 |
| 图形 | X11 / Wayland + Tk 8.6（python3-tk） |
| 中文字体 | fonts-noto-cjk / 文泉驿 / AR PL UMing，缺失时界面自动回退 |
| Word 转换与打印 | libreoffice-writer（或 WPS）+ cups-client 的 `lp` 命令 |
| 打开文档/目录 | xdg-utils 的 `xdg-open`（或 gio / WPS / LibreOffice 关联） |

## 三、Windows 专属调用的等价替换

| Windows 原实现 | 统信 / 麒麟 实现 | 代码位置 |
| --- | --- | --- |
| `win32com.client` 读 `.doc` | LibreOffice/WPS 无头转 `.docx` 后由 python-docx 读表格；`.docx` 改名 `.doc` 的文件直接读 | `main.py` `DOCExtractor` |
| `win32api.ShellExecute(..., "print", ...)` | LibreOffice/WPS 转 PDF → CUPS `lp` 提交打印队列 | `platform_compat.print_file` |
| `docx2pdf` Word→PDF | `soffice --headless --convert-to pdf` / WPS / unoconv | `platform_compat.word_to_pdf` |
| `os.startfile` / `cmd /c start` | `xdg-open` / `gio open` | `platform_compat.open_path` |
| `微软雅黑` / `仿宋_GB2312` / `宋体` | fontconfig 映射到 Noto Sans/Serif CJK、FangSong、AR PL UMing 等 | `platform_compat.ui_font` |
| `.ico` 窗口图标 | 同内容 `1.png` + `iconphoto` | `gui_app.ComplaintManagementGUI.__init__` |
| `root.state('zoomed')` | `attributes('-zoomed', True)` | `gui_app._maximize_window` |
| `*.pdf;*.docx;*.doc` 文件对话框模式 | Tk 元组形式多模式 | `gui_app.add_source_files` |
| `AppData\\` 反斜杠路径 | 自动归一化 | `gui_app.init_database` |

`platform_compat.py` 在 Windows 上仍走原逻辑（Word COM / ShellExecute / docx2pdf /
原字体名），所以同一份代码在 Windows 上行为与旧版完全一致。

## 三之二、ttkbootstrap 1.x / 2.x 兼容（`ttk_compat.py`）

ttkbootstrap 2.0 把分割面板类从 `PanedWindow` 改名为 `Panedwindow`（与
`tkinter.ttk` 一致的拼写），1.x 的驼峰别名不再导出。旧代码 `ttk.PanedWindow(...)`
在 2.x 上直接抛 `AttributeError`，主窗口建到 `create_widgets()` 就崩，表现为
「程序启动失败」弹窗（麒麟 V10 aarch64 实测 ttkbootstrap 2.2.7）。

| 位置 | 处理 |
| --- | --- |
| `ttk_compat.py` | 导入时把缺失拼写补回 `ttkbootstrap` 模块（`PanedWindow`⇄`Panedwindow`、`LabelFrame`⇄`Labelframe`），并给出 `verify_widgets()` 启动自检 |
| `gui_app.py` / `rule_management_ui.py` | 改为 `from ttk_compat import ttk`，业务代码 `.PanedWindow(...)` 不变 |
| `gui_app.main()` | 启动时记录 ttkbootstrap 版本与控件校验结果；缺控件直接给出 `pip install 'ttkbootstrap>=1.10.1,<3'` 的可执行提示 |
| 主题 | 先用两代都存在的 `litera` 建窗口，再 `theme_use(已保存主题)`；名字失效只影响配色，不再导致启动失败 |
| `requirements*.txt` / CI | 依赖上限锁 `<3`，打包前用 `ttk_compat.verify_widgets()` 断言 |

### GUI 冒烟测试（`ttk_gui_selftest.py`）

只做 `import` / `py_compile` 抓不到这类崩溃 —— 现场是 import 全过、窗口建到一半才炸。
流水线现在会在 Xvfb 里真把界面搭一遍：

| 阶段 | 命令 | 验证对象 |
| --- | --- | --- |
| 源码冒烟 | `xvfb-run -a python3 ttk_gui_selftest.py` | 源码树 |
| 产物冒烟 | `./投诉批办单管理系统 --selftest` | Release 里那个 ELF（含 `_internal/`） |
| 现场自查 | `./诊断.sh` 或 `./投诉批办单管理系统 --selftest` | 目标机器 |

自检内容：控件别名、WordML 工单提取、三栏 `PanedWindow` 布局、规则管理窗口、
编辑页底部 7 个按钮是否固定在可视区、滚轮事件绑定、10 个工具对话框（其中二维码
设置额外断言「可拖动调整 + 保存/取消贴底」）。
结论写入 `AppData/ttk_selftest.log`（`console=False` 时 stdout 为空，只能靠这个文件取证），
并以退出码返回。故意删掉 `PanedWindow` 别名时自检会 fail 并打印
`缺少控件: ['PanedWindow']`，确认它真能拦住这个回归。

## 三之三、麒麟现场三个问题的修复（2026-11-13）

| 现象 | 根因 | 修复 |
| --- | --- | --- |
| PDF 正常，Word 提示「提取数据失败」 | 「热线工单」导出的是 **Word 2003 XML（WordML）单文件**，扩展名却叫 `.docx`；`python-docx` 只认 OOXML 的 ZIP 包，直接 `PackageNotFoundError`。PDF 走 pdfplumber 所以不受影响 | `main.py` 新增 `WordMLExtractor`，按文件头（不看扩展名）识别并用 `ElementTree` 解析 `w:tbl`；`WordExtractor` / `DOCExtractor` 优先走它；输出兼容生成器的 `来文字号 / 文件标题 / 标题 / 来源 / 姓名` |
| 二维码设置窗口不能拖大、最下两个按钮看不见 | 固定 `500x580` + `resizable(False, False)`，按钮还 pack 在内容末尾，屏幕稍矮或缩放后就被挤出可视区 | `ui_layout.fit_dialog_to_screen()` 按屏幕自适应、允许拖动、居中；按钮条先 `pack(side=BOTTOM)` 固定；主体放进 `make_scrollable_frame()` 可滚动 |
| 批办单编辑页最下一排按钮被裁掉一半 | 表单放在 Canvas 里却**完全没绑滚轮**（Linux/X11 是 Button-4/5），内容超高时用户滚不动；按钮本身也在滚动内容末尾 | 按钮条移到 `edit_action_bar`，固定在 tab 底部、分两行排布；`ui_layout.bind_mousewheel()` 同时绑 `MouseWheel` 与 `Button-4/5`，并用 Toplevel 级 `bind_all` 兜住动态创建的输入框（按事件来源过滤，不干扰其它滚动区） |

`ui_layout.py` 是纯布局工具，不含业务逻辑；`ttk_gui_selftest.py` 对上述三项都有断言，
打包前跑一遍，回归会直接让流水线变红。

## 三之四、直接调用 WPS 提取数据（`wps_compat.py`）

统信 / 麒麟基本都预装 WPS。老二进制 `.doc`（OLE2）纯 Python 读不了，
必须靠 WPS（或 LibreOffice）转成 `.docx` 才能取表格。原实现有三处硬伤：
`doc_to_docx()` 只走 LibreOffice；`word_to_pdf()` 把 WPS 当 LO 用
（`wps --headless --convert-to pdf`，多数 WPS Linux 版不认）；没有路径探测、
超时和 HOME 隔离。

现在的分层策略（**不是所有 Word 都要 WPS**）：

| 文件真实格式 | 处理方式 | 是否依赖 WPS |
| --- | --- | --- |
| WordML（`<?xml ... w:wordDocument>`，现场热线工单） | `WordMLExtractor` 纯 Python 解析 | 否 |
| OOXML（标准 .docx） | `python-docx` | 否 |
| OLE2 老 `.doc` | **WPS 优先** → LibreOffice 兜底 → python-docx 读表 | 是（之一即可） |

`wps_compat.py` 提供的能力：

- **路径探测**：PATH、统信/麒麟商店版、平铺目录、龙芯/申威移植版，
  支持 `REA_WPS_COMMAND` / `WPS_COMMAND` 显式指定绝对路径
- **多参数候选**：不同 WPS 版本命令行差异大，逐一试 `--convert-to`、
  `-o`、`-f`、`-convert-to` 等写法，格式名也试 `docx:MS Word 2007 XML` / `docx`
- **结果驱动**：不看返回码，只认「目标文件真的生成且非空」；每轮清掉旧产物，
  避免超时后误读上一次的残留文件
- **环境隔离**：临时 HOME + `QT_QPA_PLATFORM=offscreen` + 剥离 DISPLAY，
  避免首次运行弹协议/首页、污染用户配置；每次转换有 45 秒超时并杀掉子进程
- **配置开关**：`COMPLAINTS_WPS_PREFERENCE=auto|wps|off`，或
  `AppData/config.json` 里 `{"wps_preference": "..."}`：
  `auto` 检测到就用（默认）、`wps` 强制、`off` 只用 LibreOffice

对外入口：`find_wps()` / `wps_available()` / `convert_with_wps(src, fmt)` /
`doc_to_docx_with_wps()` / `word_to_pdf_with_wps()` /
`extract_table_with_wps(path, reader=None)` / `diagnose()`。
`main.prefer_wps()`、`main.wps_status()` 供启动日志与 `诊断.sh` 使用。

现场验证：

```bash
./诊断.sh                       # 打印 WPS / LibreOffice 检测结果
COMPLAINTS_WPS_PREFERENCE=wps ./启动GUI.sh   # 强制只走 WPS 验证通道
```

## 三之五、与 Windows 版生成结果对齐（2026-11-20）

逐行对比 Windows 源码（`main.py` 的 `generate_from_pdf` / `generate_from_doc` /
`generate_from_word`、`set_cell_text_keep_format`）后修正四处：

| 现场现象 | 根因 | 修复 |
| --- | --- | --- |
| 来文单位不是 12345 | 麒麟的热线工单是 WordML 扩展名 `.docx`，走 `generate_from_word`，填了 `word_data['姓名']`（人名字）；Windows 的 pdf/doc 分支固定填 `"12345"` | `generate_from_word` 改为固定 `"12345"` |
| 收文途径不是信息中心 | 同上，原来填 `word_data['来源']`（渠道名） | `generate_from_word` 改为固定 `"信息中心"` |
| 来文字号为空 | `generate_from_word` 根本没写行[1]列[2]；Windows 的 doc 分支取 `table.Cell(1,2)`（第一行第二列） | 补写行[1]列[2]，取 `来文字号`（WordML 已提取工单编号），并按字符数设字号（与 pdf/doc 一致） |
| 填充内容字体是 Noto Serif CJK SC | Linux 的 `set_cell_text_keep_format` 用了 `ui_font('仿宋_GB2312')`，在麒麟被映射成 Noto Serif 写进 docx；Windows 直接写字体名字符串 | 写入 docx 的字体固定为 `'仿宋_GB2312'`（含 eastAsia/ascii/hAnsi），`ui_font` 只留给 GUI 显示 |
| 按钮图标有的显示不出来 | Tk 8.6 在麒麟用轮廓字体，不含 emoji 字形；少数符号恰好有，显得不一致 | 新增 `emoji_compat.py`：Windows 原样保留；Linux 无轮廓 emoji 字体时剥掉控件文本里的 emoji，只留中文标签，避免方框 |

`COMPLAINTS_EMOJI=auto|on|off` 可覆盖 emoji 策略（默认 auto）。
启动日志记录 `emoji=auto/剥掉（避免方框）`。

## 三之六、第二轮现场问题（2026-11-23）

| 现场现象 | 根因 | 修复 |
| --- | --- | --- |
| PDF 不填拟办意见/办理情况，Word 正常 | pdfplumber 的空单元格返回 `None`，`PDFExtractor` 直接把它当作「文件标题」传出；`RuleEngine.match_filling_rule` 内部做字符串 `in` 判断时抛 `TypeError: argument of type 'NoneType' is not iterable`，被外层 `except` 吞掉，于是规则没匹配上、两栏留空。Word 分支的 `标题` 一定非空，所以不受影响 | 三处归一化：`PDFExtractor` 把 `来文字号/文件标题` 统一成 `str.strip()`；`RuleEngine.match_filling_rule` 入口把 `filename/file_title` 的 `None` 转空串；`set_cell_text_keep_format` 对 `text=None` 兜底 |
| 按钮图标还是有的不显示 | emoji 探测把 `DejaVu` / `FreeSerif` 当成「有 emoji 字体」，但它们只覆盖极少数符号（⚙ ✓ ➡），覆盖不了 🖨 💾 📝 📋，于是判定为「可显示」却仍是方框 | 探测只认真正的轮廓 emoji 字体（Symbola / EmojiOne / Twemoji / Noto Emoji 等）；同时把补丁从「仅创建时」扩展到 `configure`/`config`，覆盖 `label.config(text="📋 …")` 这类**建好之后**更新文字的路径（收文编号就是这么更新的） |
| 右上角「下一个收文编号」末位数字显示不全 | 宽度问题：左侧「📝 请选择文件进行编辑」太长，`pack` 时把右侧 Label 挤没 | 先 `pack` 右侧让编号优先占位，左侧 `fill=X + expand` 吃掉剩余并按需截断；两侧分别设 `anchor='e'` / `'w'`，宽度不足时只裁左侧 |
| 左上角「待处理文件」的「件」字只显示一半 | 同样是宽度：左侧栏初始 240px，标题 + `width=25` 的搜索框放不下；Label 默认 `anchor=center`，宽度不足时**从两侧同时裁切**，右边的「件」就没了 | 标题行改 `grid` 且标题列 `weight=1 / sticky='w'`（从左起排不居中裁切），搜索框收窄到 `width=12`；左/中栏默认宽度 240→300，分割条下限与保存时的夹取同步调整 |

搜索结果截图核对点：左侧标题完整、收文编号完整到末位、按钮无方框。

## 三之七、打印链路修复（2026-11-23）

现场点「打印」报两个错：

```
/bin/lp：不支持的 document-format
"application/vnd.openxmlformats-officedocument.wordprocessingml.document"
```

```
WPS 文字：无法打开 "pdf"。
此文件可能被重命名、删除或移动了位置。
```

两个错是同一根源的两面：

| 环节 | 问题 | 修复 |
| --- | --- | --- |
| WPS 转换 | `wps_compat` 伪造了 LibreOffice 的 `--convert-to pdf` 参数。**WPS Linux 没有无头转换接口**，它把 `pdf` 当成要打开的文件名，于是弹「无法打开 pdf」 | 去掉伪造参数；WPS 只在**真实试转成功**时才启用，且结论缓存 —— 不支持的版本只探测一次，之后不再启动进程、不再弹窗 |
| 打印兜底 | `word_to_pdf` 返回 `None`（转换失败）后，`print_file` 里 `target` 仍是原 `.docx`，直接递给了 `lp`；CUPS 不认识 docx 的 MIME 类型，报「不支持的 document-format」 | 转换失败**直接返回失败并给出可操作提示**，绝不再把 Word 文件交给 `lp` |
| 转换优先级 | 原来 WPS 优先，每次都先启动 WPS 试错 | 改为 **LibreOffice 优先**（`--convert-to` 是官方接口，唯一可靠），WPS 仅作兜底 |

`print_file` 现在的行为：

```
.docx / .doc / .wps / .rtf / .odt  ->  必须先转 PDF，转不出来就报错中止
.pdf                              ->  直接 lp
转换失败提示：「无法把 Word 转成 PDF，打印已中止。请安装 LibreOffice（推荐）…」
```

**现场要做的**：打印 Word 需要系统里有 LibreOffice（WPS 本身不具备命令行转换能力）：

```bash
sudo apt install -y libreoffice-writer
```

装完 `./诊断.sh` 里 `LO` 那行会显示 soffice 路径，打印即可用。
若不想装 LibreOffice，可以先用「批量转PDF」把批办单转成 PDF 再打印。

CI 断言：模拟「转换失败」时必须返回失败且**不调用 lp**，防止这个 bug 回归。

### 装了 LibreOffice 仍然转换失败？

现场反馈「麒麟上已经装了 LibreOffice，打印还是报错」。查下来是转换调用的
输入文件名问题 —— 生成的批办单叫 `投诉批办单#C-26-113.docx`，而
LibreOffice 内部按 `file://` URI 处理路径，**`#` 是 URI 片段分隔符**，
`&` 是参数分隔符，空格和中文在部分版本下也会出问题。结果是
「命令返回成功、但产物名字对不上」，代码判定为失败，于是回到那条
「不支持的 document-format」老路。

现在的做法（`_try_soffice`）：

1. **先复制成安全名**：`投诉批办单#C-26-113.docx` → `src_convert.docx`，
   放进独立临时目录再转，`# & 空格 中文` 全部规避
2. **不猜产物名**：转换后扫描输出目录，取「本次新产生且非空的 `.pdf`」，
   不再依赖 `源文件名 + .pdf` 这个假设
3. **失败原因可见**：`last_convert_error()` 记录真实原因（含 LO 的
   stderr / 返回码 / 超时），打印失败的弹窗直接显示，不再是笼统的一句话
4. **超时放宽到 180 秒**：LibreOffice 首次启动要建用户配置，120 秒常不够
5. **加 `--nofirststartwizard`**：跳过首次运行向导

`诊断.sh` 也加了一步真实转换自检：拿一份 `.docx` 实际转一次 PDF，
直接告诉你「通过 / 失败 + 原因」，不用猜。

## 三之八、Word 转 PDF「一页变两页」与「能否直接打印 Word」

### 为什么会变两页

模板用的是 Windows 字体：`仿宋_GB2312`、`方正小标宋简体`。
麒麟/统信上没有这些字体，LibreOffice 会做字体替换：

```
仿宋_GB2312  ->  Noto Serif CJK SC（或其他）
```

替换字体的**度量不同**（字面高度、行距、字符宽度），而模板页边距很紧
（A4、上下 1440 twips、左右 1800 twips）。行高微增，最后一行就被挤到第 2 页。

**这不是转换错误，是字体替换的必然结果。**

### 已做的缓解（自动生效）

`font_compat.py` 给 fontconfig 写用户级别名（`~/.config/fontconfig/conf.d/`），
优先使用度量最接近的本地字体：

| 模板字体 | 映射到 |
| --- | --- |
| 仿宋_GB2312 / 仿宋 | FangSong → AR PL UMing CN → Noto Serif CJK SC |
| 宋体 / SimSun | AR PL UMing CN → Noto Serif CJK SC |
| 方正小标宋简体 | Noto Serif CJK SC |
| 黑体 / 微软雅黑 | Noto Sans CJK SC |

同时加了**页数校验**：转换后比对 Word 记录的页数与 PDF 实际页数，
不一致就弹提醒并说明缺哪些字体 —— 不再让你打完才发现多一张纸。

### 要完全一致，有两条路

1. **装上模板字体**（最彻底）
   把 `仿宋_GB2312.ttf`、`方正小标宋简体.ttf` 拷到 `~/.local/share/fonts/`
   再 `fc-cache -fv`。这两个是商业字体，需自行取得授权。

2. **不转 PDF，直接打印 Word**
   菜单「打印 → ⚙ Word 直接打印（不转 PDF）」，走 LibreOffice 的 `--print`，
   不落 PDF 文件。或「打印 → 📄 在 WPS 中打开当前文件并手动打印」，
   用 WPS 自己的排版引擎，与 Word 最接近。

也可以用环境变量切换默认行为：

```bash
COMPLAINTS_PRINT_MODE=pdf      ./启动GUI.sh   # 默认：转 PDF 再打（兼容性最好）
COMPLAINTS_PRINT_MODE=direct   ./启动GUI.sh   # 直接打印 Word
COMPLAINTS_PRINT_MODE=word     ./启动GUI.sh   # 交给 WPS 打开后手打
```

### 为什么不能把 docx 直接交给 lp

CUPS 的 `lp` 只接受 PDF / PostScript / 图片等打印就绪格式，
**不认识 OOXML**。直接递 `.docx` 会报：

```
/bin/lp：不支持的 document-format
"application/vnd.openxmlformats-officedocument.wordprocessingml.document"
```

Windows 之所以能「直接打印 Word」，是 `ShellExecute(..., "print")` 让
Word 自己去排版并输出到打印驱动 —— 真正干活的仍是 Word，不是打印子系统。
Linux 上对应的角色是 LibreOffice / WPS，所以：

- **要么**让 LibreOffice 转成 PDF 再交给 CUPS（默认，兼容性最好）
- **要么**让 LibreOffice 自己调打印子系统（`--print`，少一道转换）

## 三之九、为什么「还是两页」（真实根因，2026-11-24）

上一轮怀疑是字体，现场反馈「已经装了仿宋_GB2312，问题不在字体」——
确认属实。真正的根因是**生成文档时把模板的固定行距弄丢了**。

模板里有几个段落带固定行距：

```xml
<w:spacing w:line="260" w:lineRule="exact"/>   <!-- 13pt -->
<w:spacing w:line="340" w:lineRule="exact"/>   <!-- 17pt -->
<w:spacing w:line="560" w:lineRule="exact"/>   <!-- 28pt -->
```

而旧代码用 `cell.text = ''` 清空单元格 —— 这句不只清文字，
**连段落属性（pPr/spacing）一起清掉了**。之后新建的段落只能退化成
`w:line="240" w:lineRule="auto"`（单倍行距，按实际字体行高自动排）。

Word 对 auto 比较宽容，LibreOffice 严格按字体实际行高计算，
于是同一格在麒麟上排得更高，累积溢出到第 2 页。

**修复**：

1. 不再 `cell.text = ''`，改为保留首个段落元素、只删里面的 run，
   段落属性原样保留
2. 删除多余空段落，避免残留空段撑高单元格
3. 沿用模板固定行距；**单位坑**：python-docx 的 `spacing_line` 返回
   EMU（1pt = 12700 EMU），不是 twips 也不是磅。直接 `Pt(spacing_line)`
   会得到 215900pt 这种天文数字，反而把页面彻底撑爆
4. 模板没写行距的段落，用文档网格 `linePitch`（312 twips）作固定行距
5. `_doc_line_pitch()` 是 staticmethod，在另一个 staticmethod 里必须
   经类名调用；裸写会抛 NameError 被 except 吞掉，静默退回单倍行距

修复前后对比（同一份工单）：

```
模板        auto=0  exact=3
修复前生成  auto=7  exact=4     ← 行距被改成自动
修复后生成  auto=0  exact=11    ← 全部固定行距
```

CI 断言「排版保真」：生成文档里模板有固定行距的单元格，
行距值与行距规则必须逐项一致，否则流水线直接红。

## 三之十、LibreOffice 自身损坏（FT_Done_MM_Var）

现场报错：

```
/usr/lib/libreoffice/program/soffice.bin: symbol lookup error:
/lib/aarch64-linux-gnu/libcairo.so.2: undefined symbol: FT_Done_MM_Var
```

这是 **LibreOffice 与系统 freetype/cairo 版本不配套** 的系统级问题，
不是本程序能修的 —— 此时 LibreOffice 连启动都做不到，任何转换都会失败。

程序现在会：

- 首次遇到该错误即记住，不再反复重试（避免每次等超时）
- 弹窗给出可直接执行的修复命令
- 主动询问「是否改用 WPS 打开并手动打印」

修复 LibreOffice：

```bash
sudo apt install --reinstall libreoffice-writer libreoffice-core libfreetype6 libcairo2
# 源不可用时先卸载再装
sudo apt remove --purge libreoffice-* && sudo apt install libreoffice-writer
```

不想动系统的话，用 WPS 打印即可 —— 版式与 Word 一致，而且不需要转换：

```
「打印」菜单 → 「在 WPS 中打开当前文件并手动打印」
```

**关于诊断.sh 闪退**：原来脚本开头是 `set -u`，任一未定义变量或
`file`/`ldd` 命令缺失就会让整个脚本中途退出，什么都看不到。
现已改为 `set +u` / `set +e`，每项检查独立失败，保证一定打印出结论。

## 三之十一、打包版库路径污染（FT_Done_MM_Var 的真正根因，2026-11-25）

现场最终确认：**不是 LibreOffice 损坏，是库版本冲突**。

机制：

```
PyInstaller 打包后，程序启动时会把自己 _internal/ 塞进 LD_LIBRARY_PATH。
我们的 _internal 带着 Pillow 等依赖附带的旧版 libfreetype.so.6 /
libfontconfig.so.1。

调用 soffice 时子进程**继承**了这个 LD_LIBRARY_PATH，于是系统 libcairo
去找 FT_Done_MM_Var 时，加载的是 _internal 里的旧 freetype —— 该符号
不存在 —— 报：
    symbol lookup error: undefined symbol: FT_Done_MM_Var
```

所以重装 LibreOffice 解决不了，LO 本身是好的。

**两层修复**：

1. **打包层（根治）**：`complaints-linux.spec` 新增
   `_SYSTEM_LIB_PREFIXES` + `_filter_binaries()`，把 libfreetype /
   libfontconfig / libcairo / libpango / libharfbuzz / libX11 等
   「系统必然自带、ABI 稳定」的基础库排除出 `_internal`。
   Pillow 等改用系统库（版本更新，功能不受影响）。

2. **运行层（兜底）**：新增 `print_env.py`
   - `clean_env()`：返回剔除应用自带库路径的环境变量，供**所有**
     调用外部系统程序的地方使用（soffice / lp / unoconv / WPS /
     xdg-open / fc-list / fc-cache）
   - `sanitize_library_path()`：启动时清理本进程的 LD_LIBRARY_PATH
   - `bundled_lib_report()`：自检 `_internal` 是否仍带冲突库
   - 同时剔除 LD_PRELOAD / LD_AUDIT / PYTHONPATH / PYTHONHOME

即使将来打包漏掉排除规则，运行层也保证子进程用系统库。

**CI 校验**：构建脚本检查 `_internal` 内不再含冲突库；`--selftest`
断言 `clean_env()` 确实剔除 `_internal` 且保留用户系统路径。

**现场不需要做任何事**，装新包即可。想确认旧包状态可跑：

```bash
bash 修复打印.sh      # 只诊断，不改系统
```

它会告诉系统 freetype 有没有 `FT_Done_MM_Var`、cairo 实际加载哪个库、
`_internal` 里是否还有冲突库，并给出处置办法。

## 三之十二、docGrid（行网格）是转 PDF 多一页的元凶（2026-11-25 实测）

上一轮修了「生成时丢失固定行距」，剩下的多出那一页由 `w:docGrid` 造成。

模板的 `sectPr` 里有：

```xml
<w:docGrid w:type="lines" w:linePitch="312"/>
```

这是 Word 的**文档网格**：每行吸附到 15.6pt 的网格。Word 严格按网格排版，
模板就是照网格调出来的，所以正好一页。

LibreOffice 对网格的实现不同——它把网格当作**最小行高**，再叠加段落自身的
行距，于是每行比 Word 高出零点几磅。模板页边距很紧（A4 上下各 1440 twips），
十几行累积下来，表格后面那行「备注」就被挤到第 2 页。

**实测证据**（LibreOffice Kit 转 PDF，同一份文件只差这一行属性）：

| 文件 | 带 docGrid | 去掉 docGrid |
| --- | --- | --- |
| 空模板本身（Word 记录 1 页） | **2 页** | **1 页** |
| 生成的批办单 | **2 页** | **1 页** |

连空模板都复现，说明与内容多少无关，纯粹是网格差异。

**修复**：生成时把 `docGrid` 从 `sectPr` 里去掉。因为每个段落的行距已经
显式写成固定值（`lineRule="exact"`），版面不依赖网格；去掉后 Word 和
LibreOffice 排出来都是一页，行距不变、版式不走样（已逐页渲染核对）。

- `main.strip_docgrid(path)`：保存后就地移除，只删 `<w:docGrid .../>`，
  纸张、页边距、页眉页脚距离全部保留；操作幂等
- 三个生成函数（pdf / doc / word）保存后都调用它
- `platform_compat._try_soffice` 转 PDF 前对**临时副本**再兜一次，
  这样「待打印」里遗留的旧文件（带 docGrid）也能一页打出来，不改原文件
- `COMPLAINTS_KEEP_DOCGRID=1` 可关闭该行为（万一某模板确实依赖网格）

CI 断言：生成的文档必须不含 `docGrid`、`pgSz`/`pgMar` 必须保留、
不得存在 `lineRule="auto"` 段落。

## 三之十三、删 docGrid 后版式走样的补偿（2026-11-26，用 Word 基准标定）

只删 `docGrid` 会让排版塌陷 —— 现场反馈「虽然一页，但表格完全变了样」「页边距
看起来变大、表格离标题太近」。

原因是模板里大量元素**靠网格撑高**，网格一撤就塌：

| 位置 | Word 下 | 只删网格后 |
| --- | --- | --- |
| 表格中无 `trHeight` 的行 | 64 px | **41 px**（塌 36%） |
| 标题与表格之间的空行 | 13 pt | 几乎为 0 |
| 顶部留白 | 83.5 pt | 78.5 pt |

**校准方法**：用 Excel/Word COM 直接导出基准 PDF（`Word导出_基准.pdf`），
按像素逐项比对，反推需要补多少。最终补偿值：

- 表格中「无 `trHeight`」或「`trHeight` 过小」的行 → 补 `640` twips（`atLeast`）
- 标题段落 → 补 `before=100` twips
- 标题与表格之间的空行 → 行高提到 `430` twips

补偿后与 Word 基准逐项吻合（144dpi 下误差 ≤2px）：

| 指标 | Word 基准 | 补偿后 | 只删网格 |
| --- | --- | --- | --- |
| 顶部留白 | 83.5 pt | **83.5** | 78.5 |
| 标题→表格 | 27.5 pt | **27.5** | 19.0 |
| 表格顶 | 132.0 pt | **132.0** | 118.5 |
| 底部留白 | 79.5 pt | **78.5** | 92.5 |

## 三之十四、合并后多行来文字号 + 手动压缩行高

合并批办单时来文字号可能出现两行，需要手动把其它行调小才放得回一页。
**手动调整之后再打印，仍然是一页**，因为：

- 手工调整发生在**生成之后**，`strip_docgrid` 只在文件里**还存在 `docGrid`**
  时才动作；已经处理过的文件（无网格）再调用会直接返回「未修改」，
  绝不覆盖手工值。实测：把行高手改 500 后调用，返回 `False/0/0`，值原样保留。
- 打印时若转换，也只处理**临时副本**，不改原文件。

**但有一个坑必须提醒**：如果给「含多行内容的那一行」设成固定行高
（`hRule="exact"` 且偏小），Word 和 LibreOffice 都会把超出部分**裁掉**，
而且不报错 —— 打出来就是缺字的表格。

正确做法是把该行设为「**最小值**」而不是「固定值」：

```
Word：表格属性 → 行 → 指定高度 → 选「最小值」
```

程序已加自动检查：转换时若发现某行 `exact` 高度明显小于内容所需，
会弹窗提示「第 N 行固定行高可能过小…超出的文字会被裁掉」，
并在日志里记录，避免悄悄印出残缺表格。

实测三种情形均保持 1 页：

| 情形 | 页数 |
| --- | --- |
| 多行来文字号，不压缩其它行 | 1 |
| 多行来文字号，其它行压到 500 twips | 1 |
| 多行来文字号，其它行压到 440 twips | 1 |

## 四、目录与数据

```
投诉批办单管理系统/
├── 投诉批办单管理系统        # ELF 可执行文件
├── 启动GUI.sh               # 启动脚本
├── 投诉批办单管理系统.desktop
├── _internal/               # PyInstaller 依赖（含只读默认资源）
├── AppData/                 # 可写配置、数据库、日志、编号
├── ttk_compat.py            # ttkbootstrap 1.x/2.x 兼容层（随源码一起分发）
├── wps_compat.py            # WPS 调用层（老 .doc 转换，统信/麒麟预装）
├── ui_layout.py             # 小屏/高DPI 布局适配：窗口自适应、滚轮、贴底按钮条
├── emoji_compat.py          # 按钮 emoji 兼容：麒麟无 emoji 字体时剥掉图标避免方框
├── font_compat.py           # 字体度量兼容：Windows 字体映射 + 页数校验（防一页变两页）
├── print_env.py             # 外部程序环境清理：剔除 _internal 库路径（防 FT_Done_MM_Var）
├── 修复打印.sh               # 打印故障诊断：定位 LO 缺符号 / 冲突库（只读）
├── ttk_gui_selftest.py      # 无头 GUI 冒烟测试（--selftest 入口）
├── 诊断.sh                  # 现场诊断：Tk / ttkbootstrap 版本 / 控件自检 / 启动日志
├── 投诉批办单示例#1.docx     # Word 模板
├── 待打印/ 已打印/ 已处理/   # 工作目录
└── 使用说明.txt
```

首次启动会从 `_internal/AppData` 释放默认规则、配置与 `complaints.db` 到可写的
`AppData/`，不会覆盖已有数据。编号、路径、主题、快捷按钮、二维码等设置与 Windows 版
共用同一套 JSON 格式。

## 五、打包（GitHub Actions）

推送到 `main` 分支即自动触发 `.github/workflows/build-linux.yml`，用 matrix 并行构建两个架构，
**各自原生编译，产物不通用**：

| 架构 | 构建镜像 | 目标机器 |
| --- | --- | --- |
| `x86_64` | `amd64/ubuntu:18.04`（glibc 2.27） | Intel/AMD、兆芯、海光 |
| `aarch64` | `arm64v8/ubuntu:18.04`（QEMU 模拟，arm64 容器） | 飞腾、鲲鹏、麒麟芯、瑞芯微 |

aarch64 那条跑在 amd64 runner 上，靠 `docker/setup-qemu-action` 起 arm64 容器做原生交叉构建；
两条都产出真正的 ARM / x86-64 机器码，不是同一份包改名。

容器内流水线（`build-in-container.sh`）：

1. 编译 Python 3.10.14 到 `/opt/py310`（容器 glibc 2.27 → 产物可在 glibc ≥ 2.28 的麒麟/UOS 运行）
2. 装依赖，`ttkbootstrap` 锁 `>=1.10.1,<3`
3. `py_compile` + 全模块导入 + `ttk_compat.verify_widgets()` 控件别名断言
4. **源码 GUI 冒烟测试**：Xvfb 虚拟屏里真把三栏布局、规则窗口、9 个对话框建一遍
5. `PyInstaller complaints-linux.spec --noconfirm`（onedir，本机原生架构）
6. **产物冒烟测试**：`./投诉批办单管理系统 --selftest` 跑刚打出来的 ELF
7. 用 `file` 校验 ELF 架构；扫描所有 ELF 的 `GLIBC_*` 符号，> 2.28 直接失败
8. 补 `.desktop`、启动脚本、`诊断.sh`、`使用说明.txt`，打成
   `complaints-linux-x86_64.tar.gz` / `complaints-linux-aarch64.tar.gz` + SHA-256
9. `release` job 汇总两个 Artifacts 发 Release

任何一步失败都不会产出 Release —— 包括第 4、6 步的 GUI 冒烟，所以「打包成功但客户机上打不开」
这类问题会在流水线就红掉。

本地手动打包（在目标架构机器上执行，产物即该架构原生可执行文件）：

```bash
python3 -m venv --system-site-packages .venv
.venv/bin/python -m pip install -r requirements-linux.txt
.venv/bin/python -m PyInstaller complaints-linux.spec --noconfirm
chmod +x "dist/投诉批办单管理系统/投诉批办单管理系统"
file "dist/投诉批办单管理系统/投诉批办单管理系统"   # 确认 ARM aarch64 / x86-64
```

## 六、功能清单（与 Windows 版一致）

- 批量处理 PDF / DOCX / DOC 源文件，按文件名关键词匹配拟办意见与办理结果
- 可配置规则引擎（规则管理界面、导入导出、冲突检测）
- Word 批办单生成（保持模板字体/段落/合并单元格格式）、二维码
- 可视化编辑、输入历史联想、快捷填充按钮
- 单份打印、批量打印、批量转 PDF 合并、打印后自动归档与延迟清理
- 编号管理、SQLite 数据库检索/备份/恢复/整理、Excel 统计报表导出
- 主题切换、窗口布局记忆、路径配置、打印延迟配置
