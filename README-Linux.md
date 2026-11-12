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

自检内容：控件别名、三栏 `PanedWindow` 布局、规则管理窗口、9 个工具对话框。
结论写入 `AppData/ttk_selftest.log`（`console=False` 时 stdout 为空，只能靠这个文件取证），
并以退出码返回。故意删掉 `PanedWindow` 别名时自检会 fail 并打印
`缺少控件: ['PanedWindow']`，确认它真能拦住这个回归。

## 四、目录与数据

```
投诉批办单管理系统/
├── 投诉批办单管理系统        # ELF 可执行文件
├── 启动GUI.sh               # 启动脚本
├── 投诉批办单管理系统.desktop
├── _internal/               # PyInstaller 依赖（含只读默认资源）
├── AppData/                 # 可写配置、数据库、日志、编号
├── ttk_compat.py            # ttkbootstrap 1.x/2.x 兼容层（随源码一起分发）
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
