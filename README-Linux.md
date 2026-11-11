# 投诉批办单管理系统 · 统信 UOS / 麒麟 Kylin V10 版

本目录是 Windows 版「投诉批办单管理系统」的国产 Linux 移植版。GUI 布局、菜单、按钮、
字段、Word 模板、编号规则、数据库结构与全部业务流程保持原样，仅把 Windows 专属调用
替换成国产系统上的等价实现。

## 一、运行方式

### 1. 直接运行 GitHub Actions 打包出的可执行文件（推荐）

在 [Releases](https://github.com/littlebabyqq2019/complaints/releases) 下载
`complaints-linux-x86_64.tar.gz`，解压后：

```bash
cd 投诉批办单管理系统
./启动GUI.sh
```

或直接双击 `投诉批办单管理系统` / `启动GUI.sh`。

### 2. 源码运行

```bash
sudo apt install python3 python3-tk python3-pip \
                 fonts-noto-cjk fonts-wqy-microhei \
                 libreoffice-writer cups-client xdg-utils
python3 -m pip install -r requirements-linux.txt
./启动GUI.sh
```

命令行批处理模式：`./启动命令行.sh`（等价于 Windows 的 `python main.py`）。

## 二、系统要求

| 项目 | 要求 |
| --- | --- |
| 操作系统 | 统信 UOS 20+ / 麒麟 Kylin V10 SP1+（x86_64 或 aarch64） |
| 架构 | amd64 构建件适用于 x86_64；飞腾/鲲鹏等 aarch64 机器请在对应架构上重新执行 `pyinstaller complaints-linux.spec` |
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

## 四、目录与数据

```
投诉批办单管理系统/
├── 投诉批办单管理系统        # ELF 可执行文件
├── 启动GUI.sh               # 启动脚本
├── 投诉批办单管理系统.desktop
├── _internal/               # PyInstaller 依赖（含只读默认资源）
├── AppData/                 # 可写配置、数据库、日志、编号
├── 投诉批办单示例#1.docx     # Word 模板
├── 待打印/ 已打印/ 已处理/   # 工作目录
└── 使用说明.txt
```

首次启动会从 `_internal/AppData` 释放默认规则、配置与 `complaints.db` 到可写的
`AppData/`，不会覆盖已有数据。编号、路径、主题、快捷按钮、二维码等设置与 Windows 版
共用同一套 JSON 格式。

## 五、打包（GitHub Actions）

推送到 `main` 分支即自动触发 `.github/workflows/build-linux.yml`：

1. `py_compile` + 全模块导入校验
2. `pyinstaller complaints-linux.spec --noconfirm`（onedir）
3. 生成 `.desktop`、`启动GUI.sh`、`使用说明.txt`
4. 打成 `complaints-linux-x86_64.tar.gz` 并计算 SHA-256
5. 上传 Artifacts 并创建 Release

本地手动打包：

```bash
python3 -m pip install -r requirements-linux.txt
pyinstaller complaints-linux.spec --noconfirm
chmod +x "dist/投诉批办单管理系统/投诉批办单管理系统"
```

## 六、功能清单（与 Windows 版一致）

- 批量处理 PDF / DOCX / DOC 源文件，按文件名关键词匹配拟办意见与办理结果
- 可配置规则引擎（规则管理界面、导入导出、冲突检测）
- Word 批办单生成（保持模板字体/段落/合并单元格格式）、二维码
- 可视化编辑、输入历史联想、快捷填充按钮
- 单份打印、批量打印、批量转 PDF 合并、打印后自动归档与延迟清理
- 编号管理、SQLite 数据库检索/备份/恢复/整理、Excel 统计报表导出
- 主题切换、窗口布局记忆、路径配置、打印延迟配置
