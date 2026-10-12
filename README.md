# docx2epub — Word → EPUB 转换管线

## 介绍

docx2epub 是一个将 Microsoft Word（.docx/.doc/.rtf/.docm）转换为标准 EPUB 电子书的 Python 工具集，配套一个 Web 控制台用于批量管理和质检。

### v0.2.0（2026-10-12）变更要点

- **修订终稿化（accept_revisions）**：Word COM 统一接受所有修订标记，产出终稿再进入后续 pipeline，杜绝"部分阶段见修订版、部分阶段见终稿"的图题错乱/删除字符泄漏事故
- **公式高清化（math_hd）**：Word COM 公式转高清 PNG；不可用时回落 LO 低清
- **浮动图截图范围扩大**：纯图片浮动簇也纳入截图，防止 LO 位置漂移（实证：安心养老图三）
- **空标题/关键词行修复**：空 h1/h2 自动删或降级为 p；"关键词：…"行降级为段落，不再独占一章
- **图文混排拆段**：图夹在文字中间的浮动环绕残留统一拆出为独立居中图段；短题注（≤40字）同段也拆段居中
- **图宽锁定策略修正**：自然宽超过版心的大图（看板 1000-1200px）转纯流式 max-width:100%，小图才锁死像素防放大糊掉
- **书签锚点清除**：清除 Word _Toc 目录书签、_GoBack 光标位、OLE_LINK 复制残留等空锚点，读者端不再列成"书签"
- **etree.Element → makeelement 修复**：`body.makeelement("p", {})` 替代裸 `etree.Element("p")`，避免 `.text_content()` AttributeError
- **Web 控制台改版**：移除仪表盘页；文件夹选择支持 tkinter（本机）/ PowerShell（headless 兜底）/ 浏览器内目录浏览器三层链；试读页改为目录扫描模式
- **样式配置精简**：config.yaml 去注释行，键值去掉多余引号

### 核心流程

三个分离阶段（改样式参数只需重跑阶段3）：

1. **前置手术（Stage 1）**：接受修订→终稿化、Tab 转不间断空格、公式高清化、组合/浮动形状整块截图（Word 不可用时自动回落 LibreOffice）
2. **LO→HTML（Stage 2）**：LibreOffice headless 将处理后的文档转为 XHTML
3. **清理+组包（Stage 3）**：空标题/关键词行修复、图文拆段、书签清除、切章、注入 CSS 样式、图片压缩优化、EPUB 2.0 组包

Web 控制台（`python app/webapp/app.py`）提供配置管理、转换队列、质检面板（L1/L2 字符覆盖率闸）、试读功能。

### 技术栈

- **Python 3.11+**（依赖：`lxml`, `Pillow`, `PyYAML`, `fastapi`, `uvicorn`, `openpyxl`, `pywin32`）
- **LibreOffice**（`soffice.exe`，headless 模式）
- **Microsoft Word COM**（可选，用于修订终稿化、公式高清化、形状快照；不可用时自动回落）
- **Java 8+**（可选，用于 `epubcheck` 合规校验）

## 目录结构

```
app/                ← 只传这部分到 Git
  config.yaml       # 管线全局配置（路径、样式参数、闸门阈值）
  tools/            # 各阶段脚本与诊断工具（.py）
  webapp/           # FastAPI 后端 + 前端静态文件
    app.py          # Web 控制台主程序（启动命令见下文）
    static/         # HTML/CSS/JS + vendor/epub.min.js
out_pipeline/       # 运行时产物（EPUB 输出）— .gitignore
tests/              # 测试用例源文件（docx/doc/rtf）— .gitignore
work/               # 运行期缓存（manifest、解包、epubcheck 下载）— .gitignore
docs/               # 项目文档与验收报告 — .gitignore
```

> **Git 策略**：只提交 `app/` 目录。`out_pipeline/`、`tests/`、`work/`、`docs/`、`.workbuddy/`、`*.bak`、`__pycache__/` 均在 `.gitignore` 中排除。

## 安装依赖

### 1. 系统前置条件

- **Python 3.11+**：从 [python.org](https://www.python.org/downloads/) 安装，勾选 "Add to PATH"
- **LibreOffice 7+**：从 [downloadcenter.libreoffice.org](https://downloadcenter.libreoffice.org/) 安装（默认路径 `C:\Program Files\LibreOffice\`）
- **Microsoft Word**（可选但推荐）：用于修订终稿化、公式高清化、形状快照；不可用时各功能自动回落 LibreOffice
- **Java 8+**（可选）：仅用于 `epubcheck` 合规校验；[adoptium.net](https://adoptium.net/)

### 2. Python 依赖

用 `pip` 或 `uv` 任选一种：

```bash
# pip 方式
pip install lxml Pillow PyYAML fastapi uvicorn openpyxl pywin32

# 或 uv 方式（更快）
uv pip install lxml Pillow PyYAML fastapi uvicorn openpyxl pywin32
```

> 当前实测环境：Python 3.11.9、`lxml==6.1.0`、`pillow==12.2.0`、`PyYAML==6.0.3`、`fastapi==0.139.0`、`uvicorn==0.49.0`、`openpyxl==3.1.5`。

### 3. 配置

编辑 `app/config.yaml`，确认 `paths.soffice` 指向本机 LibreOffice 安装路径：

```yaml
paths:
  soffice: C:\Program Files\LibreOffice\program\soffice.exe
  workdir: work/pipeline
```

如需启用 epubcheck 闸门，将 `gates.epubcheck: false` 改为 `true`，并在 `work/epubcheck/` 下放 [epubcheck 5.x](https://www.w3.org/pubs/epub/epubcheck/) JAR。

## 使用

### Web 控制台（推荐）

```bash
cd D:\project\epub
python app/webapp/app.py
# 浏览器打开 http://127.0.0.1:8300
```

### 命令行批量转换

```bash
# 转换单个文件
python app/tools/pipeline.py convert input.docx --stage all -o out_pipeline/

# 仅跑闸门质检（不重新转换）
python app/tools/pipeline.py check out_pipeline/<book_id>

# 生成默认配置
python app/tools/pipeline.py init-config -o app/config.yaml
```

### 单工具调用

`app/tools/` 下各 `diag_*.py`、`v*_*.py`、`rerun_*.py` 脚本供专项诊断和批量重试使用，具体参数见各文件头部 docstring。

## 开发约定

- **备份**：修改前对改动文件做 `.bak` 副本（本仓库已自动备份后提交）
- **验证**：改后用 `python -m py_compile <file>.py` 检查语法
- **回归**：重大修改后开独立进程跑 `python app/webapp/app.py` 真实链路，不在同一进程里 query 算穷举假象
- **分支**：功能分支命名 `feat/<name>`，合入主分支后打 tag（如 `v0.2.0`）
- **commit message**：格式 `v<major>.<minor>.<patch>: 一句话摘要 — 主要模块`

## License

内部工具，未开源。
