# docx2epub — Word → EPUB 转换管线

## 介绍

docx2epub 是一个将 Microsoft Word（.docx/.doc/.rtf/.docm）转换为标准 EPUB 电子书的 Python 工具集，配套一个 Web 控制台用于批量管理和质检。

### v0.3.0（2026-10-14）变更要点

- **EPUB 导航条目补全（"导航呢？"实证修复）**：4.88 已识别正文目录并建链接，但 nav.xhtml / toc.ncx 还只按切章生成——长文档单 section 侧边栏为空。现从 4.88 导出 nav_entries，优先用真目录填导航文件，阅读器侧边栏显示全部条目
- **缓存降级固化修复**：snap 失败走 flatten 回落的 s4_ 缓存，mtime 比源新但结果是降级货（看板图缺失、L1/L2 必挂），下轮照样命中，问题接力固化。现检查 snap_err 字段，降级缓存一律作废重做
- **代码新鲜度校验（断点续跑）**：pipeline.htmlpath_poc 更新后，旧 EPUB 是旧规则产物，不应被跳过。现在比对 tools/ 源码 mtime vs EPUB mtime，源码更新时强制重转
- **表格样式内联**：平台不加载 styles.css，此前"去 width 属性，CSS 统一管"假设在平台侧不成立，表格裸奔无边框/无列宽约束。现统一内联边框、内边距、列宽比例（保留 LO 百分比，丢弃 px 绝对宽）
- **列表拍平正序保序**：DOM 改写时 li→p 搬移原倒序处理，导致表格单元格内列表顺序反转（案例集 23-26 页实证）。改为列表倒序处理外层、列表内 li 正序 addprevious，保序
- **h1_color 清空**：config.yaml 中 h1_color 由 '#8B0000' 改为空字符串，按平台统一样式规则处理

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
