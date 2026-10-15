# docx2epub — Word → EPUB 转换管线

## 介绍

docx2epub 是一个将 Microsoft Word（.docx/.doc/.rtf/.docm）转换为标准 EPUB 电子书的 Python 工具集，配套一个 Web 控制台用于批量管理和质检。

### v0.4.0（2026-10-15）变更要点

- **BUG-001：块内公式（oMathPara）导出错位**：OMaths 逐个枚举时，同段含多个 oMath 导致导出 N 张图只有 1 个替换槽位。改为按段整体导出一次（整段作为公式图片），与 surgery 侧替换粒度严格一一对应
- **BUG-002：导航/OPF 文本未转义**：标题含 `& < >` 时产出非法 XML。引入 `xml.sax.saxutils.escape` 对 nav/ncx/opf 所有文本和 href 字段转义；htmlpath_poc.py 的 `make_xhtml_doc` 同步引入并支持 `xhtml11` 开关（EPUB2 校验规则要求）
- **BUG-003：stale 死代码导致缓存永不作废**：源文件改后 stale 计算结果从未参与判断，s4_ 缓存始终命中。补齐新鲜度检查，stale 或 degraded 任一命中即作废
- **BUG-004：config.yaml 变更不触发重转**：配置中心改样式/闸门参数后，断点续跑照跳 EPUB。补 config.yaml mtime vs EPUB mtime 比对
- **BUG-005：NCX dtb:uid 与 OPF identifier 不匹配**：OPF 用 `urn:uuid:<uid>` 而 ncx 只写裸 uid，epubcheck 报"identifier does not match"。ncx 入参改为 `"urn:uuid:" + uid`
- **BUG-006（根除）：组包双实现漂移**：htmlpath_poc.py 与 pipeline.py 各有一套 make_nav/make_opf/build_epub，不同步产生差异。砍除 htmlpath_poc.py 中的组包函数，统一走 pipeline.py
- **BUG-007：worker 无 try/finally 永久卡死**：未捕获异常（磁盘写满、Popen 失败等）使 BATCH.running 永久 True，前端一切操作 409。加 try/finally 包整个 worker 循环
- **BUG-009：manifest 半行/坏行致任务历史 500**：json.loads 异常冒泡整页崩溃。逐行 try/except 跳过坏行
- **BUG-010：soffice 路径写死 C 盘**：安装在其他盘符时 LO 转换全部失败。改用 shutil.which 查找 PATH，回落到 x86 常见路径
- **BUG-012：同名图片静默互覆**：文档同时含 image.png 与 image.jpg，前者转 JPEG 后直接覆盖后者。新增 `_uniq()` 级联避让直到名字空闲
- **BUG-013：跨 attempt 输出持久化**：首次尝试的输出残留到第二次 attempt，取到的 rec 是上一次的结果。out 变量移到 attempt 循环外清空
- **BUG-015："目录"标签段污染主标题**：4.87 已排除而 4.85 漏了，居中"目录"被提升成 h1，与 4.88 目录识别抢同一段，污染 `<title>`。补全排除
- **BUG-016：before.count(b) 守卫恒真**：dict 相等计数，b 自身恒 ≥1，去重判断形同虚设。删掉此守卫，保留 (name, ctx) used 集合即可
- **EPUB2 img alt="" 补全**：epubcheck RSC-005 要求 img 必须有 alt 属性，LO 产出的 img 一律无 alt。正则自动补 `alt=""`
- **EPUB2 package xml:lang 剥离**：EPUB2 校验规则不接受 `<package>` 上的 xml:lang 属性（RSC-005），版本分流——EPUB3 保留，EPUB2 去掉（语言信息由 dc:language 承载）
- **表格样式内联（v0.3.0 延续）**：边框、内边距、列宽比例统一烤进标签，platform 不加载 styles.css 假设继续生效

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
