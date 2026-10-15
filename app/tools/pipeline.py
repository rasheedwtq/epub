# -*- coding: utf-8 -*-
"""docx2epub 正式管线 CLI（M1 工程化，A1′ 路线）

三阶段分离（改样式参数只需重跑阶段3）：
  stage1 前置手术 : tab→U+00A0、OMML→高清PNG（Word COM）→ surgery_<name>.docx（缓存）
  stage2 LO→HTML  : LibreOffice headless → <name>.html（缓存）
  stage3 cleanup+组包: 清理/切章/注入样式 CSS/EPUB3 组包 + 解包

用法:
  python pipeline.py init-config [-o config.yaml]           # 生成默认配置
  python pipeline.py convert <输入文件或目录> [-c config.yaml]
        [-o 输出目录] [--stage all|1|2|3] [--book-id X] [--no-gates]
  python pipeline.py check <输出目录> [-c config.yaml]       # 只跑闸门（L1/L2 + EPUBCheck）

约定:
  - 配置: YAML（配置中心对接；三级覆盖 = 默认配置 < 标签组 < 单本 override，本版先支持单层）
  - 任务表: work/pipeline/manifest.jsonl（每本一行 JSON，状态机 queued/running/done/failed）
  - 产物: <out>/<book_id>/ 下 转换_<名>.epub + epub解包/ + 任务.json（评分卡）
  - 阶段缓存: work/pipeline/<book_id>/{stage1,stage2}，--stage 3 时直接复用
"""
import os, re, sys, json, argparse, shutil, zipfile, time, uuid, subprocess
from urllib.parse import unquote
from lxml import html as LH
from lxml import etree

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from htmlpath_poc import surgery_tab, lo_convert, cleanup_and_split, make_xhtml_doc  # 复用已验证函数

DEFAULT_CONFIG = r"""# docx2epub 管线配置（配置中心数据源；改动后 --stage 3 即可生效）
paths:
  soffice: C:\Program Files\LibreOffice\program\soffice.exe
  workdir: work/pipeline

surgery:            # 阶段1
  tab_to_nbsp: true
  math_hd: true     # Word COM 公式高清化（Word 不可用时自动回退 LO 低清）

split:              # 阶段3 切章
  heading_tag: h1
  preamble_into_first: true

styles:             # 阶段3 CSS 注入（配置中心"样式参数"块）
  body_font: "宋体, SimSun, serif"
  body_size: 1em
  body_line_height: "1.7"
  h1_size: "1.6em"
  h2_size: "1.4em"
  h3_size: "1.2em"
  table_border: "0.5pt solid #444"
  img_max_width: 580   # 图片内联锁定宽度上限(px)：超过的转纯流式 max-width:100%，任何视口不溢出
  img_optimize: true   # 图片压缩：大图降采样 + 不透明 PNG/GIF→JPEG（epub 瘦身，适合网览）
  img_pixel_max: 1000  # 压缩后最大像素宽（1000 ≈ 580 显示宽的 1.7 倍，手机/半屏都够清晰）
  img_jpeg_quality: 82 # JPEG 压缩质量（截图类 80~85 视觉无损）

package:            # 组包（平台适配层参数位，等平台正式参数后启用）
  epub_version: "2.0"    # 平台内核 epub.js 已实测双版本兼容；大叔拍板默认 EPUB2
  filename_style: keep    # 平台确认文件名无要求，keep 定案

gates:              # 转换后闸门
  l1_min: 100.0     # L1 字符覆盖率下限（%）
  l2_strict: true   # L2 计数零差异
  epubcheck: false  # 需 Java，M0 已验证 14/14 全绿，批量时按需开
"""

def load_config(path):
    import yaml
    with open(path, encoding="utf-8") as f:
        return yaml.safe_load(f)

def _lvl_css(st, tag, size, extra_margin):
    """单级标题 CSS：颜色/加粗/斜体可配，空值=不输出该声明（走浏览器默认）"""
    color = st.get(f'{tag}_color') or ''
    bold = st.get(f'{tag}_bold', True)
    italic = st.get(f'{tag}_italic', False)
    parts = [f'font-size:{size}', extra_margin]
    if color: parts.append(f'color:{color}')
    parts.append('font-weight:' + ('bold' if bold in (True, 'true', 'True', 'bold') else 'normal'))
    if italic in (True, 'true', 'True', 'italic'): parts.append('font-style:italic')
    return ';'.join(parts)

def make_css(st):
    """阶段3：按配置生成 CSS（配置中心改参数 → 只重跑阶段3 的关键）"""
    return f"""body{{font-family:{st['body_font']};font-size:{st['body_size']};line-height:{st['body_line_height']};margin:1em;color:{st.get('body_color') or '#111'}}}
h1{{{_lvl_css(st,'h1',st['h1_size'],'text-align:center;margin:1.2em 0 .8em')}}}
h2{{{_lvl_css(st,'h2',st['h2_size'],'margin:1em 0 .6em')}}}
h3{{{_lvl_css(st,'h3',st['h3_size'],'margin:.8em 0 .5em')}}}
h4,h5,h6{{font-size:1.1em;margin:.7em 0 .4em}}
p{{margin:.35em 0}}
img.fig{{max-width:100%;height:auto}}
table{{width:100%;border-collapse:collapse;margin:.6em 0}}
td,th{{border:{st['table_border']};padding:4px 6px;vertical-align:top}}
caption{{caption-side:top;font-weight:bold;padding:4px}}
"""

def make_ncx(uid, title, secs, nav_entries=None):
    """toc.ncx（EPUB2 兼容 + EPUB3 双导航合法；平台适配层预留）。
    nav_entries：4.88 目录识别导出的条目 [(text, href)]——有则优先用真目录
    （href=section 文件相对名），阅读器侧边栏才显示全部条目；无则回落切章标题。"""
    if nav_entries:
        nps = "\n".join(
            f'    <navPoint id="np-{i}" playOrder="{i}"><navLabel><text>{(t or "文档")}</text></navLabel>'
            f'<content src="sections/{h}"/></navPoint>'
            for i, (t, h) in enumerate([e for e in nav_entries if e[1]], 1))
    else:
        nps = "\n".join(
            f'    <navPoint id="np-{i}" playOrder="{i}"><navLabel><text>{(t or "文档")}</text></navLabel>'
            f'<content src="sections/section{i:04d}.xhtml"/></navPoint>'
            for i, (t, _) in enumerate(secs, 1))
    return f'''<?xml version="1.0" encoding="utf-8"?>
<!DOCTYPE ncx PUBLIC "-//NISO//DTD ncx 2005-1//EN" "http://www.daisy.org/z3986/2005/ncx-2005-1.dtd">
<ncx xmlns="http://www.daisy.org/z3986/2005/ncx/" version="2005-1">
  <head><meta name="dtb:uid" content="{uid}"/><meta name="dtb:depth" content="1"/>
    <meta name="dtb:totalPageCount" content="0"/><meta name="dtb:maxPageNumber" content="0"/></head>
  <docTitle><text>{title}</text></docTitle>
  <navMap>
{nps}
  </navMap>
</ncx>'''

def make_nav(secs, nav_entries=None):
    """nav.xhtml（EPUB3 导航，properties="nav"）。nav_entries 有则用识别目录（同 ncx）。"""
    if nav_entries:
        lis = "\n".join(
            (f'      <li><a href="sections/{h}">{(t or "文档")}</a></li>' if h
             else f'      <li>{(t or "文档")}</li>')
            for t, h in nav_entries)
    else:
        lis = "\n".join(f'      <li><a href="sections/section{i:04d}.xhtml">{(t or "文档")}</a></li>'
                        for i, (t, _) in enumerate(secs, 1))
    return f'''<?xml version="1.0" encoding="utf-8"?>
<!DOCTYPE html>
<html xmlns="http://www.w3.org/1999/xhtml" xmlns:epub="http://www.idpf.org/2007/ops" xml:lang="zh">
<head><meta charset="utf-8"/><title>目录</title><link rel="stylesheet" type="text/css" href="styles.css"/></head>
<body>
  <nav epub:type="toc" id="toc">
    <h1>目录</h1>
    <ol>
{lis}
    </ol>
  </nav>
</body>
</html>'''

def make_opf(uid, title, secs, img_files, version="3.0"):
    if version.startswith("2"):
        man = ['    <item id="ncx" href="toc.ncx" media-type="application/x-dtbncx+xml"/>',
               '    <item id="css" href="styles.css" media-type="text/css"/>']
    else:
        man = ['    <item id="nav" href="nav.xhtml" media-type="application/xhtml+xml" properties="nav"/>',
               '    <item id="ncx" href="toc.ncx" media-type="application/x-dtbncx+xml"/>',
               '    <item id="css" href="styles.css" media-type="text/css"/>']
    for i, _ in enumerate(secs, 1):
        man.append(f'    <item id="s{i:04d}" href="sections/section{i:04d}.xhtml" media-type="application/xhtml+xml"/>')
    for j, f in enumerate(img_files):
        ext = os.path.splitext(f)[1].lower().lstrip(".") or "gif"
        mt = {"png": "image/png", "jpg": "image/jpeg", "jpeg": "image/jpeg",
              "gif": "image/gif", "svg": "image/svg+xml"}.get(ext, "application/octet-stream")
        man.append(f'    <item id="img{j:03d}" href="images/{os.path.basename(f)}" media-type="{mt}"/>')
    spine = "\n".join(f'    <itemref idref="s{i:04d}"/>' for i in range(1, len(secs) + 1))
    date_line = (f'    <dc:date>{time.strftime("%Y-%m-%d")}</dc:date>' if version.startswith("2")
                 else f'    <meta property="dcterms:modified">{time.strftime("%Y-%m-%dT%H:%M:%SZ")}</meta>')
    return f'''<?xml version="1.0" encoding="utf-8"?>
<package xmlns="http://www.idpf.org/2007/opf" version="{version}" unique-identifier="uid" xml:lang="zh">
  <metadata xmlns:dc="http://purl.org/dc/elements/1.1/">
    <dc:identifier id="uid">urn:uuid:{uid}</dc:identifier>
    <dc:title>{title}</dc:title>
    <dc:language>zh</dc:language>
{date_line}
  </metadata>
  <manifest>
{chr(10).join(man)}
  </manifest>
  <spine toc="ncx">
{spine}
  </spine>
</package>'''

def build_epub(epub_path, title, secs, img_srcmap, css, epub_version="3.0", img_renames=None, nav_entries=None):
    if os.path.exists(epub_path):
        os.remove(epub_path)
    uid = str(uuid.uuid4())
    renames = img_renames or {}
    with zipfile.ZipFile(epub_path, "w") as z:
        z.writestr(zipfile.ZipInfo("mimetype"), "application/epub+zip", zipfile.ZIP_STORED)
        z.writestr("META-INF/container.xml", '''<?xml version="1.0" encoding="utf-8"?>
<container version="1.0" xmlns="urn:oasis:names:tc:opendocument:xmlns:container">
  <rootfiles><rootfile full-path="OEBPS/content.opf" media-type="application/oebps-package+xml"/></rootfiles>
</container>''')
        ver = epub_version
        z.writestr("OEBPS/content.opf", make_opf(uid, title, secs, list(img_srcmap.keys()), ver))
        if not ver.startswith("2"):
            z.writestr("OEBPS/nav.xhtml", make_nav(secs, nav_entries))
        z.writestr("OEBPS/toc.ncx", make_ncx(uid, title, secs, nav_entries))
        z.writestr("OEBPS/styles.css", css)
        for i, (t, inner) in enumerate(secs, 1):
            def rew(m):
                # unquote 归一化后查改名表（URL 编码/裸写法通吃；quote() 会把括号编成 %28 而 LO 保留裸括号，字符串直替必漏）
                name = os.path.basename(unquote(m.group(1)))
                return f'src="../images/{renames.get(name, name)}"'
            inner2 = re.sub(r'src="([^"]+)"', rew, inner)
            z.writestr(f"OEBPS/sections/section{i:04d}.xhtml", make_xhtml_doc(t, inner2))
        for src, dst in img_srcmap.items():
            z.writestr(f"OEBPS/images/{os.path.basename(dst)}", open(src, "rb").read())

# ---------------------------------------------------------------- stages

def precheck(src):
    """前置预检：坏文件在进 LO 之前拒绝，报人话错误（M2 批量预检闸门）"""
    if src.lower().endswith((".doc", ".rtf")):
        return None  # 交给 LO，宽容转换
    try:
        z = zipfile.ZipFile(src)
    except Exception as e:
        raise RuntimeError(f"损坏的 docx（zip 打不开）: {type(e).__name__}")
    missing = [n for n in ("[Content_Types].xml", "word/document.xml") if n not in z.namelist()]
    if missing:
        raise RuntimeError(f"损坏的 docx（包内缺 {', '.join(missing)}）")
    try:
        etree.fromstring(z.read("[Content_Types].xml"))
    except etree.XMLSyntaxError as e:
        raise RuntimeError(f"损坏的 docx（[Content_Types].xml 不是合法 XML）: {e.msg}")
    return None


def _loprof(wd):
    """LO profile 目录：book_id 工作目录的哈希 → 恒为无空格 ASCII 路径（带空格路径让 LO 秒退）"""
    import hashlib
    h = hashlib.md5(wd.encode("utf-8")).hexdigest()
    prof_dir = os.path.join(os.path.dirname(wd), "_loprof", h)
    os.makedirs(prof_dir, exist_ok=True)
    return "file:///" + prof_dir.replace("\\", "/")

def _docx_has_revisions(src):
    """快检 docx 是否含修订标记（插入/删除/格式变更）；.doc/.rtf 无法快检 → 视为可能有"""
    if not src.lower().endswith((".docx", ".docm")):
        return True
    try:
        with zipfile.ZipFile(src) as z:
            xml = z.read("word/document.xml").decode("utf-8", "ignore")
    except Exception:
        return True
    return "<w:ins " in xml or "<w:del " in xml or "PrChange" in xml


def _norm_txt(s):
    return "".join((s or "").split())


def _shape_anchor_text(shp):
    """Shape 锚段文本（去空白），accept 后用它找回位置"""
    try:
        return _norm_txt(shp.Anchor.Paragraphs(1).Range.Text)[:80]
    except Exception:
        return ""


def _shape_emf(word, shp):
    try:
        shp.Select()
        return word.Selection.Range.EnhMetaFileBits
    except Exception:
        return None


def _reinsert_lost_shape(doc, shot, s1):
    """把 accept 连带删除的浮动图回插为内嵌图：EMF→PNG→插锚段等价段末尾"""
    from shapesnap import _emf_to_png     # 延迟 import（stage1 已注入 sys.path）
    emf = shot.get("emf")
    if not emf:
        return 0
    png = os.path.join(s1, "rescued_%s_%d.png" % (re.sub(r"\W+", "_", shot["name"])[:20],
                                                  abs(hash(shot["ctx"])) % 100000))
    _emf_to_png(emf, png)
    rng = None
    ctx = shot.get("ctx") or ""
    if ctx:
        paras = list(doc.Paragraphs)
        texts = [_norm_txt(p.Range.Text) for p in paras]
        for n in (60, 40, 25, 15, 10):      # 锚段可能含被删文字，accept 后变短 → 渐短前缀
            key = ctx[:n]
            if len(key) < 6:
                break
            for p, t in zip(paras, texts):
                if key in t:
                    pos = p.Range.End - 1
                    rng = doc.Range(pos, pos)
                    break
            if rng is not None:
                break
    if rng is None:
        pos = doc.Content.End - 1           # 兜底：文末
        rng = doc.Range(pos, pos)
    rng.InlineShapes.AddPicture(FileName=os.path.abspath(png), LinkToFile=False,
                                SaveWithDocument=True, Range=rng)
    return 1


def accept_revisions(src, s1):
    """P0 修订终稿化（2026-10-10，大叔拍板）：转换第一步统一接受所有修订，产出终稿 docx。

    后续所有阶段（shape截图/tab手术/flatten/math/LO渲染/编号提取/L1基线）统一消费终稿，
    杜绝"部分阶段见修订版、部分阶段见终稿"的图题错乱/删除字符泄漏（三方不同源事故）。
    .doc/.rtf 顺带用 Word 转存 docx（比 LO 预转保真，且让它们进入 .docx 主路径）。
    缓存 acc_<stem>.docx，源文件 mtime 变化自动重做；Word 不可用/打开失败原样回落（老路径照走）。
    返回 (终稿src, info)；info["accept"] ∈ cache/none/accepted(N处)/fail(...)。
    """
    stem = os.path.splitext(os.path.basename(src))[0]
    ext = src.lower().rsplit(".", 1)[-1]
    acc = os.path.join(s1, "acc_" + stem + ".docx")
    try:
        if os.path.exists(acc) and os.path.getmtime(acc) >= os.path.getmtime(src) - 1:
            return acc, {"accept": "cache"}
    except OSError:
        pass
    if ext not in ("docx", "docm", "doc", "rtf"):
        return src, {"accept": "skip"}
    if ext in ("docx", "docm") and not _docx_has_revisions(src):
        return src, {"accept": "none"}      # 无修订：零开销跳过 COM
    try:
        import win32com.client as _wc
        word = _wc.DispatchEx("Word.Application")
    except Exception as e:
        return src, {"accept": f"fail({type(e).__name__})"}
    word.Visible = False
    word.DisplayAlerts = 0
    try:
        doc = word.Documents.Open(os.path.abspath(src), ReadOnly=False, AddToRecentFiles=False)
        if doc.TrackRevisions:
            doc.TrackRevisions = False
        n_rev = doc.Revisions.Count
        if n_rev:
            # 安心养老图三事故（2026-10-10 实证）：浮动图锚在被删修订文字上时，
            # AcceptAllRevisions 会连图一起删（Shapes 6→5，架构图消失）。
            # 保护：accept 前存档全部浮动 Shape（名单+锚段文本+EMF），
            # accept 后按名单差集找出消失者，回插内嵌图到锚段等价位置。
            import collections as _col
            before = []
            for _i in range(1, doc.Shapes.Count + 1):
                try:
                    _s = doc.Shapes(_i)
                    before.append(dict(name=_s.Name, ctx=_shape_anchor_text(_s),
                                       emf=_shape_emf(word, _s)))
                except Exception:
                    continue
            doc.AcceptAllRevisions()
            after = _col.Counter()
            try:
                for _i in range(1, doc.Shapes.Count + 1):
                    after[doc.Shapes(_i).Name] += 1
            except Exception:
                pass
            lost_names = list((_col.Counter(b["name"] for b in before) - after).elements())
            if lost_names:
                used = set()
                restored = 0
                for b in before:
                    if b["name"] in lost_names and before.count(b) and (b["name"], b["ctx"]) not in used:
                        used.add((b["name"], b["ctx"]))
                        restored += _reinsert_lost_shape(doc, b, s1)
                print(f"  [accept] 警告：接受修订连带删除 {len(lost_names)} 个浮动图"
                      f"（{','.join(lost_names)}），已回插 {restored} 个内嵌图", flush=True)
        doc.SaveAs2(os.path.abspath(acc), FileFormat=16)    # 16 = wdFormatXMLDocument (.docx)
        doc.Close(False)
        return acc, {"accept": f"accepted({n_rev}处)" if n_rev else "accepted(0处)"}
    except Exception as e:
        return src, {"accept": f"fail({type(e).__name__}: {str(e)[:80]})"}
    finally:
        try:
            word.Quit()
        except Exception:
            pass


def stage1(src, wd, cfg):
    """前置手术 → conv docx（缓存：stage1/s4_<stem>.docx；s4_=Word整块截图版手术产物）

    主路径：shapesnap（Word COM）把组合/含字浮动形状整块截图成内嵌图片——小看板的
    边框/小人/高亮/排版 100% 保留，LO 再无栅格化错位/出血。.doc 也走 Word（读+存 docx
    一步到位），LO 预转降级为备用。回落路径：Word 不可用 → 旧 LO 预转 + flatten 拆解。
    gate_src：闸门基线取截图后的文档（看板文字已烤进图，基线同步排除，L1 才公平）。
    """
    precheck(src)
    s1 = os.path.join(wd, "stage1")
    os.makedirs(s1, exist_ok=True)
    info = {"tabs": 0, "math": 0, "flat": None, "skipped": False}
    # P0 修订终稿化（2026-10-10）：第一步统一接受所有修订，后续全部阶段都在终稿上分析
    src, acc_info = accept_revisions(src, s1)
    info.update(acc_info)
    info["final_src"] = src
    stem = os.path.splitext(os.path.basename(src))[0]
    snap_path = os.path.join(s1, "snap_" + stem + ".docx")
    cached = os.path.join(s1, "s4_" + stem + ".docx")
    if os.path.exists(cached):
        # 2026-10-12 修"源文件改完重转还是旧货"（问题清单 #4）：s4_ 缓存此前只看存在性、
        # 不看新鲜度——上次转换失败留下的缓存，源文件改完照样命中。规则与 stage2 对齐：
        # 缓存比当前输入（修订终稿/原文件）旧（留 2s 容忍文件系统时间精度）→ 作废重做。
        try:
            stale = os.path.getmtime(cached) < os.path.getmtime(src) - 1
        except OSError:
            stale = False
        jpath = os.path.join(s1, "stage1.json")
        degraded = False
        if os.path.exists(jpath):
            jinfo = json.load(open(jpath, encoding="utf-8"))
            # 2026-10-14 实证：snap 失败走 flatten 回落产出的 s4_ 是降级货（看板图缺失、
            # L1/L2 闸门必挂），但 mtime 比 src 新——下轮照样命中，降级结果被接力固化。
            # 缓存命中前先查 snap_err：降级缓存一律作废重做（现场文件保留供排查）。
            degraded = bool(jinfo.get("snap_err"))
            if not degraded:
                return cached, jinfo
        if not degraded:
            print("  [stage1] 缓存 s4_ 比源文件旧（源已修改）→ 作废重做", flush=True)
        else:
            print("  [stage1] 缓存 s4_ 来自降级回落（上次 snap 失败）→ 作废重做", flush=True)
        for dead in (cached, jpath):
            if os.path.exists(dead):
                os.remove(dead)
    snap_ok = False
    if cfg["surgery"].get("shape_snap", True):
        try:
            from shapesnap import snap_floating
            info["snap"] = snap_floating(src, snap_path, os.path.join(s1, "kanban"),
                                         log=lambda m: print(m, flush=True))
            snap_ok = True
            info["gate_src"] = snap_path
            cur = snap_path
        except Exception as e:
            info["snap_err"] = f"{type(e).__name__}: {e}"
            print(f"  [stage1] Word 整块截图不可用，回落旧路径：{info['snap_err'][:120]}", flush=True)
    if snap_ok:
        if cfg["surgery"].get("tab_to_nbsp", True):
            info["tabs"] = surgery_tab(cur, cached)
        else:
            shutil.copy2(cur, cached)
    elif src.lower().endswith((".doc", ".rtf")):
        tmp_out = os.path.join(s1, "lo2docx")
        os.makedirs(tmp_out, exist_ok=True)
        for f in os.listdir(tmp_out):          # 清陈货，防 lo_convert 捡到上次产物
            os.remove(os.path.join(tmp_out, f))
        conv = lo_convert(src, tmp_out, _loprof(wd), fmt="docx")
        shutil.move(conv, snap_path)
        info["doc_preconverted"] = True
        cur = snap_path
        # 二次尝试：Word 拒收原 .doc 时，LO 预转出的 docx 通常能开——再试一次整块截图
        if cfg["surgery"].get("shape_snap", True):
            try:
                from shapesnap import snap_floating
                snap2 = os.path.join(s1, "snap2_" + stem + ".docx")
                info["snap"] = snap_floating(snap_path, snap2, os.path.join(s1, "kanban"),
                                             log=lambda m: print(m, flush=True))
                os.remove(snap_path); os.rename(snap2, snap_path)
                info["gate_src"] = snap_path
                snap_ok = True
            except Exception as e:
                info["snap_err2"] = f"{type(e).__name__}: {e}"
                print(f"  [stage1] 二次截图仍失败，走 flatten：{info['snap_err2'][:120]}", flush=True)
        if snap_ok:
            if cfg["surgery"].get("tab_to_nbsp", True):
                info["tabs"] = surgery_tab(cur, cached)
            else:
                shutil.copy2(cur, cached)
        else:
            shutil.copy2(snap_path, cached)   # 截图仍失败：LO 预转产物继续走 flatten
    elif src.lower().endswith(".docx"):
        if cfg["surgery"].get("tab_to_nbsp", True):
            info["tabs"] = surgery_tab(src, cached)
        else:
            shutil.copy2(src, cached)
        cur = cached
    else:
        shutil.copy2(src, cached)
        info["skipped"] = True   # 未知格式：原样透传给 LO
    if not info["skipped"]:
        if not snap_ok:
            # 组合图形拆解（wpg 小看板：防 LO 整体栅格化导致错位/出血/丢字）
            try:
                from surgery_groups import flatten_groups
                info["flat"] = flatten_groups(cached)
            except Exception as e:
                info["flat_err"] = f"{type(e).__name__}: {e}"
        if cfg["surgery"].get("math_hd", True):
            try:
                from v13b_math_hd import export_math_pngs, surgery_docx
                pngs = export_math_pngs(cached, os.path.join(s1, "math"))
                out2 = cached + ".math.docx"
                _, info["math"] = surgery_docx(cached, out2, pngs)
                os.remove(cached); os.rename(out2, cached)
            except Exception as e:
                info["math_err"] = f"{type(e).__name__}: {e}"
    json.dump(info, open(os.path.join(s1, "stage1.json"), "w", encoding="utf-8"), ensure_ascii=False)
    return cached, info

def stage2(conv_src, wd, cfg):
    """LO→HTML（缓存：html_v4/<name>.html；v4=Word整块截图后的手术产物，旧缓存已失效）"""
    s2 = os.path.join(wd, "html_v4")
    os.makedirs(s2, exist_ok=True)
    hit = [f for f in os.listdir(s2) if f.endswith(".html")]
    if hit:
        hp = os.path.join(s2, hit[0])
        # 2026-10-10 修缓存永不过期事故：stage1 截图手术重跑后 conv_src 变新，
        # 旧 HTML 仍被复用（智慧政法看板错位查了半天其实是陈货）。
        # 规则：HTML 比 conv_src 旧（留 2s 容忍文件系统时间精度）→ 作废重转。
        if os.path.getmtime(hp) >= os.path.getmtime(conv_src) - 2:
            return hp, True
        os.remove(hp)
        print("  [stage2] 缓存 HTML 比 stage1 产物旧 → 作废重转", flush=True)
    htm = lo_convert(conv_src, s2, _loprof(wd))
    return htm, False

def build_numbering_map(docx_path):
    """从 docx 的 numbering.xml 按 Word 真实语义算出每个编号段的字面编号。

    LO 导出的自动编号是"一段一 ol"，start 续号还会错位（渲染成 1、1、2…），
    且序号后带 Word 的 tab 大空隙。本函数产出 {规范化段文本: [字面编号,...]}，
    cleanup 拿它把 ol/li 拍平成"1、xxx"式普通段落，阅读器零依赖。
    语义：按文档序遍历带 numPr 的段；同一 numId 内 ilvl=i 计数 +1 并复位更深层；
    编号 = lvlText 的 %n 替换为第 n 级计数器（含 w:start 起始值）。
    """
    try:
        z = zipfile.ZipFile(docx_path)
    except Exception:
        return {}
    names = z.namelist()
    if "word/document.xml" not in names:
        return {}
    num2abs, absfmt = {}, {}
    if "word/numbering.xml" in names:
        nx = z.read("word/numbering.xml").decode("utf-8", "ignore")
        num2abs = dict(re.findall(r'<w:num w:numId="(\d+)"[^>]*>\s*<w:abstractNumId w:val="(\d+)"', nx))
        for m in re.finditer(r'<w:abstractNum w:abstractNumId="(\d+)"[^>]*>(.*?)</w:abstractNum>', nx, re.S):
            aids, body = m.group(1), m.group(2)
            for lv in re.finditer(r'<w:lvl w:ilvl="(\d+)"[^>]*>(.*?)</w:lvl>', body, re.S):
                ilvl, lb = lv.group(1), lv.group(2)
                t = re.search(r'<w:lvlText w:val="([^"]*)"', lb)
                s = re.search(r'<w:start w:val="(\d+)"', lb)
                if t:
                    absfmt[(aids, ilvl)] = (t.group(1), int(s.group(1)) if s else 1)
    root = etree.fromstring(z.read("word/document.xml"))
    W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
    counters, marker_map = {}, {}
    for p in root.iter(W + "p"):
        numpr = p.find(f"{W}pPr/{W}numPr")
        if numpr is None:
            continue
        nid_el = numpr.find(W + "numId")
        if nid_el is None:
            continue
        nid = nid_el.get(W + "val")
        if not nid or nid == "0":          # numId=0 = 无编号
            continue
        ilvl_el = numpr.find(W + "ilvl")
        try:
            i = int(ilvl_el.get(W + "val")) if ilvl_el is not None else 0
        except (TypeError, ValueError):
            i = 0
        fmt = absfmt.get((num2abs.get(nid, ""), str(i)))
        if not fmt:
            continue
        lvltext, start = fmt
        c = counters.setdefault(nid, {})
        c[i] = c.get(i, start - 1) + 1
        for j in range(i + 1, 10):         # 复位更深层
            c.pop(j, None)
        marker = re.sub(r"%(\d)", lambda m2: str(c.get(int(m2.group(1)) - 1, 1)), lvltext)
        if marker and len(marker) == 1 and 0xE000 <= ord(marker) <= 0xF8FF:
            marker = "•"                   # Symbol/Wingdings 私有区项目符号 → 常规圆点
        txt = re.sub(r"\s+", "", "".join(p.itertext()))
        if txt and marker:
            marker_map.setdefault(txt, []).append(marker)
    return marker_map

def extract_list_numbers(docx_path):
    """Word COM 读编号权威值：Range.ListFormat.ListString 就是 Word 渲染出的编号串。
    返回全文段落序列 [(规范化文本, ListString|'")]（文档序，含无编号段）——cleanup 端
    用"游标+fuzzy"对齐 li 与段落：修订文档的 Range.Text 与 LO 渲染文本存在字词差异
    （修订显示口径不同），精确匹配不可靠；Word 无编号的段（ListString=''）不能凭空加号。"""
    import win32com.client as wc
    word = wc.DispatchEx("Word.Application")
    word.Visible = False
    word.DisplayAlerts = 0
    try:
        try:
            word.FileValidation = 1
        except Exception:
            pass
        doc = word.Documents.Open(os.path.abspath(docx_path), ReadOnly=True, AddToRecentFiles=False)
        seq = []
        try:
            for para in doc.Content.Paragraphs:
                ls = para.Range.ListFormat.ListString or ""
                if len(ls) == 1 and 0xE000 <= ord(ls) <= 0xF8FF:
                    ls = "•"               # Symbol/Wingdings 私有区项目符号 → 常规圆点
                # Range.Text 含段落符 \r、单元格符 \x07、脚注引用 \x02、域代码 \x13-\x15，全部剥掉
                t = re.sub(r"[\x02\x07\x13-\x15\s]+", "", para.Range.Text)
                seq.append([t, ls])
        finally:
            doc.Close(False)
    finally:
        word.Quit()
    return seq

def numbering_map_for(gate_src, wd):
    """编号序列（带缓存，v2）：优先 Word COM 权威值，COM 失败回落 numbering.xml 模拟。
    缓存 <wd>/stage1/listmap.json，按基线 docx 路径+mtime 失效。"""
    cache = os.path.join(wd, "stage1", "listmap.json")
    try:
        mt = os.path.getmtime(gate_src)
    except OSError:
        mt = None
    if os.path.exists(cache):
        try:
            j = json.load(open(cache, encoding="utf-8"))
            if j.get("v") == 2 and j.get("src") == gate_src and j.get("mtime") == mt:
                return j.get("seq") or [], j.get("via", "cache")
        except Exception:
            pass
    try:
        seq, via = extract_list_numbers(gate_src), "word-com"
    except Exception as e:
        d = build_numbering_map(gate_src)          # dict {文本: [编号]}
        seq = [[k, m] for k, ms in d.items() for m in ms]
        via = f"xml-fallback({type(e).__name__})"
    try:
        os.makedirs(os.path.dirname(cache), exist_ok=True)
        json.dump({"v": 2, "src": gate_src, "mtime": mt, "via": via, "seq": seq},
                  open(cache, "w", encoding="utf-8"), ensure_ascii=False)
    except Exception:
        pass
    return seq, via

def stage3(htm, wd, epub_path, unpack_dir, book_id, cfg, numbering_map=None):
    """cleanup + 切章 + 图片压缩 + 样式注入 + EPUB 组包 + 解包（解包只进 work 校验区，不占输出目录）"""
    st = cfg.get("styles", {})
    secs, img_refs, nav_entries = cleanup_and_split(htm, wd,
                    img_max=int(st.get("img_max_width", 580)),
                    numbering_map=numbering_map or {},
                    styles_cfg=st)   # h1/h2/h3 字号/加粗按配置内联固化（2026-10-12）
    htmldir = os.path.dirname(htm)
    img_srcmap = {}
    for ref in img_refs:
        p = os.path.join(htmldir, unquote(ref))
        if os.path.exists(p) and p not in img_srcmap:
            img_srcmap[p] = os.path.basename(p)
    # 图片压缩（2026-10-09）：大图降采样 + 不透明 PNG/GIF→JPEG，epub 瘦身适合网览
    renames = {}
    if st.get("img_optimize", True):
        from imgopt import optimize_images
        img_srcmap, renames = optimize_images(img_srcmap, os.path.join(wd, "imgopt"),
                                              pixel_max=int(st.get("img_pixel_max", 1000)),
                                              quality=int(st.get("img_jpeg_quality", 82)))
    build_epub(epub_path, book_id, secs, img_srcmap, make_css(cfg["styles"]),
              str(cfg.get("package", {}).get("epub_version", "3.0")), img_renames=renames,
              nav_entries=nav_entries or None)
    updir = unpack_dir
    old = updir + "_旧"
    if os.path.exists(updir):
        # 只改名腾位，绝不 rmtree（沙箱批量删除钩子会整进程强杀，后台批跑的全军覆没真凶）
        try:
            if os.path.exists(old):
                os.rename(old, old + "2")
            os.rename(updir, old)
        except OSError:
            pass
    with zipfile.ZipFile(epub_path) as z:
        z.extractall(updir)
    return dict(epub=epub_path, sections=len(secs), imgs=len(img_srcmap), unpacked=updir)

# ---------------------------------------------------------------- gates

def gate_l1l2(src, updir, cfg):
    """复用 l1l2_check 的口径做逐本闸门
    L2 口径（2026-10-09 修订）：产品图 < 期望图 = 丢图（硬失败）；
    产品图 > 期望图 = 形状/组合被 LO 栅格化的正常行为 → OK 带警告，不算 DIFF"""
    import l1l2_check as LC
    sp = LC.source_parts(src) if src.lower().endswith((".docx", ".docm")) else None
    if sp is None:
        return dict(l1=None, l2="SKIP(.doc/.rtf→V-2 三方基准口径)")
    ptxt, pimg, ptbl = LC.product_text_and_counts_from_dir(updir)
    frags = [p for p in sp["l1_frags"] if p and not LC.LIST_NO.match(p)]
    # 编辑待办备注（"麻烦将…加到此处小看板"）双侧同步剔除：cleanup 会删这类段，基线不剔会假失败
    from htmlpath_poc import NOTE_RE
    frags = [p for p in frags if not (len(p.strip()) <= 100 and NOTE_RE.fullmatch(p.strip()))]
    hit = [p for p in frags if p in ptxt]
    chars = sum(len(p) for p in frags)
    cov = sum(len(p) for p in hit)
    l1 = round(cov / chars * 100, 2) if chars else 100.0
    expect_img = sp["pic"] + sp["chart"] + sp["math"]
    tbl_ok = (ptbl == sp["tbl"])
    img_ok = (pimg >= expect_img)
    if img_ok and tbl_ok:
        l2 = "OK" if pimg == expect_img else f"OK(多{pimg - expect_img}图·形状栅格化)"
    else:
        l2 = f"DIFF img{pimg}/{expect_img} tbl{ptbl}/{sp['tbl']}"
    return dict(l1=l1, l2=l2)

def gate_epubcheck(epub, cfg):
    if not cfg["gates"].get("epubcheck"): return None
    jar = os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "work", "epubcheck", "epubcheck-5.2.1", "epubcheck.jar")
    r = subprocess.run(["java", "-jar", jar, epub], capture_output=True, timeout=300)
    out = (r.stdout + r.stderr).decode("utf-8", "ignore")
    m = re.search(r"Messages:\s*(\d+) fatals? / (\d+) errors? / (\d+) warnings?", out)
    return dict(fatal=int(m.group(1)), err=int(m.group(2)), warn=int(m.group(3))) if m else dict(raw=out[-300:])

# ---------------------------------------------------------------- main

def append_manifest(rec):
    mf = os.path.join(os.getcwd(), "work", "pipeline", "manifest.jsonl")
    os.makedirs(os.path.dirname(mf), exist_ok=True)
    with open(mf, "a", encoding="utf-8") as f:
        f.write(json.dumps(rec, ensure_ascii=False) + "\n")

def cmd_convert(args):
    cfg = load_config(args.config)
    root = cfg["paths"]["workdir"]
    if not os.path.isabs(root): root = os.path.abspath(root)
    inputs = []
    if os.path.isdir(args.input):
        for d in sorted(os.listdir(args.input)):
            p = os.path.join(args.input, d)
            if os.path.isdir(p):
                cands = [f for f in os.listdir(p) if f.lower().endswith((".docx", ".doc", ".rtf", ".docm"))]
                if cands: inputs.append(os.path.join(p, cands[0]))
    else:
        inputs = [args.input]
    for src in inputs:
        t0 = time.time()
        book_id = args.book_id or os.path.splitext(os.path.basename(src))[0]
        # 输出布局（2026-10-09 平铺版）：EPUB 直接进输出目录，不再每本建文件夹；
        # 默认输出目录 = 源文件所在目录；-o 显式指定优先
        out_root = os.path.abspath(args.out) if args.out else os.path.dirname(os.path.abspath(src))
        os.makedirs(out_root, exist_ok=True)
        wd = os.path.join(root, book_id)
        os.makedirs(wd, exist_ok=True)
        epub_path = os.path.join(out_root, f"转换_{book_id}.epub")
        unpack = os.path.join(wd, "epub校验")            # L1/L2 校验解包区，放 work 不污染输出
        rec = dict(book_id=book_id, src=src, stage=args.stage, ts=time.strftime("%F %T"),
                   status="running", out_dir=out_root)
        try:
            # 阶段1（幂等缓存）：all/1/2 需要；3 复用缓存（缺则补建）
            conv_src, s1info = stage1(src, wd, cfg)
            rec["stage1"] = s1info
            if args.stage in ("all", "1", "2"):
                if args.stage in ("all", "2"):
                    htm, cached = stage2(conv_src, wd, cfg)
                    rec["stage2_cache"] = cached
                if args.stage == "1":
                    rec["status"] = "done_stage1"
                    rec["artifacts"] = dict(conv_src=conv_src)
            else:
                htm, cached = stage2(conv_src, wd, cfg)
                rec["stage2_cache"] = cached
            if rec["status"] != "done_stage1":
                if args.stage == "2":
                    rec["status"] = "done_stage2"
                    rec["artifacts"] = dict(html=htm)
                else:
                    # 闸门/编号基线优先级：截图产物 → 修订终稿 → 原始文件（终稿化后三方必须同源）
                    base_src = s1info.get("gate_src") or s1info.get("final_src") or src
                    nmap, nvia = numbering_map_for(base_src, wd)
                    rec["numbering_via"] = nvia
                    r3 = stage3(htm, wd, epub_path, unpack, book_id, cfg, numbering_map=nmap)
                    rec.update(status="done", **r3)
                    if not args.no_gates:
                        base_src = s1info.get("gate_src") or s1info.get("final_src") or src
                        g = gate_l1l2(base_src, r3["unpacked"], cfg)
                        rec["gate_l1l2"] = g
                        ec = gate_epubcheck(r3["epub"], cfg)
                        if ec: rec["gate_epubcheck"] = ec
                        fail = (g["l1"] is not None and g["l1"] < cfg["gates"]["l1_min"]) or \
                               (cfg["gates"].get("l2_strict") and g.get("l2", "OK").startswith("DIFF"))
                        if fail:
                            rec["status"] = "failed"
                            rec["err"] = (f"闸门未过: L1={g['l1']}%（下限 {cfg['gates']['l1_min']}%）"
                                          f" L2={g['l2']}")
            # 大叔拍板 2026-10-10：转完即清中间产物——只要成功就删 stage1/html_v4/解包区，
            # 绝不留缓存（否则 Word 源文件改了，下轮还吃旧缓存转出旧货）。失败保留现场排查。
            if rec["status"] == "done":
                import shutil as _sh
                cleared = []
                for sub in ("stage1", "stage2", "html_v4", "html_v3", unpack, "imgopt"):
                    p = sub if os.path.isabs(sub) else os.path.join(wd, sub)
                    if os.path.isdir(p):
                        _sh.rmtree(p, ignore_errors=True)
                        cleared.append(os.path.basename(p.rstrip("\\/")))
                rec["cache_cleared"] = cleared
        except Exception as e:
            rec.update(status="failed", err=f"{type(e).__name__}: {e}")
        finally:
            rec["elapsed_s"] = round(time.time() - t0, 1)
            json.dump(rec, open(os.path.join(wd, "任务.json"), "w", encoding="utf-8"),
                      ensure_ascii=False, indent=1)
            append_manifest(rec)
        print(json.dumps(rec, ensure_ascii=False))

def cmd_check(args):
    cfg = load_config(args.config)
    src = None
    cands = [f for f in os.listdir(args.dir) if f.lower().endswith((".docx", ".docm"))]
    if cands: src = os.path.join(args.dir, cands[0])
    if src is None:
        print("输出目录中无源 docx，无法跑 L1/L2（.doc/.rtf 走 V-2 口径）"); return
    updir = os.path.join(args.dir, "epub解包")
    g = gate_l1l2(src, updir, cfg)
    epub = [f for f in os.listdir(args.dir) if f.endswith(".epub")]
    ec = gate_epubcheck(os.path.join(args.dir, epub[0]), {**cfg, "gates": {**cfg["gates"], "epubcheck": True}}) if epub else None
    print(json.dumps(dict(l1l2=g, epubcheck=ec), ensure_ascii=False))

def cmd_init(args):
    out = args.out or "config.yaml"
    if os.path.exists(out) and not args.force:
        print(f"{out} 已存在（--force 覆盖）"); return
    open(out, "w", encoding="utf-8").write(DEFAULT_CONFIG)
    print("生成:", out)

def main():
    ap = argparse.ArgumentParser(description="docx2epub A1′ 管线 CLI")
    sub = ap.add_subparsers(dest="cmd", required=True)
    p1 = sub.add_parser("init-config", help="生成默认配置")
    p1.add_argument("-o", "--out"); p1.add_argument("--force", action="store_true")
    p2 = sub.add_parser("convert", help="转换（三阶段 + 闸门）")
    p2.add_argument("input"); p2.add_argument("-c", "--config", default="config.yaml")
    p2.add_argument("-o", "--out", default=None,
                    help="输出目录（EPUB 平铺直入；缺省=源文件所在目录）")
    p2.add_argument("--stage", default="all", choices=["all", "1", "2", "3"])
    p2.add_argument("--book-id"); p2.add_argument("--no-gates", action="store_true")
    p3 = sub.add_parser("check", help="只跑闸门")
    p3.add_argument("dir"); p3.add_argument("-c", "--config", default="config.yaml")
    args = ap.parse_args()
    dict(init_config=cmd_init, convert=cmd_convert, check=cmd_check)[args.cmd.replace("-", "_")](args)

if __name__ == "__main__":
    main()
