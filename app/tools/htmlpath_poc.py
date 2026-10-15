# -*- coding: utf-8 -*-
"""A1′ 最小链路 PoC：前置手术 → LO→HTML → cleanup → 自建 EPUB3 组包
修复 V-11 六类问题中的 5 类：
  #1 切章   → 自组包按 h1 切分，不按渲染分页
  #2 表格   → CSS width:100% + border-collapse，无视口溢出
  #3 图表   → LO HTML 自动栅格化
  #4 文本框 → LO HTML 文字存活
  #6 tab    → 前置手术替换为空格（U+00A0）
#5 公式   → HTML 路径栅格化出现（低清，V-13b 待修）
输出：D:/project/epub/review_samples_html/<样本名>/
"""
import os, re, sys, shutil, subprocess, zipfile, time
from urllib.parse import unquote
from lxml import html as LH
from lxml import etree

ROOT = r"D:\project\epub"
S = os.path.join(ROOT, "tests")
OUT = os.path.join(ROOT, "review_samples_html")
TMP = os.path.join(ROOT, "work", "htmlpath_poc")
SOFFICE = r"C:\Program Files\LibreOffice\program\soffice.exe"
# 编辑待办备注（"麻烦将正文中脚注加到此处小看板中"）：cleanup 删段、L1 闸门同步剔除共用。
#   注意 .* 收尾——备注常以"小看板中"结尾，(?:小看板|加到此处) 后还有字，fullmatch 必须留尾巴
NOTE_RE = re.compile(r"(?:麻烦|请)将.{0,80}(?:小看板|加到此处).*", re.S)

CSS = """body{font-family:'宋体',SimSun,serif;line-height:1.7;margin:1em;color:#111}
h1{font-size:1.6em;text-align:center;margin:1.2em 0 .8em}
h2{font-size:1.4em;margin:1em 0 .6em}
h3{font-size:1.2em;margin:.8em 0 .5em}
h4,h5,h6{font-size:1.1em;margin:.7em 0 .4em}
p{margin:.35em 0}
img.fig{max-width:100%;height:auto}
table{width:100%;border-collapse:collapse;margin:.6em 0}
td,th{border:.5pt solid #444;padding:4px 6px;vertical-align:top}
caption{caption-side:top;font-weight:bold;padding:4px}
"""

def find_main(d):
    cands = [f for f in os.listdir(d) if f.lower().endswith((".docx", ".doc", ".rtf", ".docm"))]
    cands.sort(key=lambda f: (not f.startswith(os.path.basename(d)[:3]), len(f)))
    return cands[0] if cands else None

def surgery_tab(src_docx, dst_docx):
    """前置手术：把 <w:tab/> 替换为 4 个 U+00A0（保留 xml:space）。返回替换个数。"""
    zin = zipfile.ZipFile(src_docx)
    doc = zin.read("word/document.xml").decode("utf-8")
    n = doc.count("<w:tab/>")
    if n:
        doc = doc.replace("<w:tab/>", '<w:t xml:space="preserve">\u00a0\u00a0\u00a0\u00a0</w:t>')
    with zipfile.ZipFile(dst_docx, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = doc.encode("utf-8") if item.filename == "word/document.xml" else zin.read(item.filename)
            zout.writestr(item, data)
    zin.close()
    return n

def lo_convert(src, outdir, profile, fmt="html"):
    """LO 无头转换。fmt="html"→StarWriter HTML；fmt="docx"→MS Word 2007 XML（.doc 预转 docx 用，让老格式也享受 OOXML 手术）"""
    os.makedirs(outdir, exist_ok=True)
    # profile 路径经实测（2026-10-09）：裸空格→LO 秒退无产物；全 URL-quote（%3A/多字节%XX）→LO 卡死；
    # 只把空格换 %20 + 中文/冒号保持字面 → 正常。稳妥起见调用方应传无空格路径（见 pipeline.stage2 的 _loprof 哈希目录）。
    prof = str(profile).replace("\\", "/").replace(" ", "%20")
    if not prof.lower().startswith("file:///"):
        prof = "file:///" + prof
    filt = "html:HTML (StarWriter)" if fmt == "html" else "docx:MS Word 2007 XML"
    r = subprocess.run([SOFFICE, "--headless", "--norestore", f"-env:UserInstallation={prof}",
                    "--convert-to", filt, "--outdir", outdir, src],
                   capture_output=True, timeout=180)
    ext = ".html" if fmt == "html" else ".docx"
    h = [f for f in os.listdir(outdir) if f.endswith(ext)]
    if not h:
        tail = (r.stdout + b"\n" + r.stderr).decode("utf-8", "ignore").strip()[-200:]
        raise RuntimeError(f"LO 转换未产出 {'HTML' if fmt=='html' else 'DOCX'}（rc={r.returncode}）{tail}")
    return os.path.join(outdir, h[0])

def cleanup_and_split(html_path, workdir, img_max=580, numbering_map=None, styles_cfg=None):
    """解析 LO HTML → 清理 → 按 h1 切章。返回 (sections:[(title, xhtml_str)], images:[src_file], nav_entries:[(text, href)])
    img_max：图片内联锁定宽度的上限(px)，声明/实际宽超过它的图转纯流式 max-width:100%
    numbering_map：{规范化段文本: [Word 权威编号,...]}，用于 ol/li 拍平的字面编号（见 pipeline.build/extract）
    styles_cfg：config styles 块——h1/h2/h3 字号/加粗按配置内联固化（平台可能不加载 styles.css，
    固定样式=硬要求，2026-10-12 大叔拍板：目标不是还原而是按配置统一格式）"""
    raw = open(html_path, "rb").read()
    doc = LH.document_fromstring(raw.decode("utf-8", "replace"))
    body = doc.body
    # 1) 去垃圾：style/script 注释、LO 元数据 div（[if !xxx] 之类）
    for bad in body.xpath("//style|//script"):
        bad.getparent().remove(bad)
    # 1.1) 页眉/页脚整体删除（2026-10-09）：LO 把 Word 页眉页脚导出成 <div title="header"/"footer">，
    #   内含页眉横线段（空段落 border-bottom:1px solid）和页脚装饰形状图（"形状3" 554px 大 GIF）——
    #   阅读场景不需要，且是"顶部错误横线 / 大图溢出"的直接来源
    for hf in body.xpath("//div[@title='header']|//div[@title='footer']"):
        hf.getparent().remove(hf)
    # 1.6) HTML 注释节点全部删除：Word 批注经 LO 导出会变成 <!-- ... -->，成书绝不能带
    #   注意：注释节点的 tail 是同段后续正文，直接 remove 会连正文一起丢 → 先并回再删
    for c in body.xpath("//comment()"):
        p = c.getparent()
        if p is None:
            continue
        tail = c.tail or ""
        if tail:
            prev = c.getprevious()
            if prev is not None:
                prev.tail = (prev.tail or "") + tail
            else:
                p.text = (p.text or "") + tail
        p.remove(c)
    # 1.7) 编辑待办备注剔除：作者留在稿里的操作提示（如"麻烦将正文中脚注加到此处小看板中"），
    #   是给排版编辑看的话，不是书的内容——整段删除（仅匹配整段独立成段的备注，不碰含正常内容的段）。
    #   段内若带着看板截图（备注与图同段，图常嵌在 span 里），只剥文字、媒体必须保留，否则 L2 丢图
    def _strip_note_keep_media(el):
        for ch in list(el):
            if ch.tag not in ("img", "table") and not ch.xpath(".//img|.//table"):
                el.remove(ch)          # 纯文字分支（span/b/i…）：整支删掉
                continue
            if ch.tag not in ("img", "table"):
                _strip_note_keep_media(ch)   # 含图的深层分支：递归剥到底
            ch.tail = None
        el.text = None
    for p_el in list(body.xpath("//p")):
        txt = p_el.text_content().strip()
        if txt and len(txt) <= 100 and NOTE_RE.fullmatch(txt):
            if p_el.xpath(".//img|.//table"):
                _strip_note_keep_media(p_el)
            else:
                p_el.getparent().remove(p_el)
    # 1.8) 伪标题降级：h1~h6 文本超过 60 字必是正文误挂了标题样式（典型：整段"案例摘要"是 Heading 1，
    #   渲染成 1.6em 全大字，还污染 <title> 和目录）→ 降级为普通段落，保留段内原有加粗/字号标记
    for h in body.xpath("//h1|//h2|//h3|//h4|//h5|//h6"):
        if len(h.text_content().strip()) > 60:
            h.tag = "p"
    # 1.85) 标题识别修正（2026-10-12 问题清单 #1/#2 实证）：
    #   a) 空 h1 删除：Word 里误挂"标题1"样式的空段（内容只有 <br/>），LO 忠实导出成空 h1，
    #      切章时产出"未命名"垃圾章节（转换1/转换2 目录各多出 1-2 个"未命名"）；
    #      含图/表的"空标题"降级为 p 保留内容。
    #   b) "关键词：…"行降级为普通段落：作者常把关键词行误挂标题样式（转换1 实证与主标题同款
    #      style=2），渲染成大字居中还独占一章；转换2 里同位置的关键词行就是普通段落——统一按段落。
    for h in body.xpath("//h1|//h2|//h3|//h4|//h5|//h6"):
        txt = h.text_content().strip()
        if not txt:
            if h.xpath(".//img|.//table"):
                h.tag = "p"
            else:
                h.getparent().remove(h)
        elif re.match(r"^关键词[:：]", txt):
            h.tag = "p"
    # 1.8b) 列表拍平（2026-10-09 拍板）：LO 导出的自动编号是"一段一 ol"，start 续号错位
    #   （渲染成 1、1、2…），序号后还带 Word tab 的大空隙 → 按基线 docx 的 Word 权威编号
    #   （ListFormat.ListString，见 pipeline.extract_list_numbers）拍平成"1、xxx"式普通段落。
    #   对齐算法：COM 提取全文段序列（文档序，含无编号段）→ li 按文档序在窗口内 fuzzy 匹配
    #   （difflib ratio≥0.55；修订文档 Range.Text 与 LO 渲染有字词差异，精确匹配会漏）；
    #   匹配段的 ListString 为空（Word 无编号但 LO 硬导 li）→ 不加序号，宁可少不可错。
    #   两阶段：先按文档序逐 li 定 marker，再倒序做 DOM 改写（嵌套列表先内后外不乱游标）。
    num_seq = numbering_map if isinstance(numbering_map, list) else []
    from difflib import SequenceMatcher
    seq_i = 0
    def _seq_match(txt):
        nonlocal seq_i
        if not num_seq or not txt:
            return None
        best, best_r, best_j = "", 0.0, -1
        for j in range(seq_i, min(len(num_seq), seq_i + 200)):
            st = num_seq[j][0]
            if not st:
                continue
            r = SequenceMatcher(None, txt, st).ratio()
            if r > best_r:
                best, best_r, best_j = num_seq[j][1], r, j
        if best_r < 0.55:
            return None
        seq_i = best_j + 1
        return best or None          # 命中段但 Word 无编号 → None（不加号）
    li_plan = []                     # (li, marker) 文档序
    plan_by_lst = {}                 # id(lst) -> [(li, marker)...] 列表内文档序
    lst_order = []                   # 列表出现顺序（文档序）
    for lst in body.xpath("//ol|//ul"):
        if id(lst) not in plan_by_lst:
            plan_by_lst[id(lst)] = []
            lst_order.append(lst)
        for li in lst.xpath("./li"):
            marker = _seq_match(re.sub(r"\s+", "", li.text_content() or ""))
            li_plan.append((li, marker))
            plan_by_lst[id(lst)].append((li, marker))
    # DOM 改写（2026-10-14 修表格内列表乱序）：列表按文档序倒序处理（嵌套先内后外），
    # 但列表内 li 必须按文档序「正序」addprevious(lst) 逐个前插——原实现倒序搬移 +
    # 每次重算 outer.index(lst)，先搬的段落在后续搬移时被压到后面（双重反转），
    # 表1 单元格内 2. 排到 1. 前面（案例集 23-26 页表格错乱实证）。
    for lst in reversed(lst_order):
        outer = lst.getparent()      # 插到 ol 的父层级（插进 ol 内会被"删空列表"连锅端）
        if outer is None:
            continue
        for li, marker in plan_by_lst[id(lst)]:
            inner_ps = li.xpath("./p")
            if not inner_ps:
                np = li.makeelement("p", {})
                if li.text:
                    np.text = li.text
                for ch in list(li):
                    np.append(ch)
                inner_ps = [np]
            for p_el in inner_ps:
                lst.addprevious(p_el)     # 逐个前插到 ol 之前，正序保序
            if marker:
                first = inner_ps[0]
                mspan = first.makeelement("span", {})
                mspan.text = marker
                first.insert(0, mspan)
                if first.text:
                    # LO 的 p 首文本常带 \r\n\t 排版空白，编号后不能留（大叔拍板：序号后无大空隙）
                    mspan.tail = first.text.lstrip(" \t\r\n\u3000\u00a0")
                    first.text = None
            lst.remove(li)
    for lst in body.xpath("//ol|//ul"):
        if not lst.xpath("./li"):
            lst.getparent().remove(lst)
    # 1.5) 内联样式黑名单（LO 残留，破坏统一排版）：
    #   line-height 0.6cm~150% 各段打架 → 行距问题；orphans/widows/page-break-* → EPUB 无意义
    DROP_RE = re.compile(r"(?:^|;)\s*(?:line-height|orphans|widows|page-break-before|page-break-after)\s*:[^;\"]*", re.I)
    for el in body.iter():
        st = el.get("style")
        if st and DROP_RE.search(st):
            el.set("style", DROP_RE.sub(lambda m: ";" if m.group(0).startswith(";") else "", st).strip("; "))
    # 1.9) 假表格拆解：整表除图外无任何有效文本（只剩零宽符/空白）= LO 给浮动图/看板图套的布局壳。
    #   MuPDF 表格布局会按图声明宽把版心撑爆（实测 td 内 max-width:100% 循环引用失效、图原样画出），
    #   浏览器里这种表也没有表格语义 → 拆表，图并入居中段落（闸门表格计数同步剔除假表）
    _INVIS = "\u200d\ufeff\xa0 \t\r\n"
    for tbl in body.xpath("//table"):
        if tbl.xpath(".//table"):
            continue
        if tbl.text_content().strip(_INVIS) == "" and tbl.xpath(".//img"):
            # makeelement：etree.Element 造出的 _Element 无 text_content，进 body 后
            # 后续环节（4.7 缩进等）调用会崩（2026-10-12 实证）
            wrap = body.makeelement("p", {})
            wrap.set("style", "text-align:center")
            for img in tbl.xpath(".//img"):
                wrap.append(img)   # lxml 自动从原父节点摘下
            tbl.getparent().replace(tbl, wrap)
    # 2) img：统一 class=fig、去对齐属性、收集图片；
    #    保留 LO 声明的显示宽度 → 写入内联 style（平台侧不加载样式表也不出血）
    images = []
    for img in body.xpath("//img"):
        src = img.get("src", "")
        img.set("src", src)
        img.set("class", "fig")
        for a in ("align", "border", "hspace", "vspace", "name"):
            img.attrib.pop(a, None)
        w = (img.get("width") or "").strip()
        for a in ("width", "height"):
            img.attrib.pop(a, None)
        # 实际像素（防放大 + 封顶判定都要用）
        nat = None
        if src:
            try:
                from PIL import Image as _Img
                with _Img.open(os.path.join(os.path.dirname(html_path), unquote(src))) as im:
                    nat = im.size[0]
            except Exception:
                nat = None
        decl = float(w) if re.fullmatch(r"\d+(\.\d+)?", w) else None
        # 声明宽 > 实际像素 = 拉伸放大（LO 拆组后声明是组合 extent，不可信）→ 按实际像素
        if decl is not None and nat is not None and decl > nat:
            decl = nat
        px = decl if decl is not None else nat
        if px is not None and nat is not None:
            px = min(px, nat)   # 永不放大
        # 强制最佳宽度（2026-10-09 拍板）：超过版心上限封顶到 img_max，不超过的不放大，锁死 px 宽
        #   （max-width:100% 内联兜底，窄视口等比缩）
        # 【2026-10-09b Sumatra/MuPDF 出血修复】实测 MuPDF 排版引擎（SumatraPDF 内核同源）：
        #   display:block 的 img 只认 width、完全无视 max-width（553px 在 334pt 版心上原样画出 → 出血），
        #   且 margin:auto 居中也无效；改为行内 img 后 width/max-width/段落 text-align:center 全部生效。
        #   → 统一走行内图，独占段落的图改用宿主段落 text-align:center 居中（浏览器视觉不变）。
        inline = "max-width:100%;height:auto"
        # 大叔 2026-10-10d 拍板：自然宽超过版心上限的图（snap 看板 1000-1200px、LO 内容图 1000px）
        # 转纯流式 max-width:100% 跟阅读窗走（智慧政法"图还是小"实证：锁 554px 在宽窗里不跟手）；
        # 自然宽在版心内的小图才锁死像素（防放大糊掉）。
        if px is not None and px >= 20 and nat is not None and nat <= img_max:
            inline = f"width:{int(px)}px;" + inline
        # 独占段落的图居中（行内小图不动，免破坏文字流）
        cont = img.getparent()
        for _ in range(3):
            if cont is None or cont.tag in ("p", "div"):
                break
            cont = cont.getparent()
        standalone = (cont is not None and cont.tag in ("p", "div")
                      and not "".join(cont.itertext()).strip())
        old_st = (img.get("style") or "").strip().rstrip(";")
        # 剥掉 LO 残留的 width/height 声明（cm/pt 绝对单位会盖掉我们的封顶，且高度写死破纵横比）
        old_st = re.sub(r"(?:^|;)\s*(?:width|height)\s*:[^;\"]*",
                        lambda m: ";" if m.group(0).startswith(";") else "", old_st, flags=re.I).strip("; ")
        img.set("style", (inline + (";" + old_st if old_st else "")))
        if standalone:
            # 宿主段落合并 text-align:center（保留原 margin 等声明；已是 center 不重复加）
            ta_re = re.compile(r"text-align\s*:\s*\w+", re.I)
            st = (cont.get("style") or "").strip().rstrip(";")
            if ta_re.search(st):
                st = ta_re.sub("text-align:center", st)
            else:
                st = (st + ";" if st else "") + "text-align:center"
            cont.set("style", st)
        # 图+图题/图文混排的拆段统一由 4.65 处理（此处的旧 br 方案已废弃，避免残留多余 <br/>）
        if src:
            images.append(src)
    # 3) 表格固定样式内联（2026-10-14 大叔："表格还是完全不对"——根因与 4.9 标题同款：
    #    平台不加载 styles.css，此前"去 width 属性，CSS 统一管"的假设在平台侧不成立，
    #    表格裸奔 → 无边框/无列宽约束/列随内容乱挤。边框、列宽比例必须烤进标签。
    #    统一格式原则：边框/内边距按管线统一值；列宽保留 LO 导出的百分比（源文档列比例），
    #    px 绝对宽度丢弃（表格 width:100% 自适应屏宽）。
    for t_el in body.xpath("//table"):
        t_el.set("style", "width:100%;border-collapse:collapse;margin:.6em 0")
        for a in ("width", "cellpadding", "cellspacing", "border", "dir", "bgcolor", "align"):
            t_el.attrib.pop(a, None)
    for c_el in body.xpath("//td|//th"):
        w = c_el.get("width") or ""
        wst = "width:" + w + ";" if w.endswith("%") else ""   # 列宽比例（首行 td 携带，作用全列）
        c_el.set("style", wst + "border:0.5pt solid #444;padding:4px 6px;vertical-align:top")
        for a in ("width", "height", "valign", "bgcolor"):
            c_el.attrib.pop(a, None)
    for tr in body.xpath("//tr"):
        for a in ("width", "height", "style", "valign", "bgcolor"):
            tr.attrib.pop(a, None)
    # 4) 去 <font> 解包（固定样式，不留字体垃圾）
    for f in body.xpath("//font"):
        f.drop_tag()
    # 4.5) 剥离文本节点中的 CR（LO 残留，序列化会变 &#13;）
    for el in body.iter():
        if el.text and "\r" in el.text:
            el.text = el.text.replace("\r", "")
        if el.tail and "\r" in el.tail:
            el.tail = el.tail.replace("\r", "")
    # 4.6) EPUB3 合规清理（EPUBCheck RSC-005 全家桶）
    # a) colgroup/col 整体丢弃（LO 表格列定义，HTML5 不可带 width）
    for cg in body.xpath("//colgroup|//col"):
        cg.getparent().remove(cg)
    # b) sdfield（页码域等）在 EPUB 无意义，丢元素
    for sf in body.xpath("//sdfield"):
        sf.getparent().remove(sf)
    # c) a[name] → a[id]（LO 脚注锚点用 name，XHTML 断链根源）
    for a_el in body.xpath("//a[@name][not(@id)]"):
        a_el.set("id", a_el.get("name"))
        del a_el.attrib["name"]
    # c2) Word 书签空锚点清理（2026-10-10 拍板 + 2026-10-14 补丁）：
    #     空锚点（无 href 无内容）原样全清会把目录超链接的目标书签（_TocXXX）一起清掉，
    #     切章时 id_map 查无此锚 → 目录 href 全被剥（案例集问题#1 实证）。
    #     现规则：先收集全文 href="#..." 引用集；被引用的空锚点把 id 提升到宿主段落
    #     （阅读器不会再把空 <a> 列成书签，链接目标不丢），宿主已有 id 时原锚保留；
    #     未被引用的（_GoBack/OLE_LINK 残留）照旧清除。
    _href_refs = {a_el.get("href")[1:] for a_el in body.xpath('//a[starts-with(@href,"#")]')}
    for a_el in body.xpath("//a[not(@href)]"):
        if "".join(a_el.itertext()).strip() or a_el.find(".//img") is not None:
            continue
        tail = a_el.tail
        prev = a_el.getprevious()
        par = a_el.getparent()
        if par is None:
            continue
        aid = a_el.get("id")
        if aid and aid in _href_refs:
            if not par.get("id") and par is not body:
                par.set("id", aid)          # id 提升到宿主段落（p/div），锚点即可删除
            elif par is body or par.get("id"):
                continue                    # 宿主已有 id / 锚点直接挂在 body：原锚保留
        if tail:
            if prev is not None:
                prev.tail = (prev.tail or "") + tail
            else:
                par.text = (par.text or "") + tail
        par.remove(a_el)
    # d) 文本框绝对定位 span 解包（其内部含块级元素，导致 p 嵌套 p）
    for sp in body.xpath("//span[contains(@style,'position: absolute')]"):
        sp.drop_tag()
    # e) align 属性 → text-align 样式（HTML4 残留）；已有 text-align 时替换而非追加
    #   （img 居中处理在前面已写入 text-align:center，追加会变成 center;left 并存、后者生效破坏居中）
    for el in list(body.iter()):
        a = el.attrib.pop("align", None)
        if a:
            ta = {"center": "text-align:center", "left": "text-align:left",
                  "right": "text-align:right", "justify": "text-align:justify"}.get(a.lower())
            if ta:
                st = (el.get("style") or "").strip().rstrip(";")
                if st and re.search(r"text-align\s*:", st, re.I):
                    st = re.sub(r"text-align\s*:\s*\w+", ta, st, flags=re.I)
                else:
                    st = (st + ";" if st else "") + ta
                el.set("style", st)
    # f) p 含块级子元素 → 解包外层 p（文本框/表格残留结构）
    BLOCK = {"p", "div", "table", "ul", "ol", "blockquote", "h1", "h2", "h3", "h4", "h5", "h6"}
    for _ in range(3):
        changed = False
        for p_el in list(body.iter("p")):
            if any(ch.tag in BLOCK for ch in p_el):
                p_el.drop_tag()
                changed = True
        if not changed:
            break
    # 4.65) 图文拆段（2026-10-10 六书实证）：图夹在文字中间（浮动环绕残留，如"建|图|融智医"）、
    #   段首内联（绝对定位图解包后）、段尾内联（snap 看板插在锚段末）都会"位置错乱"。
    #   统一拆出：img 独立成段 text-align:center；前后文本保持原序。"图+短题注(≤40字)"同段
    #   也统一拆段（题注独立居中行，复刻 Word 视觉）；LO 环绕残留 br[clear] 一并删除。
    for br_c in body.xpath("//br[@clear]"):
        t = br_c.tail
        if t:                              # tail 安全并入（直接 remove 会丢 br 后正文——lxml 教训）
            prev = br_c.getprevious()
            if prev is not None:
                prev.tail = (prev.tail or "") + t
            else:
                par = br_c.getparent()
                par.text = (par.text or "") + t
        br_c.getparent().remove(br_c)
    def _img_frags(p_el):
        """段落内按 img 切分的文本片段（文档序；img 前/间/后的纯文本，去空白）"""
        frags = [""]
        def _walk(e):
            if e.tag == "img":
                frags.append("")
                return
            if e.text:
                frags[-1] += e.text
            for ch in e:
                _walk(ch)
                if ch.tail:
                    frags[-1] += ch.tail
        _walk(p_el)
        return [re.sub(r"\s+", "", f) for f in frags]
    for p_el in list(body.iter("p")):
        if p_el.xpath("ancestor::td|ancestor::th|ancestor::p"):
            continue
        while True:
            imgs = p_el.xpath(".//img")
            if not imgs:
                break                  # 纯文字段/已拆完
            if len(imgs) > 1:
                # 多图段（2026-10-12 问题#5 实证：img1+图1题+img2 挤一段，图1题混在图中间）：
                # 仅当所有文字片段都是短题注（≤40字）时逐图拆段；长文图文混排不动
                if not all(len(f) <= 40 for f in _img_frags(p_el)):
                    break
            img = imgs[0]
            # 注意：兄弟元素的 tail 也是正文（如 <a _GoBack>长文tail</a><img/>——去 font 后长文全在
            # a.tail 里），itertext() 不含自身 tail，漏算会把"图文混排段"误判成"纯图段"跳过
            def _stx(s):
                return "".join(s.itertext()) + (s.tail or "")
            a_txt = (img.tail or "") + "".join(_stx(s) for s in img.itersiblings())
            b_txt = (p_el.text or "") + "".join(_stx(s) for s in img.itersiblings(preceding=True))
            b_txt, a_txt = re.sub(r"\s+", "", b_txt), re.sub(r"\s+", "", a_txt)
            if not b_txt and not a_txt:
                break                  # 纯图段（standalone 已居中，不动；break 防死循环）
            # 关键：img 的 tail 常是正文纯文本（如"建<img/>融智医综合服务平台……"的"融智医…"117 字），
            # 摘 img 前必须先用 span 把 tail 留在原位，否则正文整段丢失（L1 96.92% 事故，2026-10-10）
            t = img.tail
            if t and t.strip():
                # 用 makeelement 而非 etree.Element：后者造出 _Element（无 text_content），
                # 后续环节对 body 内所有元素调 text_content() 会 AttributeError（2026-10-12 实证）
                sp = p_el.makeelement("span", {})
                sp.text = t
                img.addnext(sp)
            img.tail = None
            fig = p_el.makeelement("p", {})
            fig.set("style", "text-align:center")
            fig.append(img)
            parent = p_el.getparent()
            if not b_txt:                 # 段首图：图段放文字段之前（保持"图在前"视觉序）
                parent.insert(parent.index(p_el), fig)
                if 0 < len(a_txt) <= 40:  # 图+短题注：题注行也居中（4.7 缩进例外靠 center 判定）
                    st = (p_el.get("style") or "").strip().rstrip(";")
                    if not re.search(r"text-align\s*:", st, re.I):
                        p_el.set("style", (st + ";" if st else "") + "text-align:center")
            else:                         # 段尾图/夹中间：图段放文字段之后（夹中间时前后文本自然合并回一段）
                parent.insert(parent.index(p_el) + 1, fig)
                if not a_txt and 0 < len(b_txt) <= 40:
                    st = (p_el.get("style") or "").strip().rstrip(";")
                    if not re.search(r"text-align\s*:", st, re.I):
                        p_el.set("style", (st + ";" if st else "") + "text-align:center")
            # 拆出一图后 p_el 里可能还有图（多图题注段），循环继续拆下一张

    # 4.7) 首行缩进定版（2026-10-09 16:30 拍板；16:40 扩大到各级标题）：
    #  a) 剥段首假缩进：原文用全角空格/半角空格/Tab 手敲的缩进全部清除（跨 span 递归剥到首个可见字符）
    #  b) 正文段 + 正文各级标题一律强制首行缩进 2 字符；含 <br> 的多行段改用
    #     padding-left:2em 整段块缩进（text-indent 只缩首行，详见下方 4.7 主循环注释）：
    #     标题含整段加粗的伪标题（引言/学习目标/一、二、…，本书库标题均为加粗段而非 h 标签）
    #     例外：居中段（大标题题行/图/图注，"开头大标题"全是居中题行）、表格内段、空段；
    #     h1 按样式表本就居中（文档主标题）不缩进，h2~h6 左对齐标题照常缩进
    #     （L1 口径 norm() 不含空白，剥字不影响对账）
    _LEAD_WS = "\u3000\u00a0\u2007\u202f\u200b\ufeff\u180e \t\r\n"
    def _strip_leading_ws(p_el):
        state = [True]
        def _walk(e):
            if state[0] and e.text:
                t = e.text.lstrip(_LEAD_WS)
                if t != e.text:
                    e.text = t
                if t:
                    state[0] = False
                    return
            for ch in list(e):
                if not state[0]:
                    break
                _walk(ch)
                if state[0] and ch.tail:
                    t = ch.tail.lstrip(_LEAD_WS)
                    if t != ch.tail:
                        ch.tail = t
                    if t:
                        state[0] = False
        _walk(p_el)
    ps = body.xpath("//p") + body.xpath("//h2|//h3|//h4|//h5|//h6")
    for p_el in ps:
        _strip_leading_ws(p_el)
    center_re = re.compile(r"text-align\s*:\s*center", re.I)
    # 图/表题注行（2026-10-12 问题#5 实证：转换2 "图2 乌兹光伏项目监控室内尽调"独立成段，
    # text-align:left + 缩进 11 字符 → 渲染成"缩进 2 字符"不居中）：
    # 短段（≤40字）且以 "图N/表N"+分隔符 开头 → 一律居中，不吃正文 2em 缩进
    CAPTION_RE = re.compile(r"^[图表]\s*\d+(?:[\s．.、：:]|$)")
    for p_el in ps:
        st0 = p_el.get("style") or ""
        # LO 悬挂/左缩进残留（如"（一）中欧班列背景"正文段 margin-left:2.91cm）。
        # 大叔拍板 2026-10-10：段落统一格式，不要还原原稿的怪缩进——margin-left
        # 一律清零，首行统一 2em。此前只清 hang 命中的段，导致正文段
        # margin-left:2.91cm 残留 + 首行 2em 双重缩进（区块链中欧班列段实证）。
        _strip_leading_ws(p_el)
        st = (p_el.get("style") or "").strip().rstrip(";")
        st = re.sub(r"(?:^|;)\s*text-indent\s*:[^;\"]*",
                    lambda m: ";" if m.group(0).startswith(";") else "", st, flags=re.I).strip("; ")
        st = re.sub(r"(?:^|;)\s*margin-left\s*:[^;\"]*",
                    lambda m: ";" if m.group(0).startswith(";") else "", st, flags=re.I).strip("; ")
        st = re.sub(r"(?:^|;)\s*padding-left\s*:[^;\"]*",
                    lambda m: ";" if m.group(0).startswith(";") else "", st, flags=re.I).strip("; ")
        st = re.sub(r";{2,}", ";", st).strip("; ")
        txt = re.sub(r"\s+", "", p_el.text_content())
        in_table = bool(p_el.xpath(".//table|ancestor::td|ancestor::th"))
        if (p_el.tag == "p" and not in_table and txt and len(txt) <= 40
                and CAPTION_RE.match(p_el.text_content().strip())):
            # 图/表题注行：强制居中（text-align:left 残留一并替换），不加缩进
            ta_re = re.compile(r"text-align\s*:\s*\w+", re.I)
            if ta_re.search(st):
                st = ta_re.sub("text-align:center", st)
            else:
                st = (st + ";" if st else "") + "text-align:center"
        elif (not center_re.search(st) and not in_table and txt):
            # 缩进统一规则（2026-10-12 大叔拍板"举一反三"，原理层而非特例）：
            # 除大标题/居中段/表格内段外一律缩进 2 字符；text-indent 只缩首行，
            # 段内凡有 <br> 分行（思考题 ABCD/诗行/地址分行等）后续行必顶格——
            # 多行段一律改用 padding-left:2em 整段块缩进（每行齐进齐退），禁用 text-indent。
            if p_el.xpath(".//br"):
                st = (st + ";" if st else "") + "padding-left:2em"
            else:
                st = (st + ";" if st else "") + "text-indent:2em"
        p_el.set("style", st)
    # 4.8) 一级标题提升（2026-10-12 问题清单 #2 实证）："一、二、三…"短段是案例文档的
    #   一级标题（源文件里是普通段落/自定义样式，不挂 Word Heading，LO 导出仍是 p）→
    #   全书一级标题格式不统一、目录缺失。统一提升为 h2（class 加 sec1），剥掉段内
    #   font-weight:normal 残留（否则压掉标题加粗）并参与 5) 切章。
    #   居中的序号段同样提升（2026-10-14 大叔"46页章节标题错误居中"实证：作者误把
    #   小节标题「三、…」设了 jc=center，旧逻辑跳过 → 被 4.87 当居中块首段误提升成
    #   文章标题）。序号段=小节标题是原理级判定，与对齐方式无关；提升时剥掉居中，
    #   与同篇其他小节标题统一为左对齐缩进格式（不忠实还原原稿怪排版）。
    _SEC1_RE = re.compile(r"^[一二三四五六七八九十]{1,3}、")
    for p_el in body.xpath("//p"):
        if p_el.xpath("ancestor::td|ancestor::th|.//img|.//table"):
            continue
        t = re.sub(r"\s+", "", p_el.text_content())
        if not t or len(t) > 50 or not _SEC1_RE.match(t):
            continue
        p_el.tag = "h2"
        p_el.set("class", ((p_el.get("class") or "") + " sec1").strip())
        st = (p_el.get("style") or "").strip().rstrip(";")
        st = re.sub(r"(?:^|;)\s*text-indent\s*:[^;\"]*",
                    lambda m: ";" if m.group(0).startswith(";") else "", st, flags=re.I)
        st = re.sub(r"(?:^|;)\s*text-align\s*:[^;\"]*",
                    lambda m: ";" if m.group(0).startswith(";") else "", st, flags=re.I)
        p_el.set("style", re.sub(r";{2,}", ";", st).strip("; "))
        for sp in p_el.iter("span"):
            spst = sp.get("style") or ""
            if not spst:
                continue
            spst = re.sub(r"(?:^|;)\s*font-weight\s*:[^;\"]*",
                          lambda m: ";" if m.group(0).startswith(";") else "", spst,
                          flags=re.I).strip("; ")
            if spst:
                sp.set("style", spst)
            else:
                del sp.attrib["style"]
    # 4.85) 全文大标题补挂（2026-10-12 大叔：8/23 大标题没按格式——源文件大标题是
    #   普通居中段，没挂 Word Heading，LO 导出成 p，全文无 h1，配置的 1.6em 主标题
    #   样式落空，大标题塌成正文）。规则：全文无 h1 时，取第一个非空居中短段
    #   （≤50字、无图/表/br）提升为 h1；已有 h1 的文档不动（1/2 源挂了 Heading），
    #   作者行等后续居中段不升级。必须先于 4.9 执行，提升的 h1 才吃得到配置样式。
    if not body.xpath("//h1"):
        for el in body.xpath("//p"):
            if el.xpath("ancestor::td|ancestor::th|.//img|.//table|.//br"):
                continue
            t = re.sub(r"\s+", "", el.text_content())
            if not t or len(t) > 50:
                continue
            if not center_re.search(el.get("style") or ""):
                continue
            el.tag = "h1"
            st = (el.get("style") or "").strip().rstrip(";")
            st = re.sub(r"(?:^|;)\s*text-indent\s*:[^;\"]*",
                        lambda m: ";" if m.group(0).startswith(";") else "", st, flags=re.I)
            el.set("style", re.sub(r";{2,}", ";", st).strip("; "))
            break
    # 4.87) 文章标题提升（2026-10-14 大叔：案例集一篇文档含多篇文章，各文章标题是
    #   居中段没挂 Heading，塌成正文小字——居中保留正确，缺标题格式）。
    #   规则（原理：居中连续块的「首段」是文章标题，紧随其后的居中短段是作者行）：
    #   居中短段（≤50字、无图/表/br、非图表题注）且前一个非空兄弟不是居中段
    #   （即它开启了一个新的居中块）→ 提升为 h2（4.9 给它配置字号/加粗，居中不缩进）。
    #   「目录」标签段、已有 h1 的 4.85 主标题不在此列（主标题已提升，不再重复）。
    def _prev_nonempty_sibling(el):
        s = el.getprevious()
        while s is not None:
            if s.tag == "p" and not re.sub(r"\s+", "", s.text_content()):
                s = s.getprevious()          # 空段/占位 <br> 段跳过
                continue
            return s
        return None
    # 图/表题注排除：源文件「图1企业级模型…」图N 后未必有空格/分隔符（CAPTION_RE 拦不住），
    # 4.87 只对"已是居中段"做排除，放宽为只要 图N/表N 开头即题注（案例集误提升 8 个图题实证）
    _CAP_ANY = re.compile(r"^[图表]\s*\d+")
    for el in body.xpath("//p"):
        if el.xpath("ancestor::td|ancestor::th|.//img|.//table|.//br"):
            continue
        t = re.sub(r"\s+", "", el.text_content())
        if not t or len(t) > 50 or t in ("目录", "目錄"):
            continue
        if _CAP_ANY.match(el.text_content().strip()):
            continue                          # 图/表题注行是 4.7 的居中段，不是文章标题
        if _SEC1_RE.match(t):
            continue                          # 序号段（一、二、…）是小节标题，走 4.8 sec1 路径
                                              # （2026-10-14 双保险：4.8 已提升居中序号段，此处兜底）
        if not center_re.search(el.get("style") or ""):
            continue
        prev = _prev_nonempty_sibling(el)
        if prev is not None and prev.tag in ("p", "h1", "h2", "h3") \
                and center_re.search(prev.get("style") or ""):
            continue                          # 居中块的后续段（作者行等），不升级
        el.tag = "h2"
        el.set("class", ((el.get("class") or "") + " sec2").strip())
    # 4.88) 目录自动识别 + 整段建链 + 导航锚点（2026-10-14 大叔需求）。
    #   仅当文中存在「目录」标签段时激活——没有目录的单独文章什么都不做（不建链不挂牌）。
    #   a) 识别：「目录/目 录」段为标签；其后按文档序的段落为条目候选，条目=文本（去尾部
    #      页码）能匹配标题锚点，或已有可解析的内链且短（≤60字）；连续 3 段不匹配 → 块结束。
    #   b) 整段建链：LO 按西文 run 边界把条目链接切断（「区块链</a><span>+</span>再贷款…」
    #      实证，只有前两个字可点）——统一重建：unwrap 条目内所有 <a>，除行尾页码外整段
    #      包进一个 <a href="#锚点">。锚点=原 href 目标（仍存在时）＞文本精确匹配＞唯一前缀匹配。
    #   c) 标签：目录标签段 id=toc、条目段 id=toc-NN（检测导航/跳转用）；标题无 id 则生成。
    def _uniq_id(base, used):
        if base not in used:
            used.add(base)
            return base
        k = 2
        while f"{base}-{k}" in used:
            k += 1
        used.add(f"{base}-{k}")
        return f"{base}-{k}"
    _used_ids = set(body.xpath("//*/@id"))
    head_ids = {}                     # 规范化标题文本 → 锚点 id（首个优先）
    for h in body.xpath("//h1|//h2|//h3"):
        ht = re.sub(r"\s+", "", h.text_content())
        if not ht:
            continue
        if not h.get("id"):
            h.set("id", _uniq_id("hd", _used_ids))
        head_ids.setdefault(ht, h.get("id"))
    _all_ids = set(_used_ids)
    toc_label = None
    _toc_nav = []                     # [(条目文本, 目标锚id, 条目段id)] → 5.6 解析为导航条目
    for p_el in body.xpath("//p"):
        if re.sub(r"\s+", "", p_el.text_content()) in ("目录", "目錄"):
            toc_label = p_el
            break
    if toc_label is not None:
        def _entry_target(p_el):
            """条目 → (锚点id, 是否条目)。宁可少不可错。"""
            full = re.sub(r"\s+", "", p_el.text_content())
            if not full or len(full) > 60:
                return None, False
            t_m = re.sub(r"\d{1,4}$", "", full)          # 去尾部页码
            # a) 已有可解析内链
            for a_el in p_el.xpath(".//a[@href]"):
                href = a_el.get("href", "")
                if href.startswith("#") and href[1:] in _all_ids:
                    return href[1:], True
            # b) 文本精确匹配（去页码 → 原文 两试）
            for cand in (t_m, full):
                if cand in head_ids:
                    return head_ids[cand], True
            # c) 唯一前缀匹配（目录条目常截断/带引导符，如「北京市分行…项目」vs 标题「…项目案例」）
            for cand in (t_m, full):
                if len(cand) >= 6:
                    hits = [v for k, v in head_ids.items() if k.startswith(cand)]
                    if len(hits) == 1:
                        return hits[0], True
            # d) 去编号前缀再试（「1. 标题」型目录）
            m = re.match(r"^(?:\d{1,2}[.、．]|[一二三四五六七八九十]{1,3}、)\s*", t_m)
            if m and head_ids.get(t_m[m.end():]):
                return head_ids[t_m[m.end():]], True
            return None, False
        toc_label.set("id", _uniq_id("toc", _used_ids))
        # 条目重建链接：切到段边界（页码节点边界不干净时宁多包不拆元素）
        def _relink_entry(p_el, target):
            for a_el in p_el.xpath(".//a[@href]"):
                a_el.drop_tag()
            full = "".join(p_el.itertext())
            m = re.search(r"\s+\d{1,4}\s*$", full)       # 页码必须与标题有空白分隔，防「2025年版」误切
            cut_raw = m.start() if m else len(full)
            segs, pos = [], 0
            t = p_el.text or ""
            if t:
                segs.append([pos, pos + len(t), "text", t, ""]); pos += len(t)
            for ch in p_el:
                ct = "".join(ch.itertext())
                tl = ch.tail or ""
                segs.append([pos, pos + len(ct), "elem", ch, tl]); pos += len(ct) + len(tl)
            bounds = sorted({s for s, e, k, n, l in segs} | {e for s, e, k, n, l in segs} | {len(full)})
            cut = next((b for b in bounds if b >= cut_raw), len(full))
            a_new = p_el.makeelement("a", {"href": "#" + target})
            p_el.insert(0, a_new)
            for s, e, kind, node, tl in segs:
                if s >= cut:
                    break
                if kind == "text":
                    # p_el.text 段：内段进 a_new.text，外段挂 a_new.tail（渲染在链接后）
                    a_new.text = (a_new.text or "") + node[:cut - s]
                    if e > cut:
                        a_new.tail = (a_new.tail or "") + node[cut - s:]
                    p_el.text = None
                else:
                    node.tail = None                     # tail 由本 seg 的 tl 统一重挂
                    a_new.append(node)
                    if tl:
                        t_start = e - len(tl)            # tail 起点坐标
                        if t_start >= cut:
                            a_new.tail = (a_new.tail or "") + tl   # node 是链接内最后一个节点，a_new.tail 恰在其后
                        else:
                            node.tail = tl[:cut - t_start] or None
                            if cut - t_start < len(tl):
                                a_new.tail = (a_new.tail or "") + tl[cut - t_start:]
        n_link = 0
        _n_entry = 0
        _miss = 0
        ps_all = body.xpath("//p")
        _start = ps_all.index(toc_label) + 1
        for p_el in ps_all[_start:]:
            if not re.sub(r"\s+", "", p_el.text_content()):
                continue                                 # 空段不计 miss
            tgt, is_entry = _entry_target(p_el)
            if not is_entry:
                _miss += 1
                if _miss >= 3:
                    break
                continue
            _miss = 0
            _n_entry += 1
            if not p_el.get("id"):
                p_el.set("id", _uniq_id(f"toc-{_n_entry:02d}", _used_ids))
            efull = re.sub(r"\s+", "", p_el.text_content())
            etext = re.sub(r"\d{1,4}$", "", efull)[:60] or efull[:60]
            _toc_nav.append((etext, tgt, p_el.get("id")))
            if tgt:
                _relink_entry(p_el, tgt)
                n_link += 1
    # 4.9) 标题固定样式内联（2026-10-12 大叔复核拍板：结构对了还不够，格式必须按配置统一）：
    #   a) 字号/加粗内联固化——h1/h2/h3 此前全靠 styles.css，平台侧不加载样式表时
    #      标题塌成正文大小（图片内联锁宽同理，固定样式=硬要求）
    #   b) 除大标题（h1，居中）外一律首行缩进 2 字符——4.8 提升的 h2 晚于 4.7 缩进环节
    #      且被剥过 text-indent，这里统一补齐（"一、二"标题缩进原则）
    #   c) 标题内 span 的 font-size/weight/family 内联残留一并剥掉（内联会压掉配置值）
    _HL = {"h1": ("h1_size", "h1_bold"), "h2": ("h2_size", "h2_bold"), "h3": ("h3_size", "h3_bold")}
    _sc = styles_cfg or {}
    for h in body.xpath("//h1|//h2|//h3"):
        size_key, bold_key = _HL[h.tag]
        size = _sc.get(size_key)
        bold = _sc.get(bold_key)
        st = (h.get("style") or "").strip().rstrip(";")
        st = re.sub(r"(?:^|;)\s*font-size\s*:[^;\"]*",
                    lambda m: ";" if m.group(0).startswith(";") else "", st, flags=re.I).strip("; ")
        if size:
            st = (st + ";" if st else "") + f"font-size:{size}"
        if bold is not None:
            st = re.sub(r"(?:^|;)\s*font-weight\s*:[^;\"]*",
                        lambda m: ";" if m.group(0).startswith(";") else "", st, flags=re.I).strip("; ")
            st = (st + ";" if st else "") + ("font-weight:bold" if bold else "font-weight:normal")
        if h.tag != "h1" and not center_re.search(st):
            # 与 4.7 同一原理（2026-10-12"举一反三"）：多行段（含 <br>）禁用
            # text-indent（只缩首行），改 padding-left:2em 整段块缩进
            if h.xpath(".//br"):
                if not re.search(r"padding-left\s*:", st, re.I):
                    st = (st + ";" if st else "") + "padding-left:2em"
            elif not re.search(r"text-indent\s*:", st, re.I) \
                    and not re.search(r"padding-left\s*:", st, re.I):
                st = (st + ";" if st else "") + "text-indent:2em"
        h.set("style", re.sub(r";{2,}", ";", st).strip("; "))
        for sp in h.iter("span"):
            spst = sp.get("style") or ""
            if re.search(r"font-(?:size|weight|family)", spst, re.I):
                spst = re.sub(r"(?:^|;)\s*font-(?:size|weight|family)\s*:[^;\"]*",
                              lambda m: ";" if m.group(0).startswith(";") else "", spst,
                              flags=re.I).strip("; ")
                if spst:
                    sp.set("style", spst)
                else:
                    del sp.attrib["style"]
    # 5) 按 h1 切章——只认 h1（大标题）；一、二、三 提升的 h2 只做样式/层级，
    #    不切章（2026-10-12 大叔：不要对一、二、三强行分页，全文连续输出）。
    #    第一个切点之前的内容（申报类别/空行等前言）并入第一章，不单独成章
    children = list(body)
    sections, preamble = [], []
    cur_title, cur_nodes = None, None
    for el in children:
        if el.tag == "h1":
            if cur_nodes:
                sections.append((cur_title, cur_nodes))
            cur_title = re.sub(r"\s+", " ", el.text_content()).strip() or "未命名"
            cur_nodes = [el]
        elif cur_nodes is None:
            preamble.append(el)
        else:
            cur_nodes.append(el)
    if cur_nodes:
        sections.append((cur_title, cur_nodes))
    if preamble:
        if sections:
            sections[0] = (sections[0][0], preamble + sections[0][1])
        else:
            sections = [(None, preamble)]
    if not sections:
        first_h1 = body.xpath("//h1[1]")
        title = first_h1[0].text_content().strip() if first_h1 else None
        sections = [(title, list(body))]
    # 5.2) 无 h1 文档的标题推导（大叔 2026-10-10：阅读器书签栏全是兜底"文档"，实证 23 本通病）：
    #   取正文前两个非空段拼合（主标题 + "——副标题"，如"探索架构转型实现安全可控——分布式微服务平台案例"）
    def _derive_title(nodes):
        got = []
        for n in nodes:
            if n.tag in ("p", "h1", "h2", "h3", "h4"):
                t = re.sub(r"\s+", "", n.text_content())
                if t:
                    got.append(t)
                    if len(got) >= 2:
                        break
        if not got:
            return None
        t = got[0]
        if len(got) > 1 and got[1].startswith(("—", "――", "–")):
            t += got[1]
        return t[:60] or None
    sections = [((t if t else _derive_title(nodes)), nodes) for t, nodes in sections]
    # 5.5) 跨章锚点改写：#frag → 所在 section 文件；找不到目标则降级为纯文字（防 RSC-012）
    id_map = {}
    for si, (t, nodes) in enumerate(sections):
        for el in nodes:
            for e in ([el] + list(el.iter())):
                i = e.get("id")
                if i and i not in id_map:
                    id_map[i] = si
    for si, (t, nodes) in enumerate(sections):
        for el in nodes:
            for a_el in el.iter("a"):
                href = a_el.get("href", "")
                if href.startswith("#"):
                    frag = href[1:]
                    tgt = id_map.get(frag)
                    if tgt is None:
                        del a_el.attrib["href"]
                    elif tgt != si:
                        a_el.set("href", f"section{tgt + 1:04d}.xhtml#{frag}")
    # 5.6) 导航条目解析（2026-10-14 大叔"导航呢？"实证）：4.88 只在内容层建了链接，
    #    EPUB 导航文件（nav.xhtml/toc.ncx）还只按切章生成——单 section 长文档侧边栏是空的。
    #    把识别出的目录条目解析成 (文本, "sectionNNNN.xhtml#锚点") 导出；无目标的条目
    #    回指条目段自身锚点（toc-NN）；无目录文档导出空表（= 不建导航，单独文章场景）。
    nav_entries = []
    for etext, tgt, eid in _toc_nav:
        href = None
        if tgt and tgt in id_map:
            href = f"section{id_map[tgt] + 1:04d}.xhtml#{tgt}"
        elif eid and eid in id_map:
            href = f"section{id_map[eid] + 1:04d}.xhtml#{eid}"
        nav_entries.append((etext, href))
    # 序列化为 xhtml 片段
    out_secs = []
    for title, nodes in sections:
        wrap = etree.Element("div")
        for n in nodes:
            wrap.append(n)
        xhtml = etree.tostring(wrap, method="xml", encoding="unicode")
        out_secs.append((title, xhtml))
    return out_secs, images, nav_entries

def make_xhtml_doc(title, inner):
    t = title or "文档"
    return f'''<?xml version="1.0" encoding="utf-8"?>
<!DOCTYPE html>
<html xmlns="http://www.w3.org/1999/xhtml" xmlns:epub="http://www.idpf.org/2007/ops" xml:lang="zh">
<head><meta charset="utf-8"/><title>{t}</title><link rel="stylesheet" type="text/css" href="../styles.css"/></head>
<body>
{inner}
</body>
</html>'''

def make_nav(titles):
    lis = "\n".join(f'      <li><a href="sections/section{i:04d}.xhtml">{(t or "文档")}</a></li>'
                    for i, (t, _) in enumerate(titles, 1))
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

def make_opf(uid, title, secs, img_files, has_nav=True):
    man = ['    <item id="nav" href="nav.xhtml" media-type="application/xhtml+xml" properties="nav"/>',
           '    <item id="css" href="styles.css" media-type="text/css"/>']
    for i, _ in enumerate(secs, 1):
        man.append(f'    <item id="s{i:04d}" href="sections/section{i:04d}.xhtml" media-type="application/xhtml+xml"/>')
    seen = {}
    for j, f in enumerate(img_files):
        ext = os.path.splitext(f)[1].lower().lstrip(".") or "gif"
        mt = {"png": "image/png", "jpg": "image/jpeg", "jpeg": "image/jpeg", "gif": "image/gif", "svg": "image/svg+xml"}.get(ext, "application/octet-stream")
        key = f"img{j:03d}"
        seen[f] = key
        man.append(f'    <item id="{key}" href="images/{os.path.basename(f)}" media-type="{mt}"/>')
    spine = "\n".join(f'    <itemref idref="s{i:04d}"/>' for i in range(1, len(secs) + 1))
    return f'''<?xml version="1.0" encoding="utf-8"?>
<package xmlns="http://www.idpf.org/2007/opf" version="3.0" unique-identifier="uid" xml:lang="zh">
  <metadata xmlns:dc="http://purl.org/dc/elements/1.1/">
    <dc:identifier id="uid">urn:uuid:{uid}</dc:identifier>
    <dc:title>{title}</dc:title>
    <dc:language>zh</dc:language>
    <meta property="dcterms:modified">{time.strftime("%Y-%m-%dT%H:%M:%SZ")}</meta>
  </metadata>
  <manifest>
{chr(10).join(man)}
  </manifest>
  <spine>
{spine}
  </spine>
</package>''', seen

def build_epub(epub_path, title, secs, img_srcmap):
    if os.path.exists(epub_path):
        os.remove(epub_path)
    with zipfile.ZipFile(epub_path, "w") as z:
        z.writestr(zipfile.ZipInfo("mimetype"), "application/epub+zip", zipfile.ZIP_STORED)
        z.writestr("META-INF/container.xml", '''<?xml version="1.0" encoding="utf-8"?>
<container version="1.0" xmlns="urn:oasis:names:tc:opendocument:xmlns:container">
  <rootfiles><rootfile full-path="OEBPS/content.opf" media-type="application/oebps-package+xml"/></rootfiles>
</container>''')
        import uuid
        opf, seen = make_opf(str(uuid.uuid4()), title, secs, list(img_srcmap.keys()))
        z.writestr("OEBPS/content.opf", opf)
        z.writestr("OEBPS/nav.xhtml", make_nav(secs))
        z.writestr("OEBPS/styles.css", CSS)
        for i, (t, inner) in enumerate(secs, 1):
            # 内文图片路径改写到 images/
            def rew(m):
                old = m.group(1)
                return f'src="../images/{os.path.basename(img_srcmap.get(unquote(old), unquote(old)))}"'
            inner2 = re.sub(r'src="([^"]+)"', rew, inner)
            z.writestr(f"OEBPS/sections/section{i:04d}.xhtml", make_xhtml_doc(t, inner2))
        for src, dst in img_srcmap.items():
            z.writestr(f"OEBPS/images/{os.path.basename(dst)}", open(src, "rb").read())

def process(name):
    srcdir = os.path.join(S, name)
    mainf = find_main(srcdir)
    g = os.path.join(OUT, name)
    if os.path.isdir(g):
        shutil.rmtree(g)
    os.makedirs(g)
    if not mainf:
        open(os.path.join(g, "说明.txt"), "w", encoding="utf-8").write("无主文件\n")
        return None
    src = os.path.join(srcdir, mainf)
    shutil.copy2(src, os.path.join(g, "原始_" + mainf))
    wd = os.path.join(TMP, name)
    shutil.rmtree(wd, ignore_errors=True)
    os.makedirs(wd)
    # 前置手术（仅 docx）：tab 替换 + 公式高清化替换
    tabs, math_replaced = 0, 0
    conv_src = src
    if mainf.lower().endswith(".docx"):
        try:
            sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
            from v13b_math_hd import export_math_pngs, surgery_docx
            math_pngs = export_math_pngs(src, os.path.join(TMP, name, "math"))
            conv_src = os.path.join(wd, "surgery_" + mainf)
            tabs, math_replaced = surgery_docx(src, conv_src, math_pngs)
        except Exception as e:
            conv_src = src
            tabs, math_replaced = -1, -1
    # LO→HTML
    prof = "file:///" + os.path.join(TMP, "_prof_" + name).replace("\\", "/")
    htm = lo_convert(conv_src, os.path.join(wd, "html"), prof)
    if not htm:
        open(os.path.join(g, "说明.txt"), "w", encoding="utf-8").write("LO 转换失败\n")
        return None
    # cleanup + 切章
    secs, img_refs, _nav = cleanup_and_split(htm, wd)
    htmldir = os.path.dirname(htm)
    img_srcmap = {}
    for ref in img_refs:
        p = os.path.join(htmldir, unquote(ref))
        if os.path.exists(p) and p not in img_srcmap:
            img_srcmap[p] = os.path.basename(p)
    # 组包
    epub = os.path.join(g, "转换_" + os.path.splitext(mainf)[0] + ".epub")
    build_epub(epub, name, secs, img_srcmap, nav_entries=_nav or None)
    # 解包
    with zipfile.ZipFile(epub) as z:
        z.extractall(os.path.join(g, "epub解包"))
    note = (f"源: {mainf}\n管线: 前置手术(tab替换={tabs}, 公式高清替换={math_replaced}) → LO→HTML → cleanup(img居中/表格百分比) → 自建EPUB3\n"
            f"切章: {len(secs)} 个 section（按 h1，无 h1 则单章）\n图片: {len(img_srcmap)} 张\n"
            f"修复对照: #1切章=自控 #2表格=CSS百分比 #3图表=栅格化 #4文本框=文字存活 #5公式=Word高清 #6tab=前置替换\n")
    open(os.path.join(g, "说明.txt"), "w", encoding="utf-8").write(note)
    shutil.rmtree(wd, ignore_errors=True)
    return dict(name=name, src=mainf, secs=len(secs), imgs=len(img_srcmap), tabs=tabs, math=math_replaced, ok=True)

def main():
    names = sys.argv[1:] or ["001_文字混合_低_规范", "004_图表混合_高_很脏", "005_图片混合_高_很脏",
                             "006_表格混合_高_很脏", "007_脚注混合_中_很脏", "008_文字混合_高_极端脏",
                             "013_列表混合_低_轻微脏", "015_混合_高_极端脏"]
    os.makedirs(OUT, exist_ok=True)
    rows = []
    for n in names:
        try:
            r = process(n)
        except Exception as e:
            r = dict(name=n, ok=False, err=f"{type(e).__name__}: {e}")
        rows.append(r)
        print(r)
    with open(os.path.join(OUT, "README_检查指引.md"), "w", encoding="utf-8") as f:
        f.write("# 肉眼检查样本包 · A1′ 修复版（HTML 路径 + 自建 EPUB3）\n\n")
        f.write("对照旧包 `review_samples\\`（EPUB 滤镜直转，有 6 类问题）。\n\n")
        f.write("| 样本 | 源 | 切章 | 图片 | tab替换 | 公式高清 | 状态 |\n|---|---|---|---|---|---|---|\n")
        for r in rows:
            if r and r.get("ok"):
                f.write(f"| {r['name']} | {r['src']} | {r['secs']} | {r['imgs']} | {r['tabs']} | {r['math']} | ✅ |\n")
            else:
                f.write(f"| {r['name'] if r else '?'} | - | - | - | - | - | ❌ {r.get('err','') if r else ''} |\n")
        f.write("\n重点检查：001 不再乱切章；006 十列表第10列可见+右边框；004 图表在；005/015 文本框文字在；008 tab 缩进保留；013 列表正常。\n")

if __name__ == "__main__":
    main()
