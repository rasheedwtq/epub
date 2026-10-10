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

def cleanup_and_split(html_path, workdir, img_max=580, numbering_map=None):
    """解析 LO HTML → 清理 → 按 h1 切章。返回 (sections:[(title, xhtml_str)], images:[src_file])
    img_max：图片内联锁定宽度的上限(px)，声明/实际宽超过它的图转纯流式 max-width:100%
    numbering_map：{规范化段文本: [Word 权威编号,...]}，用于 ol/li 拍平的字面编号（见 pipeline.build/extract）"""
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
    for lst in body.xpath("//ol|//ul"):
        for li in lst.xpath("./li"):
            li_plan.append((li, _seq_match(re.sub(r"\s+", "", li.text_content() or ""))))
    for li, marker in reversed(li_plan):
        lst = li.getparent()             # 所属 ol/ul
        if lst is None:
            continue
        outer = lst.getparent()          # 插到 ol 的父层级（插进 ol 内会被"删空列表"连锅端）
        if outer is None:
            continue
        idx = outer.index(lst)
        inner_ps = li.xpath("./p")
        if not inner_ps:
            np = li.makeelement("p", {})
            if li.text:
                np.text = li.text
            for ch in list(li):
                np.append(ch)
            inner_ps = [np]
        for p_el in reversed(inner_ps):
            outer.insert(idx, p_el)
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
            wrap = etree.Element("p")
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
        if px is not None and px >= 20:
            inline = f"width:{int(min(px, img_max))}px;" + inline
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
        if src:
            images.append(src)
    # 3) 表格：去 width 属性，CSS 统一管
    for t in body.xpath("//table"):
        for a in ("width", "cellpadding", "cellspacing", "border", "style"):
            t.attrib.pop(a, None)
    for td in body.xpath("//td|//th|//tr"):
        for a in ("width", "height", "style", "valign"):
            td.attrib.pop(a, None)
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
    # 4.7) 首行缩进定版（2026-10-09 16:30 拍板；16:40 扩大到各级标题）：
    #  a) 剥段首假缩进：原文用全角空格/半角空格/Tab 手敲的缩进全部清除（跨 span 递归剥到首个可见字符）
    #  b) 正文段 + 正文各级标题一律强制首行缩进 2 字符（text-indent:2em）：
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
    for p_el in ps:
        st = (p_el.get("style") or "").strip().rstrip(";")
        st = re.sub(r"(?:^|;)\s*text-indent\s*:[^;\"]*",
                    lambda m: ";" if m.group(0).startswith(";") else "", st, flags=re.I).strip("; ")
        if (not center_re.search(st)
                and not p_el.xpath(".//table|ancestor::td|ancestor::th")
                and p_el.text_content().strip()):
            st = (st + ";" if st else "") + "text-indent:2em"
        p_el.set("style", st)
    # 5) 按 h1 切章；第一个 h1 之前的内容（文档主标题块/前言）并入第一章，不单独成章
    children = list(body)
    sections, preamble = [], []
    cur_title, cur_nodes = None, None
    for el in children:
        if el.tag == "h1":
            if cur_nodes:
                sections.append((cur_title, cur_nodes))
            cur_title = el.text_content().strip() or "未命名"
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
    # 序列化为 xhtml 片段
    out_secs = []
    for title, nodes in sections:
        wrap = etree.Element("div")
        for n in nodes:
            wrap.append(n)
        xhtml = etree.tostring(wrap, method="xml", encoding="unicode")
        out_secs.append((title, xhtml))
    return out_secs, images

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
    secs, img_refs = cleanup_and_split(htm, wd)
    htmldir = os.path.dirname(htm)
    img_srcmap = {}
    for ref in img_refs:
        p = os.path.join(htmldir, unquote(ref))
        if os.path.exists(p) and p not in img_srcmap:
            img_srcmap[p] = os.path.basename(p)
    # 组包
    epub = os.path.join(g, "转换_" + os.path.splitext(mainf)[0] + ".epub")
    build_epub(epub, name, secs, img_srcmap)
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
