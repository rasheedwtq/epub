# -*- coding: utf-8 -*-
"""
DOCX → EPUB 测试样本生成器
依据：docx2epub_测试用例规格.md

产出 15 型样本（S1~S15），覆盖 30 个失败模式（FM）。
三层技术：
  1) python-docx      —— 主体内容
  2) 手工 OOXML (lxml) —— AI 造不出的陷阱：浮动图 anchor / 文本框 / 首字下沉 /
                          OMML 公式 / 修订 / 批注 / 脚注 / 域
  3) Word COM         —— .doc / .docm / .rtf / 加密 / 原生图表 / SmartArt

用法：
  python make_samples.py            # 纯 Python 部分（默认）
  python make_samples.py --com      # 追加 Word COM 部分（.doc/.docm/加密/图表）
"""
import os
import sys
import json
import shutil
import zipfile
import random
import argparse
from pathlib import Path

from docx import Document
from docx.shared import Pt, Cm, RGBColor, Emu
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from lxml import etree
from PIL import Image, ImageDraw

# ---------------------------------------------------------------- 基础设施
ROOT = Path(__file__).resolve().parents[2]
OUT = ROOT / "tests"
ASSETS = OUT / "_assets"

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
VML = "urn:schemas-microsoft-com:vml"
M = "http://schemas.openxmlformats.org/officeDocument/2006/math"
CT = "http://schemas.openxmlformats.org/package/2006/content-types"
PR = "http://schemas.openxmlformats.org/package/2006/relationships"

DATE = "2026-10-05T00:00:00Z"
AUTHOR = "SampleBot"

IDX = []          # 样本索引（写入 index.md）
GT = []           # ground truth：标题层级标注（供 V-4 自动评估）
CURRENT = {"name": ""}
random.seed(20261005)


def gt_mark(text, label, note=""):
    """登记一条 ground truth（按段落文本匹配，避免索引口径不一致）"""
    GT.append({"sample": CURRENT["name"], "text": text,
               "label": label, "note": note})


def w(tag):
    return "{%s}%s" % (W, tag)


def vv(tag):
    return "{%s}%s" % (VML, tag)


def mm(tag):
    return "{%s}%s" % (M, tag)


def cn():
    """自然中文 filler 文本（可判定、非 lorem）"""
    pool = [
        "本章讨论文档结构与样式的关系，重点在于层叠与继承的判定顺序。",
        "转换过程中最容易被忽略的是段落的语义标记，而非其视觉表现。",
        "当源文件仅以字号区分层级时，识别器必须依赖统计特征而非固定规则。",
        "表格的跨页行为在流式介质中会被重新定义，因此需要显式约束宽度。",
        "图片的真实像素与显示尺寸往往相差一个数量级，前者是浪费。",
        "脚注与尾注在语义上等价，但在呈现上应保持一致的可达性。",
        "修订痕迹若未被显式处理，会以删除线形式混入最终正文，影响可读性。",
        "批注属于协作元数据，发布形态下应当剥离而不应保留。",
    ]
    return random.choice(pool)


def newdoc():
    d = Document()
    st = d.styles["Normal"]
    st.font.name = "SimSun"
    st.font.size = Pt(10.5)
    st.element.rPr.rFonts.set(qn("w:eastAsia"), "宋体")
    st.paragraph_format.line_spacing = 1.5
    return d


def para(doc, text="", size=None, bold=False, align=None, style=None,
         font="宋体", indent_chars=None, space_after=None):
    p = doc.add_paragraph(style=style)
    if text:
        r = p.add_run(text)
        r.bold = bold
        if size:
            r.font.size = Pt(size)
        r.font.name = font
        r._element.get_or_add_rPr().get_or_add_rFonts().set(qn("w:eastAsia"), font)
    if align is not None:
        p.alignment = align
    if indent_chars:
        pPr = p._p.get_or_add_pPr()
        ind = OxmlElement("w:ind")
        ind.set(qn("w:firstLineChars"), str(int(indent_chars * 100)))
        pPr.append(ind)
    if space_after is not None:
        p.paragraph_format.space_after = Pt(space_after)
    return p


def heading_by_size(doc, text, pt, bold=True, align=WD_ALIGN_PARAGRAPH.CENTER, gt=None):
    """仅用字号表示标题（段落样式保持 Normal）—— 核心陷阱 A1/F1-1
    gt: 传入 h1/h2/h3/body/callout 时同步登记 ground truth"""
    if gt:
        gt_mark(text, gt, "纯字号标题（无 Heading 样式）")
    return para(doc, text, size=pt, bold=bold, align=align)


def add_img(doc, path, width_cm=None, height_cm=None, alt=None):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run()
    pic = r.add_picture(str(path), width=Cm(width_cm) if width_cm else None,
                        height=Cm(height_cm) if height_cm else None)
    if alt is not None:
        docPr = pic._inline.docPr
        if alt == "":                       # 故意留空 alt
            docPr.attrib.pop("descr", None)
        else:
            docPr.set("descr", alt)
    return pic


def figcaption(doc, text):
    p = para(doc, text, size=9, align=WD_ALIGN_PARAGRAPH.CENTER, font="楷体")
    return p


# ---------------------------------------------------------------- 素材图片
def make_assets():
    ASSETS.mkdir(parents=True, exist_ok=True)
    specs = {
        "huge.jpg": (3000, 4000, "RGB", 3),        # 超大原图，显示很小 (FM-26)
        "normal.png": (800, 600, "RGB", None),
        "alpha.png": (600, 400, "RGBA", None),      # 透明 (F4-7)
        "gray16.png": (600, 400, "I;16", None),     # 16-bit (F4-7 / FM-27)
        "icon.gif": (64, 64, "P", None),            # GIF
        "wide.jpg": (2400, 300, "RGB", None),       # 极端长宽
        "chart.png": (900, 700, "RGB", None),
        "tiny.png": (48, 48, "RGBA", None),
    }
    out = {}
    for name, (wd, ht, mode, q) in specs.items():
        fp = ASSETS / name
        img = Image.new("RGBA" if mode == "RGBA" else "RGB", (wd, ht),
                        (200, 220, 245, 255) if mode == "RGBA" else (200, 220, 245))
        d = ImageDraw.Draw(img)
        step = max(1, min(wd, ht) // 8)
        for i in range(0, max(wd, ht), step):
            d.line([(i, 0), (i, ht)], fill=(90, 130, 190), width=2)
            d.line([(0, i), (wd, i)], fill=(90, 130, 190), width=2)
        d.rectangle([10, 10, wd - 10, ht - 10], outline=(30, 60, 110), width=4)
        d.text((24, 24), f"{name} {wd}x{ht}", fill=(20, 40, 80))
        if mode == "RGBA":
            d.ellipse([wd * 0.3, ht * 0.3, wd * 0.7, ht * 0.7], fill=(255, 120, 0, 160))
        if mode == "I;16":
            img = img.convert("L").convert("I;16")
        if q:
            img.convert("RGB").save(fp, "JPEG", quality=q)
        else:
            img.save(fp)
        out[name] = fp
    return out


# ---------------------------------------------------------------- OOXML 手术
def inline_to_anchor(pic_inline, wrap="square", hpos="right", vpos="page"):
    """把 wp:inline 转为 wp:anchor（浮动图）—— F4-2 / FM-05"""
    drawing = pic_inline.getparent()
    inline = pic_inline
    anchor = OxmlElement("wp:anchor")
    for k, v in dict(distT="0", distB="0", distL="114300", distR="114300",
                     simplePos="0", relativeHeight="251658240", behindDoc="0",
                     locked="0", layoutInCell="1", allowOverlap="1").items():
        anchor.set(k, v)

    sp = OxmlElement("wp:simplePos"); sp.set("x", "0"); sp.set("y", "0")
    ph = OxmlElement("wp:positionH"); ph.set("relativeFrom", "column")
    po = OxmlElement("wp:posOffset"); po.text = "0"; ph.append(po)
    pv = OxmlElement("wp:positionV"); pv.set("relativeFrom", vpos)
    pvo = OxmlElement("wp:posOffset"); pvo.text = "0"; pv.append(pvo)
    wsq = OxmlElement("wp:wrapSquare"); wsq.set("wrapText", "bothSides")

    for tag in ("wp:extent", "wp:effectExtent"):
        el = inline.find(qn(tag))
        if el is not None:
            anchor.append(el)
    anchor.append(wsq)
    for tag in ("wp:docPr", "wp:cNvGraphicFramePr"):
        el = inline.find(qn(tag))
        if el is not None:
            anchor.append(el)
    g = inline.find(qn("a:graphic"))
    if g is not None:
        anchor.append(g)
    for el in (sp, ph, pv):
        anchor.insert(0, el)
    drawing.replace(inline, anchor)
    return anchor


def add_textbox(doc, lines, width_pt=240, height_pt=70):
    """文本框 —— F4-3 / A4（流式下内容会丢，必须能提取）"""
    p = doc.add_paragraph()
    r = p.add_run()
    pict = OxmlElement("w:pict")
    shape = etree.SubElement(pict, vv("shape"))
    shape.set("id", "TextBox_%d" % random.randint(1000, 9999))
    shape.set("type", "#_x0000_t202")
    shape.set("style", "width:%dpt;height:%dpt" % (width_pt, height_pt))
    tb = etree.SubElement(shape, vv("textbox"))
    txbx = etree.SubElement(tb, w("txbxContent"))
    for ln in lines:
        tp = etree.SubElement(txbx, w("p"))
        tr = etree.SubElement(tp, w("r"))
        tt = etree.SubElement(tr, w("t"))
        tt.text = ln
    r._element.append(pict)
    return p


def add_dropcap(doc, first="飘", rest="飘然独立的排版样本，用于检验首字下沉是否被误判为标题。"):
    """首字下沉 —— F1-9（段内最大字号远大于段字号 → 必须判为正文）"""
    p = doc.add_paragraph()
    pPr = p._p.get_or_add_pPr()
    fr = OxmlElement("w:framePr")
    fr.set(qn("w:dropCap"), "drop"); fr.set(qn("w:lines"), "3")
    fr.set(qn("w:wrap"), "around"); fr.set(qn("w:vAnchor"), "text")
    fr.set(qn("w:hAnchor"), "text")
    pPr.append(fr)
    gt_mark(first + rest, "body", "首字下沉：段内字号远大但仍属正文")
    r1 = p.add_run(first); r1.font.size = Pt(58.5); r1.bold = True
    r1.font.name = "SimSun"
    r1._element.get_or_add_rPr().get_or_add_rFonts().set(qn("w:eastAsia"), "宋体")
    r2 = p.add_run(rest)
    r2.font.size = Pt(10.5)
    r2.font.name = "SimSun"
    r2._element.get_or_add_rPr().get_or_add_rFonts().set(qn("w:eastAsia"), "宋体")
    return p


def add_omml(doc, text="E = mc²"):
    """OMML 公式 —— F6-1"""
    p = doc.add_paragraph()
    omp = etree.SubElement(p._p, mm("oMathPara"))
    om = etree.SubElement(omp, mm("oMath"))
    mr = etree.SubElement(om, mm("r"))
    mt = etree.SubElement(mr, mm("t"))
    mt.text = text
    return p


def add_revisions(doc):
    """修订：插入 + 删除 + 格式变更 —— F8-1 / FM-14"""
    p = doc.add_paragraph()
    r0 = p.add_run("本段包含修订痕迹：")
    r0.font.size = Pt(10.5)
    ins = OxmlElement("w:ins")
    ins.set(qn("w:id"), "9001"); ins.set(qn("w:author"), AUTHOR); ins.set(qn("w:date"), DATE)
    ri = OxmlElement("w:r"); rit = OxmlElement("w:t"); rit.text = "这里是被插入的文字"
    ri.append(rit); ins.append(ri)
    p._p.append(ins)

    dele = OxmlElement("w:del")
    dele.set(qn("w:id"), "9002"); dele.set(qn("w:author"), AUTHOR); dele.set(qn("w:date"), DATE)
    rd = OxmlElement("w:r"); dt = OxmlElement("w:delText"); dt.text = "这里是被删除的文字"
    rd.append(dt); dele.append(rd)
    p._p.append(dele)

    r1 = p.add_run("；本段带格式变更：")
    r1.font.size = Pt(10.5)
    r2 = p.add_run("改动前")
    r2.font.size = Pt(10.5)
    rPr = r2._element.get_or_add_rPr()
    ch = OxmlElement("w:rPrChange")
    ch.set(qn("w:id"), "9003"); ch.set(qn("w:author"), AUTHOR); ch.set(qn("w:date"), DATE)
    old = OxmlElement("w:rPr")
    szv = OxmlElement("w:sz"); szv.set(qn("w:val"), "21"); old.append(szv)
    ch.append(old); rPr.append(ch)
    return p


def add_fake_footnote(doc):
    """手工假脚注：正文上标 + 文末手打块 —— F7-3 / FM-11"""
    p = doc.add_paragraph()
    r = p.add_run("这是一个带手工上标脚注的句子")
    r.font.size = Pt(10.5)
    sup = p.add_run("1")
    sup.font.superscript = True
    sup.font.size = Pt(10.5)
    r2 = p.add_run("，脚注内容以普通段落形式堆在文末。")
    r2.font.size = Pt(10.5)
    tail = doc.add_paragraph()
    tr = tail.add_run("[1] 这是手工敲上去的脚注文字，不是 Word 脚注功能。")
    tr.font.size = Pt(9)
    tr.font.name = "SimSun"
    tr._element.get_or_add_rPr().get_or_add_rFonts().set(qn("w:eastAsia"), "宋体")
    return p


def add_field(doc, instr='TOC \\o "1-3" \\h', placeholder="目录占位（域代码）"):
    p = doc.add_paragraph()
    fld = OxmlElement("w:fldSimple")
    fld.set(qn("w:instr"), instr)
    r = OxmlElement("w:r"); t = OxmlElement("w:t"); t.text = placeholder
    r.append(t); fld.append(r)
    p._p.append(fld)
    return p


def add_hyperlink_bookmark(doc, anchor_text="参见附录", bm="sec_appendix"):
    """书签 + 内部超链接 —— F13 / FM-30"""
    p = doc.add_paragraph()
    bs = OxmlElement("w:bookmarkStart"); bs.set(qn("w:id"), "9100"); bs.set(qn("w:name"), bm)
    p._p.append(bs)
    r = p.add_run(anchor_text)
    r.font.size = Pt(10.5)
    hl = OxmlElement("w:hyperlink"); hl.set(qn("w:anchor"), bm)
    hr = OxmlElement("w:r"); ht = OxmlElement("w:t"); ht.text = "（跳转）"
    hr.append(ht); hl.append(hr)
    p._p.append(hl)
    be = OxmlElement("w:bookmarkEnd"); be.set(qn("w:id"), "9100")
    p._p.append(be)
    return p


def add_tab_indent(doc, n=3):
    p = doc.add_paragraph()
    for _ in range(n):
        r = p.add_run(); r.add_tab()
    r2 = p.add_run("用 Tab 手动缩进的段落" + cn())
    r2.font.size = Pt(10.5)
    return p


# ---------------------------------------------------------------- 后处理 XML
def add_comments(path, comments):
    """注入批注（commentRangeStart/End + comments.xml + 部件声明）"""
    with zipfile.ZipFile(path, "r") as zin:
        names = zin.namelist()
        data = {n: zin.read(n) for n in names}

    docxml = etree.fromstring(data["word/document.xml"])
    body = docxml.find(w("body"))
    paras = [p for p in body.iter(w("p"))]
    targets = []
    for i, p in enumerate(paras):
        if len(p.findall(w("r"))) >= 1 and i % 7 == 3:
            targets.append(p)
    for cid, p in enumerate(targets, start=1):
        s = OxmlElement("w:commentRangeStart"); s.set(qn("w:id"), str(cid))
        e = OxmlElement("w:commentRangeEnd"); e.set(qn("w:id"), str(cid))
        p.insert(0, s)
        p.append(e)
        rr = OxmlElement("w:r")
        ref = OxmlElement("w:commentReference"); ref.set(qn("w:id"), str(cid))
        rr.append(ref)
        p.append(rr)
    data["word/document.xml"] = etree.tostring(docxml, xml_declaration=True,
                                                encoding="UTF-8", standalone=True)

    cx = etree.Element(w("comments"), nsmap={"w": W})
    for cid in range(1, len(targets) + 1):
        c = etree.SubElement(cx, w("comment"))
        c.set(w("id"), str(cid)); c.set(w("author"), AUTHOR)
        c.set(w("date"), DATE); c.set(w("initials"), "SB")
        cp = etree.SubElement(c, w("p"))
        cr = etree.SubElement(cp, w("r"))
        ct = etree.SubElement(cr, w("t"))
        ct.text = comments[(cid - 1) % len(comments)]
    data["word/comments.xml"] = etree.tostring(cx, xml_declaration=True,
                                                encoding="UTF-8", standalone=True)

    ctx = etree.fromstring(data["[Content_Types].xml"])
    ov = etree.SubElement(ctx, "{%s}Override" % CT)
    ov.set("PartName", "/word/comments.xml")
    ov.set("ContentType", "application/vnd.openxmlformats-officedocument."
                          "wordprocessingml.comments+xml")
    data["[Content_Types].xml"] = etree.tostring(ctx, xml_declaration=True,
                                                  encoding="UTF-8", standalone=True)

    rels = etree.fromstring(data["word/_rels/document.xml.rels"])
    rid = "rIdComments"
    rl = etree.SubElement(rels, "{%s}Relationship" % PR)
    rl.set("Id", rid)
    rl.set("Type", "http://schemas.openxmlformats.org/officeDocument/2006/"
                   "relationships/comments")
    rl.set("Target", "comments.xml")
    data["word/_rels/document.xml.rels"] = etree.tostring(rels, xml_declaration=True,
                                                          encoding="UTF-8", standalone=True)
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as zout:
        for n, b in data.items():
            zout.writestr(n, b)


def add_footnotes(path, notes):
    """注入真脚注（footnotes.xml + 引用 + 部件声明）
    🔴 OOXML 硬要求：footnotes.xml 必须含 separator(id=-1) 与 continuationSeparator(id=0)，
       否则 Word 报"文件可能已经损坏"；且 footnoteReference 的 id 必须与 w:footnote 的 id 对应。
       （V-3 实测：修复前 007 原样本 Word 根本打不开）"""
    with zipfile.ZipFile(path, "r") as zin:
        data = {n: zin.read(n) for n in zin.namelist()}

    docxml = etree.fromstring(data["word/document.xml"])
    body = docxml.find(w("body"))
    paras = [p for p in body.iter(w("p"))]
    targets = [p for i, p in enumerate(paras)
               if i % 5 == 2 and len(p.findall(w("r"))) >= 1]

    fid = 2                                  # 脚注 id 从 2 起（-1/0 保留给分隔符）
    n = min(len(notes), len(targets))
    for i in range(n):
        r = OxmlElement("w:r")
        fr = OxmlElement("w:footnoteReference")
        fr.set(qn("w:id"), str(fid + i))    # 引用 id 与定义 id 一一对应
        r.append(fr)
        targets[i].append(r)

    fx = etree.Element(w("footnotes"), nsmap={"w": W})
    # 特殊分隔符脚注（Word 必需）
    for sid, stype, tag in ((-1, "separator", "separator"),
                            (0, "continuationSeparator", "continuationSeparator")):
        f = etree.SubElement(fx, w("footnote"))
        f.set(w("id"), str(sid)); f.set(w("type"), stype)
        fp = etree.SubElement(f, w("p"))
        fr = etree.SubElement(fp, w("r"))
        etree.SubElement(fr, w(tag))
    # 正常脚注
    for i, note in enumerate(notes):
        f = etree.SubElement(fx, w("footnote"))
        f.set(w("id"), str(fid + i))
        fp = etree.SubElement(f, w("p"))
        fr = etree.SubElement(fp, w("r"))
        ft = etree.SubElement(fr, w("t"))
        ft.text = note

    data["word/footnotes.xml"] = etree.tostring(fx, xml_declaration=True,
                                                 encoding="UTF-8", standalone=True)
    data["word/document.xml"] = etree.tostring(docxml, xml_declaration=True,
                                                encoding="UTF-8", standalone=True)

    ctx = etree.fromstring(data["[Content_Types].xml"])
    ov = etree.SubElement(ctx, "{%s}Override" % CT)
    ov.set("PartName", "/word/footnotes.xml")
    ov.set("ContentType", "application/vnd.openxmlformats-officedocument."
                          "wordprocessingml.footnotes+xml")
    data["[Content_Types].xml"] = etree.tostring(ctx, xml_declaration=True,
                                                  encoding="UTF-8", standalone=True)
    rels = etree.fromstring(data["word/_rels/document.xml.rels"])
    rl = etree.SubElement(rels, "{%s}Relationship" % PR)
    rl.set("Id", "rIdFootnotes")
    rl.set("Type", "http://schemas.openxmlformats.org/officeDocument/2006/"
                   "relationships/footnotes")
    rl.set("Target", "footnotes.xml")
    data["word/_rels/document.xml.rels"] = etree.tostring(rels, xml_declaration=True,
                                                          encoding="UTF-8", standalone=True)
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as zout:
        for nme, b in data.items():
            zout.writestr(nme, b)


# ---------------------------------------------------------------- 表格工具
def build_tables(doc, assets):
    """表格刁钻：合并 / 嵌套 / 10列宽 / 表内图 / 表内标题 —— F3 / S6"""
    heading_by_size(doc, "表 1 合并单元格", 14)
    t1 = doc.add_table(rows=3, cols=3)
    t1.style = "Table Grid"
    t1.cell(0, 0).merge(t1.cell(0, 2))
    t1.cell(0, 0).text = "横向合并的表头"
    for i in range(1, 3):
        for j in range(3):
            t1.cell(i, j).text = f"R{i}C{j}"

    heading_by_size(doc, "表 2 嵌套表格", 14)
    outer = doc.add_table(rows=2, cols=2)
    outer.style = "Table Grid"
    outer.cell(0, 0).text = "外层 A1"
    inner = outer.cell(0, 1).add_table(rows=2, cols=2)
    inner.style = "Table Grid"
    inner.cell(0, 0).text = "内层 1-1"
    inner.cell(0, 1).text = "内层 1-2"
    inner.cell(1, 0).text = "内层 2-1"
    inner.cell(1, 1).text = "内层 2-2"
    outer.cell(1, 0).text = "外层 A2"
    outer.cell(1, 1).text = "外层 B2"

    heading_by_size(doc, "表 3 十列宽表（测试横向溢出）", 14)
    t3 = doc.add_table(rows=2, cols=10)
    t3.style = "Table Grid"
    for j in range(10):
        t3.cell(0, j).text = f"列{j + 1}"
        t3.cell(1, j).text = "溢出测试数据" * 3

    heading_by_size(doc, "表 4 单元格内含图片与标题", 14)
    t4 = doc.add_table(rows=2, cols=2)
    t4.style = "Table Grid"
    t4.cell(0, 0).paragraphs[0].add_run().add_picture(str(assets["icon.gif"]),
                                                       width=Cm(1.2))
    # 表内标题（必须不被判为章节标题 —— F1-7）
    tp = t4.cell(0, 1).paragraphs[0]
    tr = tp.add_run("表内小标题样式")
    tr.bold = True
    tr.font.size = Pt(13)
    gt_mark("表内小标题样式", "body", "表内标题：不得判为章节标题（V1/F1-7）")
    t4.cell(1, 0).text = "图在表内"
    t4.cell(1, 1).text = "标题在表内"
    return doc


# ---------------------------------------------------------------- 各型生成
def gen_S1(A):
    d = newdoc()
    para(d, "规范基准文档", size=18, bold=True, align=WD_ALIGN_PARAGRAPH.CENTER)
    d.add_heading("第一章 概述", level=1)
    para(d, cn(), indent_chars=2)
    d.add_heading("1.1 背景", level=2)
    para(d, cn(), indent_chars=2)
    d.add_heading("第二章 方法", level=1)
    para(d, cn(), indent_chars=2)
    d.add_paragraph("要点一：结构决定样式。", style="List Bullet")
    d.add_paragraph("要点二：样式决定呈现。", style="List Bullet")
    add_img(d, A["normal.png"], width_cm=8, alt="示意图：规范流程")
    figcaption(d, "图 1 规范流程示意图")
    return d


def gen_S2(A):
    CURRENT["name"] = "002"
    d = newdoc()
    _t = "纯字号标题样本"
    gt_mark(_t, "doctitle", "文档大标题/书名：不应参与章节层级映射")
    para(d, _t, size=18, bold=True, align=WD_ALIGN_PARAGRAPH.CENTER)
    for lvl, (pt, cnt) in enumerate([(16, 4), (14, 8), (12, 16)], start=1):
        for i in range(cnt):
            heading_by_size(d, f"第{lvl}级标题 {lvl}-{i + 1}：" + cn()[:12], pt,
                            gt=f"h{lvl}")
            para(d, cn(), size=10.5, indent_chars=2)
    return d


def gen_S3(A):
    CURRENT["name"] = "003"
    d = newdoc()
    _t3 = "标题刁钻样本"
    gt_mark(_t3, "doctitle", "文档大标题/书名：不应参与章节层级映射")
    para(d, _t3, size=18, bold=True, align=WD_ALIGN_PARAGRAPH.CENTER)
    heading_by_size(d, "正常一级标题（对照）", 16, gt="h1")
    para(d, cn(), indent_chars=2)
    add_dropcap(d)                                   # F1-9（自动记为 body）
    _co = "正文中间夹一句大字加粗：" + cn()[:10]
    para(d, _co, size=13, bold=True)
    gt_mark(_co, "callout", "正文强调大字：易误判为标题")
    para(d, cn(), indent_chars=2)
    heading_by_size(d, "超长标题：" + cn() * 3, 14, gt="h2")   # F1-5
    para(d, cn(), indent_chars=2)
    # 同字号多级：仅靠加粗与缩进区分 —— F1-6
    heading_by_size(d, "同字号但应视为二级标题", 14, align=WD_ALIGN_PARAGRAPH.LEFT,
                    gt="h2")
    para(d, cn(), indent_chars=2)
    # 章节编号粘连 —— F1-4
    heading_by_size(d, "1.2.3 编号与文本粘连的标题", 14, gt="h2")
    para(d, cn(), indent_chars=2)
    build_tables(d, A)                              # 含表内标题 F1-7
    return d


def gen_S5(A):
    d = newdoc()
    para(d, "图片刁钻样本", size=18, bold=True, align=WD_ALIGN_PARAGRAPH.CENTER)
    para(d, "以下内联图应紧跟本段。", indent_chars=2)
    add_img(d, A["normal.png"], width_cm=10, alt="内联图：正常尺寸")
    figcaption(d, "图 1 内联图（正常）")

    para(d, "以下浮动图应转为 inline 并紧跟本段，不得漂移。", indent_chars=2)
    pic = add_img(d, A["chart.png"], width_cm=6, alt="浮动图：环绕")
    inline_to_anchor(pic._inline)                   # F4-2
    para(d, "环绕文字：浮动图右侧应有这段文字，用于检验位置是否被破坏。" * 2, indent_chars=2)
    figcaption(d, "图 2 浮动图（环绕）")

    para(d, "以下文本框内的文字必须能被提取出来，否则视为内容丢失。", indent_chars=2)
    add_textbox(d, ["文本框第一行：这段文字在流式下极易丢失。",
                    "文本框第二行：必须被提取为正文。"])  # F4-3
    figcaption(d, "图 3 文本框")

    para(d, "以下超大原图显示很小，体积必须被压缩。", indent_chars=2)
    add_img(d, A["huge.jpg"], width_cm=4, alt="超大原图：3000x4000 显示 4cm")
    figcaption(d, "图 4 超大原图（原 3MB 级，显示 4cm）")

    para(d, "以下透明 PNG 与 16-bit PNG 不得被误转 JPEG。", indent_chars=2)
    add_img(d, A["alpha.png"], width_cm=7, alt="透明通道测试")
    add_img(d, A["gray16.png"], width_cm=7, alt="16位灰度测试")
    figcaption(d, "图 5 透明与 16bit")

    para(d, "以下极端长宽比图片。", indent_chars=2)
    add_img(d, A["wide.jpg"], width_cm=14, alt="极端长宽比")
    figcaption(d, "图 6 极端长宽比")

    para(d, "以下图片故意不加 alt，用于可访问性检查。", indent_chars=2)
    add_img(d, A["tiny.png"], width_cm=1, alt="")
    figcaption(d, "图 7 缺 alt")

    for i in range(8):
        para(d, "连续图片段落 %d：用于检验连续内联图顺序。" % (i + 1), indent_chars=2)
        add_img(d, A["icon.gif"], width_cm=1.2, alt=f"图标{i + 1}")
    return d


def gen_S6(A):
    d = newdoc()
    para(d, "表格刁钻样本", size=18, bold=True, align=WD_ALIGN_PARAGRAPH.CENTER)
    build_tables(d, A)
    para(d, cn(), indent_chars=2)
    return d


def gen_S7(A):
    d = newdoc()
    para(d, "脚注与公式样本", size=18, bold=True, align=WD_ALIGN_PARAGRAPH.CENTER)
    d.add_heading("真脚注章节", level=1)
    for i in range(6):
        para(d, f"第 {i + 1} 段含真脚注引用。" + cn(), indent_chars=2)
    add_fake_footnote(d)
    d.add_heading("公式章节", level=1)
    para(d, "下列为 OMML 公式，应可见或以图形式保留：", indent_chars=2)
    add_omml(d, "E = mc²")
    add_omml(d, "a² + b² = c²")
    return d


def gen_S8(A):
    d = newdoc()
    para(d, "极脏结构样本", size=18, bold=True, align=WD_ALIGN_PARAGRAPH.CENTER)
    heading_by_size(d, "一、修订痕迹", 16)
    add_revisions(d)
    heading_by_size(d, "二、空段落凑间距", 16)
    for i in range(6):
        para(d, f"第 {i + 1} 段。" + cn(), indent_chars=2)
        d.add_paragraph("")
        d.add_paragraph("")
    heading_by_size(d, "三、手动缩进", 16)
    add_tab_indent(d, 3)
    add_tab_indent(d, 5)
    heading_by_size(d, "四、引号与零宽字符", 16)
    para(d, '他说的"直引号"和“弯引号”混用' + cn(), indent_chars=2)
    para(d, "零宽字符测试：" + cn() + "‌‍﻿", indent_chars=2)
    heading_by_size(d, "五、场景分隔符（不得被当作空段落删除）", 16)
    para(d, "***", align=WD_ALIGN_PARAGRAPH.CENTER)
    para(d, "分隔符之后应继续正文。" + cn(), indent_chars=2)
    para(d, "***", align=WD_ALIGN_PARAGRAPH.CENTER)
    heading_by_size(d, "六、手动换行", 16)
    p = d.add_paragraph()
    r = p.add_run("第一行手动换行")
    br = OxmlElement("w:br"); r._element.append(br)
    r2 = p.add_run("第二行内容")
    return d


def gen_S9_core(A):
    """老格式的 docx 源（随后由 COM 另存为 .doc/.docm/.rtf）"""
    d = newdoc()
    para(d, "老格式转换样本", size=18, bold=True, align=WD_ALIGN_PARAGRAPH.CENTER)
    d.add_heading("第一章 老格式内容", level=1)
    para(d, cn(), indent_chars=2)
    t = d.add_table(rows=2, cols=2); t.style = "Table Grid"
    t.cell(0, 0).text = "项目"; t.cell(0, 1).text = "值"
    t.cell(1, 0).text = "样本"; t.cell(1, 1).text = "1.0"
    add_img(d, A["normal.png"], width_cm=6, alt="老格式配图")
    figcaption(d, "图 1 老格式配图")
    para(d, cn(), indent_chars=2)
    return d


def gen_S11(A):
    d = newdoc()
    para(d, "封面元数据样本", size=18, bold=True, align=WD_ALIGN_PARAGRAPH.CENTER)
    d.add_heading("短书名章节", level=1)
    para(d, cn(), indent_chars=2)
    add_img(d, A["tiny.png"], width_cm=1, alt="小图标")
    return d


def gen_S12(A):
    d = newdoc()
    para(d, "语言与字符样本", size=18, bold=True, align=WD_ALIGN_PARAGRAPH.CENTER)
    heading_by_size(d, "一、生僻字与异体字", 16)
    para(d, "䶮龘齾爩靐飝㐂䨻麤龗  生僻字用于检验字体子集化是否缺字。", indent_chars=2)
    heading_by_size(d, "二、Emoji", 16)
    para(d, "表情符号测试：📚 🔬 🧪 ⚗️ ✅ 🎯 🚀 🧭", indent_chars=2)
    heading_by_size(d, "三、繁体中文", 16)
    para(d, "繁體中文段落：這個段落使用繁體字形，供轉換驗證使用。", indent_chars=2)
    heading_by_size(d, "四、全半角混排", 16)
    para(d, "全角ＡＢＣ１２３ 与 半角 ABC123，以及，标点。！？；：（）", indent_chars=2)
    heading_by_size(d, "五、中英混排", 16)
    para(d, "混合文本 Mixed text 混排 1234567890 ABCDEFGHIJ.", indent_chars=2)
    return d


def gen_S13(A):
    d = newdoc()
    para(d, "交叉引用与域样本", size=18, bold=True, align=WD_ALIGN_PARAGRAPH.CENTER)
    add_field(d, 'TOC \\o "1-3" \\h', "（目录域占位）")
    d.add_heading("第一章 引用目标", level=1)
    para(d, cn(), indent_chars=2)
    add_hyperlink_bookmark(d)
    add_field(d, "PAGE", "1")
    return d


def gen_S14(A):
    d = newdoc()
    para(d, "性能与大文件样本", size=18, bold=True, align=WD_ALIGN_PARAGRAPH.CENTER)
    d.add_heading("第一章 规模测试", level=1)
    for i in range(160):
        para(d, f"第 {i + 1} 段。" + cn(), indent_chars=2)
    t = d.add_table(rows=60, cols=8); t.style = "Table Grid"
    for i in range(60):
        for j in range(8):
            t.cell(i, j).text = f"R{i}C{j}"
    for i in range(40):
        add_img(d, A["icon.gif"], width_cm=0.8, alt=f"图标{i + 1}")
    return d


def gen_S15(A):
    d = newdoc()
    para(d, "混合终验样本（最难）", size=18, bold=True, align=WD_ALIGN_PARAGRAPH.CENTER)
    heading_by_size(d, "第一章 脏结构", 16)
    add_revisions(d)
    for i in range(4):
        d.add_paragraph("")
    heading_by_size(d, "第二章 标题歧义", 14)
    add_dropcap(d)
    para(d, "正文夹大字：" + cn()[:8], size=13, bold=True)
    build_tables(d, A)
    heading_by_size(d, "第三章 图片", 14)
    pic = add_img(d, A["huge.jpg"], width_cm=4, alt="超大图")
    inline_to_anchor(pic._inline)
    add_textbox(d, ["文本框内容必须被提取。"])
    para(d, "脚注与公式：" + cn(), indent_chars=2)
    add_omml(d, "E=mc²")
    add_fake_footnote(d)
    heading_by_size(d, "第四章 语言字符", 14)
    para(d, "生僻字䶮龘齾爩 + emoji📚 + 繁體這個段落 + 全角ＡＢＣ", indent_chars=2)
    heading_by_size(d, "第五章 交叉引用", 14)
    add_field(d, 'TOC \\o "1-3" \\h', "（目录域）")
    add_hyperlink_bookmark(d)
    add_tab_indent(d, 4)
    para(d, "结尾。" + cn(), indent_chars=2)
    return d


# ---------------------------------------------------------------- 坏文件集
def gen_bad_files(base_src):
    bad = OUT / "010_坏文件混合_低_规范"
    bad.mkdir(parents=True, exist_ok=True)
    made = []

    # 1 空文件
    p = bad / "010_坏文件混合_低_规范_f01_empty.docx"
    p.write_bytes(b"")
    made.append((p, "空文件（0 字节）", "隔离/跳过，不中断批次"))

    # 2 截断损坏 ZIP
    p = bad / "010_坏文件混合_低_规范_f02_truncated.docx"
    raw = base_src.read_bytes()
    p.write_bytes(raw[: int(len(raw) * 0.55)])
    made.append((p, "ZIP 被截断至 55%", "走回退：Open and Repair / 抢救 media"))

    # 3 缺 document.xml 的伪 docx
    p = bad / "010_坏文件混合_低_规范_f03_no_document_xml.docx"
    with zipfile.ZipFile(base_src) as zin, zipfile.ZipFile(p, "w", zipfile.ZIP_DEFLATED) as zout:
        for n in zin.namelist():
            if n == "word/document.xml":
                continue
            zout.writestr(n, zin.read(n))
    made.append((p, "缺 word/document.xml（部件缺失）", "识别为损坏，隔离"))

    # 4 伪扩展名：内容是 .doc（由 COM 阶段改名），此处先用 docx 冒充 rtf 占位
    p = bad / "010_坏文件混合_低_规范_f04_mislabeled.doc"
    shutil.copy(base_src, p)
    made.append((p, "docx 内容但扩展名为 .doc（伪扩展名）", "按魔数识别为真 docx，不按扩展名处理"))

    # 5 Content_Types 被破坏
    p = bad / "010_坏文件混合_低_规范_f05_broken_ct.docx"
    with zipfile.ZipFile(base_src) as zin, zipfile.ZipFile(p, "w", zipfile.ZIP_DEFLATED) as zout:
        for n in zin.namelist():
            if n == "[Content_Types].xml":
                continue
            zout.writestr(n, zin.read(n))
        zout.writestr("[Content_Types].xml", "<<<BROKEN>>>")
    made.append((p, "[Content_Types].xml 损坏", "Word 会报需要修复，须隔离"))
    return made


# ---------------------------------------------------------------- 保存
def save(d, seq, cat, comp, dirty, ext="docx"):
    ddir = OUT / f"{seq:03d}_{cat}_{comp}_{dirty}"
    ddir.mkdir(parents=True, exist_ok=True)
    fp = ddir / f"{seq:03d}_{cat}_{comp}_{dirty}.{ext}"
    d.save(str(fp))
    return fp


# ---------------------------------------------------------------- Word COM
def word_com(A, base_docx):
    import win32com.client as wc
    res = []
    word = wc.Dispatch("Word.Application")
    word.Visible = False
    word.DisplayAlerts = 0
    try:
        # S9 老格式三件套
        src = str(base_docx)
        for ext, fmt, name in [("doc", 0, "S9-老doc"), ("rtf", 6, "S9-rtf"),
                               ("docm", 13, "S9-docm")]:
            try:
                ddir = OUT / f"009_{name}_中_很脏"
                ddir.mkdir(parents=True, exist_ok=True)
                tgt = str(ddir / f"009_{name}_中_很脏.{ext}")
                doc = word.Documents.Open(src, ReadOnly=True)
                doc.SaveAs2(tgt, FileFormat=fmt)
                doc.Close(0)
                res.append((Path(tgt), f"{name}（由 Word 另存为 .{ext}）", "能转换即可，能力损失只记录"))
            except Exception as e:
                res.append((None, f"{name} 失败：{e}", "跳过"))

        # S4 图表密集（原生图表 + SmartArt）
        try:
            ddir = OUT / "004_图表混合_高_很脏"
            ddir.mkdir(parents=True, exist_ok=True)
            tgt = str(ddir / "004_图表混合_高_很脏.docx")
            doc = word.Documents.Add()
            doc.Content.Text = "图表密集样本"
            for i in range(2):
                p = doc.Paragraphs.Add()
                shp = doc.InlineShapes.AddChart2(-1, 51)   # 51 = 柱状图
                p.Range.InsertAfter("原生图表 %d" % (i + 1))
            try:
                lay = doc.InlineShapes.AddSmartArt(0)
            except Exception:
                pass
            doc.SaveAs2(tgt, FileFormat=16)
            doc.Close(0)
            res.append((Path(tgt), "原生图表（AddChart2）+ SmartArt", "图表可见即可（栅格化可接受）"))
        except Exception as e:
            res.append((None, f"S4 图表生成失败：{e}", "跳过"))

        # S10 加密文件（优先 msoffcrypto 标准 ECMA-376 加密，回退 Word COM）
        enc_path = None
        try:
            ddir = OUT / "010_坏文件混合_低_规范"
            ddir.mkdir(parents=True, exist_ok=True)
            enc_path = ddir / "010_坏文件混合_低_规范_f06_encrypted.docx"
            from msoffcrypto.format.ooxml import OOXMLFile
            with open(base_docx, "rb") as fin, open(enc_path, "wb") as fout:
                OOXMLFile(fin).encrypt("test123", fout)
            res.append((enc_path, "加密（标准 ECMA-376 加密，密码 test123）",
                        "识别为加密 → 跳过并标记，不解密"))
        except Exception as e:
            try:
                doc = word.Documents.Open(str(base_docx), ReadOnly=True)
                doc.Password = "test123"
                doc.SaveAs2(str(enc_path), FileFormat=16)
                doc.Close(0)
                res.append((enc_path, "加密（Word COM 方式，密码 test123）", "同上"))
            except Exception as e2:
                res.append((None, f"加密样本失败：{e} / {e2}", "跳过"))
    finally:
        try:
            word.Quit()
        except Exception:
            pass
    return res


# ---------------------------------------------------------------- main
def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--com", action="store_true", help="追加 Word COM 阶段")
    args = ap.parse_args()

    if OUT.exists():
        shutil.rmtree(OUT)
    OUT.mkdir(parents=True)
    A = make_assets()
    print("[assets] %d 张素材图" % len(A))

    plan = [
        (1, "文字混合", "低", "规范", gen_S1, None),
        (2, "标题混合", "高", "很脏", gen_S2, None),
        (3, "标题混合", "高", "很脏", gen_S3, None),
        (5, "图片混合", "高", "很脏", gen_S5, None),
        (6, "表格混合", "高", "很脏", gen_S6, None),
        (7, "脚注混合", "中", "很脏", gen_S7, "footnotes"),
        (8, "文字混合", "高", "极端脏", gen_S8, "comments"),
        (11, "封面元数据", "低", "轻微脏", gen_S11, None),
        (12, "文字混合", "中", "很脏", gen_S12, None),
        (13, "列表混合", "低", "轻微脏", gen_S13, None),
        (14, "混合", "高", "轻微脏", gen_S14, None),
        (15, "混合", "高", "极端脏", gen_S15, "comments"),
    ]
    made = []
    for seq, cat, comp, dirty, fn, post in plan:
        try:
            d = fn(A)
            fp = save(d, seq, cat, comp, dirty)
            if post == "comments":
                add_comments(fp, ["这条批注不应出现在正文里。",
                                  "此批注用于检验批注剥离。",
                                  "批注三：图表位置建议调整。"])
            if post == "footnotes":
                add_footnotes(fp, ["脚注一：结构优先于样式。",
                                   "脚注二：脚注应保持可达。",
                                   "脚注三：流式下脚注转为章末或弹层。"])
            size = fp.stat().st_size
            made.append((fp, f"S{seq} {cat}（{dirty}）", f"{size/1024:.0f} KB"))
            print(f"[ok] {fp.name}  {size/1024:.0f} KB")
        except Exception as e:
            print(f"[FAIL] S{seq}: {e}")

    base = OUT / "002_标题混合_高_很脏" / "002_标题混合_高_很脏.docx"
    for p, desc, exp in gen_bad_files(base):
        made.append((p, desc, exp))
        print(f"[bad] {p.name}  {p.stat().st_size/1024:.1f} KB")

    if args.com:
        print("[com] 启动 Word COM 阶段…")
        for p, desc, exp in word_com(A, base):
            if p:
                made.append((p, desc, exp))
                print(f"[com] {p.name}  {p.stat().st_size/1024:.0f} KB")
            else:
                print(f"[com] {desc}")

    # 索引
    idx = OUT / "index.md"
    with open(idx, "w", encoding="utf-8") as f:
        f.write("# 测试样本索引\n\n> 由 make_samples.py 自动生成；对应 docx2epub_测试用例规格.md\n\n")
        f.write("| 文件 | 大小 | 内容 | 期望 |\n|---|---|---|---|\n")
        for p, desc, exp in made:
            f.write("| `{name}` | {kb} | {desc} | {exp} |\n".format(
                name=p.name, kb=f"{p.stat().st_size/1024:.0f} KB", desc=desc, exp=exp))

    # ground truth（供 V-4 标题聚类实测自动评估）
    gtp = OUT / "ground_truth.json"
    with open(gtp, "w", encoding="utf-8") as f:
        json.dump(GT, f, ensure_ascii=False, indent=2)
    by_sample = {}
    for g in GT:
        by_sample.setdefault(g["sample"], {}).setdefault(g["label"], 0)
        by_sample[g["sample"]][g["label"]] += 1
    print("[ground_truth] %d 条标注 → %s" % (len(GT), gtp.name))
    for s, d in sorted(by_sample.items()):
        print("   %s: %s" % (s, "  ".join("%s=%d" % (k, v) for k, v in sorted(d.items()))))

    print("\n[done] %d 个样本 → %s" % (len(made), OUT))


if __name__ == "__main__":
    main()
