# -*- coding: utf-8 -*-
"""stage1 前置手术扩展：wpg 组合图形拆解（2026-10-09）

问题：Word 里"小看板"等是 wpg:wgp 组合图形（形状带文字 + 内嵌图片）。
LO 对组合图形整体栅格化成一张 GIF → 组合内文字被烤进图（错位）、
栅格画布边界失控（右侧出血）、文字层丢字（L1 掉到 96%）。

手术：把每个 wpg:wgp 拆解为——
  - 组合内 wps 形状的文字段落 → 提升为正文段落（插在宿主段之前）
  - 组合内 pic:pic → 直接替换组合（LO 按普通图片导出，不再栅格化）
  - mc:AlternateContent 剥壳（删 VML Fallback，防 LO 双重处理）

只动 word/document.xml，其余包内部件原样保留。
"""
import os, shutil, zipfile
from lxml import etree

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
MC = "http://schemas.openxmlformats.org/markup-compatibility/2006"
A = "http://schemas.openxmlformats.org/drawingml/2006/main"
PIC = "http://schemas.openxmlformats.org/drawingml/2006/picture"
WPG = "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup"
WPS = "http://schemas.microsoft.com/office/word/2010/wordprocessingShape"
WP = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"


def _tag(ns, name):
    return f"{{{ns}}}{name}"


def _apply_pic_extent(drawing_el, pic):
    """把 pic 自身的 a:ext（组合内子变换的真实尺寸，EMU）写回宿主 drawing 的 wp:extent。
    不修的话 LO 拿组合整体 extent 当图片显示尺寸 → 竖图被声明成横版比例（655x262 vs 实际 292x633），
    产品图被等比错误放大/压扁。"""
    if drawing_el is None:
        return
    x = pic.find(f".//{_tag(A, 'xfrm')}/{_tag(A, 'ext')}")
    if x is None:
        return
    cx, cy = x.get("cx"), x.get("cy")
    if not cx or not cy:
        return
    ext = drawing_el.find(_tag(WP, "extent"))
    if ext is not None:
        ext.set("cx", cx); ext.set("cy", cy)
    for ee in drawing_el.findall(_tag(WP, "effectExtent")):
        for a in ("l", "t", "r", "b"):
            ee.set(a, "0")
    off = pic.find(f".//{_tag(A, 'xfrm')}/{_tag(A, 'off')}")
    if off is not None:          # 组内偏移归零，防 LO 把图摆到组合画布外
        off.set("x", "0"); off.set("y", "0")


def flatten_groups(docx_path):
    """拆解 wpg 组合图形。返回 dict(groups=拆解数, text_paras=提升段数, pics=保留图数)"""
    zin = zipfile.ZipFile(docx_path)
    doc = etree.fromstring(zin.read("word/document.xml"))

    n_groups = n_paras = n_pics = 0
    # 只处理顶层 wgp（嵌套组随顶层一起被拆）
    for wgp in list(doc.iter(_tag(WPG, "wgp"))):
        if any(a.tag == _tag(WPG, "wgp") for a in wgp.iterancestors()):
            continue
        # 宿主 w:p / 宿主 w:r
        host_p = None
        for a in wgp.iterancestors():
            if a.tag == _tag(W, "p"):
                host_p = a
                break
        if host_p is None:
            continue
        host_r = None
        for a in wgp.iterancestors():
            if a.tag == _tag(W, "r"):
                host_r = a
                break
        # 1) 组合内文字段落（只取 Choice 侧 wps:wsp 下的，防 Fallback 重复）
        text_paras = []
        for wsp in wgp.iter(_tag(WPS, "wsp")):
            for tp in wsp.iter(_tag(W, "txbxContent")):
                for p in tp.findall(_tag(W, "p")):
                    text_paras.append(etree.fromstring(etree.tostring(p)))
        for tp in text_paras:
            host_p.addprevious(tp)
        n_paras += len(text_paras)
        # 2) 组合内图片
        pics = [p for p in wgp.iter(_tag(PIC, "pic"))]
        gd = wgp.getparent()                       # a:graphicData
        anchor = None
        for a in wgp.iterancestors():
            if a.tag.endswith("}anchor") or a.tag.endswith("}inline"):
                anchor = a
                break
        if not pics:
            # 纯文字组合：宿主 run 整个删掉（文字段已提升）
            if host_r is not None and host_r.getparent() is not None:
                host_r.getparent().remove(host_r)
            n_groups += 1
            continue
        # 多图时先克隆宿主 run（内含完整 drawing 锚点）
        extra_runs = []
        if len(pics) > 1 and host_r is not None:
            for _ in pics[1:]:
                r2 = etree.fromstring(etree.tostring(host_r))
                extra_runs.append(r2)
        # 替换：graphicData 里 wgp → 第一张 pic
        gd.remove(wgp)
        gd.append(pics[0])
        gd.set("uri", PIC)
        _apply_pic_extent(anchor, pics[0])
        # 其余 pic：克隆 run 里同样替换（graphicData 直接子元素就是 wgp）
        for r2, pic in zip(extra_runs, pics[1:]):
            g2 = r2.find(f".//{_tag(A, 'graphicData')}")
            for w2 in list(g2):
                if w2.tag == _tag(WPG, "wgp"):
                    g2.remove(w2)
            g2.append(etree.fromstring(etree.tostring(pic)))
            g2.set("uri", PIC)
            d2 = None
            for a in g2.iterancestors():
                if a.tag.endswith("}anchor") or a.tag.endswith("}inline"):
                    d2 = a; break
            _apply_pic_extent(d2, pic)
            host_r.addnext(r2)
            host_r = r2
        n_pics += len(pics)
        # 3) mc:AlternateContent 剥壳（删 Fallback，防双重处理）
        for a in list(gd.iterancestors()):
            if a.tag == _tag(MC, "AlternateContent"):
                choice = a.find(_tag(MC, "Choice"))
                if choice is not None:
                    parent = a.getparent()
                    idx = parent.index(a)
                    for ch in list(choice):
                        parent.insert(idx, ch)
                        idx += 1
                    parent.remove(a)
                break
        n_groups += 1

    # 重写 zip
    out = docx_path + ".flat"
    new_doc = etree.tostring(doc, xml_declaration=True, encoding="UTF-8", standalone=True)
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = new_doc if item.filename == "word/document.xml" else zin.read(item.filename)
            zout.writestr(item, data)
    zin.close()
    shutil.move(out, docx_path)
    return dict(groups=n_groups, text_paras=n_paras, pics=n_pics)


if __name__ == "__main__":
    import sys, io
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")
    for f in sys.argv[1:]:
        print(f, flatten_groups(f))
