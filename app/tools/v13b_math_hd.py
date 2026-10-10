# -*- coding: utf-8 -*-
"""V-13b 公式高清化：Word COM 提取公式 → EMF → GDI 高清 PNG → 前置手术替换 OMML 为内联图片。
供 htmlpath_poc.py 调用。Word 打不开的文档返回空列表（fallback：LO 低清栅格）。
"""
import os, re, zipfile, ctypes
import win32com.client, pythoncom, win32clipboard
from emf2png import emf_bytes_to_png
from PIL import Image, ImageChops

EMU_PER_PX = 9525  # 96dpi

def _trim_white(img, pad=8):
    bg = Image.new("RGB", img.size, (255, 255, 255))
    diff = ImageChops.difference(img.convert("RGB"), bg)
    bbox = diff.getbbox()
    if not bbox:
        return img
    l = max(0, bbox[0] - pad); t = max(0, bbox[1] - pad)
    r = min(img.width, bbox[2] + pad); b = min(img.height, bbox[3] + pad)
    return img.crop((l, t, r, b))

def export_math_pngs(docx_path, outdir, scale=6):
    """Word COM 提取所有公式为高清 PNG。返回 [(png_path, emu_w, emu_h)]；失败返回 []。"""
    os.makedirs(outdir, exist_ok=True)
    out = []
    pythoncom.CoInitialize()
    try:
        w = win32com.client.gencache.EnsureDispatch("Word.Application")
        w.Visible = False; w.DisplayAlerts = 0
        try:
            d = w.Documents.Open(os.path.abspath(docx_path), ConfirmConversions=False,
                                 ReadOnly=True, AddToRecentFiles=False, Visible=False)
        except Exception:
            return []
        try:
            n = d.OMaths.Count
            for i in range(1, n + 1):
                d.OMaths.Item(i).Range.CopyAsPicture()
                win32clipboard.OpenClipboard()
                try:
                    data = win32clipboard.GetClipboardData(win32clipboard.CF_ENHMETAFILE)
                finally:
                    win32clipboard.CloseClipboard()
                if not (isinstance(data, bytes) and data[:4] == b"\x01\x00\x00\x00"):
                    continue
                img = emf_bytes_to_png(data, scale=scale)
                img = _trim_white(img)
                png = os.path.join(outdir, f"mathHD{i}.png")
                img.save(png)
                emu_w = round(img.width / scale * EMU_PER_PX)
                emu_h = round(img.height / scale * EMU_PER_PX)
                out.append((png, emu_w, emu_h))
        finally:
            d.Close(False)
    except Exception:
        return out
    finally:
        try:
            w.Quit()
        except Exception:
            pass
    return out

DRAWING_TMPL = ('<w:r><w:drawing>'
    '<wp:inline distT="0" distB="0" distL="0" distR="0" '
    'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing">'
    '<wp:extent cx="{cx}" cy="{cy}"/><wp:effectExtent l="0" t="0" r="0" b="0"/>'
    '<wp:docPr id="{did}" name="mathHD{did}"/>'
    '<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">'
    '<a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">'
    '<pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">'
    '<pic:nvPicPr><pic:cNvPr id="{did}" name="mathHD{did}"/><pic:cNvPicPr/></pic:nvPicPr>'
    '<pic:blipFill><a:blip r:embed="{rid}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>'
    '<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="{cx}" cy="{cy}"/></a:xfrm>'
    '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr>'
    '</pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>')

def surgery_docx(src_docx, dst_docx, math_pngs):
    """前置手术：tab→NBSP + 公式(OMML)→高清PNG内联图。返回 (tab数, 公式替换数)。"""
    zin = zipfile.ZipFile(src_docx)
    doc = zin.read("word/document.xml").decode("utf-8")
    tabs = doc.count("<w:tab/>")
    if tabs:
        doc = doc.replace("<w:tab/>", '<w:t xml:space="preserve">\u00a0\u00a0\u00a0\u00a0</w:t>')
    # 公式替换（oMathPara 优先，再独立 oMath）
    replaced = 0
    def drawing(i, png, emu_w, emu_h):
        return DRAWING_TMPL.format(cx=emu_w, cy=emu_h, did=1000 + i, rid=f"rIdMathHD{i}")
    if math_pngs:
        # 1) oMathPara（含内部 oMath 一并消失）
        parts = re.split(r'(<m:oMathPara>.*?</m:oMathPara>)', doc, flags=re.S)
        for k in range(1, len(parts), 2):
            if replaced < len(math_pngs):
                replaced += 1
                parts[k] = drawing(*((replaced,) + math_pngs[replaced-1]))
        doc = "".join(parts)
        # 2) 独立 oMath
        parts = re.split(r'(<m:oMath>.*?</m:oMath>)', doc, flags=re.S)
        for k in range(1, len(parts), 2):
            if replaced < len(math_pngs):
                replaced += 1
                parts[k] = drawing(*((replaced,) + math_pngs[replaced-1]))
        doc = "".join(parts)
    # rels / content-types / media
    relname = "word/_rels/document.xml.rels"
    rels = zin.read(relname).decode("utf-8") if relname in zin.namelist() else None
    adds = "".join(f'<Relationship Id="rIdMathHD{i}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/mathHD{i}.png"/>'
                   for i in range(1, replaced + 1))
    if rels and adds:
        rels = rels.replace("</Relationships>", adds + "</Relationships>")
    ct = zin.read("[Content_Types].xml").decode("utf-8")
    if 'Extension="png"' not in ct:
        ct = ct.replace("<Default Extension=", '<Default Extension="png" ContentType="image/png"/><Default Extension=', 1)
    with zipfile.ZipFile(dst_docx, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            fn = item.filename
            if fn == "word/document.xml":
                zout.writestr(item, doc.encode("utf-8"))
            elif fn == relname and rels:
                zout.writestr(item, rels.encode("utf-8"))
            elif fn == "[Content_Types].xml":
                zout.writestr(item, ct.encode("utf-8"))
            else:
                zout.writestr(item, zin.read(fn))
        for i, (png, _, _) in enumerate(math_pngs[:replaced], 1):
            zout.write(png, f"word/media/mathHD{i}.png")
    zin.close()
    return tabs, replaced
