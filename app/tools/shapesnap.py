# -*- coding: utf-8 -*-
"""shapesnap.py — Word COM 整块截图手术（2026-10-09 第四轮方案）

背景：docx 里的"小看板"是浮动形状组合（圆角边框盒 + 举牌小人 + 盒内文字，可能带
黄色高亮），LO 对其整体栅格化时画布失控（出血）、文字错位；此前"拆解组合"方案会
丢失边框/小人/排版。治本：用 Word 自己把整块形状导出为 EMF 矢量图 → Pillow 高分辨率
渲染 → 自动裁掉空白出血边 → 删除原形状、在锚点插入普通内嵌图片。LO 从此只看到图片。

规则（v1）：
  - 所有 Type==6 组合形状 → 各自整块截图（组合可能未包含叠放的小人，但当前样本
    三个组合均自带完整视觉；未组合叠放情形走 text-cluster 规则）
  - 非组合但含文字的浮动形状（TextFrame.HasText）→ 与同锚段、包围盒重叠的形状
    聚成一簇，整簇截图
  - 纯图片浮动形状不动（是正常内容图）

依赖：Windows + MS Word + pywin32 + Pillow。Word 不可用时抛 RuntimeError，
pipeline.stage1 捕获后回落旧路径（LO 预转 + flatten 拆解）。
"""
import os
import io

import win32com.client as wc
from PIL import Image, ImageChops

WD_FORMAT_XML = 12          # docx
DPI = 200                   # EMF 渲染分辨率（网页浏览足够清晰）
CROP_MARGIN = 8             # 裁边后保留的白边（px）


def _page_bbox(s, sect):
    """形状的近似页坐标包围盒（pt）。哨兵 -999995 归 0。"""
    def n(v):
        return 0.0 if v is None or v < -900000 else float(v)
    return (n(s.Left), n(s.Top), n(s.Left) + n(s.Width), n(s.Top) + n(s.Height))


def _overlap(a, b):
    return not (a[2] <= b[0] or b[2] <= a[0] or a[3] <= b[1] or b[3] <= a[1])


def _emf_to_png(emf_bytes, png_path):
    """EMF → PNG，自动裁掉四周空白（出血治理）。"""
    if isinstance(emf_bytes, memoryview):
        emf_bytes = emf_bytes.tobytes()
    elif isinstance(emf_bytes, str):
        emf_bytes = emf_bytes.encode("latin1")
    tmp = png_path + ".emf"
    open(tmp, "wb").write(emf_bytes)
    im = Image.open(tmp)
    im.load(dpi=DPI)
    os.remove(tmp)
    # 裁边：非近白像素的包围盒
    gray = im.convert("L")
    mask = gray.point(lambda p: 255 if p < 246 else 0)
    bbox = mask.getbbox()
    if bbox:
        l, t, r, b = bbox
        l = max(0, l - CROP_MARGIN); t = max(0, t - CROP_MARGIN)
        r = min(im.width, r + CROP_MARGIN); b = min(im.height, b + CROP_MARGIN)
        im = im.crop((l, t, r, b))
    im.save(png_path)
    return im.size


def vml_to_drawingml(docx_path):
    """把 docx 里的内嵌 VML 图片（w:pict/v:shape/v:imagedata）转成标准 DrawingML 内嵌图。

    背景（2026-10-09）：.doc 经 Word 另存 docx 后，AddPicture 插入的图片和原图全部写成
    VML；LO 读 VML 只能栅格化成 GIF（正文被搅乱）。同 rId 换 <w:drawing> 即可。
    只转"纯图片"pict；带 v:textbox 的复杂 VML 不动。
    """
    import zipfile, shutil, re
    from lxml import etree
    W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
    tmp = docx_path + ".vmltmp"
    shutil.copy2(docx_path, tmp)
    zin = zipfile.ZipFile(tmp)
    xml = zin.read("word/document.xml")
    root = etree.fromstring(xml)
    n = 0
    for pict in root.iter(f"{{{W}}}pict"):
        vs = pict.find(f".//{{urn:schemas-microsoft-com:vml}}shape")
        if vs is None:
            continue
        if vs.find(f".//{{urn:schemas-microsoft-com:vml}}textbox") is not None:
            continue                      # 复杂 VML（带字）不动
        imd = vs.find(f".//{{urn:schemas-microsoft-com:vml}}imagedata")
        if imd is None:
            continue
        rid = imd.get(f"{{{R}}}embed") or imd.get(f"{{{R}}}id")   # VML 里是 r:id，DrawingML 里是 r:embed
        if not rid:
            continue
        m = re.search(r"width:([\d.]+)pt", vs.get("style") or "")
        m2 = re.search(r"height:([\d.]+)pt", vs.get("style") or "")
        if not (m and m2):
            continue                      # 无尺寸信息的不冒险
        cx, cy = int(float(m.group(1)) * 12700), int(float(m2.group(1)) * 12700)
        n += 1
        drawing = (f'<w:drawing xmlns:w="{W}" xmlns:r="{R}" '
                   'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" '
                   'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
                   'xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">'
                   f'<wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="{cx}" cy="{cy}"/>'
                   f'<wp:effectExtent l="0" t="0" r="0" b="0"/><wp:docPr id="{9000+n}" name="图片 {n}"/>'
                   '<wp:cNvGraphicFramePr><a:graphicFrameLocks noChangeAspect="1"/></wp:cNvGraphicFramePr>'
                   '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">'
                   f'<pic:pic><pic:nvPicPr><pic:cNvPr id="{n}" name="图片 {n}"/><pic:cNvPicPr/></pic:nvPicPr>'
                   f'<pic:blipFill><a:blip r:embed="{rid}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>'
                   '<pic:spPr><a:xfrm><a:off x="0" y="0"/>'
                   f'<a:ext cx="{cx}" cy="{cy}"/></a:xfrm>'
                   '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr>'
                   '</pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing>')
        pict.getparent().replace(pict, etree.fromstring(drawing))
    if not n:
        zin.close(); os.remove(tmp)
        return 0
    # 重打包（其余成员原样复制）
    zout = zipfile.ZipFile(docx_path, "w", zipfile.ZIP_DEFLATED)
    for item in zin.infolist():
        data = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True) \
            if item.filename == "word/document.xml" else zin.read(item.filename)
        zout.writestr(item, data)
    zout.close(); zin.close(); os.remove(tmp)
    return n


def snap_floating(src, out_docx, img_dir, log=None):
    """把 src 文档里的组合/含字浮动形状整块截图并替换为内嵌图片，另存 out_docx。

    返回 dict(snaps=[{png,w,h,shape}], groups=N, textboxes=N)。
    """
    def say(m):
        if log: log(m)

    os.makedirs(img_dir, exist_ok=True)
    word = wc.DispatchEx("Word.Application")
    word.Visible = False
    word.DisplayAlerts = 0
    try:
        word.FileValidation = 1    # msoFileValidationSkip：绕过 OFV 校验（部分 .doc 被误拒"文件存在问题"）
    except Exception:
        pass
    doc = None
    try:
        doc = word.Documents.Open(os.path.abspath(src), ReadOnly=False)
        if doc.Revisions.Count:                    # 大叔拍板 2026-10-09：修订一律接受成终稿——
            n_rev = doc.Revisions.Count            # LO 转换会把"已删除"的修订文字画出来（"三二、"事故），
            doc.AcceptAllRevisions()               # 且基线/LO渲染/编号提取三方必须同源
            if log: log(f"  [snap] 接受全部修订 {n_rev} 处（成终稿）")
        if doc.TrackRevisions:                     # 修订模式会让 Delete 变"标记删除"（坑：Count 不变、无报错）
            doc.TrackRevisions = False
            if log: log("  [snap] 检测到修订模式，已关闭")
        n_shapes = doc.Shapes.Count
        if not n_shapes:
            doc.SaveAs2(os.path.abspath(out_docx), FileFormat=WD_FORMAT_XML)
            return dict(snaps=[], groups=0, textboxes=0)

        sect = doc.Sections(1)
        targets = []          # ([shape_refs], why, anchor_start[, top])——存 COM 引用，Delete 后索引会重排，绝不能存序号
        # 候选 = 全部非组合浮动形状（含无字图片：拖把小人这类与文字框叠放的插图要一起截）
        cands = []
        for i in range(1, n_shapes + 1):
            s = doc.Shapes(i)
            try:
                if s.Type == 6:                      # msoGroup → 整块截
                    targets.append(([s], "group", s.Anchor.Start, s.Top))
                else:
                    cands.append(s)
            except Exception:
                continue

        # 无字/含字形状按"包围盒重叠 + 锚点相近"聚类；含字的簇才是截图目标
        met = {}
        for s in cands:
            try:
                has_text = bool(s.TextFrame.HasText)
            except Exception:
                has_text = False
            met[s.ID] = (_page_bbox(s, sect), s.Anchor.Start, has_text)
        parent = {s.ID: s.ID for s in cands}
        def find(x):
            while parent[x] != x:
                parent[x] = parent[parent[x]]; x = parent[x]
            return x
        for a in cands:
            for b in cands:
                if a.ID >= b.ID: continue
                ba, bb = met[a.ID][0], met[b.ID][0]
                if _overlap(ba, bb) and abs(met[a.ID][1] - met[b.ID][1]) <= 200:
                    parent[find(a.ID)] = find(b.ID)
        clusters = {}
        for s in cands:
            clusters.setdefault(find(s.ID), []).append(s)
        for grp in clusters.values():
            # 大叔拍板 2026-10-10：浮动图片簇也一律截图转内嵌——实证（安心养老图三）：
            # LO 把"锚段下方"的浮动图输出到锚段前面，位置漂移；内嵌图 LO 处理无问题。
            # 原规则"纯图片簇不动"只对内嵌图成立（内嵌图不在 Shapes 集合，不会到这里）。
            top = min((lambda t: t if t is not None and t > -900000 else 0.0)(s.Top) for s in grp)
            why = "textbox" if any(met[s.ID][2] for s in grp) else "picture"
            targets.append((grp, why, met[grp[0].ID][1], top))

        def _cur_idx():
            """现场索引表（Shape.ID → 当前序号）：Shapes 在 Delete 后会重排"""
            m = {}
            for i in range(1, doc.Shapes.Count + 1):
                try: m[doc.Shapes(i).ID] = i
                except Exception: pass
            return m

        def _insert_rng(top, anchor):
            """位置对齐 Word 视觉（大叔 2026-10-09 反馈"位置一定要明确"）：
            Top>3pt = 形状悬在锚段文字下方 → 插锚段末尾；否则插锚段开头"""
            try:
                para = anchor.Paragraphs(1).Range
            except Exception:
                return anchor
            try:
                top = float(top)
            except Exception:
                top = 0.0
            pos = para.End - 1 if top > 3.0 else para.Start
            return doc.Range(pos, pos)

        snaps = []
        # 从文档尾部往前处理：插入内嵌图会使后续文本位置偏移，先处理后面安全
        targets.sort(key=lambda t: -t[2])
        for k, (refs, why, anchor_start, top0) in enumerate(targets, 1):
            png = os.path.join(img_dir, f"看板_{k:02d}.png")
            ids = [s.ID for s in refs]        # 先取 ID：Delete 之后再访问 COM 属性会"对象已被删除"
            try:
                im = _cur_idx()
                if len(refs) == 1:
                    shp = refs[0]
                    shp.Select()
                    emf = word.Selection.Range.EnhMetaFileBits
                    # 【2026-10-10c 撤销视觉页搬迁】微服务实证：ScrollIntoView+RangeFromPoint 探针
                    # 会被未完成的重新分页骗到（锚 p11 报"视觉页1"→ 看板被搬去第 1 页开头），
                    # 且智慧政法 XML 级对照已证明 accept 后锚点页==视觉页、按锚段插入即正确。
                    # → 一律按锚段插入，不再做任何跨页搬迁。
                    rng = _insert_rng(shp.Top, shp.Anchor)
                    shp.Delete()
                else:
                    sr = doc.Shapes.Range(tuple(im[s.ID] for s in refs))
                    g = sr.Group()                    # 临时成组 → 一张图
                    g.Select()
                    emf = word.Selection.Range.EnhMetaFileBits
                    rng = _insert_rng(g.Top, g.Anchor)
                    g.Delete()
                w_px, h_px = _emf_to_png(emf, png)
                # 内嵌插入，宽度压到版心内（≈可用宽 415pt → 553px @96dpi）
                ish = rng.InlineShapes.AddPicture(FileName=os.path.abspath(png),
                                                  LinkToFile=False, SaveWithDocument=True,
                                                  Range=rng)
                usable = sect.PageSetup.PageWidth - sect.PageSetup.LeftMargin - sect.PageSetup.RightMargin
                if ish.Width > usable:
                    ish.Width = usable
                snaps.append(dict(png=os.path.basename(png), w=w_px, h=h_px, why=why))
                say(f"  [snap] {why} id={ids} → {os.path.basename(png)} {w_px}x{h_px}")
            except Exception as e:
                say(f"  [snap失败] id={ids} {type(e).__name__}: {str(e)[:120]}")

        doc.SaveAs2(os.path.abspath(out_docx), FileFormat=WD_FORMAT_XML)
        res = dict(snaps=snaps,
                   groups=sum(1 for s in snaps if s["why"] == "group"),
                   textboxes=sum(1 for s in snaps if s["why"] == "textbox"))
    finally:
        try:
            if doc is not None: doc.Close(False)
        except Exception: pass
        try: word.Quit()
        except Exception: pass
    # VML→DrawingML 必须在 Word 退出后做：Word 存盘期间锁着 docx（Errno 13）
    if res is not None:
        try:
            nv = vml_to_drawingml(out_docx)     # Word 存 .doc→docx 会把图片写成 VML，LO 只能栅格化
            if nv and log: log(f"  [snap] VML→DrawingML：{nv} 张内嵌图转标准格式")
        except Exception as e:
            if log: log(f"  [snap] VML 转换失败（不影响流程）: {str(e)[:100]}")
    return res if res is not None else dict(snaps=[], groups=0, textboxes=0)


if __name__ == "__main__":
    import sys, io
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")
    src = sys.argv[1]; out = sys.argv[2]; imgd = sys.argv[3]
    r = snap_floating(src, out, imgd, log=print)
    print("结果:", r["groups"], "组合", r["textboxes"], "字框,", len(r["snaps"]), "张图")
