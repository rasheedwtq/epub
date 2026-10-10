# -*- coding: utf-8 -*-
"""
阶段1 · 归一化层（Normalize）

依据：docx2epub_归一化层规格.md
原则：lxml 原始 XML 最小手术 —— 只动目标节点，其余部件二进制原样复制。
       zip 原地重写（不用 os.replace，避开 Windows 文件锁；V-4 阶段已验证）。

九步：N0 装载 / N1 接受修订 / N2 删批注 / N3 清零宽+引号 / N4 删空段
      N5 清段首Tab / N6 浮动图inline化+文本框提取 / N7 标题写回Heading
      N8 采集图片EMU / N9 结构报告+无损断言

用法：
    python stage1_normalize.py                 # 批量处理 samples/ 全部 docx
    python stage1_normalize.py --one 002       # 只处理某本
"""
import argparse
import hashlib
import json
import re
import sys
import zipfile
from collections import Counter, defaultdict
from pathlib import Path

from lxml import etree

sys.path.insert(0, str(Path(__file__).resolve().parent))
from v4_heading_eval import cluster  # 复用已验证的聚类算法（承 V-4 四条修正）

ROOT = Path(__file__).resolve().parents[2]
SAMPLES = ROOT / "tests"
WORK = ROOT / "work" / "stage1"

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
WP = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
VML = "urn:schemas-microsoft-com:vml"
CT = "http://schemas.openxmlformats.org/package/2006/content-types"
PR = "http://schemas.openxmlformats.org/package/2006/relationships"
NS = {"w": W, "wp": WP, "v": VML}

ZERO_WIDTH = "‌‍﻿"
HEADING_STYLE = ('<w:style xmlns:w="%s" w:type="paragraph" w:styleId="Heading%d">'
                 '<w:name w:val="heading %d"/><w:basedOn w:val="Normal"/>'
                 '<w:next w:val="Normal"/><w:qFormat/>'
                 '<w:pPr><w:keepNext/><w:outlineLvl w:val="%d"/></w:pPr>'
                 '<w:rPr><w:b/><w:sz w:val="%d"/></w:rPr></w:style>')


def w(t):
    return "{%s}%s" % (W, t)


def ser(el):
    return etree.tostring(el, xml_declaration=True, encoding="UTF-8", standalone=True)


# ================================================================ N0 装载
def load(path: Path):
    with zipfile.ZipFile(path) as z:
        data = {n: z.read(n) for n in z.namelist()}
    doc = etree.fromstring(data["word/document.xml"])
    default_sz = 21
    if "word/styles.xml" in data:
        sr = etree.fromstring(data["word/styles.xml"])
        for st in sr.iter(w("style")):
            if st.get(w("styleId")) == "Normal":
                sz = st.find(".//" + w("sz"))
                if sz is not None and sz.get(w("val")):
                    default_sz = int(sz.get(w("val")))
                break
    return data, doc, default_sz


def paragraphs(doc, default_sz=21):
    """与 V-4 完全一致的段落视图（idx 与 doc.iter(w:p) 顺序一致）"""
    out = []
    for i, p in enumerate(doc.iter(w("p"))):
        txt = "".join(t.text or "" for t in p.iter(w("t")))
        szs, weight = [], defaultdict(int)
        for r in p.findall(w("r")):
            rPr = r.find(w("rPr"))
            if rPr is not None:
                sz = rPr.find(w("sz"))
                if sz is not None and sz.get(w("val")) and sz.get(w("val")).isdigit():
                    v = int(sz.get(w("val")))
                    szs.append(v)
                    weight[v] += max(sum(len(t.text or "") for t in r.iter(w("t"))), 1)
        size = max(weight.items(), key=lambda kv: kv[1])[0] if weight else default_sz
        in_table = any(a.tag == w("tbl") for a in p.iterancestors())
        out.append({"idx": i, "el": p, "text": txt, "size_pt": size / 2.0,
                    "max_pt": (max(szs) if szs else default_sz) / 2.0,
                    "in_table": in_table,
                    "has_draw": p.find(".//" + w("drawing")) is not None,
                    "has_obj": any(p.find(".//" + w(t)) is not None
                                   for t in ("drawing", "pict", "object"))})
    return out


# ================================================================ N1 接受修订
def n1_accept_revisions(doc, st):
    # 先清格式变更记录（否则其中的 ins/del 会被误处理）
    for tag in ("rPrChange", "pPrChange", "tblPrChange", "tcPrChange",
                "trPrChange", "sectPrChange", "cellIns", "cellDel", "cellMerge"):
        for el in list(doc.iter(w(tag))):
            par = el.getparent()
            if par is not None:
                par.remove(el)
                st["格式修订"] += 1
    # 插入 → 保留内容，unwrap
    for ins in list(doc.iter(w("ins"))):
        par = ins.getparent()
        if par is None:
            continue
        i = list(par).index(ins)
        for k, child in enumerate(list(ins)):
            par.insert(i + k, child)
        par.remove(ins)
        st["接受插入"] += 1
    # 删除 → 整体丢弃
    for tag in ("del", "moveFrom"):
        for el in list(doc.iter(w(tag))):
            par = el.getparent()
            if par is not None:
                par.remove(el)
                st["删除删除标记"] += 1
    # 移动 → 保留 moveTo
    for el in list(doc.iter(w("moveTo"))):
        par = el.getparent()
        if par is None:
            continue
        i = list(par).index(el)
        for k, child in enumerate(list(el)):
            par.insert(i + k, child)
        par.remove(el)
        st["接受移动"] += 1


# ================================================================ N2 删批注
def n2_remove_comments(data, doc, st):
    for tag in ("commentRangeStart", "commentRangeEnd"):
        for el in list(doc.iter(w(tag))):
            par = el.getparent()
            if par is not None:
                par.remove(el)
                st["批注锚点"] += 1
    for ref in list(doc.iter(w("commentReference"))):
        run = ref.getparent()
        par = run.getparent() if run is not None else None
        if par is not None:
            par.remove(run)
            st["批注引用"] += 1

    removed = []
    for name in list(data.keys()):
        if re.search(r"word/comments.*\.xml$", name):
            del data[name]
            removed.append(name.split("/")[-1])
    st["删除批注部件"] = len(removed)

    if removed and "[Content_Types].xml" in data:
        ctx = etree.fromstring(data["[Content_Types].xml"])
        for ov in list(ctx):
            pn = ov.get("PartName", "")
            if pn.rsplit("/", 1)[-1] in removed:
                ctx.remove(ov)
                st["同步删CT"] += 1
        data["[Content_Types].xml"] = ser(ctx)

    relname = "word/_rels/document.xml.rels"
    if removed and relname in data:
        rels = etree.fromstring(data[relname])
        for rl in list(rels):
            if "comments" in (rl.get("Type") or ""):
                rels.remove(rl)
                st["同步删rels"] += 1
        data[relname] = ser(rels)


# ================================================================ N3 零宽 + 引号
def n3_clean_text(doc, st):
    for p in doc.iter(w("p")):
        ts = [t for t in p.iter(w("t")) if t.text]
        changed = False
        for t in ts:
            s = "".join(c for c in t.text if c not in ZERO_WIDTH)
            if s != t.text:
                t.text = s
                changed = True
                st["清零宽字符"] += 1
        if any('"' in (t.text or "") for t in ts):
            for t in ts:
                if '"' in (t.text or ""):
                    t.text = t.text.replace('"', "“" if st.get("_q_open", True) else "”")
                    st["直引号转弯"] += 1
                    st["_q_open"] = not st.get("_q_open", True)
            changed = True
        if changed:
            pass


# ================================================================ N4 删空段
def n4_remove_empty(doc, st):
    for p in list(doc.iter(w("p"))):
        par = p.getparent()
        if par is None:
            continue
        # 🔴 表格内段落一律不删：OOXML 规定 w:tc 必须以 w:p 结尾，
        #    删掉单元格里的空段落会让该单元格只剩 w:tbl → Word 报"文件可能已损坏"
        #    （V-3 实测：006/003 正是栽在这一步）
        if any(a.tag == w("tbl") for a in p.iterancestors()):
            st["跳过表内空段"] += 1
            continue
        txt = "".join(t.text or "" for t in p.iter(w("t")))
        if txt.strip():                       # 非空（含 *** 分隔符）→ 保留
            continue
        if any(p.find(".//" + w(t)) is not None
               for t in ("drawing", "pict", "object", "tbl", "bookmarkStart")):
            continue
        if p.find(".//" + w("br")) is not None:
            continue
        par.remove(p)
        st["删空段"] += 1


# ================================================================ N5 段首 Tab
def n5_clean_leading_tabs(doc, st):
    for p in doc.iter(w("p")):
        runs = p.findall(w("r"))
        if not runs:
            continue
        for tab in runs[0].findall(w("tab")):
            runs[0].remove(tab)
            st["清段首Tab"] += 1


# ================================================================ N6 浮动图 + 文本框
def n6_inline_images(doc, st):
    for anchor in list(doc.iter("{%s}anchor" % WP)):
        inl = etree.Element("{%s}inline" % WP)
        for k in ("distT", "distB", "distL", "distR"):
            inl.set(k, "0")
        for tag in ("extent", "effectExtent", "docPr", "cNvGraphicFramePr"):
            el = anchor.find("{%s}%s" % (WP, tag))
            if el is not None:
                inl.append(el)
        g = anchor.find("{http://schemas.openxmlformats.org/drawingml/2006/main}graphic")
        if g is not None:
            inl.append(g)
        par = anchor.getparent()
        if par is not None:
            par.replace(anchor, inl)
            st["浮动图转inline"] += 1

    for tb in list(doc.iter("{%s}textbox" % VML)):
        txbx = tb.find(w("txbxContent"))
        host = tb
        while host is not None and host.tag != w("p"):
            host = host.getparent()
        if txbx is None or host is None:
            continue
        anchor_el = host
        for child in list(txbx):
            anchor_el.addnext(child)
            anchor_el = child
        tb.getparent().remove(tb)
        st["文本框提取"] += 1


# ================================================================ N7 标题写回
def n7_write_headings(data, doc, paras, labels, st, sizes):
    if "word/styles.xml" in data:
        sr = etree.fromstring(data["word/styles.xml"])
        have = {s.get(w("styleId")) for s in sr.iter(w("style"))}
        changed = False
        for lv in (1, 2, 3, 4):
            sid = "Heading%d" % lv
            if sid not in have:
                sr.append(etree.fromstring(HEADING_STYLE % (W, lv, lv, lv - 1, 32)))
                changed = True
                st["注入Heading样式"] += 1
        if changed:
            data["word/styles.xml"] = ser(sr)

    for p in paras:
        lvl = labels.get(p["idx"])
        if not lvl or not lvl.startswith("h"):
            continue
        el = p["el"]
        pPr = el.find(w("pPr"))
        if pPr is None:
            pPr = etree.Element(w("pPr"))
            el.insert(0, pPr)                      # pPr 必须是 p 的第一个子元素
        pStyle = pPr.find(w("pStyle"))
        if pStyle is None:
            pStyle = etree.Element(w("pStyle"))
            pPr.insert(0, pStyle)                  # pStyle 必须是 pPr 的第一个子元素
        pStyle.set(w("val"), "Heading" + lvl[1])
        st["写回Heading"] += 1


# ================================================================ N8 图片 EMU
def n8_images(doc, st):
    out = []
    for blip in doc.iter("{http://schemas.openxmlformats.org/drawingml/2006/main}blip"):
        rid = blip.get("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed")
        anc = blip
        ext = None
        while anc is not None:
            e = anc.find("{%s}extent" % WP)
            if e is not None:
                ext = e
                break
            anc = anc.getparent()
        if ext is None:
            continue
        cx, cy = int(ext.get("cx", 0)), int(ext.get("cy", 0))
        out.append({"rid": rid, "emu_w": cx, "emu_h": cy,
                    "cm_w": round(cx / 360000.0, 2), "cm_h": round(cy / 360000.0, 2),
                    "target_px": min(1200, max(330, int(cx / 9525)))})
    st["图片数"] = len(out)
    return out


# ================================================================ N9 无损断言
def assert_lossless(before, after, st, expect_del=()):
    def parts(d):
        return {k: hashlib.sha256(v).hexdigest()
                for k, v in d.items() if not k.startswith("word/media/")}
    pb, pa = parts(before), parts(after)
    lost = [k for k in pb if k not in pa]
    added = [k for k in pa if k not in pb]
    changed = [k for k in pb if k in pa and pb[k] != pa[k]]
    media_b = {k for k in before if k.startswith("word/media/")}
    media_a = {k for k in after if k.startswith("word/media/")}
    # N2 删批注会**预期**移除 comments*.xml 部件 —— 不算"丢失"
    unexpected = [k for k in lost
                  if k.split("/")[-1] not in expect_del
                  and "comments" not in k.lower()]
    st["部件丢失"] = unexpected
    st["预期删除部件"] = sorted(lost)
    st["部件新增"] = added
    st["部件变更"] = changed
    st["media丢失"] = sorted(media_b - media_a)
    st["无损"] = (not unexpected and not (media_b - media_a)
                  and set(changed) <= {"word/document.xml", "word/styles.xml",
                                       "[Content_Types].xml",
                                       "word/_rels/document.xml.rels"})
    return st["无损"]


# ================================================================ 主流程
def normalize(path: Path, outdir: Path):
    st = defaultdict(int)
    st["_q_open"] = True
    before, doc, default_sz = load(path)
    paras = paragraphs(doc, default_sz)
    sizes_txt = Counter(p["size_pt"] for p in paras).most_common(3)

    n1_accept_revisions(doc, st)
    n2_remove_comments(before, doc, st)
    n3_clean_text(doc, st)
    n4_remove_empty(doc, st)
    n5_clean_leading_tabs(doc, st)
    n6_inline_images(doc, st)

    # 聚类 → 写回（清洗后再聚类，避免把已删内容算进基准）
    paras2 = paragraphs(doc, default_sz)
    labels, steps = cluster(paras2)
    n7_write_headings(before, doc, paras2, labels, st, sizes_txt)
    imgs = n8_images(doc, st)

    before["word/document.xml"] = ser(doc)
    outdir.mkdir(parents=True, exist_ok=True)
    dst = outdir / path.name
    with zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED) as z:
        for n, b in before.items():
            z.writestr(n, b)

    with zipfile.ZipFile(path) as z:
        orig = {n: z.read(n) for n in z.namelist()}
    with zipfile.ZipFile(dst) as z:
        prod = {n: z.read(n) for n in z.namelist()}
    ok = assert_lossless(orig, prod, st,
                         expect_del=("comments.xml", "commentsExtended.xml",
                                     "commentsIds.xml", "commentsExtensible.xml"))

    st.pop("_q_open", None)
    report = {"file": path.name, "无损": ok, "清洗统计": dict(st),
              "正文基准pt": steps.get("正文基准B_pt"),
              "分桶": steps.get("桶"),
              "低置信标题": steps.get("低置信标题_需抽检", []),
              "图片": imgs}
    return dst, report


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--one", help="只处理样本编号，如 002")
    args = ap.parse_args()

    targets = []
    for d in sorted(SAMPLES.iterdir()):
        if not d.is_dir():
            continue
        if args.one and not d.name.startswith(args.one + "_"):
            continue
        for f in sorted(d.iterdir()):
            if f.suffix == ".docx" and not f.name.startswith("~$") \
                    and "encrypted" not in f.name:
                try:
                    with zipfile.ZipFile(f) as z:
                        if "word/document.xml" in z.namelist():
                            targets.append(f)
                except Exception:
                    pass

    WORK.mkdir(parents=True, exist_ok=True)
    reports, ok_n = [], 0
    for p in targets:
        try:
            dst, rep = normalize(p, WORK)
            reports.append(rep)
            ok_n += 1 if rep["无损"] else 0
            flag = "✓" if rep["无损"] else "✗"
            print(f"[{flag}] {p.name}  ->  {dst.name}")
        except Exception as e:
            print(f"[FAIL] {p.name}: {type(e).__name__}: {e}")

    rp = WORK / "_stage1_report.json"
    rp.write_text(json.dumps(reports, ensure_ascii=False, indent=2), encoding="utf-8")
    print(f"\n处理 {len(reports)} 本，无损通过 {ok_n} 本，报告 → {rp}")


if __name__ == "__main__":
    main()