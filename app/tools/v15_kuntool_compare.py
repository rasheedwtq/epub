# -*- coding: utf-8 -*-
"""V-15: KunTool vs A1' 管线同口径对照（结构 + 文本保真 + 哨兵对象）."""
import os, re, zipfile

ROOT = r"D:\project\epub"
KUN = os.path.join(ROOT, "KUN TOOL")
OURS = os.path.join(ROOT, "review_samples_html")
S = os.path.join(ROOT, "tests")

SENTINELS = {
    "005_图片混合_高_很脏": ["文本框"],
    "015_混合_高_极端脏": ["文本框内容必须被提取"],
    "006_表格混合_高_很脏": ["列10"],
    "013_列表混合_低_轻微脏": ["sec_appendix", "跳转"],
}


def epub_text(path):
    z = zipfile.ZipFile(path)
    t = ""
    for n in z.namelist():
        if n.endswith((".xhtml", ".html", ".htm")):
            t += z.read(n).decode("utf-8", "replace")
    # 去 style/script
    t = re.sub(r"<(style|script)[^>]*>.*?</\1>", "", t, flags=re.S)
    plain = re.sub(r"<[^>]+>", "", t)
    imgs = len(re.findall(r"<img[\s>]", t))
    links = len(re.findall(r"<a[\s>]", t))
    tables = len(re.findall(r"<table[\s>]", t))
    return plain, imgs, links, tables


def src_text(docx_path):
    z = zipfile.ZipFile(docx_path)
    doc = z.read("word/document.xml").decode("utf-8", "replace")
    parts = re.findall(r"<w:t[^>]*>([^<]*)</w:t>", doc)
    return "".join(parts)


def norm(s):
    return re.sub(r"\s+", "", s)


rows = []
for f in sorted(os.listdir(KUN)):
    if not f.endswith(".epub"):
        continue
    name = os.path.splitext(f)[0]
    kun_p = os.path.join(KUN, f)
    our_dir = os.path.join(OURS, name)
    our_p = None
    if os.path.isdir(our_dir):
        for x in os.listdir(our_dir):
            if x.endswith(".epub") and x.startswith("转换_"):
                our_p = os.path.join(our_dir, x)
    kp, ki, kl, kt = epub_text(kun_p)
    if our_p:
        op, oi, ol, ot = epub_text(our_p)
    else:
        op, oi, ol, ot = "", -1, -1, -1
    row = dict(name=name, kun=dict(chars=len(norm(kp)), imgs=ki, links=kl, tables=kt),
               ours=dict(chars=len(norm(op)), imgs=oi, links=ol, tables=ot))
    # 源文本基准（仅 docx）
    srcd = os.path.join(S, name)
    if os.path.isdir(srcd):
        dx = [x for x in os.listdir(srcd) if x.endswith(".docx")]
        if dx:
            st = norm(src_text(os.path.join(srcd, dx[0])))
            if st:
                row["src_chars"] = len(st)
                row["kun_cov"] = round(len(norm(kp)) / len(st) * 100, 1)
                row["ours_cov"] = round(len(norm(op)) / len(st) * 100, 1) if op else 0
    rows.append(row)

print(f"{'样本':<28} {'源字':>7} | {'KUN字':>7} {'KUN图':>5} {'KUN表':>5} {'KUN链':>5} {'KUN覆盖':>7} | {'A1字':>7} {'A1图':>5} {'A1表':>5} {'A1链':>5} {'A1覆盖':>7}")
print("-" * 120)
for r in rows:
    k, o = r["kun"], r["ours"]
    print(f"{r['name']:<28} {r.get('src_chars','-'):>7} | {k['chars']:>7} {k['imgs']:>5} {k['tables']:>5} {k['links']:>5} {str(r.get('kun_cov','-')):>7} | "
          f"{o['chars']:>7} {o['imgs']:>5} {o['tables']:>5} {o['links']:>5} {str(r.get('ours_cov','-')):>7}")

# 哨兵对象检查
print("\n===== 哨兵对象存活检查 =====")
for name, kws in SENTINELS.items():
    kun_p = os.path.join(KUN, name + ".epub")
    if not os.path.exists(kun_p):
        continue
    kt, _, _, _ = epub_text(kun_p)
    our_dir = os.path.join(OURS, name)
    op = ""
    for x in os.listdir(our_dir):
        if x.endswith(".epub") and x.startswith("转换_"):
            our_p = os.path.join(our_dir, x)
    ot, _, _, _ = epub_text(our_p)
    for kw in kws:
        print(f"  {name} | '{kw}': KUN={'存活' if kw in kt else '丢失'} | A1={'存活' if kw in ot else '丢失'}")

# 008 tab 缩进（NBSP 计数）
print("\n===== 008 tab 缩进 =====")
for label, path in (("KUN", os.path.join(KUN, "008_文字混合_高_极端脏.epub")),
                    ("A1", os.path.join(OURS, "008_文字混合_高_极端脏", "转换_008_文字混合_高_极端脏.epub"))):
    z = zipfile.ZipFile(path)
    n = 0
    for x in z.namelist():
        if x.endswith(".xhtml"):
            n += z.read(x).decode("utf-8", "replace").count("\u00a0\u00a0\u00a0")
    print(f"  {label}: 3连NBSP 块 = {n}（源 8 个 tab）")
