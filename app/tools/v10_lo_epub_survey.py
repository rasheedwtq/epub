# -*- coding: utf-8 -*-
"""
V-10 · LibreOffice docx→epub 输出结构摸底

背景：总纲 A1 决策（LO 作阶段2 引擎）的"LO 能渲染原生图表/输出结构可后处理"
属于设计推断，M0 原清单没有任务验证它。本脚本补上：
  1) 对复杂样本跑 soffice --convert-to epub
  2) 解包产物，统计：XHTML 切分方式、标题/段落/图/表/脚注计数、
     内联样式规模、nav 结构、图片组织、图表是否栅格化（SVG/PNG 化）
  3) 产出 work/v10_lo_epub/report.json + report.md，供 A1 决策复核

用法：
  "C:/Users/Administrator/AppData/Local/Programs/Python/Python311/python.exe" tools/v10_lo_epub_survey.py
"""

import json
import re
import shutil
import subprocess
import sys
import zipfile
from collections import Counter
from pathlib import Path
from lxml import etree

PROJ = Path(r"D:/project/epub")
SAMPLES = PROJ / "tests"
OUT = PROJ / "work" / "v10_lo_epub"

# 覆盖图表/图片/表格/脚注/极脏的 6 本
TARGETS = [
    "004_图表混合_高_很脏",
    "005_图片混合_高_很脏",
    "006_表格混合_高_很脏",
    "007_脚注混合_中_很脏",
    "008_文字混合_高_极端脏",
    "015_混合_高_极端脏",
]

XH = "{http://www.w3.org/1999/xhtml}"
EPUB_NS = "{http://www.idpf.org/2007/opf}"
DC_NS = "{http://purl.org/dc/elements/1.1/}"


def find_soffice():
    for p in [Path(r"C:/Program Files/LibreOffice/program/soffice.exe"),
              Path(r"C:/Program Files (x86)/LibreOffice/program/soffice.exe")]:
        if p.exists():
            return p
    f = shutil.which("soffice")
    return Path(f) if f else None


def lo_convert(lo, src, outdir):
    outdir.mkdir(parents=True, exist_ok=True)
    args = [str(lo), "--headless", "--norestore",
            f"-env:UserInstallation=file:///{(outdir/'_prof').as_posix()}",
            "--convert-to", "epub", "--outdir", str(outdir), str(src)]
    r = subprocess.run(args, capture_output=True, text=True, timeout=300)
    prod = outdir / (src.stem + ".epub")
    if not prod.exists():
        raise RuntimeError(f"无产物: rc={r.returncode} {r.stdout} {r.stderr}")
    return prod


def survey_epub(epub: Path) -> dict:
    rec = {"epub": epub.name, "size_kb": round(epub.stat().st_size / 1024, 1)}
    with zipfile.ZipFile(epub) as z:
        names = z.namelist()
        rec["file_count"] = len(names)
        rec["tree_top"] = sorted(set(n.split("/")[0] for n in names))
        rec["images"] = {
            "count": sum(1 for n in names if n.startswith(("OEBPS/", "Images", "OEBPS/Images")) and re.search(r"\.(png|jpe?g|gif|svg)$", n, re.I)),
            "by_ext": dict(Counter(n.rsplit(".", 1)[-1].lower() for n in names if re.search(r"\.(png|jpe?g|gif|svg|webp)$", n, re.I))),
        }
        xhtmls = [n for n in names if n.endswith((".xhtml", ".html"))]
        rec["xhtml_count"] = len(xhtmls)

        agg = Counter()
        inline_style_chars = 0
        style_attr_count = 0
        svg_inline = 0
        img_refs = 0
        figcaption = 0
        nav_found = None
        for n in xhtmls:
            try:
                root = etree.fromstring(z.read(n))
            except etree.XMLSyntaxError:
                rec.setdefault("xhtml_parse_errors", []).append(n)
                continue
            for el in root.iter():
                tag = etree.QName(el).localname
                agg[tag] += 1
                if tag == "svg":
                    svg_inline += 1
                if tag == "figcaption":
                    figcaption += 1
                st = el.get("style")
                if st:
                    style_attr_count += 1
                    inline_style_chars += len(st)
            if "nav" in root.iter() or root.find(f".//{XH}nav") is not None:
                nav_found = nav_found or n

        rec["tag_top20"] = dict(agg.most_common(20))
        rec["style_attr_count"] = style_attr_count
        rec["inline_style_chars"] = inline_style_chars
        rec["inline_svg_count"] = svg_inline
        rec["figcaption_count"] = figcaption
        rec["nav_file"] = nav_found
        rec["img_refs_in_xhtml"] = img_refs

        # 标题层级
        heads = {h: agg.get(h, 0) for h in ["h1", "h2", "h3", "h4", "h5", "h6"]}
        rec["headings"] = heads

        # OPF 元数据
        opf = [n for n in names if n.endswith(".opf")]
        if opf:
            opf_root = etree.fromstring(z.read(opf[0]))
            rec["opf"] = opf[0]
            rec["meta_title"] = (opf_root.findtext(f".//{DC_NS}title") or "").strip()
            rec["manifest_items"] = len(opf_root.findall(f".//{EPUB_NS}item"))
            rec["cover_prop"] = any(i.get("properties") and "cover-image" in i.get("properties", "")
                                    for i in opf_root.findall(f".//{EPUB_NS}item"))

        # 抽一个正文 xhtml 的前 600 字符供人工看脏度
        body_file = next((n for n in xhtmls if n != nav_found), xhtmls[0] if xhtmls else None)
        if body_file:
            rec["sample_body_file"] = body_file
            rec["sample_body_head"] = z.read(body_file)[:600].decode("utf-8", "replace")
    return rec


def main():
    lo = find_soffice()
    if not lo:
        print("❌ LibreOffice 未安装")
        return 1
    ver = subprocess.run([str(lo), "--version"], capture_output=True, text=True).stdout.strip()
    print("LO:", ver)
    reports = []
    for name in TARGETS:
        d = SAMPLES / name
        srcs = [p for p in d.iterdir() if p.suffix.lower() in (".docx",) ] if d.exists() else []
        if not srcs:
            reports.append({"sample": name, "error": "无 docx"})
            continue
        src = srcs[0]
        print(f"[V-10] {src.name} …", flush=True)
        try:
            epub = lo_convert(lo, src, OUT / name)
            rec = survey_epub(epub)
            rec["sample"] = name
            reports.append(rec)
        except Exception as e:
            reports.append({"sample": name, "error": f"{type(e).__name__}: {e}"})

    OUT.mkdir(parents=True, exist_ok=True)
    payload = {"lo_version": ver, "reports": reports}
    (OUT / "report.json").write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    print(json.dumps(payload, ensure_ascii=False, indent=2))
    return 0


if __name__ == "__main__":
    sys.exit(main())
