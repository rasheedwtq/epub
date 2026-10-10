# -*- coding: utf-8 -*-
"""
样本校验器：验证 22 个样本的陷阱是否真的注入成功（防止"假样本"）
输出：samples/_verify_report.md
"""
import zipfile
from pathlib import Path
from lxml import etree

ROOT = Path(__file__).resolve().parents[2]
OUT = ROOT / "tests"

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
NS = {
    "w": W,
    "wp": "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
    "v": "urn:schemas-microsoft-com:vml",
    "m": "http://schemas.openxmlformats.org/officeDocument/2006/math",
}

def w(tag):
    return "{%s}%s" % (W, tag)


MAGIC = {
    b"PK\x03\x04": "ZIP(docx)",
    b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1": "CFB(旧doc/加密docx)",
    b"{\\rtf": "RTF",
}


def magic(p: Path) -> str:
    try:
        head = p.read_bytes()[:8]
    except Exception:
        return "EMPTY"
    for k, v in MAGIC.items():
        if head.startswith(k):
            return v
    return "UNKNOWN"


FEATURES = [
    ("修订痕迹",      "//w:ins | //w:del",                 "word/document.xml"),
    ("批注锚点",      "//w:commentRangeStart",            "word/document.xml"),
    ("批注部件",      None,                              "__COMMENTS__"),
    ("脚注引用",      "//w:footnoteReference",            "word/document.xml"),
    ("脚注部件",      None,                              "__FOOTNOTES__"),
    ("浮动图anchor",  "//wp:anchor",                     "word/document.xml"),
    ("文本框",        "//v:textbox",                     "word/document.xml"),
    ("OMML公式",      "//m:oMath",                       "word/document.xml"),
    ("首字下沉",      "//w:framePr[@w:dropCap]",         "word/document.xml"),
    ("域代码",        "//w:fldSimple",                   "word/document.xml"),
    ("书签",          "//w:bookmarkStart",               "word/document.xml"),
    ("Tab缩进",       "//w:tab",                         "word/document.xml"),
    ("合并单元格",    "//w:gridSpan | //w:vMerge",       "word/document.xml"),
    ("嵌套表格",      "//w:tbl//w:tbl",                  "word/document.xml"),
    ("OLE/图表",      "//w:object | //c:chart",          "word/document.xml"),
    ("图片",          "//w:drawing",                     "word/document.xml"),
    ("16bit/大图",    None,                              "__BIGIMG__"),
]


def inspect(p: Path) -> dict:
    info = {
        "name": p.name,
        "kb": p.stat().st_size / 1024,
        "magic": magic(p),
        "zip": zipfile.is_zipfile(p),
        "feats": set(),
        "note": "",
    }
    if not info["zip"]:
        if info["magic"].startswith("CFB"):
            if p.suffix.lower() == ".docx":
                info["feats"].add("加密(标准加密/非ZIP)")
                info["note"] = "CFB 容器 + .docx 扩展名 → **标准加密文件**，应识别为加密并跳过"
            else:
                info["note"] = "CFB 容器 → 旧版 .doc 二进制格式（Word 97-2003）"
        elif info["magic"] == "RTF":
            info["note"] = "RTF 文本格式"
        elif info["magic"] == "EMPTY":
            info["note"] = "空文件（应被隔离）"
        return info

    try:
        with zipfile.ZipFile(p) as z:
            names = z.namelist()
            xml = z.read("word/document.xml")
            rels = ""
            if "word/_rels/document.xml.rels" in names:
                rels = z.read("word/_rels/document.xml.rels").decode("utf8", "ignore")
    except Exception as e:
        info["note"] = f"ZIP 打开失败：{e}"
        return info

    doc = etree.fromstring(xml)

    # ---------- 内容级检测（校验"看不见"的样本属性）----------
    try:
        szs = {int(e.get(w("val"))) for e in doc.iter(w("sz"))
               if e.get(w("val")) and e.get(w("val")).isdigit()}
        if len(szs) >= 3:
            info["feats"].add("多级字号(%d级: %spt)" % (
                len(szs), ",".join(str(s // 2) for s in sorted(szs, reverse=True))))
    except Exception as e:
        print(f"  [warn] 字号检测失败 {p.name}: {e}")
    try:
        hs = doc.xpath("//w:pStyle[starts-with(@w:val,'Heading')]", namespaces=NS)
        if hs:
            info["feats"].add("使用Heading样式(%d处)" % len(hs))
        else:
            info["feats"].add("无Heading样式(纯字号)")
    except Exception:
        pass
    try:
        txt = "".join(t.text or "" for t in doc.iter(w("t")))
        RARE = "䶮龘齾爩靐飝㐂䨻麤龗"
        if any(c in txt for c in RARE):
            info["feats"].add("生僻字(扩展区)")
        EMOJI = "📚🔬🧪⚗✅🎯🚀🧭"
        if any(c in txt for c in EMOJI):
            info["feats"].add("Emoji")
        if any(c in txt for c in "繁體這個"):
            info["feats"].add("繁体中文")
        if any(0xFF01 <= ord(c) <= 0xFF5E for c in txt):
            info["feats"].add("全角字符")
        if "‌" in txt or "﻿" in txt:
            info["feats"].add("零宽字符")
        if '"' in txt and "“" in txt:
            info["feats"].add("直引号+弯引号混用")
        if "***" in txt:
            info["feats"].add("场景分隔符")
    except Exception as e:
        print(f"  [warn] 文本检测失败 {p.name}: {e}")

    for label, xpath, part in FEATURES:
        if part == "__COMMENTS__":
            if any(n.endswith("comments.xml") for n in names):
                info["feats"].add(label)
            continue
        if part == "__FOOTNOTES__":
            if any(n.endswith("footnotes.xml") for n in names):
                info["feats"].add(label)
            continue
        if part == "__BIGIMG__":
            big = [n for n in names if n.startswith("word/media/")]
            if big:
                tot = sum(z.getinfo(n).file_size for n in big)
                if tot > 400_000:
                    info["feats"].add(label)
            continue
        try:
            if doc.xpath(xpath, namespaces=NS):
                info["feats"].add(label)
        except Exception:
            pass
    if rels:
        if "comments" in rels:
            info["feats"].add("批注关系")
        if "footnotes" in rels:
            info["feats"].add("脚注关系")
    return info


def main():
    rows = []
    for p in sorted(OUT.rglob("*")):
        if p.is_file() and p.suffix.lower() in (".docx", ".doc", ".rtf", ".docm"):
            rows.append(inspect(p))

    print(f"{'文件':<44}{'大小':>9}  {'魔数':<18} 特征")
    print("-" * 130)
    for r in rows:
        feats = ",".join(sorted(r["feats"])) or "-"
        print(f"{r['name']:<44}{r['kb']:>7.0f}KB  {r['magic']:<18} {feats}")

    rep = OUT / "_verify_report.md"
    with open(rep, "w", encoding="utf-8") as f:
        f.write("# 样本校验报告\n\n> 由 verify_samples.py 自动生成\n\n")
        f.write("| 文件 | 大小 | 真实格式 | 已注入陷阱 | 备注 |\n|---|---|---|---|---|\n")
        for r in rows:
            f.write("| `{}` | {:.0f} KB | {} | {} | {} |\n".format(
                r["name"], r["kb"], r["magic"],
                "、".join(sorted(r["feats"])) or "—", r["note"]))
    print("\n[report] %s" % rep)


if __name__ == "__main__":
    main()
