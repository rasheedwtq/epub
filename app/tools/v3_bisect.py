# -*- coding: utf-8 -*-
"""V-3 辅助：二分定位是哪一步手术导致 Word 打不开"""
import sys
import zipfile
from collections import defaultdict
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
import stage1_normalize as S
from v4_heading_eval import cluster
from lxml import etree

TARGET = sys.argv[1] if len(sys.argv) > 1 else "006"
ROOT = S.ROOT
OUT = ROOT / "work" / "bisect"

src = None
for d in sorted((ROOT / "tests").iterdir()):
    if d.is_dir() and d.name.startswith(TARGET + "_"):
        for f in d.iterdir():
            if f.suffix == ".docx":
                src = f
                break
    if src:
        break

OUT.mkdir(parents=True, exist_ok=True)
before, doc, ds = S.load(src)
st = defaultdict(int)
st["_q_open"] = True
snaps = []


def snap(tag):
    data = dict(before)
    data["word/document.xml"] = S.ser(doc)
    p = OUT / f"{TARGET}_{tag}.docx"
    with zipfile.ZipFile(p, "w", zipfile.ZIP_DEFLATED) as z:
        for n, b in data.items():
            z.writestr(n, b)
    snaps.append((tag, p))


snap("00_原始")
S.n1_accept_revisions(doc, st); snap("01_n1接受修订")
S.n2_remove_comments(before, doc, st); snap("02_n2删批注")
S.n3_clean_text(doc, st); snap("03_n3零宽引号")
S.n4_remove_empty(doc, st); snap("04_n4删空段")
S.n5_clean_leading_tabs(doc, st); snap("05_n5清Tab")
S.n6_inline_images(doc, st); snap("06_n6浮动图")
paras = S.paragraphs(doc, ds)
labels, steps = cluster(paras)
S.n7_write_headings(before, doc, paras, labels, st, None); snap("07_n7标题写回")

print(f"已生成 {len(snaps)} 个快照 → {OUT}")

# 用 Word 验证
import win32com.client as wc
word = wc.Dispatch("Word.Application")
word.Visible = False
word.DisplayAlerts = 0
print()
prev = True
try:
    for tag, p in snaps:
        try:
            d = word.Documents.Open(str(p), ConfirmConversions=False,
                                     ReadOnly=True, AddToRecentFiles=False,
                                     OpenAndRepair=False)
            mark = "✓"
            d.Close(0)
        except Exception:
            mark = "✗ 打不开"
        flag = "" if (mark == "✓") == prev else "   ←← 这里开始坏"
        print(f"  [{mark}] {tag}{flag}")
        prev = (mark == "✓")
finally:
    try:
        word.Quit()
    except Exception:
        pass