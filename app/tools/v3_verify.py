# -*- coding: utf-8 -*-
"""
V-3 · 手术合法性验证（用 Word 作独立第三方验证器）

思路：不自己说自己合法，而是让 **Microsoft Word** 打开产物。
      关键参数 `OpenAndRepair=False` —— Word **不会**自动修复；
      若文件损坏/打包不一致会直接抛异常，从而暴露问题。

顺带用 Word 读出的文本与原 docx 的文本做交叉比对 → 独立验证 L1 内容一致。

用法：python v3_verify.py
"""
import json
import re
import zipfile
from pathlib import Path

from lxml import etree
import win32com.client as wc

ROOT = Path(__file__).resolve().parents[2]
STAGE1 = ROOT / "work" / "stage1"
SAMPLES = ROOT / "tests"
OUT = ROOT / "work" / "stage1" / "_v3_verify.json"

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


def w(t):
    return "{%s}%s" % (W, t)


def xml_text(path: Path):
    """从原 docx（lxml）提取纯文本，作为内容基准"""
    try:
        with zipfile.ZipFile(path) as z:
            doc = etree.fromstring(z.read("word/document.xml"))
        return "".join(t.text or "" for t in doc.iter(w("t")))
    except Exception:
        return ""


def norm(s: str):
    """归一化：去空白，便于比较（内容一致而非排版一致）"""
    return re.sub(r"\s+", "", s or "")


def find_origin(name: str):
    for d in sorted(SAMPLES.iterdir()):
        if d.is_dir() and (d / name).exists():
            return d / name
    return None


def main():
    if not STAGE1.exists():
        print("请先运行 stage1_normalize.py")
        return

    products = [f for f in sorted(STAGE1.glob("*.docx"))]
    word = wc.Dispatch("Word.Application")
    word.Visible = False
    word.DisplayAlerts = 0
    results = []
    try:
        for p in products:
            rec = {"product": p.name}
            # --- 对照：原样本 Word 能否打开 ---
            origin = find_origin(p.name)
            try:
                d0 = word.Documents.Open(str(origin), ConfirmConversions=False,
                                         ReadOnly=True, AddToRecentFiles=False,
                                         OpenAndRepair=False)
                rec["origin_openable"] = True
                d0.Close(0)
            except Exception as e:
                rec["origin_openable"] = False
                rec["origin_err"] = str(e)[:120]
            # --- 产物：Word 能否在不修复的情况下打开 ---
            try:
                d = word.Documents.Open(str(p), ConfirmConversions=False,
                                        ReadOnly=True, AddToRecentFiles=False,
                                        OpenAndRepair=False)
                rec["opened"] = True
                rec["段落数"] = d.Paragraphs.Count
                rec["字符数"] = d.Characters.Count
                rec["修订数"] = d.Revisions.Count
                rec["批注数"] = d.Comments.Count
                rec["内嵌形状"] = d.InlineShapes.Count
                rec["浮动形状"] = d.Shapes.Count
                rec["表格数"] = d.Tables.Count
                rec["脚注数"] = d.Footnotes.Count
                rec["域数量"] = d.Fields.Count
                rec["word_text"] = norm(d.Content.Text)
                d.Close(0)
            except Exception as e:
                rec["opened"] = False
                rec["err"] = str(e)[:200]
                results.append(rec)
                print(f"[✗] {p.name}  Word 打不开：{str(e)[:80]}")
                continue

            # --- L1 内容交叉比对（Word 读的 vs 原 docx 的）---
            if origin:
                src = norm(xml_text(origin))
                got = rec.get("word_text", "")
                rec["src_len"] = len(src)
                rec["got_len"] = len(got)
                if src:
                    common = len(set(src) & set(got))
                    rec["字符集重合率"] = round(common / max(len(set(src)), 1) * 100, 2)
                    # 差异定位
                    import difflib
                    sm = difflib.SequenceMatcher(None, src, got)
                    rec["相似度"] = round(sm.quick_ratio() * 100, 2)
            results.append(rec)
            print(f"[✓] {p.name}  段落{rec['段落数']} 字数{rec['字符数']} "
                  f"修订{rec['修订数']} 批注{rec['批注数']} "
                  f"内嵌图{rec['内嵌形状']} 浮动图{rec['浮动形状']} "
                  f"表{rec['表格数']} 脚注{rec['脚注数']}")
    finally:
        try:
            word.Quit()
        except Exception:
            pass

    OUT.write_text(json.dumps(results, ensure_ascii=False, indent=2), encoding="utf-8")

    ok = sum(1 for r in results if r.get("opened"))
    no_repair_needed = sum(1 for r in results if r.get("opened"))
    zero_rev = sum(1 for r in results if r.get("修订数") == 0)
    zero_com = sum(1 for r in results if r.get("批注数") == 0)
    print(f"\n{'='*62}")
    print(f"V-3 结果：Word 免修复打开 {ok}/{len(results)} 本")
    print(f"  修订残留为 0：{zero_rev}/{len(results)}  （验证 N1 接受修订）")
    print(f"  批注残留为 0：{zero_com}/{len(results)}  （验证 N2 删批注）")
    print(f"  报告 → {OUT}")


if __name__ == "__main__":
    main()