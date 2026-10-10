# -*- coding: utf-8 -*-
"""
V-2 · .doc 静默降级 + L1 校验基准验证（M0 硬门槛）

目的：验证"LO 转换不报错 = 没丢内容"这一假设是否成立。
方法（对齐 docx2epub_M0原型验证清单.md V-2）：
  对每个样本，建立三条独立文本基准并交叉比对：
    A. Word COM  Document.Content.Text            （权威基准）
    B. LO  --convert-to docx 后抽文本              （我们管线阶段0 的实际路径）
    C. LO  --convert-to txt                       （LO 独立提取路径）
  再做 L1 核心比对：以 A（原始输入）为基准 vs B（LO 转出的 docx），
  找出"无报错但丢内容"的样本。

用法：
  "C:/Users/Administrator/AppData/Local/Programs/Python/Python311/python.exe" tools/v2_text_baseline.py

输出：
  work/v2_text_baseline/<样本名>/… 中间产物
  work/v2_text_baseline/report.json
  work/v2_text_baseline/report.md
"""

import json
import os
import re
import shutil
import subprocess
import sys
from collections import Counter
from pathlib import Path

import win32com.client
import pythoncom

PROJ = Path(r"D:/project/epub")
SAMPLES = PROJ / "tests"
OUT = PROJ / "work" / "v2_text_baseline"
SOFFICE = None  # 运行时探测

# 管线将要处理的样本（S9 老格式三件 + 对照 docx 两件）
TARGETS = [
    SAMPLES / "009_S9-老doc_中_很脏" / "009_S9-老doc_中_很脏.doc",
    SAMPLES / "009_S9-rtf_中_很脏" / "009_S9-rtf_中_很脏.rtf",
    SAMPLES / "009_S9-docm_中_很脏" / "009_S9-docm_中_很脏.docm",
    SAMPLES / "002_标题混合_高_很脏" / "002_标题混合_高_很脏.docx",   # 对照：正常 docx
    SAMPLES / "015_混合_高_极端脏" / "015_混合_高_极端脏.docx",       # 对照：最脏 docx
]

W_NS = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"


def find_soffice():
    """探测 soffice.exe（先装后跑，此处只找路径）。"""
    for p in [
        Path(r"C:/Program Files/LibreOffice/program/soffice.exe"),
        Path(r"C:/Program Files (x86)/LibreOffice/program/soffice.exe"),
    ]:
        if p.exists():
            return p
    found = shutil.which("soffice")
    return Path(found) if found else None


def norm_text(s: str) -> str:
    """归一化：去空白、统一引号/连字符。比较用，不改变判定实质。"""
    if not s:
        return ""
    s = s.replace("\r", "\n").replace("\x07", "")  # Word COM 的单元格结束符
    s = re.sub(r"[\u200b\u200c\u200d\ufeff]", "", s)  # 零宽
    s = s.replace("\u2018", "'").replace("\u2019", "'")
    s = s.replace("\u201c", '"').replace("\u201d", '"')
    s = s.replace("\uff0c", ",").replace("\u3002", ".").replace("\uff1a", ":")
    s = re.sub(r"\s+", "", s)
    return s


def para_list_from_docx(path: Path) -> list:
    """lxml 抽 docx 全部段落文本（含表格单元格，w:t 全文拼接即可）。"""
    from lxml import etree
    import zipfile
    texts = []
    with zipfile.ZipFile(path) as z:
        xml = z.read("word/document.xml")
    root = etree.fromstring(xml)
    for p in root.iter(f"{W_NS}p"):
        t = "".join(p.itertext())
        texts.append(t)
    return texts


def docx_all_text(path: Path) -> str:
    return "\n".join(para_list_from_docx(path))


# ---------------------------------------------------------------- Word COM 侧
def word_baseline(word, src: Path, tmpdir: Path) -> dict:
    """Word COM 打开原件：抽 Content.Text（基准A），并 SaveAs docx（word_ref）。"""
    out = {}
    doc = word.Documents.Open(
        str(src), ConfirmConversions=False, ReadOnly=True,
        AddToRecentFiles=False, OpenAndRepair=False, Visible=False,
    )
    try:
        out["wordcom_text"] = doc.Content.Text
        ref = tmpdir / "word_ref.docx"
        # FileFormat=12 → wdFormatXMLDocument（16 在旧版可能存回 .doc，显式 12）
        doc.SaveAs2(str(ref), FileFormat=12)
        out["word_ref_docx"] = ref
    finally:
        doc.Close(SaveChanges=False)
    return out


# ---------------------------------------------------------------- LO 侧
def lo_convert(lo: Path, src: Path, outdir: Path, target: str, extra=None) -> Path:
    outdir.mkdir(parents=True, exist_ok=True)
    profile = outdir / "_lo_profile"
    args = [str(lo), "--headless", "--norestore",
            f"-env:UserInstallation=file:///{profile.as_posix()}",
            "--convert-to", target, "--outdir", str(outdir), str(src)]
    if extra:
        args.insert(-2, extra)
    r = subprocess.run(args, capture_output=True, text=True, timeout=180)
    ext = target.split(":")[0]
    prod = outdir / (src.stem + "." + ext)
    if not prod.exists():
        raise RuntimeError(f"LO 转换无产物: {r.stdout} {r.stderr}")
    return prod


# ---------------------------------------------------------------- 比对
def diff_norm(a: str, b: str) -> dict:
    """字符多重集差 + 段落差。返回摘要。"""
    na, nb = norm_text(a), norm_text(b)
    ca, cb = Counter(na), Counter(nb)
    missing = ca - cb   # A 有 B 没有
    extra = cb - ca     # B 有 A 没有
    return {
        "norm_chars_A": len(na),
        "norm_chars_B": len(nb),
        "missing_chars": sum(missing.values()),
        "missing_top": "".join(sorted(missing.elements()))[:80],
        "extra_chars": sum(extra.values()),
        "extra_top": "".join(sorted(extra.elements()))[:80],
        "identical_after_norm": na == nb,
    }


def para_diff(a_paras: list, b_paras: list) -> dict:
    """段落数差（去空段后）。"""
    an = [p for p in (norm_text(p) for p in a_paras) if p]
    bn = [p for p in (norm_text(p) for p in b_paras) if p]
    return {"paras_A": len(an), "paras_B": len(bn), "para_delta": len(bn) - len(an)}


# ---------------------------------------------------------------- 主流程
def run_one(lo: Path, src: Path, word) -> dict:
    name = src.parent.name
    tmp = OUT / name
    if tmp.exists():
        shutil.rmtree(tmp)
    tmp.mkdir(parents=True)
    rec = {"sample": name, "src": str(src), "fmt": src.suffix}

    # --- 基准 A：Word COM（gencache 静态绑定——本机动态 Dispatch 的 Document 缺 Close） ---
    a_text = None
    try:
        wb = word_baseline(word, src, tmp)
        a_text = wb["wordcom_text"]
        a_paras = para_list_from_docx(wb["word_ref_docx"])
        # 自检：Word 自己存 docx 后抽文本应与 Content.Text 一致（归一化后）
        rec["A_selfcheck"] = diff_norm(a_text, "\n".join(a_paras))
        rec["A_wordcom_chars"] = len(norm_text(a_text))
        rec["A_wordcom_paras"] = len([p for p in map(norm_text, a_paras) if p])
        (tmp / "A_wordcom.txt").write_text(a_text, encoding="utf-8")
        shutil.copy(wb["word_ref_docx"], tmp / "A_word_ref.docx")
    except Exception as e:
        rec["A_error"] = f"{type(e).__name__}: {e}"

    # --- 基准 B：LO docx ---
    b_text = None
    try:
        lo_docx = lo_convert(lo, src, tmp / "lo", "docx")
        b_paras = para_list_from_docx(lo_docx)
        b_text = "\n".join(b_paras)
        shutil.copy(lo_docx, tmp / "B_lo_conv.docx")
    except Exception as e:
        rec["B_error"] = f"{type(e).__name__}: {e}"

    # --- 基准 C：LO txt ---
    try:
        lo_txt = lo_convert(lo, src, tmp / "lo_txt", "txt:Text (encoded):UTF8")
        c_text = lo_txt.read_text(encoding="utf-8", errors="replace")
        shutil.copy(lo_txt, tmp / "C_lo_conv.txt")
    except Exception as e:
        rec["C_error"] = f"{type(e).__name__}: {e}"
        c_text = None

    # --- 核心比对 ---
    if a_text is not None and b_text is not None:
        rec["A_vs_B"] = diff_norm(a_text, b_text)         # L1 核心：原始 vs LO 转出
        rec["A_vs_B"].update(para_diff(a_paras, b_paras))
        rec["A_vs_B_docx_paras"] = para_diff(a_paras, b_paras)
        rec["silent_loss_suspect"] = (
            rec["A_vs_B"]["missing_chars"] > max(5, rec["A_vs_B"]["norm_chars_A"] * 0.01)
        )
    if a_text is not None and c_text is not None:
        rec["A_vs_C"] = diff_norm(a_text, c_text)
    if b_text is not None and c_text is not None:
        rec["B_vs_C"] = diff_norm(b_text, c_text)
    return rec


def main():
    global SOFFICE
    SOFFICE = find_soffice()
    OUT.mkdir(parents=True, exist_ok=True)
    report = []
    pythoncom.CoInitialize()
    word = win32com.client.gencache.EnsureDispatch("Word.Application")
    word.Visible = False
    word.DisplayAlerts = 0
    try:
        for src in TARGETS:
            if not src.exists():
                report.append({"sample": src.parent.name, "error": "样本不存在"})
                continue
            print(f"[V-2] 处理 {src.name} …", flush=True)
            report.append(run_one(SOFFICE, src, word))
    finally:
        word.Quit()

    (OUT / "report.json").write_text(
        json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")
    print(json.dumps(report, ensure_ascii=False, indent=2))
    return 0


if __name__ == "__main__":
    sys.exit(main())
