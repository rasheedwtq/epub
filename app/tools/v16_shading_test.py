# -*- coding: utf-8 -*-
"""V-16 底纹存活测试：段落底纹 + 文字底纹（run 级）走 A1' 全管线，检查 HTML/EPUB 产物存活情况"""
import os, sys, subprocess, glob, shutil, re
from pathlib import Path
ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "tools"))

# ---------- 1. 造测试 docx ----------
from docx import Document
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

doc = Document()
doc.add_heading("底纹存活测试", level=1)

def add_shaded_para(text, fill, run_level=False, bold=False):
    p = doc.add_paragraph()
    if not run_level:
        pPr = p._p.get_or_add_pPr()
        shd = OxmlElement("w:shd")
        shd.set(qn("w:val"), "clear"); shd.set(qn("w:color"), "auto"); shd.set(qn("w:fill"), fill)
        pPr.append(shd)
    r = p.add_run(text + ("（run级）" if run_level else "（段落级）"))
    if bold: r.bold = True
    if run_level:
        rPr = r._r.get_or_add_rPr()
        shd = OxmlElement("w:shd")
        shd.set(qn("w:val"), "clear"); shd.set(qn("w:color"), "auto"); shd.set(qn("w:fill"), fill)
        rPr.append(shd)
    return p

add_shaded_para("灰色段落底纹 D9D9D9", "D9D9D9")
add_shaded_para("深灰段落底纹 808080", "808080")
add_shaded_para("浅黄段落底纹 FFFF00", "FFFF00")
add_shaded_para("灰色文字底纹 C0C0C0", "C0C0C0", run_level=True)
add_shaded_para("灰色文字底纹混合普通文字", "D9D9D9", run_level=True)
# 同段混合：一半有底纹一半没有
p = doc.add_paragraph()
r1 = p.add_run("正常文字+")
r2 = p.add_run("灰色底纹文字E0E0E0")
rPr = r2._r.get_or_add_rPr()
shd = OxmlElement("w:shd"); shd.set(qn("w:val"), "clear"); shd.set(qn("w:color"), "auto"); shd.set(qn("w:fill"), "E0E0E0")
rPr.append(shd)
# 标题带底纹
h = doc.add_heading("带底纹的二级标题", level=2)
pPr = h._p.get_or_add_pPr()
shd = OxmlElement("w:shd"); shd.set(qn("w:val"), "clear"); shd.set(qn("w:color"), "auto"); shd.set(qn("w:fill"), "BFBFBF")
pPr.append(shd)
doc.add_heading("普通收尾章", level=1)
doc.add_paragraph("结尾正文，无底纹。")

test_dir = ROOT / "work" / "v16_shading"
test_dir.mkdir(parents=True, exist_ok=True)
src = test_dir / "底纹测试.docx"
doc.save(src)
print("docx 已生成:", src)

# ---------- 2. 走管线（全量） ----------
import yaml
cfg_path = ROOT / "config.yaml"
out = test_dir / "out"
shutil.rmtree(out, ignore_errors=True)
r = subprocess.run([sys.executable, str(ROOT/"tools"/"pipeline.py"), "convert", str(src),
                    "-c", str(cfg_path), "-o", str(out), "--book-id", "底纹测试"],
                   cwd=str(ROOT), capture_output=True, timeout=600)
line = r.stdout.strip().splitlines()[-1] if r.stdout.strip() else ""
print("管线结果:", line[:300])
if r.returncode != 0:
    print("STDERR:", r.stderr.decode("utf-8","ignore")[-500:]); sys.exit(1)

# ---------- 3. 检查产物 ----------
unpacked = out / "底纹测试" / "epub解包" / "OEBPS"
secs = sorted(glob.glob(str(unpacked / "sections" / "*.xhtml")))
allhtml = "".join(open(f, encoding="utf-8").read() for f in secs)
css = open(unpacked / "styles.css", encoding="utf-8").read()

probes = ["灰色段落底纹 D9D9D9", "深灰段落底纹 808080", "灰色文字底纹 C0C0C0",
          "正常文字+", "带底纹的二级标题", "结尾正文，无底纹。"]
print("\n== 文本存活 ==")
for t in probes:
    print(f"  {'OK ' if t in re.sub(r'<[^>]+>','',allhtml) else 'MISS'} | {t}")

print("\n== 样式存活 ==")
bg_styles = re.findall(r'style="[^"]*background[^"]*"', allhtml)
print("  产物内联 background 声明:", len(bg_styles), bg_styles[:6])
print("  产物 styles.css 中 background:", "background" in css)
# LO 中间 HTML（stage2 缓存）里 background 是否存在——判断丢在哪个阶段
lo_html_files = glob.glob(str(ROOT / "work" / "pipeline" / "底纹测试" / "**" / "*.html"), recursive=True)
for f in lo_html_files[:2]:
    t = open(f, encoding="utf-8", errors="ignore").read()
    print(f"  中间HTML {Path(f).name}: background 出现 {len(re.findall('background', t))} 次")
