# -*- coding: utf-8 -*-
"""肉眼检查样本包生成器
对 samples/ 下每个样本目录：
  1. 找主文件（docx/doc/rtf/docm）
  2. LO --convert-to epub
  3. 解包 EPUB 到 epub解包/
  4. 复制原始文件，写 说明.txt
输出：D:/project/epub/review_samples/<样本名>/
"""
import os, sys, shutil, subprocess, zipfile, time, re

ROOT = r"D:\project\epub"
SAMPLES = os.path.join(ROOT, "tests")
OUT = os.path.join(ROOT, "review_samples")
SOFFICE = r"C:\Program Files\LibreOffice\program\soffice.exe"
PROF_BASE = r"D:\project\epub\work\review_profiles"

EXTS = (".docx", ".doc", ".rtf", ".docm")

def find_main(d):
    cands = [f for f in os.listdir(d) if f.lower().endswith(EXTS)]
    if not cands:
        return None
    # 优先与目录同名的文件
    cands.sort(key=lambda f: (not f.startswith(os.path.basename(d)[:3]), len(f)))
    return cands[0]

def lo_convert(src, outdir, profile):
    os.makedirs(outdir, exist_ok=True)
    cmd = [SOFFICE, "--headless", "--norestore",
           f"-env:UserInstallation={profile}",
           "--convert-to", "epub", "--outdir", outdir, src]
    r = subprocess.run(cmd, capture_output=True, timeout=180)
    eps = [f for f in os.listdir(outdir) if f.endswith(".epub")]
    return (eps[0] if eps else None), r.returncode

def main():
    only = sys.argv[1] if len(sys.argv) > 1 else None
    dirs = sorted(d for d in os.listdir(SAMPLES)
                  if os.path.isdir(os.path.join(SAMPLES, d)) and d[0].isdigit())
    if only:
        dirs = [d for d in dirs if d.startswith(only)]
    rows = []
    for name in dirs:
        srcdir = os.path.join(SAMPLES, name)
        mainf = find_main(srcdir)
        g = os.path.join(OUT, name)
        if os.path.isdir(g):
            shutil.rmtree(g)
        os.makedirs(g)
        if not mainf:
            with open(os.path.join(g, "说明.txt"), "w", encoding="utf-8") as nf:
                nf.write("无主文件（纯辅助样本）\n内容文件：\n" + "\n".join(os.listdir(srcdir)))
            rows.append((name, "-", "无主文件"))
            continue
        src = os.path.join(srcdir, mainf)
        shutil.copy2(src, os.path.join(g, "原始_" + mainf))
        prof = "file:///" + os.path.join(PROF_BASE, name).replace("\\", "/")
        t0 = time.time()
        try:
            epub, rc = lo_convert(src, os.path.join(g, "_tmp"), prof)
        except subprocess.TimeoutExpired:
            epub, rc = None, -1
        dt = time.time() - t0
        if epub:
            epub_path = os.path.join(g, "_tmp", epub)
            final_epub = os.path.join(g, "转换_" + os.path.splitext(mainf)[0] + ".epub")
            shutil.move(epub_path, final_epub)
            # 解包
            unpack = os.path.join(g, "epub解包")
            with zipfile.ZipFile(final_epub) as z:
                z.extractall(unpack)
            # 统计
            xhtmls, imgs = [], 0
            for dp, _, fs in os.walk(unpack):
                for f in fs:
                    if f.endswith((".xhtml", ".html")):
                        xhtmls.append(os.path.join(dp, f))
                    if f.lower().endswith((".png", ".jpg", ".jpeg", ".gif", ".svg")):
                        imgs += 1
            note = (f"源: {mainf}\n转换: {os.path.basename(final_epub)} ({os.path.getsize(final_epub)//1024} KB, {dt:.1f}s)\n"
                    f"解包 XHTML 页数: {len(xhtmls)}, 图片数: {imgs}\n"
                    f"看什么: 打开 epub解包/index.xhtml 或 *.xhtml 肉眼比对原始文件的层级/图表/表格/脚注。\n")
            open(os.path.join(g, "说明.txt"), "w", encoding="utf-8").write(note)
            rows.append((name, mainf, f"OK {os.path.getsize(final_epub)//1024}KB xhtml={len(xhtmls)} img={imgs} {dt:.1f}s"))
        else:
            open(os.path.join(g, "说明.txt"), "w", encoding="utf-8").write(
                f"源: {mainf}\n转换: ❌ 失败 (rc={rc}, {dt:.1f}s)\n这也是检查点：LO 对该样本的兼容性问题。\n")
            rows.append((name, mainf, f"❌转换失败 rc={rc}"))
        shutil.rmtree(os.path.join(g, "_tmp"), ignore_errors=True)
        print(rows[-1])

    print("\n===== 汇总 =====")
    ok = sum(1 for r in rows if r[2].startswith("OK"))
    print(f"共 {len(rows)} 组，成功 {ok}，失败 {len(rows)-ok}")
    # 顶层 README
    with open(os.path.join(OUT, "README_检查指引.md"), "w", encoding="utf-8") as f:
        f.write("# 肉眼检查样本包\n\n每组 = 原始文件 + 转换.epub + epub解包/ + 说明.txt\n\n")
        f.write("| 样本 | 源 | 结果 |\n|---|---|---|\n")
        for n, s, r in rows:
            f.write(f"| {n} | {s} | {r} |\n")
        f.write("\n检查建议重点：标题层级是否保留、图表是否丢失（004/015 已知 EPUB 滤镜会丢）、表格结构、脚注、列表缩进、封面（011）。\n")

if __name__ == "__main__":
    main()
