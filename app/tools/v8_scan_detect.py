# -*- coding: utf-8 -*-
"""V-8 半扫描/混合件探测规则
信号（docx 侧可直接计算，无需渲染）：
  chars        : w:t 总字符数
  est_pages    : 估算页数 = 整页图张数 + chars/800（经验每页字符量）
  img_ratio    : 图片显示面积合计 / (A4面积 × est_pages)
  char_density : chars / est_pages
样本：14 现有样本（文本件参照） + 合成纯扫描/半扫描件
输出：work/v8_scan_detect/v8_metrics.csv + 阈值建议
"""
import os, re, csv, glob, zipfile, math
from lxml import etree

ROOT = r"D:\project\epub"
S = os.path.join(ROOT, "tests")
OUT = os.path.join(ROOT, "review_samples_html")
W = os.path.join(ROOT, "work", "v8_scan_detect")
os.makedirs(W, exist_ok=True)
NS = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
      'wp': 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing'}
A4_W, A4_H = 210*36000, 297*36000          # EMU
A4_AREA = A4_W * A4_H
CHARS_PER_PAGE = 800                        # 经验值：中文正文每页约 800 字

def metrics(docx):
    with zipfile.ZipFile(docx) as z:
        root = etree.fromstring(z.read('word/document.xml'))
    chars = sum(len(t.text or '') for t in root.iter('{%s}t' % NS['w']))
    n_fullpage = 0; img_area = 0
    for ext in root.iter('{%s}extent' % NS['wp']):
        cx, cy = int(ext.get('cx')), int(ext.get('cy'))
        img_area += cx * cy
        if cy >= 0.7 * A4_H: n_fullpage += 1
    est_pages = max(1, n_fullpage + chars / CHARS_PER_PAGE)
    return dict(chars=chars, n_img=root.string.count('') if False else
                sum(1 for _ in root.iter('{%s}extent' % NS['wp'])),
                n_fullpage=n_fullpage, est_pages=round(est_pages, 1),
                img_ratio=round(img_area / (A4_AREA * est_pages), 2),
                char_density=round(chars / est_pages))

def synth(dirpath):
    """合成：纯扫描件 / 半扫描件 / 轻度半扫描件"""
    import docx as pydocx
    from docx.shared import Cm
    big = os.path.join(S, '_assets', 'huge.jpg')
    cases = []
    d = pydocx.Document(); d.add_picture(big, width=Cm(17), height=Cm(25))  # 整页图
    p = os.path.join(dirpath, 'syn_纯扫描.docx'); d.save(p); cases.append(p)
    d = pydocx.Document()
    for i in range(3):  # 文字页 + 整页图交替
        for j in range(20): d.add_paragraph(f'第{i}组第{j}段。' + '文档结构与样式关系的讨论文本。' * 5)
        d.add_picture(big, width=Cm(17), height=Cm(25))
    p = os.path.join(dirpath, 'syn_半扫描.docx'); d.save(p); cases.append(p)
    d = pydocx.Document()
    for i in range(40): d.add_paragraph(f'第{i}段。' + '文档结构与样式关系的讨论文本。' * 5)
    d.add_picture(big, width=Cm(17), height=Cm(25))
    p = os.path.join(dirpath, 'syn_轻度半扫描.docx'); d.save(p); cases.append(p)
    return cases

def main():
    rows = []
    for sd in sorted(glob.glob(os.path.join(S, '0*'))):
        name = os.path.basename(sd)
        rdir = os.path.join(OUT, name)
        if not os.path.isdir(rdir): continue
        srcf = next((g[0] for c in ('docx','docm') for g in [glob.glob(os.path.join(sd,'*.'+c))] if g), None)
        if srcf is None:
            rows.append([name, '(doc/rtf 跳过)', '', '', '', '', '']); continue
        m = metrics(srcf)
        rows.append([name, m['chars'], m['n_img'], m['n_fullpage'], m['est_pages'], m['img_ratio'], m['char_density']])
    for p in synth(W):
        m = metrics(p)
        rows.append(['[合成]' + os.path.basename(p), m['chars'], m['n_img'], m['n_fullpage'], m['est_pages'], m['img_ratio'], m['char_density']])
    with open(os.path.join(W, 'v8_metrics.csv'), 'w', newline='', encoding='utf-8-sig') as f:
        w = csv.writer(f); w.writerow(['样本','chars','图片数','整页图数','估算页数','图面积占比','字/页'])
        w.writerows(rows)
    print(f"{'样本':<28}{'chars':>7}{'图':>4}{'整页':>4}{'页':>7}{'图占比':>8}{'字/页':>7}")
    for r in rows:
        if r[1] == '': print(f"{r[0]:<28}  (跳过)"); continue
        print(f"{r[0]:<28}{r[1]:>7}{r[2]:>4}{r[3]:>4}{r[4]:>7}{r[5]:>8}{r[6]:>7}")

if __name__ == '__main__':
    main()
