# -*- coding: utf-8 -*-
"""L1/L2 正式验收检查器（M1 出口标准）
L1 内容一致：原始输入文本（排除项见 EXCLUDE）vs 产物 XHTML 文本，段落命中率 + 字符覆盖率
L2 结构一致：图片/图表/公式/文本框/表格/脚注 源↔产物计数对账

排除口径（依据总纲 §19 V-2 结论）：
  - m:t（公式内文字）→ 已栅格化，不计 L1，归 L2 公式计数
  - w:del / w:delText（修订删除）→ 非正文内容
  - w:hdr / w:ftr（页眉页脚）→ 页面装饰，A1′ 不转换
  - 文本框文字 → 计入 L1（A1′ 保字），文本框另计 L2
用法：python l1l2_check.py [样本目录名 ...]（缺省跑全部）
"""
import os, re, sys, glob, csv, zipfile
from lxml import etree, html as LH

ROOT = r"D:\project\epub"
S = os.path.join(ROOT, "tests")
OUT = os.path.join(ROOT, "review_samples_html")
W = os.path.join(ROOT, "work", "l1l2")
os.makedirs(W, exist_ok=True)
WNS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
MNS = 'http://schemas.openxmlformats.org/officeDocument/2006/math'
WP = 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing'
PIC = 'http://schemas.openxmlformats.org/drawingml/2006/picture'
CHT = 'http://schemas.openxmlformats.org/drawingml/2006/chart'
WPS = 'http://schemas.microsoft.com/office/word/2010/wordprocessingShape'
STRIP = re.compile(r'[\s\u00A0\u200B\uFEFF]+')

def norm(s):
    return STRIP.sub('', s or '')

def run_fragments(root, wns=WNS):
    """run 粒度文本片段（排除 w:del、m:t）；匹配粒度=片段，避免修订交织误报"""
    frags = []
    for t in root.iter():
        q = etree.QName(t)
        if q.localname in ('t', 'delText') and q.namespace == WNS:
            anc = t.getparent(); deleted = False
            while anc is not None:
                if etree.QName(anc).localname == 'del': deleted = True; break
                anc = anc.getparent()
            if deleted: continue
            v = norm(t.text)
            if v: frags.append(v)
    return frags

LIST_NO = re.compile(r'^\d{1,3}([.、)．)]?)$')  # 手打编号段落 → 列表等价转换豁免

def para_texts(root, wns=WNS):
    """每个 w:p 的聚合文本（排除 w:del、m:t）"""
    out = []
    for p in root.iter('{%s}p' % wns):
        buf = []
        for t in p.iter():
            tag = etree.QName(t).localname
            if tag == 't' and etree.QName(t).namespace == WNS:
                anc = t.getparent(); deleted = False
                while anc is not None:
                    if etree.QName(anc).localname == 'del': deleted = True; break
                    anc = anc.getparent()
                if not deleted: buf.append(t.text or '')
            elif tag == 'tab':
                buf.append(' ')
        out.append(norm(''.join(buf)))
    return out

def _tbl_effective(tbl_el, WNS):
    """表格有效文本（剔除零宽符/空白）。假表格（纯摆图壳）两侧口径一致地不计入 tbl 数"""
    invis = "\u200d\ufeff\xa0 \t\r\n"
    txt = ''.join(t.text or '' for t in tbl_el.iter('{%s}t' % WNS))
    return txt.strip(invis)

def source_parts(path):
    """返回 dict: l1_paras, l2 counts。.doc/.rtf 返回 None"""
    if not path.lower().endswith(('.docx', '.docm')): return None
    z = zipfile.ZipFile(path)
    names = z.namelist()
    doc = etree.fromstring(z.read('word/document.xml'))
    parts = {'l1_frags': [], 'pic': 0, 'chart': 0, 'math': 0, 'textbox': 0, 'tbl': 0, 'fn': 0}
    parts['l1_frags'] += run_fragments(doc)
    # pic/chart 按元素直数（2026-10-09 修订）：旧口径按 wp:extent 数，组合图形(wpg)内多图只算 1 → 期望图漏数
    parts['pic'] = sum(1 for _ in doc.iter('{%s}pic' % PIC))
    parts['chart'] = sum(1 for _ in doc.iter('{%s}chart' % CHT))
    for ext in doc.iter('{%s}extent' % WP):
        d = ext.getparent()
        if d.find('.//{%s}txbx' % WPS) is not None: parts['textbox'] += 1
    parts['math'] = sum(1 for _ in doc.iter('{%s}oMath' % MNS))
    parts['tbl'] = sum(1 for t in doc.iter('{%s}tbl' % WNS) if _tbl_effective(t, WNS))
    ref_ids = set(re.findall(r'footnoteReference w:id="(-?\d+)"',
                             z.read('word/document.xml').decode('utf-8', 'ignore')))
    if 'word/footnotes.xml' in names:
        fx = etree.fromstring(z.read('word/footnotes.xml'))
        fns = [f for f in fx.iter('{%s}footnote' % WNS)
               if f.get('{%s}type' % WNS) is None
               and f.get('{%s}id' % WNS) in ref_ids]  # 只计被正文引用的脚注（孤儿脚注源无引用，LO 不输出属正确）
        parts['fn'] = len(fns)
        for f in fns: parts['l1_frags'] += run_fragments(f)
    if 'word/endnotes.xml' in names:
        ex = etree.fromstring(z.read('word/endnotes.xml'))
        ens = [e for e in ex.iter('{%s}endnote' % WNS)
               if e.get('{%s}type' % WNS) is None]
        for e in ens: parts['l1_frags'] += run_fragments(e)
    return parts

_INVIS = "\u200d\ufeff\xa0 \t\r\n"

def product_text_and_counts(sample_dir):
    txt_all, imgs, tables = [], 0, 0
    for xh in sorted(glob.glob(os.path.join(sample_dir, 'epub解包', 'OEBPS', 'sections', '*.xhtml'))):
        doc = LH.fromstring(open(xh, 'rb').read())
        txt_all.append(doc.text_content())
        imgs += sum(1 for _ in doc.iter('img'))
        tables += sum(1 for t in doc.iter('table') if t.text_content().strip(_INVIS))
    return norm(''.join(txt_all)), imgs, tables

def product_text_and_counts_from_dir(updir):
    """直接给 epub解包 目录的版本（pipeline.py 闸门用）"""
    txt_all, imgs, tables = [], 0, 0
    for xh in sorted(glob.glob(os.path.join(updir, 'OEBPS', 'sections', '*.xhtml'))):
        doc = LH.fromstring(open(xh, 'rb').read())
        txt_all.append(doc.text_content())
        imgs += sum(1 for _ in doc.iter('img'))
        tables += sum(1 for t in doc.iter('table') if t.text_content().strip(_INVIS))
    return norm(''.join(txt_all)), imgs, tables

def main():
    targets = sys.argv[1:] or None
    rows = []
    for sd in sorted(glob.glob(os.path.join(OUT, '0*'))):
        name = os.path.basename(sd)
        if targets and not any(t in name for t in targets): continue
        srcf = next((g[0] for c in ('docx', 'docm') for g in [glob.glob(os.path.join(S, name, '*.' + c))] if g), None)
        if srcf is None:
            rows.append([name, 'SKIP(.doc/.rtf)', '', '', '', '', '', '', '', '', '', '']); continue
        sp = source_parts(srcf)
        ptxt, pimg, ptbl = product_text_and_counts(sd)
        frags = [p for p in sp['l1_frags'] if p]
        exempt = [p for p in frags if LIST_NO.match(p)]      # 手打编号→列表等价转换豁免
        check = [p for p in frags if not LIST_NO.match(p)]
        hit = [p for p in check if p in ptxt]
        chars = sum(len(p) for p in check)
        cov = sum(len(p) for p in hit)
        miss = [p[:36] for p in check if p not in ptxt][:8]
        # L2 对账：产物 img 应 = pic + chart + math（textbox 走文字路径不产图）
        expect_img = sp['pic'] + sp['chart'] + sp['math']
        l2 = 'OK' if pimg == expect_img and ptbl == sp['tbl'] else 'DIFF'
        rows.append([name, f"{len(check)}(+{len(exempt)}编号豁免)", f"{len(hit)}", f"{cov}/{chars}",
                     f"{cov/chars*100:.2f}%" if chars else "n/a",
                     f"pic{sp['pic']}+cht{sp['chart']}+m{sp['math']}={expect_img}",
                     str(pimg), f"tbl {sp['tbl']}→{ptbl}", f"fn {sp['fn']}",
                     f"txbx {sp['textbox']}", l2, ' | '.join(miss)])
    with open(os.path.join(W, 'l1l2_report.csv'), 'w', newline='', encoding='utf-8-sig') as f:
        w = csv.writer(f)
        w.writerow(['样本', '校验片段数', '命中', '命中字符', 'L1字符覆盖率', '期望img(pic+cht+math)', '产物img', '表(源→产物)', '脚注(源,被引用)', '文本框(源)', 'L2', '未命中片段示例'])
        w.writerows(rows)
    print(f"{'样本':<26}{'L1覆盖':>9}  {'期望img':>9} {'实际':>5}  {'表':>7} {'L2':>4}  未命中示例")
    for r in rows:
        if r[1].startswith('SKIP'): print(f"{r[0]:<26}  SKIP"); continue
        print(f"{r[0]:<26}{r[4]:>9}  {r[5]:>9} {r[6]:>5}  {r[7]:>7} {r[10]:>4}  {r[11][:60]}")
    print('\n明细: work/l1l2/l1l2_report.csv')

if __name__ == '__main__':
    main()
