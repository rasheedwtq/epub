# -*- coding: utf-8 -*-
"""V-5 V1 四指标报警式筛查摸底
指标（报警式，不判合格）：
  A 图注分离   : 图注关键词块(图N/表N/Fig/Table)距最近 img ≥2 块 → 报警
  B 图文脱节   : img 相邻块为空段/另一 img（连续图群）→ 报警（浮动图群高发区）
  C 超大表     : 表格列数 ≥8 或嵌套表 → 报警（溢出高发区）
  D 跨章图     : 章节末/首元素是 img → 报警（切章切断图流）
输出：work/v5_v1_indicators/v5_alarms.csv + v5_context.txt（人工复核用）
"""
import os, re, csv, glob
from lxml import html as LH

ROOT = r"D:\project\epub"
OUT = os.path.join(ROOT, "review_samples_html")
W = os.path.join(ROOT, "work", "v5_v1_indicators")
os.makedirs(W, exist_ok=True)
CAP_RE = re.compile(r'^\s*(图|表|Fig(?:ure)?\.?|Table)\s*\d+', re.I)

def blocks(xhtml):
    """展开为 [(tag, text)] 的块级序列（表格记为 table:Ncol）"""
    doc = LH.fromstring(open(xhtml, 'rb').read())
    seq = []
    for el in doc.body.iter():
        if el.tag == 'table':
            rows = el.findall('.//tr')
            ncol = max((len(r.findall('.//td')) + len(r.findall('.//th')) for r in rows), default=0)
            nested = el.find('.//table') is not None
            seq.append(('table', f'{ncol}{"|nested" if nested else ""}'))
    # 上面的 iter 顺序含嵌套重复，改用线性遍历
    seq = []
    def walk(node):
        for ch in node:
            tag = LH.tostring(ch, method='text', encoding='unicode').strip()
            if ch.tag == 'table':
                rows = ch.findall('.//tr')
                ncol = max((len(r.findall('.//td')) + len(r.findall('.//th')) for r in rows), default=0)
                nested = ch.find('.//table') is not None
                seq.append(('table', f'{ncol}{"|nested" if nested else ""}'))
            elif ch.tag == 'p':
                has_img = ch.find('.//img') is not None
                txt = (ch.text_content() or '').strip()
                if has_img and not txt:
                    seq.append(('img-only', ''))
                elif has_img:
                    seq.append(('p+img', txt))
                else:
                    seq.append(('p', txt))
            elif ch.tag in ('h1','h2','h3','h4','h5','h6','blockquote','pre','ul','ol','figure','div'):
                if ch.tag == 'div' and ch.find('.//p') is not None:  # 容器不下沉
                    walk(ch)
                else:
                    has_img = ch.find('.//img') is not None
                    txt = (ch.text_content() or '').strip()
                    seq.append((ch.tag + ('+img' if has_img else ''), txt))
    walk(doc.body)
    return seq

def imgs_in(seq):
    return [i for i,(t,_) in enumerate(seq) if 'img' in t or t=='img-only']

def scan_sample(name):
    alarms = []  # (sample, indicator, section, detail, context)
    for xh in sorted(glob.glob(os.path.join(OUT, name, 'epub解包', 'OEBPS', 'sections', '*.xhtml'))):
        sec = os.path.basename(xh)
        seq = blocks(xh)
        ii = imgs_in(seq)
        ti = [i for i,(t,_) in enumerate(seq) if t=='table']
        # A 图注/表注分离（分开判定：图注贴 img，表注贴 table）
        for bi, (t, txt) in enumerate(seq):
            if t not in ('p','h2','h3','h4'): continue
            m = CAP_RE.match(txt)
            if not m: continue
            kind = m.group(1)
            if kind in ('图', 'Fig', 'Figure'):
                dist = min((abs(bi-j) for j in ii), default=99)
                obj = 'img'
            else:
                dist = min((abs(bi-j) for j in ti), default=99)
                obj = 'table'
            if dist >= 2:
                alarms.append((name, 'A图注分离', sec, f'{kind}注距最近{obj} {dist}块', txt[:30].replace(chr(10),'')))
        # B 连续图群（折叠：连续 img-only 游程记 1 条）
        j = 0
        while j < len(seq):
            if seq[j][0] == 'img-only':
                k = j
                while k < len(seq) and seq[k][0] == 'img-only': k += 1
                run = k - j
                prev_t = seq[j-1][0] if j > 0 else 'HEAD'
                if run >= 2:
                    alarms.append((name, 'B连续图群', sec, f'第{j}~{k-1}块连续{run}图(前邻{prev_t})', ''))
                j = k
            else:
                j += 1
        # C 超大表
        for bi, (t, txt) in enumerate(seq):
            if t == 'table':
                ncol = int(txt.split('|')[0])
                if ncol >= 8 or 'nested' in txt:
                    alarms.append((name, 'C超大表', sec, f'{ncol}列{"|嵌套" if "nested" in txt else ""}', ''))
        # D 跨章图（仅 img-only 落在章首/尾才算；p+img 属正常段落带图）
        if seq and ii:
            if seq[0][0] == 'img-only' or seq[-1][0] == 'img-only':
                pos = '章首' if seq[0][0] == 'img-only' else '章尾'
                alarms.append((name, 'D跨章图', sec, f'{pos}孤图 (len={len(seq)})', ''))
    return alarms

def main():
    all_alarms = []
    for sd in sorted(glob.glob(os.path.join(OUT, '0*'))):
        name = os.path.basename(sd)
        all_alarms += scan_sample(name)
    with open(os.path.join(W, 'v5_alarms.csv'), 'w', newline='', encoding='utf-8-sig') as f:
        w = csv.writer(f); w.writerow(['样本','指标','章节','详情','上下文'])
        w.writerows(all_alarms)
    # 统计
    from collections import Counter
    cnt = Counter(a[1] for a in all_alarms)
    n_samples = len(glob.glob(os.path.join(OUT, '0*')))
    print(f'样本数 {n_samples}，报警总数 {len(all_alarms)}')
    for k, v in sorted(cnt.items()): print(f'  {k}: {v}')
    # 上下文文件（人工复核）
    with open(os.path.join(W, 'v5_context.txt'), 'w', encoding='utf-8') as f:
        cur = None
        for name, ind, sec, det, ctx in all_alarms:
            if (name, sec) != cur:
                f.write(f'\n===== {name} / {sec} =====\n'); cur = (name, sec)
            f.write(f'[{ind}] {det}  {ctx}\n')
    print('明细: work/v5_v1_indicators/v5_alarms.csv + v5_context.txt')

if __name__ == '__main__':
    main()
