# -*- coding: utf-8 -*-
"""V-6 图片尺寸链路一致性
三段链路：源 docx wp:extent (EMU) → 目标像素(96dpi) → 产物 EPUB 实际像素
A1′ 产物侧 img 无写死宽高（.fig: max-width:100%; height:auto），
显示尺寸 = 容器宽 × 图片纵横比，故核心判定 = 纵横比保真（<5%）。
对齐方式：源 drawings（文档序）与产物 img（章节序）按【序 + 纵横比指纹】贪心匹配。
输出：work/v6_image_chain/v6_report.csv + 控制台汇总
"""
import os, re, csv, zipfile, glob
from lxml import etree
from PIL import Image

ROOT = r"D:\project\epub"
S = os.path.join(ROOT, "tests")
OUTDIR = os.path.join(ROOT, "review_samples_html")
W = os.path.join(ROOT, "work", "v6_image_chain")
os.makedirs(W, exist_ok=True)
NS = {
    'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
    'wp': 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing',
    'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
    'pic': 'http://schemas.openxmlformats.org/drawingml/2006/picture',
    'c': 'http://schemas.openxmlformats.org/drawingml/2006/chart',
    'wps': 'http://schemas.microsoft.com/office/word/2010/wordprocessingShape',
    'v': 'urn:schemas-microsoft-com:vml',
}

def doc_kind(d):
    """分类 drawing：picture/chart/textbox/other"""
    if d.find('.//pic:pic', NS) is not None: return 'picture'
    if d.find('.//c:chart', NS) is not None: return 'chart'
    if d.find('.//wps:txbx', NS) is not None or d.find('.//v:textbox', NS) is not None: return 'textbox'
    return 'other'

def source_drawings(docx):
    """返回 [(kind, cx_emu, cy_emu)] 按文档序；无 extent 的记 0"""
    out = []
    with zipfile.ZipFile(docx) as z:
        try: xml = z.read('word/document.xml')
        except KeyError:
            # .doc/.rtf 无 OOXML，由调用方跳过
            return None
        root = etree.fromstring(xml)
    for d in root.iter('{%s}inline' % NS['wp'], '{%s}anchor' % NS['wp']):
        ext = d.find('wp:extent', NS)
        cx = int(ext.get('cx', 0)) if ext is not None else 0
        cy = int(ext.get('cy', 0)) if ext is not None else 0
        out.append((doc_kind(d), cx, cy))
    return out

def product_imgs(sample_dir):
    """返回 [(fname, w_px, h_px, src)] 按章节/文档序（含同章内序）"""
    out = []
    for xh in sorted(glob.glob(os.path.join(sample_dir, 'epub解包', 'OEBPS', 'sections', '*.xhtml'))):
        txt = open(xh, encoding='utf-8').read()
        for m in re.finditer(r'<img src="\.\./images/([^"]+)"', txt):
            f = os.path.join(sample_dir, 'epub解包', 'OEBPS', 'images', m.group(1))
            if os.path.exists(f):
                try:
                    with Image.open(f) as im: w, h = im.size
                except Exception: w = h = 0
                out.append((m.group(1), w, h, xh))
    return out

def greedy_match(srcs, imgs, tol=0.05):
    """按文档序贪心匹配： kind 同类或 picture 类放宽 + 纵横比差 < tol"""
    pairs, used_i = [], set()
    for kind, cx, cy in srcs:
        if not cx or not cy: continue
        ra = cy / cx
        best = None
        for i, (fn, w, h, _) in enumerate(imgs):
            if i in used_i or not w or not h: continue
            rb = h / w
            if abs(rb - ra) / ra < tol:
                best = i
                break  # 顺序优先：取第一个纵横比吻合的
        if best is not None:
            used_i.add(best)
            pairs.append((kind, cx, cy, *imgs[best]))
    return pairs, len(imgs) - len(used_i)

def main():
    rows, notes = [], []
    total_src = total_img = matched = 0
    ar_devs, scale_rs = [], []
    for sd in sorted(glob.glob(os.path.join(S, '0*'))):
        name = os.path.basename(sd)
        rdir = os.path.join(OUTDIR, name)
        if not os.path.isdir(rdir): continue
        srcf = None
        for cand in ('docx', 'docm'):
            g = glob.glob(os.path.join(sd, '*.' + cand))
            if g: srcf = g[0]; break
        if srcf is None:  # 009 老 doc / rtf
            notes.append(f'{name}: 源为 .doc/.rtf，无 OOXML extent，跳过源段（{len(product_imgs(rdir))} 张产物图）')
            continue
        srcs = source_drawings(srcf)
        if srcs is None: continue
        imgs = product_imgs(rdir)
        pairs, unmatched = greedy_match(srcs, imgs)
        total_src += len(srcs); total_img += len(imgs); matched += len(pairs)
        for kind, cx, cy, fn, w, h, sec in pairs:
            exp_w = cx / 9525.0  # 96dpi 目标像素
            exp_h = cy / 9525.0
            ra_s, ra_p = cy / cx, h / w
            dev = abs(ra_p - ra_s) / ra_s * 100
            scale = (w / exp_w) if exp_w else 0
            ar_devs.append(dev); scale_rs.append(scale)
            rows.append([name, kind, round(exp_w, 1), round(exp_h, 1), w, h,
                         round(ra_s, 4), round(ra_p, 4), round(dev, 2), round(scale, 3), fn])
        if unmatched or len(srcs) != len(pairs):
            notes.append(f'{name}: 源 drawings {len(srcs)}，产物图 {len(imgs)}，匹配 {len(pairs)}，未对齐 {unmatched + len(srcs)-len(pairs)}')
    with open(os.path.join(W, 'v6_report.csv'), 'w', newline='', encoding='utf-8-sig') as f:
        wcsv = csv.writer(f)
        wcsv.writerow(['样本', '对象类型', '目标px宽', '目标px高', '产物px宽', '产物px高',
                       '源纵横比', '产物纵横比', '纵横比偏差%', '分辨率缩放率', '产物文件'])
        wcsv.writerows(rows)
    print(f'源 drawings 合计: {total_src}，产物图合计: {total_img}，顺序+纵横比对齐: {matched}')
    if ar_devs:
        big = [d for d in ar_devs if d > 5]
        print(f'纵横比偏差: 平均 {sum(ar_devs)/len(ar_devs):.2f}%，中位 {sorted(ar_devs)[len(ar_devs)//2]:.2f}%，>5% 条数 {len(big)}/{len(ar_devs)}')
        print(f'分辨率缩放率: min {min(scale_rs):.2f}x / max {max(scale_rs):.2f}x / 平均 {sum(scale_rs)/len(scale_rs):.2f}x')
    for n in notes: print('  -', n)
    bad = [r for r in rows if r[8] > 5]
    if bad:
        print(f'\n[偏差>5% 明细 {len(bad)} 条]')
        for r in bad[:15]:
            print(f'  {r[0]} {r[1]} 源比{r[6]} 产物比{r[7]} 偏差{r[8]}% {r[10]}')

if __name__ == '__main__':
    main()
