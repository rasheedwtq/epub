# -*- coding: utf-8 -*-
"""阶段3 图片优化（2026-10-09，大叔拍板：适合网络浏览即可，不必保打印级分辨率）

- 宽 > pixel_max → 等比降采样到 pixel_max（LANCZOS）
- 不透明 PNG/GIF 且体积达标 → 转 JPEG（截图类体积省 50~80%）
- 带透明通道的图保留原格式（JPEG 无 alpha，转了会糊底）
- 无需处理返回原路径，epub 里原样打包
"""
import os
from PIL import Image


def _opaque(im):
    """是否完全不透明（无有效透明像素）"""
    if im.mode == 'P':
        im = im.convert('RGBA')
    if im.mode in ('RGBA', 'LA'):
        lo, _ = im.getchannel('A').getextrema()
        return lo >= 255
    return im.mode in ('L', 'RGB', 'CMYK', 'YCbCr', 'I', 'F')


def _uniq(outdir, name):
    """2026-10-15 修 BUG-012：outdir 已存在同名文件时原样 save 会静默互覆——
    典型场景：文档同时含 image.png 与 image.jpg，前者转 JPEG 时把后者冲掉。
    级联避让直到名字空闲。"""
    cand = os.path.join(outdir, name)
    if not os.path.exists(cand):
        return cand
    stem, ext = os.path.splitext(name)
    k = 2
    while os.path.exists(os.path.join(outdir, f"{stem}_{k}{ext}")):
        k += 1
    return os.path.join(outdir, f"{stem}_{k}{ext}")


def optimize_image(path, outdir, pixel_max=1000, quality=82, transcode_min_kb=30):
    """返回优化后的文件路径（可能换扩展名 .jpg）；无需优化返回原路径。"""
    try:
        im = Image.open(path)
        im.load()
        fmt = (im.format or '').upper()
    except Exception:
        return path
    base = os.path.basename(path)
    stem, _ = os.path.splitext(base)
    os.makedirs(outdir, exist_ok=True)
    resized = False
    if im.size[0] > pixel_max:
        r = pixel_max / im.size[0]
        im = im.resize((pixel_max, max(1, round(im.size[1] * r))), Image.LANCZOS)
        resized = True
    kb = os.path.getsize(path) / 1024
    if fmt in ('PNG', 'GIF') and _opaque(im) and kb >= transcode_min_kb:
        new = _uniq(outdir, stem + '.jpg')
        try:
            im.convert('RGB').save(new, 'JPEG', quality=quality, optimize=True)
            return new
        except Exception:
            pass
    if resized:
        new = _uniq(outdir, base)
        try:
            if fmt == 'PNG':
                im.save(new, 'PNG', optimize=True)
            elif fmt == 'JPEG':
                im.save(new, 'JPEG', quality=quality, optimize=True)
            else:
                im.save(new)
            return new
        except Exception:
            pass
    return path


def optimize_images(img_srcmap, outdir, pixel_max=1000, quality=82):
    """批量优化 {绝对路径: 包内名} 映射。返回 (新映射, {旧包内名: 新包内名})"""
    new_map, renames = {}, {}
    for p, dst in img_srcmap.items():
        p2 = optimize_image(p, outdir, pixel_max=pixel_max, quality=quality)
        dst2 = os.path.basename(p2)
        new_map[p2] = dst2
        if dst2 != dst:
            renames[dst] = dst2
    return new_map, renames


def rewrite_secs(secs, renames):
    """切章产物里的图片引用改名（raw + URL 编码双写法都替换）"""
    if not renames:
        return secs
    from urllib.parse import quote
    out = []
    for t, inner in secs:
        for old, new in renames.items():
            inner = inner.replace(old, new).replace(quote(old), quote(new))
        out.append((t, inner))
    return out
