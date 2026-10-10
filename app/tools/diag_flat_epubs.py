# -*- coding: utf-8 -*-
"""验证平铺产物 epub 的排版质量"""
import sys, io, zipfile, re
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")
names = [
    "1人工智能平台案例（20231206需要加脚注版本）",
    "1裕农通平台-图片修改版本20231206",
    "1【智慧政法以此为准】金融科技践行社会主义核心价值观——智慧政法资金管理平台案例20231205",
    "1218以金融科技助力土地制度改革数字化——土地三权分置相关平台搭建案例 魏溥亭",
]
for name in names:
    z = zipfile.ZipFile(rf"d:/金融科技案例1207收/转换_{name}.epub")
    h = z.read("OEBPS/sections/section0001.xhtml").decode("utf-8")
    lh = len(re.findall(r"line-height", h))
    ti = re.findall(r"text-indent:\s*([^;\"]+)", h)
    imgw = len(re.findall(r"width:\d+px", h))
    imgs = h.count("<img")
    print(f"{name[:24]:26} | line-height残留:{lh:3} | text-indent:{len(ti):3} 值:{set(ti)} | 内联宽图:{imgw}/{imgs}")
