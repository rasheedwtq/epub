# -*- coding: utf-8 -*-
"""诊断：产物 XHTML 内联样式分布 + 指定段落上下文"""
import re, sys, io
from collections import Counter
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")

paths = {
    "云工作室": r"out_pipeline/1云工作室案例（20231121图片修改及内容微调版本）/epub解包/OEBPS/sections/section0001.xhtml",
    "智慧政法": r"out_pipeline/1【智慧政法以此为准】金融科技践行社会主义核心价值观——智慧政法资金管理平台案例20231205/epub解包/OEBPS/sections/section0001.xhtml",
    "人工智能": r"out_pipeline/1人工智能平台案例（20231206需要加脚注版本）/epub解包/OEBPS/sections/section0001.xhtml",
}
for name, p in paths.items():
    h = open(p, encoding="utf-8").read()
    print(f"########## {name} ##########")
    c = Counter()
    for m in re.finditer(r"line-height:\s*([^;\"]+)", h):
        c["LH=" + m.group(1).strip()] += 1
    for m in re.finditer(r"text-indent:\s*([^;\"]+)", h):
        c["TI=" + m.group(1).strip()] += 1
    c["position:absolute"] = h.count("position: absolute") + h.count("position:absolute")
    c["page-break-before"] = h.count("page-break-before")
    c["margin-left"] = h.count("margin-left")
    c["<div"] = h.count("<div")
    c["nbsp段"] = h.count("\u00a0\u00a0")
    for k, v in c.most_common(30):
        print(f"  {k!r}: {v}")
