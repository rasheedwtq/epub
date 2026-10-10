# -*- coding: utf-8 -*-
"""诊断：产物 xhtml 里的 div 结构（小看板文字版落点）"""
import sys, io, re
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")

paths = {
    "云工作室": r"out_pipeline/1云工作室案例（20231121图片修改及内容微调版本）/epub解包/OEBPS/sections/section0001.xhtml",
    "智慧政法": r"out_pipeline/1【智慧政法以此为准】金融科技践行社会主义核心价值观——智慧政法资金管理平台案例20231205/epub解包/OEBPS/sections/section0001.xhtml",
    "人工智能": r"out_pipeline/1人工智能平台案例（20231206需要加脚注版本）/epub解包/OEBPS/sections/section0001.xhtml",
}
for name, p in paths.items():
    h = open(p, encoding="utf-8").read()
    print(f"########## {name} ##########")
    for m in re.finditer(r"<div.{0,2500}?</div>", h, re.S):
        s = m.group(0)
        # 压缩文字
        txt = "".join(re.findall(r">([^<>]+)<", s))[:120]
        head = s[:s.find(">") + 1]
        print("  <div> 开标签:", head[:200])
        print("  内文字:", txt.replace("\n", " ")[:120])
        print("  ---")
