# -*- coding: utf-8 -*-
"""诊断：docx 图文框/形状结构 + L1/L2 口径对账"""
import sys, io, zipfile, re
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")
sys.path.insert(0, "tools")
import l1l2_check as LC

books = {
    "云工作室": r"d:/金融科技案例1207收/1云工作室案例（20231121图片修改及内容微调版本）.docx",
    "人工智能": r"d:/金融科技案例1207收/1人工智能平台案例（20231206需要加脚注版本）.docx",
    "智慧政法": r"d:/金融科技案例1207收/1【智慧政法以此为准】金融科技践行社会主义核心价值观——智慧政法资金管理平台案例20231205.docx",
}
for name, path in books.items():
    z = zipfile.ZipFile(path)
    doc = z.read("word/document.xml").decode("utf-8", "ignore")
    print(f"########## {name} ##########")
    print("  framePr(图文框):", doc.count("<w:framePr"))
    print("  txbxContent(文本框):", doc.count("<w:txbxContent"))
    print("  wps:wsp(形状):", doc.count("<wps:wsp"))
    print("  v:shape:", doc.count("<v:shape "), " v:imagedata:", doc.count("<v:imagedata"))
    print("  w:drawing:", doc.count("<w:drawing>"), " pic:pic:", doc.count("<pic:pic"), " chart:", doc.count("c:chart"), " math:", doc.count("<m:oMath"))
    print("  w:object:", doc.count("<w:object"), " OLEObject:", doc.count("OLEObject"))
    sp = LC.source_parts(path)
    print("  gate口径 pic/chart/math/tbl =", sp["pic"], sp["chart"], sp["math"], sp["tbl"])
    # 文本框内文字抽样（判断小看板是文本还是图）
    m = re.search(r"<w:txbxContent.{0,4000}?</w:txbxContent>", doc, re.S)
    if m:
        txts = re.findall(r"<w:t[^>]*>([^<]{2,30})</w:t>", m.group(0))
        print("  首个文本框文字样例:", txts[:8])
