# -*- coding: utf-8 -*-
"""V-7 字体子集化缺字验证
A1′ 不嵌字体、不做子集化 → 缺字风险=①转换链路丢字 ②阅读器字体覆盖。
本脚本验证 ①：造生僻字/异体字/emoji/全角符号文档 → LO→HTML → 逐字符比对。
②给出策略结论（见结论文档）。
输出：work/v7_charset/ 下测试 docx + html + 比对报告
"""
import os, re, subprocess, shutil
from lxml import html as LH

ROOT = r"D:\project\epub"
W = os.path.join(ROOT, "work", "v7_charset")
os.makedirs(W, exist_ok=True)
LO = r"C:\Program Files\LibreOffice\program\soffice.exe"

GROUPS = {
    "常用对照": "汉字常用字测试：永国机器人民主和谐",
    "生僻字GB*": "龘齉爩鱻麤龗䲜饢讞钀",
    "CJK扩展B": "𠀀𠀁𠮷𪚥𪜀𫝀",
    "异体/日文汉字": "峠塩畑辻込戶戸",
    "单位/兼容字": "兙兛兞㐀㲢旡",
    "带圈兼容": "①⑴㊀㈠⒈㎡㎏℃",
    "emoji": "😀🀄🈚✅⚡",
    "全角标点": "「」『』～……——·〔〕【】",
}

def make_docx(path):
    import docx
    d = docx.Document()
    d.add_heading("V-7 字符保真测试", level=1)
    for g, chars in GROUPS.items():
        d.add_heading(g, level=2)
        d.add_paragraph("原始：" + chars)
        d.add_paragraph("字符数：" + str(len(chars)))
    d.save(path)

def html_text(p):
    doc = LH.fromstring(open(p, 'rb').read())
    return doc.text_content()

def main():
    dx = os.path.join(W, "v7_charset_test.docx")
    make_docx(dx)
    # LO 转换（独立 profile）
    prof = "file:///" + os.path.join(W, "_prof").replace("\\", "/")
    r = subprocess.run([LO, "--headless", f"-env:UserInstallation={prof}",
                        "--convert-to", "html:HTML (StarWriter)", "--outdir", W, dx],
                       capture_output=True, timeout=120)
    htm = os.path.join(W, "v7_charset_test.html")
    if not os.path.exists(htm):
        print("LO 转换失败:", r.stderr[:300]); return
    txt = html_text(htm)
    total = miss = repl = 0
    lines = []
    for g, chars in GROUPS.items():
        got = []
        for ch in chars:
            total += 1
            if ch in txt:
                got.append(ch)
            else:
                miss += 1
                got.append(f"❌[{ch} U+{ord(ch):04X}]")
        status = "OK" if len(got) == len(chars) else "有丢失"
        lines.append(f"[{status}] {g}: 原始 {len(chars)} 字 → 存活 {len(got)-sum(1 for x in got if x.startswith('❌'))}")
        if status != "OK":
            lines.append("        丢失: " + " ".join(x for x in got if x.startswith("❌")))
    print(f"总字符 {total}，丢失 {miss}，保真率 {(total-miss)/total*100:.1f}%")
    print("\n".join(lines))
    # 顺带看 LO 给的字体
    m = re.search(r'font-family[^;}]*', open(htm, encoding='utf-8', errors='ignore').read())
    if m: print("\nLO 字体声明示例:", m.group(0)[:80])
    open(os.path.join(W, "报告.txt"), "w", encoding="utf-8").write(
        f"总字符 {total} 丢失 {miss}\n" + "\n".join(lines))

if __name__ == "__main__":
    main()
