# -*- coding: utf-8 -*-
"""补跑批跑失败的 8 本（stderr 落盘可查；输出=源文件目录平铺）"""
import os, sys, io, json, time, subprocess
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", line_buffering=True)
PY = r"C:/Users/Administrator/AppData/Local/Programs/Python/Python311/python.exe"
ROOT = r"D:/project/epub"
SRC = r"d:/金融科技案例1207收"
BOOKS = [
    "1云工作室案例（20231121图片修改及内容微调版本）",
    "1人工智能平台案例（20231206需要加脚注版本）",
    "1案例6：成都事业群安心养老 (2)(3)",
    "1聚焦用户需求 提供精准服务——建融慧学校园综合服务平台20231205",
    "1裕农通平台-图片修改版本20231206",
    "1让技术重回“人本”属性——手机银行同质化破局的“减法之道”（1207）",
    "1（1207校对清稿）科技赋能，一撮即合——“建行全球撮合家”智能撮合平台案例",
]
import glob
# 区块链那本刚才已单独补跑成功，不重跑；先探测文件
for i, bid in enumerate(BOOKS, 1):
    exts = [e for e in (".docx", ".doc") if os.path.exists(os.path.join(SRC, bid + e))]
    f = os.path.join(SRC, bid + exts[0])
    t0 = time.time()
    r = subprocess.run([PY, "tools/pipeline.py", "convert", f, "-c", "config.yaml",
                        "-o", SRC, "--book-id", bid],
                       capture_output=True, cwd=ROOT, timeout=1200)
    rec = {}
    for ln in r.stdout.decode("utf-8", "ignore").strip().splitlines():
        try: rec = json.loads(ln)
        except Exception: pass
    st = rec.get("status", "?")
    g = rec.get("gate_l1l2") or {}
    mark = "✅" if st == "done" else "❌"
    print(f"[{i}/{len(BOOKS)}] {mark} {st.upper():6} L1={g.get('l1')} L2={g.get('l2')} "
          f"图{rec.get('imgs')} {round(time.time()-t0,1)}s | {bid[:44]}"
          + (f" | ERR: {(rec.get('err') or '')[:100]}" if rec.get("err") else ""))
    if rec.get("err") and not rec.get("status"):
        errtxt = r.stderr.decode("utf-8", "ignore")[-500:]
        open(os.path.join(ROOT, "work", "rerun_stderr.log"), "a", encoding="utf-8").write(
            f"==== {bid} ====\n{errtxt}\n")
print("补跑完成")
