# -*- coding: utf-8 -*-
"""金融科技案例批 全量重跑（2026-10-09 修复版管线：组合图拆解 + profile 修复 + 排版清理）"""
import glob, os, sys, io, json, time, subprocess
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", line_buffering=True)

PY = r"C:/Users/Administrator/AppData/Local/Programs/Python/Python311/python.exe"
ROOT = r"D:/project/epub"
SRC_DIR = r"d:/金融科技案例1207收"
EXTS = (".docx", ".doc", ".docm", ".rtf")

files = [f for f in sorted(glob.glob(os.path.join(SRC_DIR, "*"))) if f.lower().endswith(EXTS)]
print(f"共 {len(files)} 本，开始批跑")
ok = fail = 0
for i, f in enumerate(files, 1):
    bid = os.path.splitext(os.path.basename(f))[0]
    t0 = time.time()
    r = subprocess.run([PY, "tools/pipeline.py", "convert", f, "-c", "config.yaml",
                        "-o", "out_pipeline", "--book-id", bid],
                       capture_output=True, cwd=ROOT, timeout=1200)
    rec = {}
    for ln in r.stdout.decode("utf-8", "ignore").strip().splitlines():
        try: rec = json.loads(ln)
        except Exception: pass
    st = rec.get("status", "?")
    g = rec.get("gate_l1l2") or {}
    mark = "✅" if st == "done" else "❌"
    if st == "done": ok += 1
    else: fail += 1
    print(f"[{i}/{len(files)}] {mark} {st.upper():6} L1={g.get('l1')} L2={g.get('l2')} "
          f"章{rec.get('sections')} 图{rec.get('imgs')} {round(time.time()-t0,1)}s | {bid[:44]}"
          + (f" | ERR: {rec.get('err','')[:90]}" if rec.get("err") else ""))
print(f"\n批跑完成：done={ok} failed={fail}")
