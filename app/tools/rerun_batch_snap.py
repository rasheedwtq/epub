# -*- coding: utf-8 -*-
"""全量重跑（shapesnap 版）。用法: python rerun_batch_snap.py <start> <end>  （按 glob 序切片，含头不含尾）"""
import glob, os, subprocess, sys, json, time, io

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', line_buffering=True)
SRC = r'd:/金融科技案例1207收'
PY = r'C:/Users/Administrator/AppData/Local/Programs/Python/Python311/python.exe'
a, b = int(sys.argv[1]), int(sys.argv[2])
books = sorted(glob.glob(SRC + '/*.docx')) + sorted(glob.glob(SRC + '/*.doc'))
books = [f for f in books if not os.path.basename(f).startswith('~$')]
books = sorted(books)[a:b]
ok = n = 0
for f in books:
    bid = os.path.splitext(os.path.basename(f))[0]
    if os.path.exists(SRC + f'/转换_{bid}.epub') and '云工作室' in bid:
        continue
    t0 = time.time()
    r = subprocess.run([PY, 'tools/pipeline.py', 'convert', f, '-c', 'config.yaml',
                        '-o', SRC, '--book-id', bid],
                       capture_output=True, cwd='D:/project/epub', timeout=1800)
    rec = {}
    for ln in r.stdout.decode('utf-8', 'ignore').strip().splitlines():
        try: rec = json.loads(ln)
        except Exception: pass
    good = rec.get('status') == 'done'
    ok += good; n += 1
    print(('✅' if good else '❌'), bid[:44], round(time.time()-t0, 1), 's',
          (rec.get('err') or rec.get('gate_l1l2', {}).get('l2', ''))[:70], flush=True)
print(f'== 本批 {ok}/{n} ==', flush=True)
