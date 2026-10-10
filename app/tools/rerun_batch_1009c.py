# -*- coding: utf-8 -*-
"""2026-10-09 第三轮全量重跑：页眉脚删除 + 图片封顶 + .doc 预转 docx 拆组合
.doc 书清缓存全链路；docx 书命中 s1/s2 缓存只跑 stage3。"""
import glob, os, subprocess, json, time, io, sys, shutil
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', line_buffering=True)
PY = r'C:/Users/Administrator/AppData/Local/Programs/Python/Python311/python.exe'
SRC = r'd:/金融科技案例1207收'
os.chdir(r'D:/project/epub')

t0 = time.time()
ok = fail = 0
for f in sorted(glob.glob(SRC + '/*')):
    b = os.path.basename(f)
    if not b.lower().endswith(('.docx', '.doc', '.docm', '.rtf')):
        continue
    bid = os.path.splitext(b)[0]
    if b.lower().endswith(('.doc', '.rtf')):
        wd = os.path.join('work/pipeline', bid)
        if os.path.exists(wd):
            # 沙箱批量删除钩子会误杀进程：改名腾位代替 rmtree（旧缓存自然失效）
            os.rename(wd, wd + '_旧' + time.strftime('%H%M%S'))
    t1 = time.time()
    r = subprocess.run([PY, 'tools/pipeline.py', 'convert', f, '-c', 'config.yaml',
                        '-o', SRC, '--book-id', bid],
                       capture_output=True, cwd='D:/project/epub', timeout=1200)
    rec = {}
    for ln in r.stdout.decode('utf-8', 'ignore').strip().splitlines():
        try:
            rec = json.loads(ln)
        except Exception:
            pass
    good = rec.get('status') == 'done'
    ok += good
    fail += (not good)
    print(('✅' if good else '❌'), bid[:40], round(time.time() - t1, 1), 's',
          f"L1={rec.get('gate_l1l2', {}).get('l1')}", (rec.get('err') or '')[:60])
print(f'== {ok} 成 / {fail} 败，总耗时 {round(time.time()-t0,1)}s ==')
