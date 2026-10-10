# -*- coding: utf-8 -*-
"""补跑剩余书（前台专用，沙箱对后台批跑的子进程会误杀）"""
import os, subprocess, json, time, io, sys
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', line_buffering=True)
PY = r'C:/Users/Administrator/AppData/Local/Programs/Python/Python311/python.exe'
SRC = r'd:/金融科技案例1207收'
os.chdir(r'D:/project/epub')

# 上轮已成功的 5 本 + 2 坏文件跳过
SKIP = {'1万物在线带来的效率、成本、安全革命——物联网平台案例 张晏玮',
        '1218以金融科技助力土地制度改革数字化——土地三权分置相关平台搭建案例 魏溥亭',
        '1_新版金融科技助力土地制度改革数字化——魏溥亭',
        '1“双管齐下”让平台“精准定位”——建融智医综合服务平台案例20231205',
        '1【智慧政法以此为准】金融科技践行社会主义核心价值观——智慧政法资金管理平台案例20231205',
        '1金融科技部-《金融科技实践案例集》（征求意见稿）三资监管平台20231205',
        '1金融科技部-《金融科技实践案例集》（征求意见稿）智慧村务平台20231205'}

t0 = time.time()
ok = fail = 0
for b in sorted(os.listdir(SRC)):
    if not b.lower().endswith(('.docx', '.doc', '.docm', '.rtf')):
        continue
    bid = os.path.splitext(b)[0]
    if bid in SKIP:
        continue
    t1 = time.time()
    r = subprocess.run([PY, 'tools/pipeline.py', 'convert', os.path.join(SRC, b), '-c', 'config.yaml',
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
          f"L1={rec.get('gate_l1l2', {}).get('l1')}", (rec.get('err') or '')[:70])
print(f'== {ok} 成 / {fail} 败，总耗时 {round(time.time()-t0,1)}s ==')
