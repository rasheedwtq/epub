# -*- coding: utf-8 -*-
"""批量任务汇总：读 out_pipeline/*/任务.json 输出状态表"""
import json, glob, os, sys

rows = []
for f in sorted(glob.glob("out_pipeline/*/任务.json")):
    r = json.load(open(f, encoding="utf-8"))
    g = r.get("gate_l1l2") or {}
    ec = r.get("gate_epubcheck")
    rows.append((r["book_id"], r["status"], g.get("l1"), g.get("l2"),
                 r.get("sections"), r.get("imgs"), r.get("elapsed_s"),
                 (r.get("err") or "")[:40], ec))

print(f"{'book_id':<30}{'状态':<14}{'L1':>7}  {'L2':<28}{'章':>3} {'图':>3} {'耗时':>6}  EPUBCheck")
ok = fail = skip = 0
for b, st, l1, l2, sec, img, el, err, ec in rows:
    mark = {'done': '✅', 'failed': '❌'}.get(st, '⏳')
    if st == 'done': ok += 1
    elif st == 'failed': fail += 1
    else: skip += 1
    ecs = f"{ec['fatal']}F/{ec['err']}E/{ec['warn']}W" if isinstance(ec, dict) and 'fatal' in ec else ('—' if not ec else str(ec)[:20])
    l2s = str(l2)[:28]
    print(f"{mark}{b:<28}{st:<14}{l1 if l1 is not None else '—':>7}  {l2s:<28}{sec or '—':>3} {img or '—':>3} {el if el else '—':>6}s  {ecs}{('  ERR:'+err) if err else ''}")
print(f"\n合计 {len(rows)} 本：✅{ok} ❌{fail} ⏳{skip}")
