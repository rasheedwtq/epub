# -*- coding: utf-8 -*-
"""V-1: LibreOffice headless 批量稳定性与并发数标定.

三档并发(1/2/4) x 全量样本，每 worker 独立 UserInstallation profile。
记录: 每 doc 转换耗时 / 成功率 / 超时 / 轮次墙钟时间 / 进程残留。
产出: work/v1_stability/report.json + 控制台汇总。
"""
import os, sys, json, time, subprocess, threading, queue
from concurrent.futures import ThreadPoolExecutor

SO = r"C:/Program Files/LibreOffice/program/soffice.exe"
S = r"D:/project/epub/samples"
OUT = r"D:/project/epub/work/v1_stability"
PROFILE_ROOT = os.path.join(OUT, "profiles")
EXT_PRIORITY = [".docx", ".doc", ".docm", ".rtf"]
DOC_TIMEOUT = 120  # seconds per document


def find_main(sample_dir):
    cands = []
    for f in os.listdir(sample_dir):
        ext = os.path.splitext(f)[1].lower()
        if ext in EXT_PRIORITY:
            cands.append((EXT_PRIORITY.index(ext), f))
    if not cands:
        return None
    cands.sort()
    return os.path.join(sample_dir, cands[0][1])


def kill_soffice():
    subprocess.run(["taskkill", "/F", "/IM", "soffice.exe", "/T"],
                   capture_output=True)
    subprocess.run(["taskkill", "/F", "/IM", "soffice.bin", "/T"],
                   capture_output=True)
    time.sleep(2)


def convert_one(task):
    src, outdir, prof, tag = task
    os.makedirs(outdir, exist_ok=True)
    t0 = time.time()
    cmd = [SO, "--headless", "--norestore", "--nolockcheck", "--nodefault",
           f"-env:UserInstallation=file:///{prof.replace(chr(92), '/')}",
           "--convert-to", "epub", "--outdir", outdir, src]
    rc, timed_out, stderr = None, False, ""
    try:
        p = subprocess.run(cmd, capture_output=True, timeout=DOC_TIMEOUT)
        rc = p.returncode
        stderr = p.stderr.decode("utf-8", "replace")[:200]
    except subprocess.TimeoutExpired:
        timed_out = True
        subprocess.run(["taskkill", "/F", "/IM", "soffice.bin", "/T"],
                       capture_output=True)
        time.sleep(1)
    dt = time.time() - t0
    # 判定成功：outdir 出现同名 .epub
    base = os.path.splitext(os.path.basename(src))[0]
    epub = os.path.join(outdir, base + ".epub")
    ok = (not timed_out) and os.path.exists(epub) and os.path.getsize(epub) > 1000
    return dict(src=os.path.basename(src), tag=tag, rc=rc, timed_out=timed_out,
                seconds=round(dt, 1), ok=ok, err=stderr if not ok else "")


def run_round(n_workers, samples, round_tag):
    prof_dir = os.path.join(PROFILE_ROOT, f"round{n_workers}")
    outdir = os.path.join(OUT, f"out_w{n_workers}")
    tasks = []
    for i, src in enumerate(samples):
        prof = os.path.join(prof_dir, f"p{i % n_workers}")
        tasks.append((src, outdir, prof, f"w{i % n_workers}"))
    q = queue.Queue()
    results = []
    lock = threading.Lock()
    t0 = time.time()
    with ThreadPoolExecutor(max_workers=n_workers) as ex:
        futures = [ex.submit(convert_one, t) for t in tasks]
        for f in futures:
            r = f.result()
            with lock:
                results.append(r)
                status = "OK " if r["ok"] else ("TMO" if r["timed_out"] else "FAIL")
                print(f"  [{status}] {r['seconds']:>6.1f}s  {r['src']}", flush=True)
    wall = time.time() - t0
    return dict(workers=n_workers, wall=round(wall, 1),
                ok=sum(1 for r in results if r["ok"]),
                timeout=sum(1 for r in results if r["timed_out"]),
                fail=sum(1 for r in results if not r["ok"] and not r["timed_out"]),
                results=results)


def main():
    only = sys.argv[1:] if len(sys.argv) > 1 else None
    os.makedirs(OUT, exist_ok=True)
    samples = []
    for d in sorted(os.listdir(S)):
        p = os.path.join(S, d)
        if not os.path.isdir(p):
            continue
        m = find_main(p)
        if m:
            if only and not any(o in d for o in only):
                continue
            samples.append(m)
    print(f"样本数: {len(samples)}", flush=True)

    report = []
    for n in (1, 2, 4):
        print(f"===== 并发 {n} =====", flush=True)
        kill_soffice()
        r = run_round(n, samples, n)
        r["leftover_soffice"] = leftover_soffice()
        report.append(r)
        print(f"  并发{n}: 墙钟 {r['wall']}s | 成功 {r['ok']} | 超时 {r['timeout']} "
              f"| 失败 {r['fail']} | 残留进程 {r['leftover_soffice']}", flush=True)
    kill_soffice()
    with open(os.path.join(OUT, "report.json"), "w", encoding="utf-8") as f:
        json.dump(report, f, ensure_ascii=False, indent=1)
    print("DONE -> work/v1_stability/report.json", flush=True)


def leftover_soffice():
    p = subprocess.run(["tasklist", "/FI", "IMAGENAME eq soffice.bin"],
                       capture_output=True)
    out = (p.stdout or b"").decode("gbk", "replace")
    return out.count("soffice.bin")


if __name__ == "__main__":
    main()
