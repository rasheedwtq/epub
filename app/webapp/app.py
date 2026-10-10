# -*- coding: utf-8 -*-
"""docx2epub Web 控制台后端（FastAPI）
模块3 配置中心 + 模块1 仪表盘 + 模块2 转换控制台（后台 worker）+ 模块4 质检面板 + 模块6 试读
原则（设计文档铁律）：UI 只发参数；转换绝不跑在 web 请求线程里（独立 worker 子进程）。

启动: python webapp/app.py  →  http://127.0.0.1:8300
"""
import os, sys, json, shutil, subprocess, time, threading
from pathlib import Path
from fastapi import FastAPI, Body
from fastapi.responses import FileResponse, JSONResponse
from fastapi.staticfiles import StaticFiles
import uvicorn, yaml

ROOT = Path(__file__).resolve().parents[2]        # D:\project\epub（app/webapp/app.py → 上两级）
APP = ROOT / "app"                                # 程序目录（config/tools/webapp）
CFG = APP / "config.yaml"
VERS = ROOT / "work" / "config_versions"
SAMPLES = ROOT / "tests"                          # 测试用例库（2026-10-10 整理：samples 并入 tests）
MANIFEST = ROOT / "work" / "pipeline" / "manifest.jsonl"
OUT = ROOT / "out_pipeline"
OUT.mkdir(exist_ok=True)                          # /files 静态服务要求目录存在（输出平铺后此目录仅作兜底）
PY = sys.executable

app = FastAPI(title="docx2epub 配置中心")
app.mount("/files", StaticFiles(directory=str(OUT)), name="files")   # 产物 EPUB 静态服务（供 epub.js 拉取）
app.mount("/static", StaticFiles(directory=str(APP / "webapp" / "static")), name="static")  # vendor/epub.min.js 等

def load_cfg():
    return yaml.safe_load(CFG.read_text(encoding="utf-8"))

def save_cfg(data: dict):
    VERS.mkdir(parents=True, exist_ok=True)
    snap = VERS / f"config_{time.strftime('%Y%m%d_%H%M%S')}.yaml"
    shutil.copy2(CFG, snap)                      # 存旧版
    CFG.write_text(yaml.safe_dump(data, allow_unicode=True, sort_keys=False), encoding="utf-8")
    return snap.name

@app.get("/")
def index():
    return FileResponse(APP / "webapp" / "static" / "index.html")

@app.get("/reader")
def reader():
    return FileResponse(APP / "webapp" / "static" / "reader.html")

@app.get("/api/epubs")
def epubs():
    """列出所有可试读的 EPUB（manifest 驱动：支持任意输出目录，平铺/旧版目录布局均兼容）"""
    from urllib.parse import quote
    out = []
    for bid, r in manifest_latest().items():
        ep = r.get("epub")
        if r.get("status") == "done" and ep and os.path.exists(ep):
            out.append(dict(book_id=bid, url=f"/api/epub-file?path={quote(ep)}", path=ep))
    out.sort(key=lambda x: x["book_id"])
    return out

@app.get("/api/epub-file")
def epub_file(path: str):
    """按绝对路径回源 EPUB（平铺输出目录不再固定，静态挂载覆盖不到）"""
    p = os.path.abspath(path)
    if not p.lower().endswith(".epub") or not os.path.exists(p):
        return JSONResponse(dict(ok=False, err="EPUB 不存在"), status_code=404)
    return FileResponse(p, media_type="application/epub+zip",
                        filename=os.path.basename(p))

@app.get("/api/config")
def get_config():
    return load_cfg()

@app.put("/api/config")
def put_config(data: dict = Body(...)):
    snap = save_cfg(data)
    return dict(ok=True, snapshot=snap)

@app.get("/api/versions")
def versions():
    if not VERS.exists(): return []
    return sorted([f.name for f in VERS.glob("config_*.yaml")], reverse=True)

@app.post("/api/rollback/{name}")
def rollback(name: str):
    p = VERS / name
    if not p.exists(): return JSONResponse(dict(ok=False, err="版本不存在"), status_code=404)
    cur = VERS / f"config_{time.strftime('%Y%m%d_%H%M%S')}_rollback_src.yaml"
    shutil.copy2(CFG, cur)                       # 回滚前也存快照（可逆）
    shutil.copy2(p, CFG)
    return dict(ok=True, restored=name, backup=cur.name)

@app.get("/api/samples")
def samples():
    out = []
    if SAMPLES.exists():
        for d in sorted(SAMPLES.iterdir()):
            if d.is_dir() and not d.name.startswith("_"):
                mains = [f for f in os.listdir(d) if f.lower().endswith((".docx", ".doc", ".rtf", ".docm"))]
                if mains: out.append(dict(book_id=d.name, src=str(d / mains[0])))
    return out

SCAN_EXTS = (".docx", ".doc", ".rtf", ".docm")
SCAN_LIMIT = 3000

@app.post("/api/scan")
def scan(payload: dict = Body(...)):
    """扫描目录（递归）下所有 Word 文件，供转换控制台目录模式选择"""
    d = (payload.get("dir") or "").strip().strip('"')
    if not d or not os.path.isdir(d):
        return JSONResponse(dict(ok=False, err=f"目录不存在: {d}"), status_code=400)
    files, total = [], 0
    for root, _dirs, names in os.walk(d):
        for n in names:
            if n.lower().endswith(SCAN_EXTS):
                total += 1
                if len(files) >= SCAN_LIMIT: continue
                p = os.path.join(root, n)
                try: size_kb = round(os.path.getsize(p) / 1024, 1)
                except OSError: size_kb = -1
                files.append(dict(src=p, book_id=os.path.splitext(n)[0],
                                  rel=os.path.relpath(p, d), size_kb=size_kb))
    if len(files) > SCAN_LIMIT:
        files = files[:SCAN_LIMIT]
    dup = {}
    for f in files: dup.setdefault(f["book_id"], []).append(f["rel"])
    dups = {k: v for k, v in dup.items() if len(v) > 1}
    return dict(ok=True, dir=d, total=total, files=files,
                truncated=total > SCAN_LIMIT, duplicate_names=dups)

@app.get("/api/root")
def root_info():
    """前端默认值用：项目根目录"""
    return dict(root=str(ROOT), default_out=str(OUT))

@app.get("/api/tasks")
def tasks(limit: int = 30):
    if not MANIFEST.exists(): return []
    lines = MANIFEST.read_text(encoding="utf-8").strip().splitlines()
    return [json.loads(l) for l in lines[-limit:]]

@app.post("/api/run")
def run(payload: dict = Body(...)):
    book_id = payload["book_id"]
    src = payload.get("src")
    stage = payload.get("stage", "3")
    out_dir = (payload.get("out_dir") or "").strip() or str(OUT)
    if not src:
        hits = [s for s in samples() if s["book_id"] == book_id]
        if not hits: return JSONResponse(dict(ok=False, err="样本不存在"), status_code=404)
        src = hits[0]["src"]
    r = subprocess.run([PY, str(APP / "tools" / "pipeline.py"), "convert", src,
                        "-c", str(CFG), "-o", out_dir, "--book-id", book_id, "--stage", stage],
                       cwd=str(ROOT), capture_output=True, timeout=600)
    try:
        rec = json.loads(r.stdout.strip().splitlines()[-1])
    except Exception:
        rec = dict(ok=False, err=(r.stderr or b"").decode("utf-8", "ignore")[-500:])
    return rec

# ---------------------------------------------------------------- 模块1 仪表盘 + 模块4 质检：manifest / 评分卡数据

def manifest_latest() -> dict:
    """book_id -> manifest 最新一条记录"""
    if not MANIFEST.exists(): return {}
    out = {}
    for l in MANIFEST.read_text(encoding="utf-8").strip().splitlines():
        try:
            r = json.loads(l)
            out[r.get("book_id")] = r
        except Exception:
            pass
    return out

@app.get("/api/stats")
def stats():
    """模块1 仪表盘：manifest 全量聚合（book_id 去重取最新）"""
    m = manifest_latest()
    by_status, fmt, fails = {}, {}, []
    total_t = 0.0; n_t = 0
    for bid, r in m.items():
        st = r.get("status", "?")
        by_status[st] = by_status.get(st, 0) + 1
        src = r.get("src", "")
        if st == "failed":
            kind = "坏文件/加密" if "_f0" in bid or "坏文件" in src else "闸门未过"
        else:
            kind = Path(src).suffix.lower().lstrip(".") or "未知"
        fmt[kind] = fmt.get(kind, 0) + 1
        if r.get("elapsed_s"):
            total_t += r["elapsed_s"]; n_t += 1
        if st == "failed":
            fails.append(dict(book_id=bid, err=(r.get("err") or "")[:160]))
    recent = sorted(m.values(), key=lambda r: r.get("ts", ""))[-12:][::-1]
    return dict(total=len(m), by_status=by_status, input_fmt=fmt,
                avg_elapsed=round(total_t / n_t, 1) if n_t else 0,
                fails=fails, recent=recent)

@app.get("/api/scores")
def scores():
    """模块4 质检面板：每本一行评分卡（manifest 最新状态 + work 任务.json，兼容旧版 out_pipeline 布局）"""
    m = manifest_latest()
    rows = []
    for bid, r in sorted(m.items()):
        tj = ROOT / "work" / "pipeline" / bid / "任务.json"
        if not tj.exists():
            legacy = OUT / bid / "任务.json"           # 旧版每本一夹布局
            if legacy.exists():
                tj = legacy
        try:
            t = json.loads(tj.read_text(encoding="utf-8"))
        except Exception:
            t = r                                       # 没有 任务.json 就用 manifest 记录兜底
        g = t.get("gate_l1l2") or {}
        ep = t.get("epub") or r.get("epub")
        rows.append(dict(book_id=bid, status=r.get("status") or t.get("status") or "?",
                         l1=g.get("l1"), l2=g.get("l2"),
                         sections=t.get("sections"), imgs=t.get("imgs"),
                         elapsed=t.get("elapsed_s"),
                         out_dir=r.get("out_dir") or (str(Path(ep).parent) if ep else ""),
                         err=(r.get("err") or t.get("err") or "")[:160],
                         epub=bool(ep and os.path.exists(ep))))
    return rows

# ---------------------------------------------------------------- 模块2 转换控制台：后台 worker（独立子进程，绝不占用 web 线程）

BATCH = dict(running=False, cancel=False, proc=None, thread=None,
             items=[], logs=[], started=None, finished=None)

def _kill_tree(pid):
    """Windows 杀整棵进程树（python pipeline + 其拉起的 soffice）"""
    subprocess.run(["taskkill", "/PID", str(pid), "/T", "/F"], capture_output=True)

def _batch_worker(items: list, out_dir: str = ""):
    out_root = out_dir or str(OUT)
    skip_done = BATCH.get("skip_done", True)
    done_map = manifest_latest() if skip_done else {}
    for item in items:
        src = item["src"]
        if BATCH["cancel"]: break
        bid = item["book_id"]
        if skip_done and done_map.get(bid, {}).get("status") == "done":
            item.update(status="skipped", summary="断点续跑：已完成，跳过")
            continue
        item.update(status="running")
        rec = {}
        for attempt in (1, 2):   # V-1 结论：失败自动重试 1 次（消化瞬时环境故障）
            proc = subprocess.Popen([PY, str(APP / "tools" / "pipeline.py"), "convert", src,
                                     "-c", str(CFG), "-o", out_root, "--book-id", bid],
                                    cwd=str(ROOT), stdout=subprocess.PIPE, stderr=subprocess.STDOUT)
            BATCH["proc"] = proc
            rec = {}
            try:
                out, _ = proc.communicate(timeout=900)
                for ln in out.decode("utf-8", "ignore").strip().splitlines():
                    try: rec = json.loads(ln)
                    except Exception: pass
            except subprocess.TimeoutExpired:
                _kill_tree(proc.pid)
                rec = dict(status="failed", err="TimeoutExpired: 单本超时 900s 强杀")
            if rec.get("status") == "done" or BATCH["cancel"]:
                break
            if attempt == 1:
                BATCH["logs"].append(f"[重] {bid} — 首次失败（{rec.get('err') or '退出码 ' + str(proc.returncode)}），自动重试")
        if rec.get("status") == "done":
            g = rec.get("gate_l1l2") or {}
            s = f"L1={g.get('l1')}% L2={g.get('l2')} 章{rec.get('sections')} 图{rec.get('imgs')}"
            item.update(status="done", elapsed=rec.get("elapsed_s"), summary=s)
            BATCH["logs"].append(f"[OK] {time.strftime('%H:%M:%S')} {bid} — {s}（{rec.get('elapsed_s')}s）")
        elif BATCH["cancel"]:
            item.update(status="canceled", summary="用户取消")
            BATCH["logs"].append(f"[停] {bid} — 用户取消，剩余任务终止")
            break
        else:
            g = rec.get("gate_l1l2") or {}
            err = rec.get("err") or (f"闸门未过: L1={g.get('l1')}% L2={g.get('l2')}" if g
                                     else f"退出码 {proc.returncode}")
            tail = (out.decode("utf-8", "ignore") if 'out' in dir() else "")[-400:].strip()
            item.update(status="failed", summary=err[:120])
            BATCH["logs"].append(f"[ERR] {time.strftime('%H:%M:%S')} {bid} — {err[:160]}")
            if tail and not rec.get("status"):
                BATCH["logs"].append(f"[dbg] {bid} 输出尾部: {tail[-280:]}")
    BATCH.update(running=False, proc=None, finished=time.strftime("%H:%M:%S"))

def _normalize_srcs(srcs):
    """srcs 支持 str 或 {src, book_id} 两种形态 → 统一 items 结构（同名文件可显式 book_id 防覆盖）"""
    items = []
    used = {}
    for s in srcs or []:
        if isinstance(s, dict):
            src, bid = s.get("src"), s.get("book_id") or Path(s["src"]).stem
        else:
            src, bid = s, Path(s).stem
        # book_id 防撞：同名时追加序号（否则产物互相覆盖）
        if bid in used:
            used[bid] += 1
            bid = f"{bid}~{used[bid]}"
        else:
            used[bid] = 1
        items.append(dict(book_id=bid, src=src, status="pending", summary="", elapsed=None))
    return items

@app.post("/api/batch/start")
def batch_start(payload: dict = Body(...)):
    if BATCH["running"]:
        return JSONResponse(dict(ok=False, err="已有批次在跑，请先取消或等待完成"), status_code=409)
    srcs = payload.get("srcs")
    if payload.get("all_samples"):
        srcs = [s["src"] for s in samples()]
    if not srcs:
        return JSONResponse(dict(ok=False, err="未指定输入"), status_code=400)
    items = _normalize_srcs(srcs)
    out_dir = (payload.get("out_dir") or "").strip()
    BATCH.update(running=True, cancel=False, proc=None, items=items,
                 logs=[f"[起] {time.strftime('%H:%M:%S')} 批次启动：{len(items)} 本，"
                       f"输出目录={out_dir or '源文件所在目录（平铺）'}，断点续跑={'开' if payload.get('skip_done', True) else '关'}"],
                 started=time.strftime("%H:%M:%S"), finished=None,
                 skip_done=payload.get("skip_done", True))
    t = threading.Thread(target=_batch_worker, args=(items, out_dir), daemon=True)
    BATCH["thread"] = t
    t.start()
    return dict(ok=True, total=len(items))

@app.get("/api/batch/status")
def batch_status():
    return dict(running=BATCH["running"], started=BATCH["started"], finished=BATCH["finished"],
                items=BATCH["items"], logs=BATCH["logs"][-200:],
                progress=dict(done=sum(1 for i in BATCH["items"] if i["status"] in ("done", "skipped")),
                              failed=sum(1 for i in BATCH["items"] if i["status"] == "failed"),
                              total=len(BATCH["items"])))

@app.post("/api/batch/cancel")
def batch_cancel():
    if not BATCH["running"]: return dict(ok=False, err="没有在跑的批次")
    BATCH["cancel"] = True
    if BATCH["proc"] and BATCH["proc"].poll() is None:
        _kill_tree(BATCH["proc"].pid)
    return dict(ok=True)

@app.post("/api/batch/retry")
def batch_retry(payload: dict = Body(...)):
    """单本重试：独立小批次提交（走同一 worker，不占 web 线程）"""
    book_id = payload["book_id"]
    src = payload.get("src")
    if not src:
        m = manifest_latest().get(book_id, {})
        src = m.get("src")
    if not src or not Path(src).exists():
        return JSONResponse(dict(ok=False, err=f"找不到源文件: {src}"), status_code=404)
    return batch_start(dict(srcs=[src], skip_done=False))

if __name__ == "__main__":
    uvicorn.run(app, host="127.0.0.1", port=8300, log_level="warning")
