# -*- coding: utf-8 -*-
"""诊断：epub.js 渲染下图片溢出测量（复现小看板右侧出血）"""
import base64, json, subprocess, sys, time, urllib.request
import websocket

URL = sys.argv[1] if len(sys.argv) > 1 else "http://127.0.0.1:8300/reader?book=1云工作室案例（20231121图片修改及内容微调版本）"
SHOT = sys.argv[2] if len(sys.argv) > 2 else "work/reader_shots/kanban.png"
WAIT = float(sys.argv[3]) if len(sys.argv) > 3 else 12
PORT = 9334
CHROME = r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe"

proc = subprocess.Popen([CHROME, "--headless=new", "--disable-gpu", "--no-first-run",
    f"--remote-debugging-port={PORT}", "--remote-allow-origins=*", "--window-size=900,700", "about:blank"],
    stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
try:
    ws_url = None
    for _ in range(40):
        time.sleep(0.5)
        try:
            targets = json.loads(urllib.request.urlopen(f"http://127.0.0.1:{PORT}/json").read())
            pages = [t for t in targets if t["type"] == "page"]
            if pages:
                ws_url = pages[0]["webSocketDebuggerUrl"]; break
        except Exception:
            pass
    assert ws_url, "CDP 未就绪"
    ws = websocket.create_connection(ws_url, timeout=30)
    mid = 0
    def call(method, **params):
        global mid
        mid += 1
        ws.send(json.dumps({"id": mid, "method": method, "params": params}))
        while True:
            msg = json.loads(ws.recv())
            if msg.get("id") == mid:
                return msg.get("result", {})
    def js(expr, await_promise=False):
        r = call("Runtime.evaluate", expression=expr, returnByValue=True,
                 awaitPromise=await_promise)
        return r.get("result", {}).get("value")

    call("Page.enable")
    call("Page.navigate", url=URL)
    time.sleep(WAIT)
    # 翻到含图页并测量所有图片的溢出情况
    measure = """(() => {
      const f = document.querySelector('#viewer iframe');
      if (!f) return 'NO_IFRAME';
      const fd = f.contentDocument;
      const cw = fd.documentElement.clientWidth;
      const imgs = [...fd.querySelectorAll('img')];
      return {
        colWidth: cw,
        cssRules: [...fd.styleSheets].map(s => { try { return s.href || 'inline' } catch(e){ return 'CORS' } }),
        figCssApplied: (() => { try {
            const r = [...fd.styleSheets].flatMap(s => { try { return [...s.cssRules] } catch(e){ return [] } })
                .find(r => r.selectorText && r.selectorText.includes('.fig'));
            return r ? r.style.cssText : null } catch(e){ return 'ERR' } })(),
        imgs: imgs.map(im => {
          const b = im.getBoundingClientRect();
          return { src: (im.getAttribute('src')||'').slice(-30), nat: im.naturalWidth + 'x' + im.naturalHeight,
                   disp: Math.round(b.width) + 'x' + Math.round(b.height), right: Math.round(b.right),
                   overflow: b.right > cw ? 'YES+' + Math.round(b.right - cw) : 'no' };
        })
      };
    })()"""
    # 逐页翻找含图页（最多翻 12 页）
    found = None
    for i in range(12):
        st = js(measure)
        if isinstance(st, dict) and st.get("imgs"):
            found = (i, st)
            break
        js("rendition && rendition.next(); 'ok'")
        time.sleep(2.5)
    print("页面#", found[0] if found else "-")
    print(json.dumps(found[1] if found else "未找到含图页", ensure_ascii=False, indent=1))
    data = call("Page.captureScreenshot", format="png")
    open(SHOT, "wb").write(base64.b64decode(data["data"]))
    print("截图:", SHOT)
    ws.close()
finally:
    proc.kill()
