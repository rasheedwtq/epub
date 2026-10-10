# -*- coding: utf-8 -*-
"""用 Chrome DevTools Protocol 驱动真实 headless Chrome 实测 epub.js 渲染 A1' 产物。
用法: python tools/reader_cdp_test.py <url> <输出截图> [等待秒]"""
import base64, json, subprocess, sys, time, urllib.request
import websocket

URL = sys.argv[1] if len(sys.argv) > 1 else "http://127.0.0.1:8300/reader"
SHOT = sys.argv[2] if len(sys.argv) > 2 else "work/reader_shots/cdp.png"
WAIT = float(sys.argv[3]) if len(sys.argv) > 3 else 12
PORT = 9333
CHROME = r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe"

proc = subprocess.Popen([CHROME, "--headless=new", "--disable-gpu", "--no-first-run",
    f"--remote-debugging-port={PORT}", "--remote-allow-origins=*", "--window-size=1200,900", "about:blank"],
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
    state = js("""(() => ({
        info: document.getElementById('info')?.textContent,
        meta: document.getElementById('meta')?.textContent,
        iframes: document.querySelectorAll('#viewer iframe').length,
        iframeText: (() => { const f = document.querySelector('#viewer iframe');
            try { return f ? f.contentDocument.body.innerText.slice(0, 300) : null; }
            catch(e) { return 'ERR:'+e.message; } })(),
        tables: (() => { const f = document.querySelector('#viewer iframe');
            try { return f ? f.contentDocument.querySelectorAll('table').length : 0; }
            catch(e) { return -1; } })(),
        imgs: (() => { const f = document.querySelector('#viewer iframe');
            try { return f ? f.contentDocument.querySelectorAll('img').length : 0; }
            catch(e) { return -1; } })(),
        bgs: (() => { const f = document.querySelector('#viewer iframe');
            try { return f ? f.contentDocument.querySelectorAll('[style*=background]').length : 0; }
            catch(e) { return -1; } })(),
    }))()""")
    print(json.dumps(state, ensure_ascii=False, indent=1))
    # 翻页一次再抓状态
    js("rendition && rendition.next(); 'ok'")
    time.sleep(4)
    page2 = js("""(() => { const f = document.querySelector('#viewer iframe');
        try { return f ? f.contentDocument.body.innerText.slice(0, 300) : null; }
        catch(e) { return 'ERR:'+e.message; } })()""")
    print("翻页后文本:", page2)
    data = call("Page.captureScreenshot", format="png")
    open(SHOT, "wb").write(base64.b64decode(data["data"]))
    print("截图:", SHOT)
    ws.close()
finally:
    proc.kill()
