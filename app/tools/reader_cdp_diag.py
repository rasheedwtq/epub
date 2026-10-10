# -*- coding: utf-8 -*-
"""一次性诊断：查 epub.js 内部 navPath/ncxPath/navigation 状态"""
import json, time, urllib.request
import websocket

PORT = 9333
CHROME = r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe"
BOOK = "006_%E8%A1%A8%E6%A0%BC%E6%B7%B7%E5%90%88_%E9%AB%98_%E5%BE%88%E8%84%8F"

proc = subprocess_proc = None
import subprocess
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
            if pages: ws_url = pages[0]["webSocketDebuggerUrl"]; break
        except Exception: pass
    ws = websocket.create_connection(ws_url, timeout=30)
    mid = 0
    def call(method, **params):
        global mid
        mid += 1
        ws.send(json.dumps({"id": mid, "method": method, "params": params}))
        while True:
            msg = json.loads(ws.recv())
            if msg.get("id") == mid: return msg.get("result", {})
    def js(expr):
        r = call("Runtime.evaluate", expression=expr, returnByValue=True, awaitPromise=True)
        return r.get("result", {}).get("value")
    call("Page.enable")
    call("Page.navigate", url=f"http://127.0.0.1:8300/reader?book={BOOK}")
    time.sleep(14)
    out = js("""(async () => {
      const b = book;
      const toc2 = await b.loaded.toc;
      return JSON.stringify({
        navPath: b.packaging && b.packaging.navPath,
        ncxPath: b.packaging && b.packaging.ncxPath,
        spineToc: b.spine && b.spine.toc,
        navigation: b.navigation ? (b.navigation.toc||[]).length + '条' : 'null',
        loadedToc: Array.isArray(toc2) ? toc2.length : String(toc2),
        nav: b.archived ? undefined : undefined,
      })})()""")
    print("内部状态:", out)
    # 直接用 jszip 看 ncx/nav 是否真的在包里可读
    out2 = js("""(async () => {
      const nav = await b.load('nav.xhtml').catch(e=>'ERR:'+e.message);
      const ncx = await b.load('toc.ncx').catch(e=>'ERR:'+e.message);
      return JSON.stringify({
        navType: nav && nav.constructor.name, navLen: nav ? (nav.outerHTML||String(nav)).length : null,
        ncxType: ncx && ncx.constructor.name, ncxLen: ncx ? (ncx.outerHTML||String(ncx)).length : null,
        ncxNavPoints: ncx && ncx.querySelectorAll ? ncx.querySelectorAll('navPoint').length : null,
      })})()""")
    print("nav/ncx 可读性:", out2)
    ws.close()
finally:
    proc.kill()
