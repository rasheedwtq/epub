# -*- coding: utf-8 -*-
"""配置中心 API 冒烟测试：PUT 存配置 → 版本快照 → 回滚"""
import json, urllib.request

B = "http://127.0.0.1:8300"

def call(path, method="GET", data=None):
    req = urllib.request.Request(B + path,
                                 data=json.dumps(data).encode() if data is not None else None,
                                 headers={"Content-Type": "application/json"}, method=method)
    return json.loads(urllib.request.urlopen(req).read().decode())

cfg = call("/api/config")
old_h1 = cfg["styles"]["h1_size"]
cfg["styles"]["h1_size"] = "1.8em"
print("PUT:", call("/api/config", "PUT", cfg))
vs = call("/api/versions")
print("版本快照:", vs[:2])
print("当前 h1_size:", call("/api/config")["styles"]["h1_size"])
print("回滚:", call("/api/rollback/" + vs[0], "POST"))
print("回滚后 h1_size:", call("/api/config")["styles"]["h1_size"], "(应为", old_h1, ")")
