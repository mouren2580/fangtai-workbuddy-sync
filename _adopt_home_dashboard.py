# -*- coding: utf-8 -*-
"""把家里包的新看板接入办公室工作区：
1) 替换缓存穿透段为「按版本号去重」的修复版
2) 更新 BUILD 号
3) 写出 dashboard_offline.html / dashboard.html / deploy_cs/index.html
4) 同步 version.json
"""
import json
import os
import re

HERE = r"D:\WorkBuddy"
SRC = os.path.join(HERE, "_pkg_1105", "fangtai-dashboard-update", "dashboard.html")
OLD = os.path.join(HERE, "dashboard_offline.html")
NEW_BUILD = "20261106-1003"
CUT = "2026-11-04"

s = open(SRC, encoding="utf-8").read()

# 取办公室已修复的缓存穿透段
old_html = open(OLD, encoding="utf-8").read()
m_s = old_html.index("/*__CACHEBUST_START__*/")
m_e = old_html.index("/*__CACHEBUST_END__*/") + len("/*__CACHEBUST_END__*/")
fixed = old_html[m_s:m_e]
fixed = re.sub(r'(var BUILD = ")[^"]*(")', r"\g<1>%s\g<2>" % NEW_BUILD, fixed, count=1)
assert 'var BUILD = "%s"' % NEW_BUILD in fixed

# 替换新文件的旧段
n_s = s.index("/*__CACHEBUST_START__*/")
n_e = s.index("/*__CACHEBUST_END__*/") + len("/*__CACHEBUST_END__*/")
s2 = s[:n_s] + fixed + s[n_e:]

# BUILD 全局只应出现一次
assert s2.count('var BUILD = "%s"' % NEW_BUILD) == 1, s2.count('var BUILD = ')

out = os.path.join(HERE, "dashboard_offline.html")
open(out, "w", encoding="utf-8", newline="\n").write(s2)
print("已写 dashboard_offline.html  %.2f MB" % (len(s2.encode()) / 1048576))

ver = {"v": NEW_BUILD, "cut": CUT, "t": "2026-11-06T10:03:00"}
js = json.dumps(ver, ensure_ascii=False)
for p in (os.path.join(HERE, "version.json"), os.path.join(HERE, "deploy_cs", "version.json")):
    open(p, "w", encoding="utf-8").write(js)
print("version.json →", js)
