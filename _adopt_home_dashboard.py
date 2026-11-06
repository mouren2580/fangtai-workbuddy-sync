# -*- coding: utf-8 -*-
"""把家里包的新看板接入办公室工作区：
1) 替换缓存穿透段为「按版本号去重」的修复版（家里原版仍是旧的 _cbN 计数版）
2) 更新 BUILD 号
3) 写出 dashboard_offline.html / dashboard.html / deploy_cs/index.html
4) 同步 version.json（含 deploy_cs）

用法： python _adopt_home_dashboard.py 20261106-1032 2026-11-05
"""
import json
import os
import re
import shutil
import sys
import time

HERE = r"D:\WorkBuddy"
SRC = os.path.join(HERE, "_pkg_1105", "fangtai-dashboard-update", "dashboard.html")
OLD = os.path.join(HERE, "dashboard_offline.html")
NEW_BUILD = sys.argv[1] if len(sys.argv) > 1 else time.strftime("%Y%m%d-%H%M")
CUT = sys.argv[2] if len(sys.argv) > 2 else "2026-11-05"

s = open(SRC, encoding="utf-8").read()

# 取办公室已修复的缓存穿透段（从上一版工作区看板里抠出来）
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

# BUILD 全局只应出现一次；旧缓存计数字段必须已消失
assert s2.count('var BUILD = "%s"' % NEW_BUILD) == 1, s2.count("var BUILD = ")
assert "_cbTried" in s2 and "_cbN'" not in s2

out = os.path.join(HERE, "dashboard_offline.html")
open(out, "w", encoding="utf-8", newline="\n").write(s2)
print("已写 dashboard_offline.html  %.2f MB" % (len(s2.encode()) / 1048576))

for dst in (os.path.join(HERE, "dashboard.html"), os.path.join(HERE, "deploy_cs", "index.html")):
    shutil.copyfile(out, dst)
    print("已写", dst)

ver = {"v": NEW_BUILD, "cut": CUT, "t": time.strftime("%Y-%m-%dT%H:%M:%S")}
js = json.dumps(ver, ensure_ascii=False)
for p in (os.path.join(HERE, "version.json"), os.path.join(HERE, "deploy_cs", "version.json")):
    open(p, "w", encoding="utf-8").write(js)
print("version.json →", js)
