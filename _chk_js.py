# -*- coding: utf-8 -*-
"""检查 dashboard 内联 JS 语法（node --check 逐块）"""
import os
import re
import subprocess
import sys
import tempfile

p = sys.argv[1] if len(sys.argv) > 1 else r"D:\WorkBuddy\dashboard_offline.html"
s = open(p, encoding="utf-8").read()
blocks = re.findall(r"<script(?![^>]*src=)[^>]*>(.*?)</script>", s, re.S)
bad = 0
n = 0
for i, b in enumerate(blocks):
    if not b.strip():
        continue
    n += 1
    f = os.path.join(tempfile.gettempdir(), "_syn_chk_%d.js" % i)
    open(f, "w", encoding="utf-8").write(b)
    r = subprocess.run(["node", "--check", f], capture_output=True, text=True)
    if r.returncode != 0:
        bad += 1
        print("✘ 块 %d 语法错误：%s" % (i, (r.stderr or "").strip().splitlines()[:3]))
    os.remove(f)
print("内联脚本 %d 块，语法错误 %d 块 → %s" % (n, bad, "✔ 全部通过" if not bad else "✘"))
sys.exit(1 if bad else 0)
