# -*- coding: utf-8 -*-
"""build 后回归校验：历史月是否保留 / MDATA 块数量 / 新功能是否还在 / W 归属是否异常"""
import re, json, sys
sys.path.insert(0, r"D:\WorkBuddy")
import build_month as B

html = open(r"D:\WorkBuddy\dashboard_offline.html", encoding="utf-8").read()
print("BUILD:", re.search(r'(?:var|let)\s+BUILD\s*=\s*"([^"]+)"', html).group(1))
print("CURRENT_MONTH:", re.search(r'(?:var|let)\s+CURRENT_MONTH\s*=\s*"([^"]+)"', html).group(1))
print("version.json:", open(r"D:\WorkBuddy\version.json", encoding="utf-8").read())

mds = re.findall(r"(?:var|let|const)\s+(MDATA_\d{4}_\d{2}_[DTBWE])\s*=", html)
print("\nMDATA 块数:", len(mds))
from collections import Counter
print("  按月:", sorted(Counter(m.rsplit("_", 2)[0][6:].replace("_", "-") for m in mds).items()))

mf = json.loads(B.extract(html, "MONTHLY_FOUR"))
print("\nMONTHLY_FOUR.months:", sorted(mf["months"].keys()))
print("year 键:", sorted(mf.get("year", {}).keys())[:12])

et = json.loads(B.extract(html, "EXTEND_TIME_DATA"))
cl = json.loads(B.extract(html, "CLEANING_DATA"))
for nm, o in (("EXTEND_TIME", et), ("CLEANING", cl)):
    print("\n%s months:" % nm, {k: len(v) for k, v in sorted(o["months"].items())})

# 10 月块抽查
d = json.loads(B.extract(html, "MDATA_2026_10_D"))
print("\nD 2026-10 meta:", json.dumps(d.get("meta"), ensure_ascii=False)[:220])
b = json.loads(B.extract(html, "MDATA_2026_10_B"))
print("B 2026-10 meta:", json.dumps(b.get("meta"), ensure_ascii=False)[:220])

w = json.loads(B.extract(html, "MDATA_2026_10_W"))
bad = set()
for day, nets in w.get("net", {}).items():
    for n, v in nets.items():
        if isinstance(v, dict) and v.get("office") not in (
            "新疆办事处", "宁夏办事处", "兰州办事处", "包头办事处", "呼和浩特办事处", "西宁办事处"):
            bad.add(str(v.get("office")))
print("\nW 异常办事处值:", bad if bad else "无 ✅")
print("W 周:", [x["label"] for x in w.get("weeks", [])])

print("\n功能保留自检:")
for k in [".iPieScope", 'id="itemPie"', "PIE_COLORS", "renderItemPie", "renderItemScopeSeg",
          "锁定规则 2026-09-21", "toFixed(2) + '%'", "const pct = x =>"]:
    print("   %-24s %s" % (k, "✅" if k in html else "❌ 缺失"))
print("   toFixed(1) 残留:", len(re.findall(r"toFixed\(1\)", html)), "处")
