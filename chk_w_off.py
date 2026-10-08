# -*- coding: utf-8 -*-
import sys, json
sys.path.insert(0, r"D:\WorkBuddy")
import build_month as B
html = open(r"D:\WorkBuddy\dashboard_offline.html", encoding="utf-8").read()
w = json.loads(B.extract(html, "MDATA_2026_10_W"))
print("office 维度聚合:")
for wk in sorted(w["office"]):
    tot = sum(v["total"] for v in w["office"][wk].values())
    bad = {o: v for o, v in w["office"][wk].items()
           if o not in ("新疆办事处", "宁夏办事处", "兰州办事处", "包头办事处", "呼和浩特办事处", "西宁办事处")}
    print("  %s  周合计=%.2f  异常=%s" % (wk, tot, json.dumps(bad, ensure_ascii=False)))
print("\nnet 维度异常网点:")
n = 0
for wk in sorted(w["net"]):
    for nm, v in w["net"][wk].items():
        if v.get("office") not in ("新疆办事处", "宁夏办事处", "兰州办事处", "包头办事处", "呼和浩特办事处", "西宁办事处"):
            n += 1
            if n <= 12:
                print("   %s | %s | office=%s total=%s" % (wk, nm, v.get("office"), v.get("total")))
print("   异常网点行数:", n)
