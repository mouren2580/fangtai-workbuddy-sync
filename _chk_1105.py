# -*- coding: utf-8 -*-
"""核对新看板的关键数（年度累计 + 当月/上月块）"""
import json
import sys
sys.path.insert(0, r"D:\WorkBuddy")
from _cmp_blocks import decls

p = sys.argv[1] if len(sys.argv) > 1 else r"D:\WorkBuddy\_pkg_1105\fangtai-dashboard-update\dashboard.html"
d = decls(p)
mf = json.loads(d["MONTHLY_FOUR"])
y = mf.get("year", {})
print("== MONTHLY_FOUR.year ==")
for k, v in y.items():
    if isinstance(v, (int, float, str)):
        print("  %s = %s" % (k, v))
for mk in ("2026-10", "2026-11"):
    b = mf["months"][mk]
    wo = b["workorder"]["meta"]
    print("\n== %s ==" % mk)
    print("  工单 %s | 截止 %s | 网点 %s | 工程师 %s" % (wo.get("总工单数"), wo.get("截止日"), wo.get("网点数"), wo.get("工程师数")))
    print("  防火阀 %s%% | 增值产品 %s | 配件当月 %s | 大保养 %s%%" % (
        wo.get("防火阀占比"), b["valueadded"]["meta"].get("全区总计数"),
        b["valueparts"]["meta"].get("全区总计数_当月"), b["bigcare"]["meta"].get("全区占比")))

# 当月收入块
D = json.loads(d["MDATA_2026_11_D"])
print("\n== MDATA_2026_11_D ==")
print("  截止 %s | 网点 %d | 合计 ¥%.1f" % (D["meta"]["统计截止日"], len(D["网点"]), sum(n["合计"] for n in D["网点"])))
T = json.loads(d["MDATA_2026_11_T"])
print("== T ==  人数 %s | 合计 ¥%s" % (T["meta"].get("工程师数") or T["meta"].get("人数"), T["meta"].get("合计") if "合计" in T["meta"] else sum(r.get("合计", 0) for r in T.get("技师", T.get("techs", [])))))
E = json.loads(d["EXTEND_TIME_DATA"])
e = E["months"]["2026-11"]["total"] if isinstance(E.get("months", {}).get("2026-11"), dict) and "total" in E["months"]["2026-11"] else E["months"]["2026-11"]
print("== 延保 2026-11 ==", e)
C = json.loads(d["CLEANING_DATA"])
cm = C["months"]["2026-11"]
print("== 清洗 2026-11 == 单数 %d" % len(cm))
U = json.loads(d["UNIT_REVENUE_DATA"])
for mk in ("2026-10", "2026-11"):
    u = U.get("months", {}).get(mk)
    if u:
        print("== 单均 %s ==" % mk, json.dumps({k: v for k, v in u.items() if isinstance(v, (int, float, str))}, ensure_ascii=False)[:300])
