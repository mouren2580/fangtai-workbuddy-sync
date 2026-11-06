# -*- coding: utf-8 -*-
"""核对看板内 10/11 月关键指标与去年同期"""
import json
import sys

sys.path.insert(0, r"D:\WorkBuddy")
from build_month import extract  # noqa: E402

s = open(r"D:\WorkBuddy\dashboard_offline.html", encoding="utf-8").read()
mf = json.loads(extract(s, "MONTHLY_FOUR"))
for mk in ("2026-09", "2026-10", "2026-11"):
    b = mf["months"][mk]
    wo, va = b["workorder"]["meta"], b["valueadded"]["meta"]
    print("%s: 工单=%s 网点=%s 工程师=%s 止回阀占比=%.2f%% 增值产品=%s 增值配件当月=%s 大保养占比=%.2f%%"
          % (mk, wo.get("总工单数"), wo.get("网点数"), wo.get("工程师数"), wo.get("防火阀占比", 0) or 0,
             va.get("全区总计数"), b["valueparts"]["meta"].get("全区总计数_当月"),
             b["bigcare"]["meta"].get("全区占比", 0) or 0))

y = mf["year"]["workorder"]["meta"]
print()
print("年度累计: 总工单=%s 截止=%s" % (y.get("总工单数"), y.get("截止") or y.get("统计截止日")))

ly = json.loads(extract(s, "LASTYEAR_BY_MONTH"))
print()
print("去年同期 2026-11:", json.dumps(ly.get("2026-11"), ensure_ascii=False)[:300])
print("去年同期 2026-10:", json.dumps(ly.get("2026-10"), ensure_ascii=False)[:200])

ur = json.loads(extract(s, "UNIT_REVENUE_DATA"))
print()
print("UNIT_REVENUE_DATA keys:", list(ur.keys())[:8])
for k in list(ur.keys())[-3:]:
    v = ur[k]
    print("  %s → %s" % (k, json.dumps(v, ensure_ascii=False)[:200]))
