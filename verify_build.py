# -*- coding: utf-8 -*-
"""build 后回归校验：历史月是否原样保留 / 新月份块是否齐全 / 功能是否还在 / W 归属是否异常"""
import re, json, sys
sys.path.insert(0, r"D:\WorkBuddy")
import build_month as B

NEW = r"D:\WorkBuddy\dashboard_offline.html"
PREV = r"D:\WorkBuddy\_prev.html"
html = open(NEW, encoding="utf-8").read()
prev = open(PREV, encoding="utf-8").read() if __import__("os").path.exists(PREV) else None

print("BUILD:", re.search(r'(?:var|let)\s+BUILD\s*=\s*"([^"]+)"', html).group(1))
CUR = re.search(r'(?:var|let)\s+CURRENT_MONTH\s*=\s*"([^"]+)"', html).group(1)
print("CURRENT_MONTH:", CUR)
print("version.json:", open(r"D:\WorkBuddy\version.json", encoding="utf-8").read())

mds = re.findall(r"(?:var|let|const)\s+(MDATA_\d{4}_\d{2}_[DTBWE])\s*=", html)
print("\nMDATA 块数:", len(mds), "（11 月起应为 50）")
from collections import Counter
print("  按月:", sorted(Counter(m.rsplit("_", 2)[0][6:].replace("_", "-") for m in mds).items()))

mf = json.loads(B.extract(html, "MONTHLY_FOUR"))
print("\nMONTHLY_FOUR.months:", sorted(mf["months"].keys()))
et = json.loads(B.extract(html, "EXTEND_TIME_DATA"))
cl = json.loads(B.extract(html, "CLEANING_DATA"))
for nm, o in (("EXTEND_TIME", et), ("CLEANING", cl)):
    print("%s months:" % nm, {k: len(v) for k, v in sorted(o["months"].items())})

# ---------- 历史月原样性（关键：跨月 build 不得把上月退回底板旧快照）----------
print("\n===== 历史月原样性（与 _prev.html 逐月比对）=====")
if prev:
    pmf = json.loads(B.extract(prev, "MONTHLY_FOUR"))
    pet = json.loads(B.extract(prev, "EXTEND_TIME_DATA"))
    pcl = json.loads(B.extract(prev, "CLEANING_DATA"))
    bad = 0
    for mk in sorted(pmf["months"]):
        same = json.dumps(pmf["months"][mk], sort_keys=True) == json.dumps(mf["months"].get(mk), sort_keys=True)
        if not same:
            bad += 1
            print("   ❌ MONTHLY_FOUR[%s] 变了" % mk)
        else:
            print("   ✅ MONTHLY_FOUR[%s] 原样保留" % mk)
    for mk in sorted(pet["months"]):
        if json.dumps(pet["months"][mk], sort_keys=True) != json.dumps(et["months"].get(mk), sort_keys=True):
            bad += 1; print("   ❌ EXTEND_TIME_DATA[%s] 变了" % mk)
    for mk in sorted(pcl["months"]):
        if json.dumps(pcl["months"][mk], sort_keys=True) != json.dumps(cl["months"].get(mk), sort_keys=True):
            bad += 1; print("   ❌ CLEANING_DATA[%s] 变了" % mk)
    pnames = re.findall(r"(?:var|let|const)\s+(MDATA_\d{4}_\d{2}_[DTBWE])\s*=", prev)
    for n in pnames:
        if json.dumps(json.loads(B.extract(prev, n)), sort_keys=True) != json.dumps(json.loads(B.extract(html, n)), sort_keys=True):
            bad += 1; print("   ❌ %s 变了" % n)
    print("   → 历史月差异项: %d %s" % (bad, "✅" if bad == 0 else "⚠"))
else:
    print("   （无 _prev.html，跳过）")

# ---------- 当月新块 ----------
TAG = CUR.replace("-", "_")
print("\n===== 当月 %s 数据块 =====" % CUR)
for suf in "DTBWE":
    v = json.loads(B.extract(html, "MDATA_%s_%s" % (TAG, suf)))
    print("   MDATA_%s_%s ✅ keys=%s" % (TAG, suf, sorted(v.keys())[:6]))

wo = mf["months"][CUR]["workorder"]["meta"]
print("\n   工单: 总数=%s 网点=%s 工程师=%s 止回阀=%s 占比=%s%% 截止=%s 周期=%s"
      % (wo.get("总工单数"), wo.get("网点数"), wo.get("工程师数"), wo.get("止回阀总数"),
         wo.get("防火阀占比"), wo.get("截止"), wo.get("周期")))
va = mf["months"][CUR]["valueadded"]["meta"]
vp = mf["months"][CUR]["valueparts"]["meta"]
bc = mf["months"][CUR]["bigcare"]["meta"]
print("   增值产品: 总计数=%s 有销量=%s" % (va.get("全区总计数"), va.get("有销量项数")))
print("   增值配件: 当月=%s 昨日=%s 近7日=%s" % (vp.get("全区总计数_当月"), vp.get("全区总计数_昨日"), vp.get("全区总计数_近7日")))
print("   大保养: 总单=%s 分母=%s 大保养=%s 占比=%s%%" % (bc.get("清洗保养总单"), bc.get("分母"), bc.get("大保养总数"), bc.get("全区占比")))

W = json.loads(B.extract(html, "MDATA_%s_W" % TAG))
bad_off = set()
for day, nets in W.get("net", {}).items():
    for n, v in nets.items():
        if isinstance(v, dict) and v.get("office") not in (
                "新疆办事处", "宁夏办事处", "兰州办事处", "包头办事处", "呼和浩特办事处", "西宁办事处"):
            bad_off.add(str(v.get("office")))
print("\n   W 异常办事处值:", bad_off if bad_off else "无 ✅")
print("   W 周:", [(x["label"], round(sum((x.get("office") or {}).get(o, {}).get("total", 0) for o in []), 2)) for x in W.get("weeks", [])])

print("\n功能保留自检:")
for k in [".iPieScope", 'id="itemPie"', "PIE_COLORS", "renderItemPie", "renderItemScopeSeg",
          "锁定规则 2026-09-21", "toFixed(2) + '%'", "const pct = x =>", "__CACHEBUST_START__"]:
    print("   %-24s %s" % (k, "✅" if k in html else "❌ 缺失"))
print("   toFixed(1) 残留:", len(re.findall(r"toFixed\(1\)", html)), "处")
print("   MONTH_DATA 含当月:", "✅" if '"%s"' % CUR in B.extract(html, "MONTH_DATA") else "❌")
