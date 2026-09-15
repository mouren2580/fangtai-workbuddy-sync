# -*- coding: utf-8 -*-
"""构建后完整性验证"""
import re, json, sys, subprocess, os
sys.path.insert(0, r"D:\WorkBuddy")
import build_month as BM

h = open("dashboard_offline.html", encoding="utf-8").read()
def g(v): return json.loads(BM.extract(h, v))

print("BUILD:", re.search(r'var BUILD = "([^"]*)"', h).group(1))
print("version.json:", open("version.json", encoding="utf-8").read())

cl = g("CLEANING_DATA")
print("\nCLEANING_DATA months:", {k: len(v) for k, v in cl["months"].items()})
et = g("EXTEND_TIME_DATA")
print("EXTEND_TIME_DATA months:", {k: len(v) for k, v in et["months"].items()})
mf = g("MONTHLY_FOUR")
print("MONTHLY_FOUR months:", sorted(mf["months"].keys()))
print("MONTHLY_FOUR.year 键:", list(mf.get("year", {}).keys())[:6] if isinstance(mf.get("year"), dict) else type(mf.get("year")).__name__)

names = set(re.findall(r'(?m)^(?:const|let)\s+(MDATA_20\d\d_\d\d_[DTBWE])\s*=', h))
print("\nMDATA 块:", sorted(names))

D = g("MDATA_2026_09_D"); T = g("MDATA_2026_09_T"); B = g("MDATA_2026_09_B")
W = g("MDATA_2026_09_W"); E = g("MDATA_2026_09_E")
print("\nD: 网点 %d 合计 %.2f time=%s" % (len(D["网点"]), sum(x["合计"] for x in D["网点"]), json.dumps(D["time"], ensure_ascii=False)))
print("T: 工程师 %d 工单消耗 %.2f" % (T["count"], T["total"]))
print("B: 总单 %s 分母 %s 大保养 %s 占比 %s%%" % (B["meta"]["清洗保养总单"], B["meta"]["分母"], B["meta"]["大保养总数"], B["meta"]["全区占比"]))
print("W: 周数 %d  周标签 %s" % (len(W["weeks"]), [x["label"] for x in W["weeks"]]))
print("E: 工程师 %d  y1 %.2f y2 %.2f y3 %.2f" % (E["meta"]["工程师数"], E["total"]["y1"]["amt"], E["total"]["y2"]["amt"], E["total"]["y3"]["amt"]))
wo = mf["months"]["2026-09"]["workorder"]
print("\nworkorder: 工单 %s 网点 %s 工程师 %s" % (wo["meta"].get("总工单", wo["meta"].get("工单总数")), wo["meta"].get("网点数"), wo["meta"].get("工程师数")))
print("止回阀/烟机安装/防火阀占比:", wo["meta"].get("止回阀总数"), wo["meta"].get("烟机安装数"), wo["meta"].get("防火阀占比"))
va = mf["months"]["2026-09"]["valueadded"]; vp = mf["months"]["2026-09"]["valueparts"]
print("valueadded 全区总计数:", va["meta"]["全区总计数"])
print("valueparts 当月总计数:", vp["meta"]["全区总计数_当月"])
print("延保明细 %d 条  清洗明细 %d 条" % (len(et["months"]["2026-09"]), len(cl["months"]["2026-09"])))

# 三模块止回阀一致性
ah = None
for p in vp["products"]:
    if p["name"] == "止回阀": ah = p["qty_month"]
print("\n[互校] 止回阀  工单=%s  增值配件=%s" % (wo["meta"].get("止回阀总数"), ah))
print("[互校] 清洗保养总单  B=%s  bigcare=%s" % (B["meta"]["清洗保养总单"], mf["months"]["2026-09"]["bigcare"]["meta"]["清洗保养总单"]))

# JS 语法
blocks = re.findall(r'<script[^>]*>(.*?)</script>', h, re.S)
big = max(blocks, key=len)
open("_syntax_check.js", "w", encoding="utf-8").write(big)
node = r"C:/Users/40973/.workbuddy/binaries/node/versions/22.12.0/node.exe"
r = subprocess.run([node, "--check", "_syntax_check.js"], capture_output=True, text=True)
print("\nJS 语法(node --check):", "通过 ✅" if r.returncode == 0 else "失败 ⚠ " + r.stderr[:300])
print("script 块:", len(blocks), " 文件大小: %.2f MB" % (os.path.getsize("dashboard_offline.html") / 1048576))
