# -*- coding: utf-8 -*-
"""打印 CSM配件 中配件类行的完整 38 列，判断各列含义。"""
from openpyxl import load_workbook
EXCEL = r"D:\WorkBuddy\2026年10月西北服务产品.xlsx"
wb = load_workbook(EXCEL, read_only=True, data_only=True)
ws = wb["CSM配件"]
rows = list(ws.iter_rows(values_only=True))
hdr = rows[1]
data = rows[2:]
CARE = {"防烟宝", "排烟管", "灶具保养液", "净水机", "出风管", "波纹管", "装饰罩", "养护套装",
        "其他烟机", "其他热水器", "其他灶具", "止回阀"}
shown = 0
for r in data:
    if str(r[16]) in CARE and shown < 6:
        print("----- 销售分类=%s -----" % r[16])
        for i, h in enumerate(hdr):
            if r[i] not in (None, "", 0):
                print("   %2d %-16s = %s" % (i, h, str(r[i])[:40]))
        shown += 1
print("\n===== 配件类行小计 =====")
from collections import Counter
c = Counter(str(r[16]) for r in data if str(r[16]) in CARE)
print(c)
wb.close()
