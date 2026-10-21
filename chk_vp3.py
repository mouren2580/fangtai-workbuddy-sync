# -*- coding: utf-8 -*-
"""判断：新 Excel 里耗材配件行是真的减少（源数据回溯），还是解析漏匹配"""
import sys, datetime, collections, re
sys.path.insert(0, r'D:\WorkBuddy')
import gen_v2_lib as G
import openpyxl

EXCEL = r'D:\WorkBuddy\2026年10月西北服务产品.xlsx'
CUT = datetime.date(2026, 10, 15)
SVC = datetime.date(2026, 9, 28)
KEY = ['洗碗粉', '洗碗盐', '漂洗剂', '清洁膏', '蒸烤', '厨小护', '洗碗液', '软水盐', '洗碗块']

wb = openpyxl.load_workbook(EXCEL, read_only=True, data_only=True)
ws = wb['CSM配件']
tot = collections.Counter()      # 配件名 -> 数量
cnt = collections.Counter()      # 配件名 -> 行数
win_rows = 0
all_names = collections.Counter()
for r in G._rows(ws, 2):
    if not r or len(r) <= G.PJ_DATE:
        continue
    d = G._date(r[G.PJ_DATE])
    if d is None or d > CUT or d < SVC:
        continue
    win_rows += 1
    nm = G._s(r[G.PJ_NAME])
    all_names[nm] += 1
    if any(k in nm for k in KEY):
        tot[nm] += G._num(r[G.PJ_QTY])
        cnt[nm] += 1

print('CSM配件 窗口内行数 =', win_rows)
print('\n=== 窗口内含关键词的配件（名称 / 行数 / 数量）===')
for nm in sorted(tot, key=lambda x: -tot[x]):
    print('  %-42s %4d 行  %6.0f 件' % (nm, cnt[nm], tot[nm]))
print('  —— 合计 %.0f 件' % sum(tot.values()))

# 按天看耗材件数
byday = collections.Counter()
ws2 = wb['CSM配件']
for r in G._rows(ws2, 2):
    if not r or len(r) <= G.PJ_DATE:
        continue
    d = G._date(r[G.PJ_DATE])
    if d is None or d > CUT or d < SVC:
        continue
    nm = G._s(r[G.PJ_NAME])
    if any(k in nm for k in KEY):
        byday[d] += G._num(r[G.PJ_QTY])
print('\n=== 耗材按日（CSM配件，窗口内）===')
for d in sorted(byday):
    print('  %s %6.0f' % (d, byday[d]))

# 自购表耗材
ws3 = wb['工单自购配件明细']
Z = G._cols(ws3, 1, G.ZG_HEADERS, tag='自购')
zg = collections.Counter()
zg_rows = 0
for r in G._rows(ws3, 1):
    if not r or len(r) <= Z['date']:
        continue
    d = G._date(r[Z['date']])
    if d is None or d > CUT or d < SVC:
        continue
    zg_rows += 1
    nm = G._s(r[Z['name']])
    if any(k in nm for k in KEY):
        zg[nm] += G._num(r[Z['qty']])
print('\n=== 自购配件表 窗口内行数 = %d；耗材命中 ===' % zg_rows)
for nm in sorted(zg, key=lambda x: -zg[x]):
    print('  %-42s %6.0f 件' % (nm, zg[nm]))
print('  耗材合计 %.0f 件' % sum(zg.values()))
wb.close()
