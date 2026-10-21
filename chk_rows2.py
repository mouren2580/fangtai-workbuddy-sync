# -*- coding: utf-8 -*-
"""统计 自购配件表 消失的 244 行里，有多少命中增值产品"""
import sys, datetime, collections
sys.path.insert(0, r'D:\WorkBuddy')
import gen_v2_lib as G
import openpyxl

OLD = r'D:\WorkBuddy\_old_1015.xlsx'
NEW = r'D:\WorkBuddy\2026年10月西北服务产品.xlsx'
CUT = datetime.date(2026, 10, 15); SVC = datetime.date(2026, 9, 28)

code2name, pname2name = {}, {}
for n, codes, pns in G.VA_PRODUCTS:
    for c in codes: code2name[str(c)] = n
    for pn in pns: pname2name[str(pn)] = n

def load(path):
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    ws = wb['工单自购配件明细']
    C = G._cols(ws, 1, G.ZG_HEADERS, tag='')
    bag = collections.Counter()
    for r in G._rows(ws, 1):
        if not r or len(r) <= C['date']: continue
        d = G._date(r[C['date']])
        if d is None or d > CUT or d < SVC: continue
        bag[(G._s(r[C['tech']]), G._s(r[C['tcode']]), G._s(r[C['code']]),
             G._s(r[C['name']]), G._num(r[C['qty']]), d)] += 1
    wb.close()
    return bag

o, n = load(OLD), load(NEW)
gone = o - n
hit = collections.Counter(); hitqty = collections.Counter()
byday = collections.Counter(); byprod = collections.Counter()
for k, v in gone.items():
    nm = code2name.get(k[2]) or pname2name.get(k[3])
    if nm:
        hit[k[3] or k[2]] += v
        hitqty[k[3] or k[2]] += k[4] * v
        byday[k[5]] += k[4] * v
        byprod[nm] += k[4] * v

print('消失行数 = %d，其中命中增值产品 %d 种键，件数合计 %.0f' %
      (sum(gone.values()), len(hit), sum(hitqty.values())))
print('\n=== 命中增值产品的消失行（按物料/名称）===')
for k, v in sorted(hitqty.items(), key=lambda x: -x[1]):
    print('  %-22s %3d 行  %5.0f 件' % (k, hit[k], v))
print('\n=== 按产品 ===')
for k, v in sorted(byprod.items(), key=lambda x: -x[1]):
    print('  %-14s %+.0f 件' % (k, v))
print('\n=== 消失行的日期分布（全部行 / 命中增值产品的行）===')
allday = collections.Counter()
for k, v in gone.items(): allday[k[5]] += v
for d in sorted(allday):
    print('  %s  消失 %3d 行   其中增值 %.0f 件' % (d, allday[d], byday.get(d, 0)))
