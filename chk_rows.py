# -*- coding: utf-8 -*-
"""逐行对比 旧(10-15) vs 新(10-20) Excel 的 CSM配件 / 自购配件表 在 9.28~10.15 窗口内的行"""
import sys, datetime, collections
sys.path.insert(0, r'D:\WorkBuddy')
import gen_v2_lib as G
import openpyxl

OLD = r'D:\WorkBuddy\_old_1015.xlsx'
NEW = r'D:\WorkBuddy\2026年10月西北服务产品.xlsx'
CUT = datetime.date(2026, 10, 15)
SVC = datetime.date(2026, 9, 28)

def load_pj(path, sheet, hdr_skip, spec):
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    ws = wb[sheet]
    C = G._cols(ws, hdr_skip, spec, tag='')
    bag = collections.Counter()
    for r in G._rows(ws, hdr_skip):
        if not r or len(r) <= C['date']:
            continue
        d = G._date(r[C['date']])
        if d is None or d > CUT or d < SVC:
            continue
        key = (G._s(r[C['tech']]), G._s(r[C['tcode']]), G._s(r[C['code']]),
               G._s(r[C['name']]), G._num(r[C['qty']]), d)
        bag[key] += 1
    wb.close()
    return bag

PJ = {"office": ("办事处", G.PJ_OFFICE), "net": ("服务网点", G.PJ_NET),
      "tech": ("服务工程师", G.PJ_TECH), "tcode": ("服务工程师编码", G.PJ_TCODE),
      "class": ("销售分类", G.PJ_CLASS), "name": ("配件名称", G.PJ_NAME),
      "code": ("配件编码", G.PJ_CODE), "qty": ("数量", G.PJ_QTY),
      "amt": ("上缴金额", G.PJ_AMT), "date": ("录入完成时间", G.PJ_DATE)}

for sheet, spec, skip in (('CSM配件', PJ, 2), ('工单自购配件明细', G.ZG_HEADERS, 1)):
    o = load_pj(OLD, sheet, skip, spec)
    n = load_pj(NEW, sheet, skip, spec)
    print('\n' + '=' * 78)
    print('%s  窗口 %s~%s' % (sheet, SVC, CUT))
    print('  旧行数 %d   新行数 %d' % (sum(o.values()), sum(n.values())))
    only_o = o - n
    only_n = n - o
    print('  仅旧有: %d 行（合计 %.0f 件）' % (sum(only_o.values()),
          sum(k[4] * v for k, v in only_o.items())))
    print('  仅新有: %d 行（合计 %.0f 件）' % (sum(only_n.values()),
          sum(k[4] * v for k, v in only_n.items())))
    if only_o:
        print('  --- 仅旧有（最多 25 条）---')
        for k, v in sorted(only_o.items(), key=lambda x: -x[0][4] * x[1])[:25]:
            print('    %s %s(%s) %s | %s | %.0f件 x%d' % (k[5], k[0], k[1], k[2], k[3][:26], k[4], v))
    if only_n:
        print('  --- 仅新有（最多 25 条）---')
        for k, v in sorted(only_n.items(), key=lambda x: -x[0][4] * x[1])[:25]:
            print('    %s %s(%s) %s | %s | %.0f件 x%d' % (k[5], k[0], k[1], k[2], k[3][:26], k[4], v))
