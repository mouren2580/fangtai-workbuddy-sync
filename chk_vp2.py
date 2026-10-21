# -*- coding: utf-8 -*-
"""判断 valueparts 差异: 源表行撤回 还是 列读错"""
import sys, datetime, collections
sys.path.insert(0, r'D:\WorkBuddy')
import gen_v2_lib as G
import openpyxl

EXCEL = r'D:\WorkBuddy\2026年10月西北服务产品.xlsx'
CUT = datetime.date(2026, 10, 15)
SVC = datetime.date(2026, 9, 28)

WATCH = ['于立良', '冯灵军', '张凯祥', '薛雷', '马锋', '汤效龙', '黄钢', '贾龙龙']
CODES = {'JS000005154': '于立良', 'JS000011222': '冯灵军', 'JS000013384': '张凯祥',
         'JS000020413': '薛雷', 'JS000008694': '马锋', 'JS000011563': '汤效龙',
         'JS000013397': '黄钢', 'JS000014319': '贾龙龙'}

wb = openpyxl.load_workbook(EXCEL, read_only=True, data_only=True)

# ---------- 自购配件表 ----------
ws = wb['工单自购配件明细']
hdr = list(ws.iter_rows(min_row=1, max_row=3, values_only=True))
print('=== 工单自购配件明细 前3行 ===')
for i, h in enumerate(hdr):
    print(' row%d: %s' % (i, [str(x)[:14] if x is not None else '' for x in h[:30]]))

Z = G._cols(ws, 1, G.ZG_HEADERS, tag='自购配件表')
print('\n解析列号:', Z)

hit = collections.Counter()
watch_rows = []
for r in G._rows(ws, 1):
    if not r or len(r) <= Z['date']:
        continue
    d = G._date(r[Z['date']])
    if d is None or d > CUT or d < SVC:
        continue
    tech, code = G._s(r[Z['tech']]), G._s(r[Z['tcode']])
    nm = G._s(r[Z['name']])
    if tech in WATCH or code in CODES:
        watch_rows.append((d, tech, code, nm, G._num(r[Z['qty']])))
    hit[tech] += 1

print('\n=== 自购表窗口内 关注技师的命中行 ===')
for x in sorted(watch_rows):
    print('  %s %s(%s) %s x%.0f' % x)
print('  关注技师命中行数 =', len(watch_rows))

# 全表（不过滤日期）搜这些技师
allrows = collections.Counter()
for r in G._rows(ws, 1):
    if not r or len(r) <= Z['date']:
        continue
    tech, code = G._s(r[Z['tech']]), G._s(r[Z['tcode']])
    if tech in WATCH or code in CODES:
        d = G._date(r[Z['date']])
        allrows[(tech, code, '窗口内' if (d and SVC <= d <= CUT) else ('窗口外:%s' % d))] += 1
print('\n=== 自购表全表范围（含窗口外）关注技师 ===')
for k, v in sorted(allrows.items(), key=lambda x: str(x[0])):
    print('  %s  %d 行' % (k, v))

# ---------- CSM配件 ----------
ws2 = wb['CSM配件']
PJ_SPEC = {"office": ("办事处", G.PJ_OFFICE), "net": ("服务网点", G.PJ_NET),
           "tech": ("服务工程师", G.PJ_TECH), "tcode": ("服务工程师编码", G.PJ_TCODE),
           "class": ("销售分类", G.PJ_CLASS), "name": ("配件名称", G.PJ_NAME),
           "code": ("配件编码", G.PJ_CODE), "qty": ("数量", G.PJ_QTY),
           "amt": ("上缴金额", G.PJ_AMT), "date": ("录入完成时间", G.PJ_DATE)}
P = G._cols(ws2, 2, PJ_SPEC, tag='CSM配件')
pj_watch = collections.Counter()
pj_win = 0
for r in G._rows(ws2, 2):
    if not r or len(r) <= P['date']:
        continue
    d = G._date(r[P['date']])
    if d is None or d > CUT or d < SVC:
        continue
    pj_win += 1
    tech, code = G._s(r[P['tech']]), G._s(r[P['tcode']])
    if tech in WATCH or code in CODES:
        pj_watch[(tech, code, G._s(r[P['name']]))] += G._num(r[P['qty']])
print('\n=== CSM配件 窗口内行数 = %d；关注技师命中 ===' % pj_win)
for k, v in sorted(pj_watch.items()):
    print('  %s  %.0f' % (k, v))
wb.close()
