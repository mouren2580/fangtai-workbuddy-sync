# -*- coding: utf-8 -*-
"""精确定位 valueparts 在同一截止日(10-15)下的差异：工程师级对比"""
import json, sys, datetime, collections
sys.path.insert(0, r'D:\WorkBuddy')
import gen_v2_lib as G

EXCEL = r'D:\WorkBuddy\2026年10月西北服务产品.xlsx'
PREV  = r'D:\WorkBuddy\_prev.html'
CUT   = datetime.date(2026, 10, 15)
SVC   = datetime.date(2026, 9, 28)

def grab_js_obj(src, name):
    import re
    m = re.search(r'(?:var|let|const)\s+' + name + r'\s*=', src)
    if not m:
        return None
    i = m.end()
    k = src.find('{', i)
    depth = 0
    for p in range(k, len(src)):
        ch = src[p]
        if ch == '{':
            depth += 1
        elif ch == '}':
            depth -= 1
            if depth == 0:
                return json.loads(src[k:p+1])
    return None

src = open(PREV, encoding='utf-8').read()
mf = grab_js_obj(src, 'MONTHLY_FOUR')
print('MONTHLY_FOUR months:', sorted(mf.get('months', {}).keys()) if mf else None)
old = mf['months']['2026-10']['valueparts']

new = G.build_valueparts(EXCEL, CUT)

print('\n=== meta 关键项 ===')
for k in ('全区总计数_昨日', '全区总计数_近3日', '全区总计数_近7日', '全区总计数_当月'):
    print('  %-16s 旧 %.0f   新 %.0f   差 %+.0f' % (k, old['meta'].get(k, 0), new['meta'].get(k, 0),
                                                    new['meta'].get(k, 0) - old['meta'].get(k, 0)))

def flat(vp):
    """返回 {产品: {技师键: month}}"""
    d = {}
    for p in vp['products']:
        d[p['name']] = {e['name']: e['month'] for e in p['engineers']}
    return d

O, N = flat(old), flat(new)
print('\n=== 产品级（qty_month）===')
print('  %-24s %10s %10s %8s' % ('产品', '旧', '新', '差'))
for nm in sorted(set(list(O) + list(N)), key=lambda x: -(N.get(x, {}).get('__x', 0) if False else 0)):
    pass
oldq = {p['name']: p['qty_month'] for p in old['products']}
newq = {p['name']: p['qty_month'] for p in new['products']}
for nm in sorted(set(list(oldq) + list(newq)), key=lambda x: -max(oldq.get(x, 0), newq.get(x, 0))):
    o, n = oldq.get(nm, 0), newq.get(nm, 0)
    if abs(o - n) > 0.5:
        print('  %-24s %10.0f %10.0f %8.0f ❌' % (nm, o, n, n - o))
print('  （未列出的产品完全一致）')

print('\n=== 工程师级差异（仅列 month 不同的）===')
cnt = 0
for nm in sorted(set(list(O) + list(N))):
    o, n = O.get(nm, {}), N.get(nm, {})
    for who in sorted(set(list(o) + list(n))):
        a, b = o.get(who, 0), n.get(who, 0)
        if abs(a - b) > 0.5:
            print('  [%s] %-30s 旧 %6.0f → 新 %6.0f  (%+.0f)' % (nm, who, a, b, b - a))
            cnt += 1
print('  共 %d 处工程师级差异' % cnt)

print('\n=== 仅在一侧出现的技师 ===')
for nm in sorted(set(list(O) + list(N))):
    o, n = O.get(nm, {}), N.get(nm, {})
    only_o = set(o) - set(n); only_n = set(n) - set(o)
    if only_o or only_n:
        print('  [%s]' % nm)
        for w in sorted(only_o): print('      仅旧: %s (%.0f)' % (w, o[w]))
        for w in sorted(only_n): print('      仅新: %s (%.0f)' % (w, n[w]))
