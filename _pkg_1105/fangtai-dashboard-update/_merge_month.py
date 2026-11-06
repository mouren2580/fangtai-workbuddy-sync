# -*- coding: utf-8 -*-
"""把「指定服务月」的四块面板中间文件合并进 _monthly_four.json，并重算 year 累计。

与 _merge_9.py 等价，但月份参数化（_merge_9.py 写死 2026-09）。
用法:
  python _merge_month.py --month 2026-10
前置:
  _mf_<月>_wo.json / _mf_<月>_va.json / _mf_<月>_vp.json / _mf_<月>_bc.json 已生成
"""
import json, os, sys, argparse

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import _gen_monthly_four as G

ap = argparse.ArgumentParser()
ap.add_argument("--month", required=True, help="服务月键，如 2026-10")
a = ap.parse_args()
mk = a.month

mf = json.load(open('_monthly_four.json', encoding='utf-8'))
print("现有 months:", sorted(mf['months'].keys()))
print("原 year 工单=%d 防火阀占比=%.2f%% 增值产品=%d 增值配件当月=%.0f 大保养占比=%.2f%%" % (
    mf['year']['workorder']['meta']['总工单数'], mf['year']['workorder']['meta']['防火阀占比'],
    mf['year']['valueadded']['meta']['全区总计数'], mf['year']['valueparts']['meta']['全区总计数_当月'],
    mf['year']['bigcare']['meta']['全区占比']))

blk = {
    'workorder': json.load(open('_mf_%s_wo.json' % mk, encoding='utf-8')),
    'valueadded': json.load(open('_mf_%s_va.json' % mk, encoding='utf-8')),
    'valueparts': json.load(open('_mf_%s_vp.json' % mk, encoding='utf-8')),
    'bigcare': json.load(open('_mf_%s_bc.json' % mk, encoding='utf-8')),
}
for k in ('workorder', 'valueadded', 'valueparts', 'bigcare'):
    assert isinstance(blk[k], dict), k
    print("  %s keys=%s" % (k, sorted(blk[k].keys())))
print("  %s 工单=%s 大保养=%s" % (mk, blk['workorder']['meta'].get('总工单数'),
                               blk['bigcare']['meta'].get('全区占比')))
mf['months'][mk] = blk

months_list_wo = [mf['months'][m]['workorder'] for m in sorted(mf['months'])]
months_list_va = [mf['months'][m]['valueadded'] for m in sorted(mf['months'])]
months_list_vp = [mf['months'][m]['valueparts'] for m in sorted(mf['months'])]
months_list_bc = [mf['months'][m]['bigcare'] for m in sorted(mf['months'])]
mf['year'] = {
    'workorder': G.merge_wo(months_list_wo),
    'valueadded': G.merge_va(months_list_va),
    'valueparts': G.merge_vp(months_list_vp),
    'bigcare': G.merge_bc(months_list_bc),
}

# year 块的 meta 文案在 merge_* 里是写死的（1-8 月口径），按实际月份动态修正
ms = sorted(mf['months'].keys())
y0, y1 = ms[0].split('-')[0], ms[-1].split('-')[0]
rng = "%s ~ %s" % (ms[0], ms[-1])
try:   # 真实统计截止日以月份包 D 块为准（≠ 周期结束日：服务月首周时两者不同）
    _md = json.load(open('_month_%s.json' % mk, encoding='utf-8'))
    cut = _md['D']['meta'].get('统计截止日') or blk['workorder']['meta'].get('截止') or mk
except Exception:
    cut = blk['workorder']['meta'].get('截止') or mk
print("  year 元数据: 区间=%s  统计截止日=%s" % (rng, cut))
yfix = {
    'workorder': ('工单表+CSM配件表(累计 %s)' % rng, rng, cut),
    'valueadded': (cut, "%s年" % y1, rng, ),
    'valueparts': (cut, rng),
    # ⚠️ bigcare 的 meta 文案原先是 _gen_monthly_four 里写死的「2026-08-26 / 2025-12-28~2026-08-26」，
    #    数值是活的、文案是死的（年度视图会显示「截止 08-26」）。2026-11-06 补上动态修正。
    'bigcare': (cut, rng),
}
m = mf['year']['bigcare']['meta']
m['截止日'] = yfix['bigcare'][0]; m['周期'] = yfix['bigcare'][1]
m = mf['year']['workorder']['meta']
m['来源'] = yfix['workorder'][0]; m['周期'] = yfix['workorder'][1]; m['截止'] = yfix['workorder'][2]
m = mf['year']['valueadded']['meta']
m['截止日'] = yfix['valueadded'][0]; m['月份'] = yfix['valueadded'][1]; m['服务月'] = yfix['valueadded'][2]
m['口径'] = m['口径'].split('全年')[0] + '累计=%s。' % rng
m = mf['year']['valueparts']['meta']
m['截止日'] = yfix['valueparts'][0]; m['服务月'] = yfix['valueparts'][1]
m['口径'] = m['口径'].split('全年')[0] + '累计=%s。' % rng
m['时间窗口'] = {'昨日': cut, '近3日': '~' + cut, '近7日': '~' + cut, '当月': rng}

json.dump(mf, open('_monthly_four.json', 'w', encoding='utf-8'), ensure_ascii=False, separators=(',', ':'))
print("合并后 months:", sorted(mf['months'].keys()))
print("新 year 工单=%d 防火阀占比=%.2f%% 增值产品=%d 增值配件当月=%.0f 大保养占比=%.2f%%" % (
    mf['year']['workorder']['meta']['总工单数'],
    mf['year']['workorder']['meta']['防火阀占比'],
    mf['year']['valueadded']['meta']['全区总计数'], mf['year']['valueparts']['meta']['全区总计数_当月'],
    mf['year']['bigcare']['meta']['全区占比']))
print("[OK] 写出 _monthly_four.json")
