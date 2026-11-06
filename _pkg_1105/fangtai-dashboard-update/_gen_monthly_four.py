# -*- coding: utf-8 -*-
"""生成 1-7 月(本次导出) + 8 月(现有基准) + 全年(1-8 累计) 的四块数据，
   输出 _monthly_four.json，供 dashboard.html 注入后做月份切换。

四块：工单量(WORKORDER) / 防火阀占比(在 WORKORDER.meta) / 增值产品(VALUE_ADDED)
      / 增值配件明细(VALUE_PARTS) / 大保养占比(BIG_CARE)。

周期(铁律服务月边界)：
  1月 25.12.28~26.1.27 | 2月 26.1.28~26.2.25 | 3月 26.2.26~26.3.27
  4月 26.3.28~26.4.27  | 5月 26.4.28~26.5.27  | 6月 26.5.28~26.6.27
  7月 26.6.28~26.7.27  | 8月 26.7.28~26.8.26(现有基准截止日)
全年 = 1-8 逐月合并。
"""
import subprocess, json, os, sys
from collections import defaultdict

PY = sys.executable  # 用当前解释器，避免硬编码 python 路径（家里电脑路径不同也能跑）
HERE = os.path.dirname(os.path.abspath(__file__))
NEW_XLSX = r"E:/方太2026/2026年增值创新/服务月报/WorkBuddy/26年1-7月数据.xlsx"
OLD_XLSX = os.path.join(HERE, "_新8月数据.xlsx")
TEMPLATE = os.path.join(HERE, "增值产品填写模板.xlsx")

PERIODS = {
    '2026-01': ('2025-12-28', '2026-01-27'),
    '2026-02': ('2026-01-28', '2026-02-25'),
    '2026-03': ('2026-02-26', '2026-03-27'),
    '2026-04': ('2026-03-28', '2026-04-27'),
    '2026-05': ('2026-04-28', '2026-05-27'),
    '2026-06': ('2026-05-28', '2026-06-27'),
    '2026-07': ('2026-06-28', '2026-07-27'),
    '2026-08': ('2026-07-28', '2026-08-26'),   # 现有基准截止日
}
# 8 月用旧文件；1-7 用新导出文件
XLSX_FOR = {m: (OLD_XLSX if m == '2026-08' else NEW_XLSX) for m in PERIODS}

def run(cmd):
    r = subprocess.run(cmd, cwd=HERE, capture_output=True, text=True, encoding='utf-8')
    if r.returncode != 0:
        sys.stderr.write("CMD FAIL: %s\n%s\n%s\n" % (" ".join(cmd), r.stdout, r.stderr))
        raise SystemExit("生成失败: " + " ".join(cmd))
    return r.stdout

def gen_month(m):
    lo, hi = PERIODS[m]
    xlsx = XLSX_FOR[m]
    wo = os.path.join(HERE, "_mf_%s_wo.json" % m)
    va = os.path.join(HERE, "_mf_%s_va.json" % m)
    vp = os.path.join(HERE, "_mf_%s_vp.json" % m)
    bc = os.path.join(HERE, "_mf_%s_bc.json" % m)
    run([PY, os.path.join(HERE, "_gen_workorder.py"), xlsx, "--ps", lo, "--pe", hi, "--out", wo])
    run([PY, os.path.join(HERE, "_gen_valueadded.py"), "--lo", lo, "--hi", hi,
         "--excel", xlsx, "--template", TEMPLATE, "--out", va])
    run([PY, os.path.join(HERE, "_gen_valueparts.py"), "--svc-lo", lo, "--cut", hi,
         "--excel", xlsx, "--out", vp])
    run([PY, os.path.join(HERE, "_gen_bigcare.py"), "--excel", xlsx, "--ps", lo, "--pe", hi, "--out", bc])
    return {
        "workorder": json.load(open(wo, encoding='utf-8')),
        "valueadded": json.load(open(va, encoding='utf-8')),
        "valueparts": json.load(open(vp, encoding='utf-8')),
        "bigcare": json.load(open(bc, encoding='utf-8')),
    }

# ---------------- 合并工具 ----------------
def merge_wo(months):
    offices = {}
    pgroup = defaultdict(int); wotype = defaultdict(int); sitem = defaultdict(int)
    total = 0; total_zh = 0; total_yj = 0
    all_nets = defaultdict(set); all_techs = set()
    for m in months:
        for o in m['office']:
            if o['name'].startswith('[合计]'): continue
            on = o['name']
            if on not in offices: offices[on] = {'name': on, 'cnt': 0, 'zh': 0, 'yj': 0, 'nets': {}}
            oo = offices[on]
            oo['cnt'] += o['cnt']; oo['zh'] += o['zh']; oo['yj'] += o['yj']
            for n in o['nets']:
                nn = n['name']
                if nn not in oo['nets']: oo['nets'][nn] = {'name': nn, 'cnt': 0, 'zh': 0, 'yj': 0, 'techs': {}}
                no = oo['nets'][nn]
                no['cnt'] += n['cnt']; no['zh'] += n['zh']; no['yj'] += n['yj']
                all_nets[on].add(nn)
                for t in n['techs']:
                    tk = (t['name'], t['code'])
                    all_techs.add(tk)
                    if tk not in no['techs']: no['techs'][tk] = {'name': t['name'], 'code': t['code'], 'cnt': 0, 'zh': 0, 'yj': 0}
                    nt = no['techs'][tk]
                    nt['cnt'] += t['cnt']; nt['zh'] += t['zh']; nt['yj'] += t['yj']
        for d in (m.get('pgroup') or []): pgroup[d['name']] += d['cnt']
        for d in (m.get('wotype') or []): wotype[d['name']] += d['cnt']
        for d in (m.get('sitem') or []): sitem[d['name']] += d['cnt']
        total += m['meta']['总工单数']; total_zh += m['meta']['止回阀总数']; total_yj += m['meta']['烟机安装量']
    out_offices = []
    for on, o in offices.items():
        nets = []
        for nn, n in o['nets'].items():
            techs = [{'name': t['name'], 'code': t['code'], 'cnt': t['cnt'], 'zh': t['zh'], 'yj': t['yj']} for t in n['techs'].values()]
            techs.sort(key=lambda x: -x['cnt'])
            nets.append({'name': nn, 'cnt': n['cnt'], 'zh': n['zh'], 'yj': n['yj'], 'techs': techs})
        nets.sort(key=lambda x: -x['cnt'])
        out_offices.append({'name': on, 'cnt': o['cnt'], 'zh': o['zh'], 'yj': o['yj'], 'nets': nets})
    out_offices.sort(key=lambda x: -x['cnt'])
    def rank(d): return sorted(({'name': k, 'cnt': v} for k, v in d.items()), key=lambda x: -x['cnt'])
    meta = {
        '来源': '工单表+CSM配件表(全年累计 1-8月)', '周期': '2025-12-28~2026-08-26',
        '截止': '2026-08-26', '时间口径': '服务完成时间',
        '总工单数': total, '办事处数': len(out_offices),
        '网点数': sum(len(s) for s in all_nets.values()), '工程师数': len(all_techs),
        '维度': '办事处→网点→服务工程师（工单编号去重计数）',
        '产品组数': len(pgroup), '工单类型数': len(wotype), '服务项目数': len(sitem),
        '止回阀总数': total_zh, '烟机安装量': total_yj,
        '防火阀占比': round(total_zh / total_yj * 100, 2) if total_yj else 0.0,
    }
    return {'meta': meta, 'office': out_offices, 'pgroup': rank(pgroup), 'wotype': rank(wotype), 'sitem': rank(sitem)}

def merge_va(months):
    by_name = {}
    for m in months:
        for p in m['products']:
            n = p['name']
            if n not in by_name:
                by_name[n] = {'name': n, 'code': p.get('code'), 'pname': p.get('pname'),
                              'special': p.get('special', False), 'total': 0, 'office': {}, 'net': {}, 'tech': {}}
            t = by_name[n]
            t['total'] += p.get('total', 0)
            for k, v in (p.get('office') or {}).items(): t['office'][k] = t['office'].get(k, 0) + v
            for k, v in (p.get('net') or {}).items(): t['net'][k] = t['net'].get(k, 0) + v
            for k, v in (p.get('tech') or {}).items(): t['tech'][k] = t['tech'].get(k, 0) + v
    products = list(by_name.values())
    products.sort(key=lambda x: x['name'])
    grand = sum(p['total'] for p in products)
    sold = sum(1 for p in products if p['total'] > 0)
    meta = {
        '截止日': '2026-08-26', '月份': '2026全年', '服务月': '2025-12-28 ~ 2026-08-26',
        '来源表': 'CSM配件', '口径': '普通产品按 配件编码=产品编码 计数（服务月内）；止回阀按 销售分类=止回阀 计数（与防火阀占比模块同口径）。全年=1-8月累计。',
        '产品项数': len(products), '有销量项数': sold, '全区总计数': grand, '校验': None,
    }
    return {'meta': meta, 'products': products}

def merge_vp(months):
    by_name = {}
    for m in months:
        for p in m['products']:
            n = p['name']
            if n not in by_name:
                by_name[n] = {'name': n, 'special': p.get('special', False), 'pname': p.get('pname'), 'code': p.get('code'),
                              'qty_yest': 0, 'qty_last3': 0, 'qty_last7': 0, 'qty_month': 0, 'engineers': {}}
            t = by_name[n]
            t['qty_yest'] += p.get('qty_yest', 0); t['qty_last3'] += p.get('qty_last3', 0)
            t['qty_last7'] += p.get('qty_last7', 0); t['qty_month'] += p.get('qty_month', 0)
            for e in p.get('engineers', []):
                ek = e['name']
                if ek not in t['engineers']:
                    t['engineers'][ek] = {'name': ek, 'office': e.get('office', ''), 'yest': 0, 'last3': 0, 'last7': 0, 'month': 0}
                te = t['engineers'][ek]
                te['office'] = e.get('office', '') or te['office']
                te['yest'] += e.get('yest', 0); te['last3'] += e.get('last3', 0); te['last7'] += e.get('last7', 0); te['month'] += e.get('month', 0)
    products = []
    for t in by_name.values():
        eng = sorted(t['engineers'].values(), key=lambda x: (-x['month'], -x['last3'], -x['yest']))
        products.append({'name': t['name'], 'special': t['special'], 'pname': t['pname'], 'code': t['code'],
                         'qty_yest': round(t['qty_yest'], 2), 'qty_last3': round(t['qty_last3'], 2),
                         'qty_last7': round(t['qty_last7'], 2), 'qty_month': round(t['qty_month'], 2), 'engineers': eng})
    products.sort(key=lambda p: (-p['qty_month'], p['name']))
    tot_yest = sum(p['qty_yest'] for p in products); tot_last3 = sum(p['qty_last3'] for p in products)
    tot_last7 = sum(p['qty_last7'] for p in products); tot_month = sum(p['qty_month'] for p in products)
    sold = sum(1 for p in products if p['qty_month'] > 0)
    meta = {
        '截止日': '2026-08-26', '服务月': '2025-12-28 ~ 2026-08-26',
        '口径': '普通产品：CSM配件表(配件名称/配件编码) + 工单自购配件明细(物料名称/物料编码)，数量列求和；止回阀：CSM配件表 销售分类=止回阀。全年=1-8月累计。',
        '来源表': 'CSM配件 + 工单自购配件明细',
        '时间窗口': {'昨日': '2026-08-26', '近3日': '2026-08-24 ~ 2026-08-26', '近7日': '2026-08-20 ~ 2026-08-26', '当月': '2025-12-28 ~ 2026-08-26'},
        '产品项数': len(products), '有销量项数': sold,
        '全区总计数_昨日': round(tot_yest, 2), '全区总计数_近3日': round(tot_last3, 2),
        '全区总计数_近7日': round(tot_last7, 2), '全区总计数_当月': round(tot_month, 2), '校验': None,
    }
    return {'meta': meta, 'products': products}

def merge_bc(months):
    def merge_list(key):
        agg = {}
        for m in months:
            for row in m[key]:
                if row['name'].startswith('[合计]'): continue
                nm = row['name']
                if nm not in agg:
                    agg[nm] = {k: row[k] for k in row if k not in ('big', 'den', 'pct')}
                    agg[nm]['big'] = 0; agg[nm]['den'] = 0
                agg[nm]['big'] += row['big']; agg[nm]['den'] += row['den']
        out = []
        for nm, row in agg.items():
            b = row['big']; d = row['den']
            item = dict(row); item['pct'] = round(b / d * 100, 2) if d else 0.0
            out.append(item)
        out.sort(key=lambda x: -x['big'])
        return out
    office = merge_list('office'); net = merge_list('net'); eng = merge_list('eng')
    tb = sum(o['big'] for o in office); td = sum(o['den'] for o in office)
    office.append({'name': '[合计]全区', 'big': tb, 'den': td, 'pct': round(tb / td * 100, 2) if td else 0.0})
    tb_n = sum(n['big'] for n in net); td_n = sum(n['den'] for n in net)
    net.append({'name': '[合计]全部网点', 'office': '—', 'big': tb_n, 'den': td_n, 'pct': round(tb_n / td_n * 100, 2) if td_n else 0.0})
    tb_e = sum(e['big'] for e in eng); td_e = sum(e['den'] for e in eng)
    eng.append({'name': '[合计]全部工程师', 'id': '—', 'office': '—', 'net': '—', 'big': tb_e, 'den': td_e, 'pct': round(tb_e / td_e * 100, 2) if td_e else 0.0})
    meta = {
        '截止日': '2026-08-26', '来源表': 'CSM服务项目', '周期': '2025-12-28~2026-08-26',
        '清洗保养总单': sum(m['meta']['清洗保养总单'] for m in months),
        '剔除项': ['保养超范围收费', '灶具保养'], '大保养项': ['油烟机大保养', '油烟机大保养升级包'],
        '分母': td, '大保养总数': tb, '全区占比': round(tb / td * 100, 2) if td else 0.0,
    }
    return {'meta': meta, 'office': office, 'net': net, 'eng': eng}

def main():
    months = {}
    for m in PERIODS:
        print("=== 生成 %s ===" % m)
        months[m] = gen_month(m)
    # 全年(1-8)
    year = {
        'workorder': merge_wo([months[m]['workorder'] for m in PERIODS]),
        'valueadded': merge_va([months[m]['valueadded'] for m in PERIODS]),
        'valueparts': merge_vp([months[m]['valueparts'] for m in PERIODS]),
        'bigcare': merge_bc([months[m]['bigcare'] for m in PERIODS]),
    }
    out = {'months': months, 'year': year}
    with open(os.path.join(HERE, '_monthly_four.json'), 'w', encoding='utf-8') as f:
        json.dump(out, f, ensure_ascii=False, separators=(',', ':'))
    # 摘要
    print("\n========== 摘要 ==========")
    for m in PERIODS:
        d = months[m]
        print("%s: 工单=%d 防火阀占比=%.2f%% 增值产品=%d件 增值配件当月=%.0f 大保养占比=%.2f%%"
              % (m, d['workorder']['meta']['总工单数'], d['workorder']['meta']['防火阀占比'],
                 d['valueadded']['meta']['全区总计数'], d['valueparts']['meta']['全区总计数_当月'],
                 d['bigcare']['meta']['全区占比']))
    print("%s: 工单=%d 防火阀占比=%.2f%% 增值产品=%d件 增值配件当月=%.0f 大保养占比=%.2f%%"
          % ("全年", year['workorder']['meta']['总工单数'], year['workorder']['meta']['防火阀占比'],
             year['valueadded']['meta']['全区总计数'], year['valueparts']['meta']['全区总计数_当月'],
             year['bigcare']['meta']['全区占比']))
    sz = os.path.getsize(os.path.join(HERE, '_monthly_four.json'))
    print("写出 _monthly_four.json (%.2f MB)" % (sz / 1048576))

if __name__ == '__main__':
    main()
