# -*- coding: utf-8 -*-
"""大保养占比 数据生成器（支持任意服务月周期）。

源表：CSM服务项目（表头第2行，数据第3行起）。
列（0-based）：办事处=1, 服务网点=3, 服务工程师=5, 服务工程师编号=6,
     录入完成时间=7, 服务收费项目=13, 销售分类=16。

口径：筛 销售分类=清洗保养 →
     分母 = 剔除『保养超范围收费』『灶具保养』；
     分子 = 『油烟机大保养』+『油烟机大保养升级包』（只计单子数量，与金额无关）。
     占比 = 分子 ÷ 分母。

周期过滤：录入完成时间 落在 [--ps, --pe] 内才计入（铁律服务月边界）。
"""
import argparse, json
from collections import defaultdict
import openpyxl
from datetime import datetime, date

EXCEL = "_新8月数据.xlsx"
OUT = "_bigcare_data.json"
SC, FEE, OFF, NET, ENG, ENGID, LCT = 16, 13, 1, 3, 5, 6, 7
EXCLUDE = {'保养超范围收费', '灶具保养'}
BIG = {'油烟机大保养', '油烟机大保养升级包'}

def to_date(v):
    if isinstance(v, datetime): return v.date()
    if isinstance(v, date): return v
    return None

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--excel", default=EXCEL)
    ap.add_argument("--ps", required=True, help="服务月起始 YYYY-MM-DD")
    ap.add_argument("--pe", required=True, help="服务月结束 YYYY-MM-DD")
    ap.add_argument("--out", default=OUT)
    ap.add_argument("--cut", help="统计截止日 YYYY-MM-DD（仅决定 meta['截止日'] 展示值；不给则取服务月结束日）")
    a = ap.parse_args()
    ps = datetime.strptime(a.ps, "%Y-%m-%d").date()
    pe = datetime.strptime(a.pe, "%Y-%m-%d").date()
    cut = datetime.strptime(a.cut, "%Y-%m-%d").date() if a.cut else pe

    wb = openpyxl.load_workbook(a.excel, read_only=True, data_only=True)
    ws = [w for w in wb.worksheets if w.title == 'CSM服务项目'][0]

    clean = []
    for row in ws.iter_rows(min_row=3, values_only=True):
        if len(row) <= SC or row[SC] != '清洗保养':
            continue
        d = to_date(row[LCT])
        if d is None or not (ps <= d <= pe):
            continue
        clean.append(row)

    denom = [r for r in clean if r[FEE] not in EXCLUDE]

    def agg(idx):
        big = defaultdict(int); den = defaultdict(int)
        for r in denom:
            k = r[idx] if r[idx] else '(空)'; den[k] += 1
        for r in clean:
            if r[FEE] in BIG:
                k = r[idx] if r[idx] else '(空)'; big[k] += 1
        return big, den

    big_o, den_o = agg(OFF)
    big_n, den_n = agg(NET)
    meta_n = {}
    for r in denom:
        n = r[NET] if r[NET] else '(空)'
        meta_n.setdefault(n, r[OFF])

    big_e = defaultdict(int); den_e = defaultdict(int); meta_e = {}
    for r in denom:
        name = r[ENG] if r[ENG] else '(空)'; eid = r[ENGID] if r[ENGID] else ''
        k = (name, eid); den_e[k] += 1
        meta_e.setdefault(k, {'office': r[OFF], 'net': r[NET]})
    for r in clean:
        if r[FEE] in BIG:
            name = r[ENG] if r[ENG] else '(空)'; eid = r[ENGID] if r[ENGID] else ''
            k = (name, eid); big_e[k] += 1

    office = []
    for o in sorted(den_o, key=lambda k: -den_o[k]):
        b = big_o.get(o, 0); d = den_o[o]
        office.append({'name': o, 'big': b, 'den': d, 'pct': round(b / d * 100, 2) if d else 0.0})
    tb = sum(big_o.values()); td = sum(den_o.values())
    office.append({'name': '[合计]全区', 'big': tb, 'den': td, 'pct': round(tb / td * 100, 2) if td else 0.0})

    net = []
    for n in sorted(den_n, key=lambda k: (-big_n.get(k, 0), -den_n[k])):
        b = big_n.get(n, 0); d = den_n[n]
        net.append({'name': n, 'office': meta_n.get(n, ''), 'big': b, 'den': d,
                    'pct': round(b / d * 100, 2) if d else 0.0})
    tb_n = sum(big_n.values()); td_n = sum(den_n.values())
    net.append({'name': '[合计]全部网点', 'office': '—', 'big': tb_n, 'den': td_n,
                'pct': round(tb_n / td_n * 100, 2) if td_n else 0.0})

    eng = []
    for k in sorted(den_e, key=lambda k: (-big_e.get(k, 0), -den_e[k])):
        b = big_e.get(k, 0); d = den_e[k]; name, eid = k
        eng.append({'name': name, 'id': eid, 'office': meta_e[k]['office'], 'net': meta_e[k]['net'],
                    'big': b, 'den': d, 'pct': round(b / d * 100, 2) if d else 0.0})
    tb_e = sum(big_e.values()); td_e = sum(den_e.values())
    eng.append({'name': '[合计]全部工程师', 'id': '—', 'office': '—', 'net': '—', 'big': tb_e, 'den': td_e,
                'pct': round(tb_e / td_e * 100, 2) if td_e else 0.0})

    data = {
        'meta': {
            '截止日': cut.isoformat(), '来源表': 'CSM服务项目',
            '周期': '%s~%s' % (ps.isoformat(), pe.isoformat()),
            '清洗保养总单': len(clean), '剔除项': list(EXCLUDE), '大保养项': list(BIG),
            '分母': len(denom), '大保养总数': tb, '全区占比': round(tb / td * 100, 2) if td else 0.0,
        },
        'office': office, 'net': net, 'eng': eng,
    }
    with open(a.out, 'w', encoding='utf-8') as f:
        json.dump(data, f, ensure_ascii=False, indent=0)
    print("大保养占比: 周期%s~%s 清洗保养总单=%d 分母=%d 大保养总数=%d 全区占比=%.2f%%"
          % (ps.isoformat(), pe.isoformat(), len(clean), len(denom), tb, data['meta']['全区占比']))
    print("  office行:%d net行:%d eng行:%d" % (len(office), len(net), len(eng)))

if __name__ == "__main__":
    main()
