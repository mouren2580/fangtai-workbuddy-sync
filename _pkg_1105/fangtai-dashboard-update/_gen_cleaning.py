# -*- coding: utf-8 -*-
"""清洗保养 原始记录生成器（按服务月分组）。

源表：CSM服务项目（表头第2行，数据第3行起）。
筛选：销售分类(Q列/16) = 清洗保养。
每行 = 1 单：数量 = 行数；上缴金额 = Y列(24)。
输出 _cleaning_data.json：
  { months: { "2026-09": [ {o,n,e,ei,d,it,amt}, ... ], ... },
    meta: {...} }
时间窗口（昨日/近3/近7/当月）与维度聚合均在看板 JS 端按统计截止日实时计算，
故此处只输出原始记录，按月分组。
"""
import openpyxl, json, argparse, sys, os
from datetime import datetime, date
from collections import defaultdict

EXCEL = "E:/方太2026/2026年增值创新/服务月报/WorkBuddy/2026年9月西北服务产品.xlsx"
OUT = "_cleaning_data.json"

# 主保养白名单：免费/收费统计口径仅含这四项；其余（保养超范围收费、灶具保养等）不计。
# 与「大保养占比」分母剔除口径一致，且符合业务上「免费/收费只看主保养四项」的口径。
MAIN_ITEMS = {'油烟机小保养', '油烟机大保养', '集成灶小保养', '集成灶大保养'}

def service_month_of(d):
    """自定义服务月：上月28日 ~ 当月27日。d.day>=28 -> 归属下月服务月。"""
    if d.day >= 28:
        m = d.month + 1 if d.month < 12 else 1
        y = d.year if d.month < 12 else d.year + 1
        return "%04d-%02d" % (y, m)
    return "%04d-%02d" % (d.year, d.month)

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--excel", default=EXCEL)
    ap.add_argument("--out", default=OUT)
    a = ap.parse_args()
    wb = openpyxl.load_workbook(a.excel, read_only=True, data_only=True)
    ws = [w for w in wb.worksheets if w.title == 'CSM服务项目'][0]
    header = [c.value for c in ws[2]]
    def ci(name):
        for i, h in enumerate(header):
            if h and str(h).strip() == name:
                return i
        return None
    c_cat = ci('销售分类'); c_item = ci('服务收费项目'); c_off = ci('办事处'); c_net = ci('服务网点')
    c_eng = ci('服务工程师'); c_eid = ci('服务工程师编号'); c_in = ci('录入完成时间'); c_amt = ci('上缴金额')
    missing = [n for n, v in {'销售分类':c_cat,'服务收费项目':c_item,'办事处':c_off,'服务网点':c_net,
                              '服务工程师':c_eng,'录入完成时间':c_in,'上缴金额':c_amt}.items() if v is None]
    if missing:
        print("ERR missing columns:", missing); sys.exit(1)

    months = defaultdict(list)
    total = 0
    for r in ws.iter_rows(min_row=3, values_only=True):
        if len(r) <= c_cat or r[c_cat] is None:
            continue
        if str(r[c_cat]).strip() != '清洗保养':
            continue
        if str(r[c_item]).strip() not in MAIN_ITEMS:
            continue
        d = r[c_in]
        if not isinstance(d, datetime):
            continue
        dd = d.date()
        sm = service_month_of(dd)
        try:
            amt = round(float(r[c_amt]), 2) if r[c_amt] is not None else 0.0
        except Exception:
            amt = 0.0
        months[sm].append({
            'o': str(r[c_off]).strip() if r[c_off] else '(空)',
            'n': str(r[c_net]).strip() if r[c_net] else '(空)',
            'e': str(r[c_eng]).strip() if r[c_eng] else '(空)',
            'ei': str(r[c_eid]).strip() if r[c_eid] else '',
            'd': dd.isoformat(),
            'it': str(r[c_item]).strip() if r[c_item] else '(空)',
            'amt': amt,
        })
        total += 1

    # 保留历史服务月（供前端跨月窗口：月初近3/近7需回看上一服务月末）
    if os.path.exists(a.out):
        try:
            old = json.load(open(a.out, encoding='utf-8'))
            for k, v in (old.get('months') or {}).items():
                if k not in months:
                    months[k] = v
        except Exception as e:
            print("  [warn] 读取既有 %s 失败：%s" % (a.out, e))
    out = {
        'months': dict(months),
        'meta': {
            '来源表': 'CSM服务项目',
            '筛选': '销售分类=清洗保养',
            '生成': datetime.now().strftime('%Y-%m-%d %H:%M'),
            '总记录': total,
            '主项': ['油烟机小保养', '油烟机大保养', '集成灶小保养', '集成灶大保养'],
        },
    }
    json.dump(out, open(a.out, 'w', encoding='utf-8'), ensure_ascii=False, separators=(',', ':'))
    for k in sorted(months):
        print("  %s: %d 单" % (k, len(months[k])))
    print("OK 写出 %s（共 %d 单）" % (a.out, total))

if __name__ == "__main__":
    main()
