# -*- coding: utf-8 -*-
"""重生成 EXTEND_TIME_DATA 的 months 结构（仅 2026-09 块），保留 2026-08 块原样。
源: CSM服务项目 表, 筛选 销售分类=延保服务, 按 录入完成时间 落在 服务月 8.28~cutoff 内拆分到每日,
年限由 服务收费项目(一年/两年/三年) 决定 yr。
输出 _extend_time.json（覆盖）。"""
import openpyxl, json, sys, datetime, argparse
from collections import Counter, defaultdict
sys.stdout.reconfigure(encoding='utf-8')

# ⚠️ 截止日必须显式传入（或服务月起点），禁止硬编码——历史上多次因上次截止日写死而漏掉新增日
ap = argparse.ArgumentParser()
ap.add_argument('--excel', default=r"E:/方太2026/2026年增值创新/服务月报/WorkBuddy/2026年9月西北服务产品.xlsx")
ap.add_argument('--svc-lo', default='2026-08-28', help='服务月起始日')
ap.add_argument('--cut', required=True, help='统计截止日 (= 更新当日 - 1)，如 2026-09-02')
ap.add_argument('--month', default='2026-09', help='写入 months 的哪个键')
a = ap.parse_args()

EXCEL = a.excel
SVC_LO = datetime.date(*[int(x) for x in a.svc_lo.split('-')])
CUT    = datetime.date(*[int(x) for x in a.cut.split('-')])
MONTH_KEY = a.month

def yr_of(s):
    s = str(s)
    if '一年' in s or '1年' in s: return 'y1'
    if '两年' in s or '二年' in s or '2年' in s: return 'y2'
    if '三年' in s or '3年' in s: return 'y3'
    return None

def num(v):
    try: return round(float(v), 2)
    except Exception: return 0.0

wb = openpyxl.load_workbook(EXCEL, read_only=True, data_only=True)
ws = wb['CSM服务项目']
# 表头在第2行
header = None
for ri in range(1, 5):
    r = next(ws.iter_rows(min_row=ri, max_row=ri, values_only=True))
    if any(str(x).strip() == '销售分类' for x in r if x):
        header = r; break
def ci(name):
    for i, h in enumerate(header):
        if h and str(h).strip() == name: return i
    return None
c_cat = ci('销售分类'); c_item = ci('服务收费项目'); c_tech = ci('服务工程师')
c_off = ci('办事处'); c_pay = ci('上缴金额'); c_in = ci('录入完成时间')

recs = []
per_day = Counter()
by_day_yr = defaultdict(lambda: Counter())
for r in ws.iter_rows(min_row=3, values_only=True):
    if c_cat is None or r[c_cat] is None: continue
    if str(r[c_cat]).strip() != '延保服务': continue
    d = r[c_in]
    if not isinstance(d, datetime.datetime): continue
    dd = d.date()
    if not (SVC_LO <= dd <= CUT): continue
    yk = yr_of(r[c_item])
    if yk is None: continue
    name = str(r[c_tech]).strip() if r[c_tech] else '（未知）'
    off = str(r[c_off]).strip() if r[c_off] else ''
    amt = num(r[c_pay])
    recs.append({'d': dd.isoformat(), 'o': off, 't': name, 'c': '', 'yr': yk, 'amt': amt})
    per_day[dd.isoformat()] += 1
    by_day_yr[dd.isoformat()][yk] += 1

# 排序
recs.sort(key=lambda x: (x['d'], x['o'], x['t'], x['yr']))
print("2026-09 延保记录数:", len(recs))
for d in sorted(per_day):
    print("  %s: %d 单  %s" % (d, per_day[d], dict(by_day_yr[d])))

# 载入旧 _extend_time.json，保留其它月份（如 2026-08）
old = json.load(open('_extend_time.json', encoding='utf-8'))
out = {'months': dict(old.get('months', {}))}
out['months'][MONTH_KEY] = recs
print("保留其它月份:", {k: len(v) for k, v in out['months'].items() if k != MONTH_KEY})
json.dump(out, open('_extend_time.json', 'w', encoding='utf-8'), ensure_ascii=False, separators=(',', ':'))
print("[OK] 写出 _extend_time.json (%s=%d)" % (MONTH_KEY, len(recs)))
print("[校验] 合计 %d 单 / %.2f 元  (应与 D 块延保一致)" % (len(recs), sum(x['amt'] for x in recs)))
