# -*- coding: utf-8 -*-
"""通用月份数据生成器：读任意月份月报 Excel（4 张源表），产出该月五大块数据 JSON。
   用法: python _gen_month.py <月报.xlsx> --date YYYY-MM-DD --out _month_YYYY-MM.json
   产出: {"D":DASHBOARD_DATA,"T":TECH_DATA,"B":BIG_CARE_DATA,"E":EXTEND_WARRANTY_DATA,"W":WEEKLY_DATA}
   复用技能 build_dashboard.build + 本地 _refresh_data.build_tech + bigcare/extend/weekly 聚合逻辑。
"""
import sys, os, json, argparse, datetime, collections
for s in (sys.stdout, sys.stderr):
    try: s.reconfigure(encoding="utf-8", errors="replace")
    except Exception: pass

# 本地优先：脚本同目录若有 build_dashboard.py 就用它（家里电脑没有 WorkBuddy 技能目录也能跑）；
# 找不到再回退到办公室技能目录。
_HERE = os.path.dirname(os.path.abspath(__file__))
SKILL = _HERE if os.path.exists(os.path.join(_HERE, "build_dashboard.py")) \
        else r"C:/Users/ABC/.workbuddy/skills/fangtai-service-achievement-dashboard/scripts"
sys.path.insert(0, SKILL)
from build_dashboard import build, service_period, service_month_of
import openpyxl

ITEMS = ["清洗保养","延保服务","清洁耗材","止回阀","装饰罩","水/气路金属软管",
         "排烟出风管","水气路通用辅材","其他产品线专用","其他水路辅材","其他气路辅材",
         "橱柜类辅材","燃气管快速接头","工具","电路辅材"]

def _fix(s):
    if isinstance(s, str):
        try: return s.encode('latin-1').decode('gbk')
        except Exception: return s
    return s
def _num(v):
    try: return round(float(v), 2)
    except Exception: return 0.0

# ---------- TECH_DATA（复制自 _refresh_data.build_tech） ----------
def build_tech(src):
    wb = openpyxl.load_workbook(src, read_only=True, data_only=True)
    ws = wb["服务产品收入统计"]
    techs, source_split = {}, collections.defaultdict(float)
    full_total = 0.0
    for r in ws.iter_rows(min_row=3, values_only=True):
        if not r or len(r) < 25:
            continue
        amount = _num(r[24])
        full_total += amount
        src_name = _fix(r[8]) if r[8] else ""
        source_split[src_name] += amount
        code = _fix(r[7]) if r[7] else ""
        if not code:
            continue
        t = techs.get(code)
        if t is None:
            t = {"id": code, "name": _fix(r[6]) or "", "office": _fix(r[2]) or "",
                 "center": _fix(r[1]) or "", "nets": set(), "items": [0.0]*15, "amount": 0.0}
            techs[code] = t
        t["amount"] += amount
        for i, ci in enumerate(range(9, 24)):
            t["items"][i] += _num(r[ci])
        if r[3]: t["nets"].add(_fix(r[3]))
        if not t["office"] and r[2]: t["office"] = _fix(r[2])
        if not t["center"] and r[1]: t["center"] = _fix(r[1])
    tech_list, offices = [], set()
    for t in techs.values():
        if t["office"]: offices.add(t["office"])
        tech_list.append({
            "id": t["id"], "name": t["name"], "office": t["office"], "center": t["center"],
            "net": "、".join(sorted(t["nets"])) if t["nets"] else "",
            "cnt": sum(1 for v in t["items"] if v > 0),
            "amount": round(t["amount"], 2),
            "items": {ITEMS[i]: round(v, 2) for i, v in enumerate(t["items"])}})
    tech_list.sort(key=lambda x: x["amount"], reverse=True)
    total = round(sum(t["amount"] for t in tech_list), 2)
    full_total = round(full_total, 2)
    other = round(full_total - total, 2)
    return {"techs": tech_list, "offices": sorted(offices), "items": ITEMS,
            "total": total, "count": len(tech_list),
            "meta": {"口径": "服务产品收入统计·按服务工程师(姓名+编码)汇总『合计』(工单消耗)",
                     "来源": "服务产品收入统计", "截止": "报表全量",
                     "报表总合计": full_total, "其他未归属": other,
                     "工程师工单消耗": total, "网点买断": other}}, dict(source_split)

# ---------- BIG_CARE_DATA ----------
def build_bigcare(wb, cutoff):
    ws = [w for w in wb.worksheets if w.title == 'CSM服务项目'][0]
    SC, FEE, OFF, NET, ENG, ENGID = 16, 13, 1, 3, 5, 6
    clean = [r for r in ws.iter_rows(min_row=3, values_only=True)
             if len(r) > SC and r[SC] == '清洗保养']
    EXCLUDE = {'保养超范围收费', '灶具保养'}
    BIG = {'油烟机大保养', '油烟机大保养升级包'}
    denom = [r for r in clean if r[FEE] not in EXCLUDE]
    def agg(idx):
        big = collections.defaultdict(int); den = collections.defaultdict(int)
        for r in denom:
            k = r[idx] if r[idx] else '(空)'; den[k] += 1
        for r in clean:
            if r[FEE] in BIG:
                k = r[idx] if r[idx] else '(空)'; big[k] += 1
        return big, den
    big_o, den_o = agg(OFF); big_n, den_n = agg(NET)
    meta_n = {}
    for r in denom:
        n = r[NET] if r[NET] else '(空)'; meta_n.setdefault(n, r[OFF])
    big_e = collections.defaultdict(int); den_e = collections.defaultdict(int); meta_e = {}
    for r in denom:
        name = r[ENG] if r[ENG] else '(空)'; eid = r[ENGID] if r[ENGID] else ''
        k = (name, eid); den_e[k] += 1; meta_e.setdefault(k, {'office': r[OFF], 'net': r[NET]})
    for r in clean:
        if r[FEE] in BIG:
            name = r[ENG] if r[ENG] else '(空)'; eid = r[ENGID] if r[ENGID] else ''
            k = (name, eid); big_e[k] += 1
    office = []
    for o in sorted(den_o, key=lambda k: -den_o[k]):
        b = big_o.get(o, 0); d = den_o[o]
        office.append({'name': o, 'big': b, 'den': d, 'pct': round(b/d*100, 2) if d else 0.0})
    tb = sum(big_o.values()); td = sum(den_o.values())
    office.append({'name': '[合计]全区', 'big': tb, 'den': td, 'pct': round(tb/td*100, 2) if td else 0.0})
    net = []
    for n in sorted(den_n, key=lambda k: (-big_n.get(k, 0), -den_n[k])):
        b = big_n.get(n, 0); d = den_n[n]
        net.append({'name': n, 'office': meta_n.get(n, ''), 'big': b, 'den': d,
                    'pct': round(b/d*100, 2) if d else 0.0})
    tb_n = sum(big_n.values()); td_n = sum(den_n.values())
    net.append({'name': '[合计]全部网点', 'office': '—', 'big': tb_n, 'den': td_n,
                'pct': round(tb_n/td_n*100, 2) if td_n else 0.0})
    eng = []
    for k in sorted(den_e, key=lambda k: (-big_e.get(k, 0), -den_e[k])):
        b = big_e.get(k, 0); d = den_e[k]; name, eid = k
        eng.append({'name': name, 'id': eid, 'office': meta_e[k]['office'], 'net': meta_e[k]['net'],
                    'big': b, 'den': d, 'pct': round(b/d*100, 2) if d else 0.0})
    tb_e = sum(big_e.values()); td_e = sum(den_e.values())
    eng.append({'name': '[合计]全部工程师', 'id': '—', 'office': '—', 'net': '—', 'big': tb_e, 'den': td_e,
                'pct': round(tb_e/td_e*100, 2) if td_e else 0.0})
    return {'meta': {'截止日': cutoff, '来源表': 'CSM服务项目',
                     '清洗保养总单': len(clean), '剔除项': list(EXCLUDE), '大保养项': list(BIG),
                     '分母': len(denom), '大保养总数': tb, '全区占比': round(tb/td*100, 2) if td else 0.0},
            'office': office, 'net': net, 'eng': eng}

# ---------- EXTEND_WARRANTY_DATA ----------
def build_extend(wb, cutoff):
    target = next((s for s in wb.sheetnames if ("服务项目" in s) or ("CSM" in s.upper())), None)
    ws = wb[target]
    KW = ["销售分类", "服务收费项目", "服务工程师", "服务工程师编码", "上缴金额", "办事处"]
    hdr_row = None; hdr = None
    for ri in range(1, 7):
        row = list(next(ws.iter_rows(min_row=ri, max_row=ri, values_only=True)))
        flat = [str(x) for x in row if x is not None]
        if any(k in f for f in flat for k in KW):
            hdr_row = ri; hdr = row; break
    idx = {}
    for ci, v in enumerate(hdr):
        if v is None: continue
        s = str(v).strip()
        if s in KW: idx[s] = ci
    c_cat, c_item, c_tech, c_code, c_pay, c_off = (
        idx["销售分类"], idx["服务收费项目"], idx["服务工程师"],
        idx.get("服务工程师编码"), idx["上缴金额"], idx.get("办事处"))
    maxc = max(x for x in (c_cat, c_item, c_tech, c_pay, c_code, c_off) if x is not None)
    def yr_of(s):
        s = str(s)
        if "一年" in s or "1年" in s: return "y1"
        if "两年" in s or "二年" in s or "2年" in s: return "y2"
        if "三年" in s or "3年" in s: return "y3"
        return None
    techs = {}; offices = set()
    tot = {"y1": {"qty": 0, "amt": 0.0}, "y2": {"qty": 0, "amt": 0.0}, "y3": {"qty": 0, "amt": 0.0}}
    n = 0
    for row in ws.iter_rows(min_row=hdr_row + 1, values_only=True):
        if not row or len(row) <= maxc: continue
        if str(row[c_cat]).strip() != "延保服务": continue
        yk = yr_of(row[c_item])
        if yk is None: continue
        name = str(row[c_tech]).strip() if row[c_tech] else "（未知）"
        code = str(row[c_code]).strip() if (c_code is not None and row[c_code]) else ""
        office = str(row[c_off]).strip() if (c_off is not None and row[c_off]) else ""
        if office: offices.add(office)
        key = (name, code)
        t = techs.get(key)
        if t is None:
            t = {"id": (code or name), "name": name, "office": office,
                 "y1": {"qty": 0, "amt": 0.0}, "y2": {"qty": 0, "amt": 0.0}, "y3": {"qty": 0, "amt": 0.0}}
            techs[key] = t
        elif office and not t["office"]: t["office"] = office
        amt = _num(row[c_pay])
        t[yk]["qty"] += 1; t[yk]["amt"] += amt
        tot[yk]["qty"] += 1; tot[yk]["amt"] += amt
        n += 1
    tech_list = list(techs.values())
    for t in tech_list:
        t["tqty"] = t["y1"]["qty"] + t["y2"]["qty"] + t["y3"]["qty"]
        t["tamt"] = round(t["y1"]["amt"] + t["y2"]["amt"] + t["y3"]["amt"], 2)
    tech_list.sort(key=lambda x: x["tamt"], reverse=True)
    return {"meta": {"来源": "CSM服务项目", "筛选": "销售分类=延保服务", "金额口径": "上缴金额",
                     "年限区分": "服务收费项目列(含一年/两年/三年)", "截止": cutoff+"（同主数据）",
                     "工程师数": len(tech_list)},
            "offices": sorted(offices), "techs": tech_list, "total": tot}

# ---------- WEEKLY_DATA（自然周，覆盖服务月周期） ----------
def find(wb, p):
    for ws in wb.worksheets:
        if p in ws.title: return ws
    return None
def g(r, c): return r[c] if c < len(r) else None
def to_date(v):
    if v is None: return None
    if isinstance(v, datetime.datetime): return v.date()
    if isinstance(v, datetime.date): return v
    return None
def week_key(dt):
    if not hasattr(dt, 'year'): return None
    return (dt - datetime.timedelta(days=dt.weekday())).strftime('%Y-%m-%d')
def month_weeks(ps, pe):
    d = ps - datetime.timedelta(days=ps.weekday())
    weeks = []
    while d <= pe:
        we = d + datetime.timedelta(days=6)
        if we >= ps:
            weeks.append({'key': d.isoformat(), 'label': f"{d.month:02d}/{d.day:02d}-{we.month:02d}/{we.day:02d}",
                          'start': d.isoformat(), 'end': we.isoformat()})
        d = we + datetime.timedelta(days=1)
    return weeks
def build_weekly(path, as_of):
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    y, m = service_month_of(as_of)
    ps, pe = service_period(y, m)
    weeks = month_weeks(ps, pe)
    wkeys = {w['key'] for w in weeks}
    office = collections.defaultdict(lambda: collections.defaultdict(lambda: {'清洗保养':0.0,'延保服务':0.0,'清洁耗材':0.0,'total':0.0}))
    net = collections.defaultdict(lambda: collections.defaultdict(lambda: {'office':'—','清洗保养':0.0,'延保服务':0.0,'清洁耗材':0.0,'total':0.0}))
    m2o = {}
    # CSM服务项目
    ws = find(wb, 'CSM服务项目'); rows = list(ws.iter_rows(values_only=True))
    for r in rows[2:]:
        cls = g(r, 16)
        if cls not in ('清洗保养', '延保服务'): continue
        wk = week_key(to_date(g(r, 7)))
        if wk not in wkeys: continue
        off = g(r, 1) or '—'; nm = g(r, 3) or '未命名网点'; amt = _num(g(r, 24))
        office[wk][off][cls] += amt; office[wk][off]['total'] += amt
        net[wk][nm]['office'] = off
        net[wk][nm][cls] += amt; net[wk][nm]['total'] += amt
        if nm not in m2o: m2o[nm] = off
    # CSM配件
    ws = find(wb, 'CSM配件'); rows = list(ws.iter_rows(values_only=True))
    for r in rows[2:]:
        if g(r, 17) != '清洁耗材': continue
        wk = week_key(to_date(g(r, 36)))
        if wk not in wkeys: continue
        off = g(r, 0) or '—'; nm = g(r, 3) or '未命名网点'; amt = _num(g(r, 27))
        office[wk][off]['清洁耗材'] += amt; office[wk][off]['total'] += amt
        if net[wk][nm]['office'] in ('—', None): net[wk][nm]['office'] = off
        net[wk][nm]['清洁耗材'] += amt; net[wk][nm]['total'] += amt
    # WMS买断
    ws = find(wb, 'WMS网点买断配件明细'); rows = list(ws.iter_rows(values_only=True))
    for r in rows[2:]:
        if g(r, 10) != '清洁耗材': continue
        wk = week_key(to_date(g(r, 11)))
        if wk not in wkeys: continue
        nm = g(r, 3) or '未命名网点'; off = m2o.get(nm, '—'); amt = _num(g(r, 8))
        office[wk][off]['清洁耗材'] += amt; office[wk][off]['total'] += amt
        if net[wk][nm]['office'] in ('—', None): net[wk][nm]['office'] = off
        net[wk][nm]['清洁耗材'] += amt; net[wk][nm]['total'] += amt
    wb.close()
    office = {w['key']: {k: dict(v) for k, v in office[w['key']].items()} for w in weeks if w['key'] in office}
    net = {w['key']: {k: dict(v) for k, v in net[w['key']].items()} for w in weeks if w['key'] in net}
    note = ('真实周数据：清洗保养/延保服务取自CSM服务项目(录入完成时间,上缴金额)；'
            '清洁耗材取自CSM配件(录入完成时间,上缴金额)+WMS网点买断配件明细(发货时间,实收/上缴金额)。'
            '网点维度为源表维修商/经销部名。周期为服务月 %s–%s。' % (ps.isoformat(), pe.isoformat()))
    return {'meta': {'sample': False, 'note': note}, 'weeks': weeks, 'office': office, 'net': net}

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("excel")
    ap.add_argument("--date", required=True, help="数据截止日 YYYY-MM-DD（服务月周期见 SERVICE_PERIOD_OVERRIDE：2月→1.28–2.25、3月→2.26–3.27，其余 (m-1).28–m.27）")
    ap.add_argument("--out", default="_month.json")
    a = ap.parse_args()
    y, m, d = map(int, a.date.split("-"))
    as_of = datetime.date(y, m, d)
    D = build(a.excel, as_of)
    # 数据更新时间：仅在「用户发来新 Excel 并确认同步」时落地（铁律：结构/UI 改动不加）。
    # 格式 YYYY-MM-DD，供看板「数据更新于」展示；未设置则只显示「数据覆盖至」。
    D['meta']['_stamp'] = datetime.date.today().isoformat()
    T, _tech_split = build_tech(a.excel)  # build_tech 返回 (result_dict, source_split_dict)，只取前者
    wb = openpyxl.load_workbook(a.excel, read_only=True, data_only=True)
    B = build_bigcare(wb, a.date)
    E = build_extend(wb, a.date)
    W = build_weekly(a.excel, as_of)
    wb.close()
    out = {"D": D, "T": T, "B": B, "E": E, "W": W}
    json.dump(out, open(a.out, "w", encoding="utf-8"), ensure_ascii=False, separators=(",", ":"))
    # 诊断
    t = D["time"]
    print("✅ 月份=%s 截止=%s 周期=%s 进度=%d/%d=%.1f%%" % (
        D["meta"]["月份"], a.date, t["周期标签"], t["日"], t["月度天数"], t["日"]/t["月度天数"]*100))
    print("   网点=%d 实际合计=¥%s" % (len(D["网点"]), format(round(sum(n["合计"] for n in D["网点"])), ",")))
    print("   品项数=%d 首品=%s" % (len(D["meta"]["品项"]), D["meta"]["品项"][0]))
    print("   大保养: 分母=%d 大保养=%d 占比=%.2f%%" % (B["meta"]["分母"], B["meta"]["大保养总数"], B["meta"]["全区占比"]))
    print("   延保: 工程师=%d 三年=%d单/%.2f元" % (E["meta"]["工程师数"], E["total"]["y3"]["qty"], E["total"]["y3"]["amt"]))
    print("   周数=%d: %s" % (len(W["weeks"]), [w["label"] for w in W["weeks"]]))
    # 四表勾稽（三品）
    nets3 = sum(n["品项"].get("清洗保养",0)+n["品项"].get("延保服务",0)+n["品项"].get("清洁耗材",0) for n in D["网点"])
    print("   三品合计(NETS)=%.2f" % nets3)
    print("→ 写出 %s" % a.out)

if __name__ == "__main__":
    main()
