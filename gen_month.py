# -*- coding: utf-8 -*-
"""重新生成 2026-09 看板数据块（MDATA_2026_09_D/T/B/W/E）并注入 dashboard_offline.html。
复用 sync-kit/sync.py 中自包含的 build_tech/gen_extend/gen_weekly（D、B 为本次重建）。
唯一依赖：openpyxl。"""
import os, re, json, datetime, importlib.util
from collections import defaultdict
import openpyxl

HERE = os.path.dirname(os.path.abspath(__file__))
EXCEL = r"D:/xwechat_files/xuyuerong6203_786d/temp/RWTemp/2026-09/9e20f478899dc29eb19741386f9343c8/2026年9月西北服务产品(2).xlsx"
HTML = os.path.join(HERE, "dashboard_offline.html")
CUTOFF = datetime.date(2026, 9, 10)          # 统计截止日（铁律：更新日 9-11 前一天）
WEEKS_END = datetime.date(2026, 9, 27)       # 周数据脚手架：整服务月（8/28–9/27）
MONTH = "2026-09"

ITEMS = ["清洗保养", "延保服务", "清洁耗材", "止回阀", "装饰罩", "水/气路金属软管",
         "排烟出风管", "水气路通用辅材", "其他产品线专用", "其他水路辅材", "其他气路辅材",
         "橱柜类辅材", "燃气管快速接头", "工具", "电路辅材"]


# ---------------- 通用工具（同 sync.py）----------------
def _fix(s):
    if isinstance(s, str):
        try:
            return s.encode("latin-1").decode("gbk")
        except Exception:
            return s
    return s


def _num(v):
    try:
        return round(float(v), 2)
    except Exception:
        return 0.0


def _date(v):
    if isinstance(v, datetime.datetime):
        return v.date()
    if isinstance(v, datetime.date):
        return v
    if isinstance(v, str) and v.strip():
        for fmt in ("%Y-%m-%d %H:%M:%S", "%Y-%m-%d %H:%M", "%Y-%m-%d"):
            try:
                return datetime.datetime.strptime(v.strip(), fmt).date()
            except Exception:
                continue
    return None


# ---------------- 服务工程师（T）----------------  (取自 sync.py build_tech)
def build_tech(src):
    wb = openpyxl.load_workbook(src, read_only=True, data_only=True)
    ws = wb["服务产品收入统计"]
    techs, source_split = {}, defaultdict(float)
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
                 "center": _fix(r[1]) or "", "nets": set(), "items": [0.0] * 15, "amount": 0.0}
            techs[code] = t
        t["amount"] += amount
        for i, ci in enumerate(range(9, 24)):
            t["items"][i] += _num(r[ci])
        if r[3]:
            t["nets"].add(_fix(r[3]))
        if not t["office"] and r[2]:
            t["office"] = _fix(r[2])
        if not t["center"] and r[1]:
            t["center"] = _fix(r[1])
    tech_list, offices = [], set()
    for t in techs.values():
        if t["office"]:
            offices.add(t["office"])
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


# ---------------- 延保（E）----------------  (取自 sync.py gen_extend)
def gen_extend(excel, cutoff):
    wb = openpyxl.load_workbook(excel, read_only=True, data_only=True)
    target = next((s for s in wb.sheetnames if ("服务项目" in s) or ("CSM" in s.upper())), None)
    ws = wb[target]
    KW = ["销售分类", "服务收费项目", "服务工程师", "服务工程师编码", "上缴金额", "办事处"]
    hdr_row = hdr = None
    for ri in range(1, 7):
        row = list(next(ws.iter_rows(min_row=ri, max_row=ri, values_only=True)))
        flat = [str(x) for x in row if x is not None]
        if any(k in f for f in flat for k in KW):
            hdr_row, hdr = ri, row
            break
    idx = {}
    for ci, v in enumerate(hdr):
        if v is None:
            continue
        s = str(v).strip()
        if s in KW:
            idx[s] = ci
    c_cat, c_item, c_tech = idx["销售分类"], idx["服务收费项目"], idx["服务工程师"]
    c_code, c_pay, c_off = idx.get("服务工程师编码"), idx["上缴金额"], idx.get("办事处")
    maxc = max(x for x in (c_cat, c_item, c_tech, c_pay, c_code, c_off) if x is not None)

    def yr_of(s):
        s = str(s)
        if "一年" in s or "1年" in s:
            return "y1"
        if "两年" in s or "二年" in s or "2年" in s:
            return "y2"
        if "三年" in s or "3年" in s:
            return "y3"
        return None

    techs, offices = {}, set()
    tot = {"y1": {"qty": 0, "amt": 0.0}, "y2": {"qty": 0, "amt": 0.0}, "y3": {"qty": 0, "amt": 0.0}}
    for row in ws.iter_rows(min_row=hdr_row + 1, values_only=True):
        if not row or len(row) <= maxc:
            continue
        if str(row[c_cat]).strip() != "延保服务":
            continue
        yk = yr_of(row[c_item])
        if yk is None:
            continue
        name = str(row[c_tech]).strip() if row[c_tech] else "（未知）"
        code = str(row[c_code]).strip() if (c_code is not None and row[c_code]) else ""
        office = str(row[c_off]).strip() if (c_off is not None and row[c_off]) else ""
        if office:
            offices.add(office)
        key = (name, code)
        t = techs.get(key)
        if t is None:
            t = {"id": (code or name), "name": name, "office": office,
                 "y1": {"qty": 0, "amt": 0.0}, "y2": {"qty": 0, "amt": 0.0}, "y3": {"qty": 0, "amt": 0.0}}
            techs[key] = t
        elif office and not t["office"]:
            t["office"] = office
        amt = _num(row[c_pay])
        t[yk]["qty"] += 1
        t[yk]["amt"] += amt
        tot[yk]["qty"] += 1
        tot[yk]["amt"] += amt
    tech_list = list(techs.values())
    for t in tech_list:
        t["tqty"] = t["y1"]["qty"] + t["y2"]["qty"] + t["y3"]["qty"]
        t["tamt"] = round(t["y1"]["amt"] + t["y2"]["amt"] + t["y3"]["amt"], 2)
    tech_list.sort(key=lambda x: x["tamt"], reverse=True)
    return {"meta": {
                "来源": "CSM服务项目", "筛选": "销售分类=延保服务", "金额口径": "上缴金额",
                "年限区分": "服务收费项目列(含一年/两年/三年)",
                "截止": cutoff.strftime("%Y-%m-%d"), "工程师数": len(tech_list)},
            "offices": sorted(offices), "techs": tech_list, "total": tot}


# ---------------- 周数据（W）----------------  (取自 sync.py gen_weekly，增加数据截止过滤)
def _svc_start(month, year):
    if month == 1:
        return datetime.date(year - 1, 12, 28)
    return datetime.date(year, month - 1, 28)


def gen_weekly(excel, cutoff, data_cutoff):
    wb = openpyxl.load_workbook(excel, data_only=True, read_only=True)

    def find_sheet(p):
        for ws in wb.worksheets:
            if p in ws.title:
                return ws

    svc = _svc_start(cutoff.month, cutoff.year)
    ws0 = svc
    while ws0.weekday() != 0:
        ws0 -= datetime.timedelta(days=1)

    def week_key(dt):
        if not hasattr(dt, "year"):
            return None
        if isinstance(dt, datetime.datetime):
            dt = dt.date()
        monday = dt - datetime.timedelta(days=dt.weekday())
        if monday < ws0:
            return None
        return monday.strftime("%Y-%m-%d")

    weeks = []
    d = ws0
    while d <= cutoff:
        if d.weekday() == 0:
            start, end = d, d + datetime.timedelta(days=6)
            weeks.append({"key": start.strftime("%Y-%m-%d"),
                          "label": "%s-%s" % (start.strftime("%m/%d"), end.strftime("%m/%d")),
                          "start": start.strftime("%Y-%m-%d"), "end": end.strftime("%Y-%m-%d")})
        d += datetime.timedelta(days=1)

    office = defaultdict(lambda: defaultdict(lambda: {"清洗保养": 0.0, "延保服务": 0.0, "清洁耗材": 0.0, "total": 0.0}))
    net = defaultdict(lambda: defaultdict(lambda: {"office": "—", "清洗保养": 0.0, "延保服务": 0.0, "清洁耗材": 0.0, "total": 0.0}))
    m2o = {}

    def add_map(name, off):
        if name and off and name not in m2o:
            m2o[name] = off

    def g(r, c):
        return r[c] if c < len(r) else None

    # 清洗保养 / 延保服务 : CSM服务项目
    ws = find_sheet("CSM服务项目")
    for r in list(ws.iter_rows(values_only=True))[2:]:
        cls = g(r, 16)
        if cls not in ("清洗保养", "延保服务"):
            continue
        dt = _date(g(r, 7))
        if dt is None or dt > data_cutoff:
            continue
        wk = week_key(dt)
        if not wk:
            continue
        off = g(r, 1) or "—"
        nm = g(r, 3) or "未命名网点"
        amt = _num(g(r, 24))
        office[wk][off][cls] += amt
        office[wk][off]["total"] += amt
        net[wk][nm]["office"] = off
        net[wk][nm][cls] += amt
        net[wk][nm]["total"] += amt
        add_map(nm, off)
    # 清洁耗材 : CSM配件
    # 注意：CSM配件表 col0 是「大区」而非「办事处」，不能用于登记 网点→办事处 映射
    #（否则 WMS 网点买断行会被错误归到大区名下；与参考看板一致：映射只取自 CSM服务项目）
    ws = find_sheet("CSM配件")
    for r in list(ws.iter_rows(values_only=True))[2:]:
        if g(r, 17) != "清洁耗材":
            continue
        dt = _date(g(r, 36))
        if dt is None or dt > data_cutoff:
            continue
        wk = week_key(dt)
        if not wk:
            continue
        off = g(r, 0) or "—"
        nm = g(r, 3) or "未命名网点"
        amt = _num(g(r, 27))
        office[wk][off]["清洁耗材"] += amt
        office[wk][off]["total"] += amt
        if net[wk][nm]["office"] in ("—", None):
            net[wk][nm]["office"] = off
        net[wk][nm]["清洁耗材"] += amt
        net[wk][nm]["total"] += amt
    # 清洁耗材 : WMS网点买断配件明细
    ws = find_sheet("WMS网点买断配件明细")
    for r in list(ws.iter_rows(values_only=True))[2:]:
        if g(r, 10) != "清洁耗材":
            continue
        dt = _date(g(r, 11))
        if dt is None or dt > data_cutoff:
            continue
        wk = week_key(dt)
        if not wk:
            continue
        nm = g(r, 3) or "未命名网点"
        off = m2o.get(nm, "—")
        amt = _num(g(r, 8))
        office[wk][off]["清洁耗材"] += amt
        office[wk][off]["total"] += amt
        if net[wk][nm]["office"] in ("—", None):
            net[wk][nm]["office"] = off
        net[wk][nm]["清洁耗材"] += amt
        net[wk][nm]["total"] += amt

    office_out = {wk: {o: dict(v) for o, v in office[wk].items()} for wk in office}
    net_out = {wk: {n: dict(v) for n, v in net[wk].items()} for wk in net}
    return {"meta": {
                "sample": False,
                "note": "真实周数据：清洗保养/延保服务取自CSM服务项目(录入完成时间,上缴金额)；"
                        "清洁耗材取自CSM配件(录入完成时间,上缴金额)+WMS网点买断配件明细(发货时间,实收/上缴金额)。"
                        "首周(服务月28日起)已跳过。网点维度为源表维修商/经销部名。数据截至 %s。" % data_cutoff.strftime("%Y-%m-%d"),
                "截止": data_cutoff.strftime("%Y-%m-%d")},
            "weeks": weeks, "office": office_out, "net": net_out}


# ---------------- 主数据（D）----------------  (本次重建)
def _month_time(cutoff):
    svc = _svc_start(cutoff.month, cutoff.year)          # 服务月起点（上月28日）
    end = svc + datetime.timedelta(days=30)              # 当月27日
    md = (end - svc).days + 1                            # 月度天数（31）
    day = (cutoff - svc).days + 1                        # 截止日所处服务月第几天
    week = (day + 6) // 7
    total_weeks = (md + 6) // 7
    label = "%d.%d–%d.%d（%d天）" % (svc.month, svc.day, end.month, end.day, md)
    return {"年": cutoff.year, "月": cutoff.month, "日": day, "月度天数": md,
            "周": week, "总周数": total_weeks,
            "周期开始": svc.strftime("%Y-%m-%d"), "周期结束": end.strftime("%Y-%m-%d"),
            "周期标签": label}


def build_main(excel):
    wb = openpyxl.load_workbook(excel, read_only=True, data_only=True)
    ws = wb["服务产品收入统计"]
    groups = {}
    order = []
    for r in ws.iter_rows(min_row=3, values_only=True):
        if not r or len(r) < 25:
            continue
        office = _fix(r[2]) if r[2] else ""
        center = _fix(r[1]) if r[1] else ""
        net = _fix(r[3]) if r[3] else ""
        src = _fix(r[8]) if r[8] else ""
        if not net:
            continue
        key = (office, center, net)
        g = groups.get(key)
        if g is None:
            g = {"品项": [0.0] * 15, "合计": 0.0, "来源": set()}
            groups[key] = g
            order.append(key)
        for i in range(15):
            g["品项"][i] += _num(r[9 + i])
        g["合计"] += _num(r[24])
        if src:
            g["来源"].add(src)
    nets = []
    for key in order:
        office, center, net = key
        g = groups[key]
        items = {ITEMS[i]: round(v, 2) for i, v in enumerate(g["品项"])}
        nets.append({"办事处": office, "服务中心": center, "服务网点": net,
                     "合计": round(g["合计"], 2), "品项": items, "来源": sorted(g["来源"])})
    return {"meta": {"大区": "西北大区部", "月份": MONTH, "品项": ITEMS,
                     "来源表": "服务产品收入统计", "统计截止日": CUTOFF.strftime("%Y-%m-%d"),
                     "_stamp": "2026-09-11"},
            "网点": nets,
            "time": _month_time(CUTOFF)}


# ---------------- 大保养（B）----------------  (本次重建)
def build_bigcare(excel):
    wb = openpyxl.load_workbook(excel, read_only=True, data_only=True)
    ws = wb["CSM服务项目"]
    EXCLUDE = ["保养超范围收费", "灶具保养"]
    BIG = ["油烟机大保养升级包", "油烟机大保养"]
    office_clean, office_excl, office_big = defaultdict(int), defaultdict(int), defaultdict(int)
    net_clean, net_excl, net_big = defaultdict(int), defaultdict(int), defaultdict(int)
    net_office = {}
    total_clean = total_excl = total_big = 0
    for r in ws.iter_rows(min_row=3, values_only=True):
        if not r or len(r) < 17:
            continue
        item = (str(r[13]).strip() if r[13] is not None else "")
        sclass = (str(r[16]).strip() if r[16] is not None else "")
        office = (str(r[1]).strip() if r[1] is not None else "")
        net = (str(r[3]).strip() if r[3] is not None else "")
        # 清洗保养总单：销售分类(col16)=清洗保养（含小保养/大保养/灶具保养/超范围等子类）
        if sclass == "清洗保养":
            office_clean[office] += 1; net_clean[net] += 1; total_clean += 1
            if net and office:
                net_office[net] = office
        # 剔除项 / 大保养项：服务收费项目(col13) 命中（与清洗保养总单可重叠，故独立 if）
        if item in EXCLUDE:
            office_excl[office] += 1; net_excl[net] += 1; total_excl += 1
            if net and office:
                net_office[net] = office
        if item in BIG:
            office_big[office] += 1; net_big[net] += 1; total_big += 1
            if net and office:
                net_office[net] = office
    offices = set(office_clean) | set(office_excl) | set(office_big)
    office_arr = []
    for o in offices:
        den = office_clean[o] - office_excl[o]
        big = office_big[o]
        office_arr.append({"name": o, "big": big, "den": den, "pct": round(big / den * 100, 2) if den else 0.0})
    office_arr.sort(key=lambda x: (-x["den"], x["name"]))
    total_den = sum(o["den"] for o in office_arr)
    office_arr.append({"name": "[合计]全区", "big": total_big, "den": total_den,
                       "pct": round(total_big / total_den * 100, 2) if total_den else 0.0})

    nets = set(net_clean) | set(net_excl) | set(net_big)
    net_arr = []
    for n in nets:
        den = net_clean[n] - net_excl[n]
        big = net_big[n]
        net_arr.append({"name": n, "office": net_office.get(n, ""), "big": big, "den": den,
                        "pct": round(big / den * 100, 2) if den else 0.0})
    net_arr.sort(key=lambda x: (-x["big"], -x["den"], x["name"]))
    return {"meta": {"截止日": CUTOFF.strftime("%Y-%m-%d"), "来源表": "CSM服务项目",
                     "清洗保养总单": total_clean, "剔除项": EXCLUDE, "大保养项": BIG,
                     "分母": total_den, "大保养总数": total_big,
                     "全区占比": round(total_big / total_den * 100, 2) if total_den else 0.0},
            "office": office_arr, "net": net_arr}


# ---------------- 注入 HTML ----------------
def replace_block(html, varname, obj):
    pat = re.compile(r"(?m)^(const|let)\s+" + re.escape(varname) + r"\s*=.*$")
    if not pat.search(html):
        raise SystemExit("未找到 %s 单行声明" % varname)
    new = "const %s = %s;" % (varname, json.dumps(obj, ensure_ascii=False, separators=(",", ":")))
    return pat.sub(new, html, count=1)


def main():
    if not os.path.exists(EXCEL):
        raise SystemExit("找不到 Excel")
    html = open(HTML, encoding="utf-8").read()

    D = build_main(EXCEL)
    T, _ = build_tech(EXCEL)
    B = build_bigcare(EXCEL)
    W = gen_weekly(EXCEL, WEEKS_END, CUTOFF)
    E = gen_extend(EXCEL, CUTOFF)

    html = replace_block(html, "MDATA_2026_09_D", D)
    html = replace_block(html, "MDATA_2026_09_T", T)
    html = replace_block(html, "MDATA_2026_09_B", B)
    html = replace_block(html, "MDATA_2026_09_W", W)
    html = replace_block(html, "MDATA_2026_09_E", E)

    # 校验：抽取并 JSON 解析每个块
    for v in ("MDATA_2026_09_D", "MDATA_2026_09_T", "MDATA_2026_09_B", "MDATA_2026_09_W", "MDATA_2026_09_E"):
        m = re.search(r"(?m)^(const|let)\s+" + re.escape(v) + r"\s*=\s*(.*);$", html)
        if not m:
            raise SystemExit("注入后找不到 %s" % v)
        json.loads(m.group(2))

    # 版本号 / version.json / BUILD 常量
    build_id = "20260911-" + datetime.datetime.now().strftime("%H%M")
    html = re.sub(r'var BUILD = "[^"]+";', 'var BUILD = "%s";' % build_id, html, count=1)
    open(HTML, "w", encoding="utf-8").write(html)

    vj = {"v": build_id, "cut": CUTOFF.strftime("%Y-%m-%d"),
          "t": datetime.datetime.now().strftime("%Y-%m-%dT%H:%M:%S")}
    open(os.path.join(HERE, "version.json"), "w", encoding="utf-8").write(json.dumps(vj, ensure_ascii=False))

    # 汇总
    actual = round(sum(n["合计"] for n in D["网点"]), 2)
    print("[OK] 已重建并注入 2026-09 数据块 ->", os.path.abspath(HTML))
    print("   BUILD =", build_id, " 截止日 =", CUTOFF.strftime("%Y-%m-%d"))
    print("   主数据: 网点=%d 总实际=¥%s" % (len(D["网点"]), format(actual, ",")))
    print("   工程师: %d 人 工单消耗=¥%s" % (T["count"], format(T["total"], ",")))
    print("   大保养: 总单=%d 剔除=%d 分母=%d 大保养=%d 占比=%.2f%%"
          % (B["meta"]["清洗保养总单"], B["meta"]["清洗保养总单"] - B["meta"]["分母"],
             B["meta"]["分母"], B["meta"]["大保养总数"], B["meta"]["全区占比"]))
    print("   周数据: %d 周 (脚手架至 %s)" % (len(W["weeks"]), WEEKS_END.strftime("%m-%d")))
    print("   延保: 一年 ¥%.2f / 二年 ¥%.2f / 三年 ¥%.2f"
          % (E["total"]["y1"]["amt"], E["total"]["y2"]["amt"], E["total"]["y3"]["amt"]))


if __name__ == "__main__":
    main()
