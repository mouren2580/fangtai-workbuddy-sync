# -*- coding: utf-8 -*-
"""生成/校验 2026-09 看板各数据块（对照参考 dashboard.html 的逻辑）。

用法:
  python gen_v2.py --check 2026-09-12     # 校验模式：按参考截止日重算，打印对照值
  python gen_v2.py --build 2026-09-14     # 构建模式：重建并注入 HTML
"""
import os, re, sys, json, datetime
from collections import defaultdict, Counter, OrderedDict
import openpyxl

HERE = os.path.dirname(os.path.abspath(__file__))
EXCEL = os.path.join(HERE, "2026年9月西北服务产品(3).xlsx")

# ---------- 基础工具 ----------
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

def _s(v):
    if v is None:
        return ""
    s = _fix(v)
    return str(s).strip()

def _iq(x):
    """整数则转 int，否则保留 2 位小数（与参考看板一致）"""
    x = round(float(x), 2)
    return int(x) if float(x).is_integer() else x


def svc_start(month, year):
    return datetime.date(year - 1, 12, 28) if month == 1 else datetime.date(year, month - 1, 28)

# ---------- 产品定义（12 项，取自参考看板 meta.校验）----------
VA_PRODUCTS = [
    ("厨小护",         ["1506040100565"], ["厨小护养护四件套-/-象山"]),
    ("蒸烤清洁剂",      ["1506040100577"], ["蒸烤产品油污清洗套装-500ml-/"]),
    ("锅支架清洁膏",    ["1506040100619"], ["锅支架清洁膏(正式装)-250g膏体-/"]),
    ("洗碗粉",         ["1506040100621"], ["洗碗机专用洗涤粉-1KG/瓶(含包装盒)-/FW"]),
    ("漂洗剂",         ["1506040100585"], ["洗碗机专用漂洗剂-500ml/瓶-/"]),
    ("洗碗盐",         ["1506040100586"], ["洗碗机专用盐-2KG/瓶-/"]),
    ("洗碗液",         ["1512041700025"], ["洗碗液-700ml-酷盘"]),
    ("不锈钢出风管",    ["1125000700000"], ["不锈钢出风管_FG-B-Φ180-S100"]),
    ("排烟管弯头组件",  ["1125000700030"], ["排烟管弯头组件_FG-B-Φ180-S100"]),
    ("排烟直管组件",    ["1125000700010"], ["排烟直管组件_FG-B-Φ180-S100"]),
    ("油杯垫",         ["1125000200170", "1125000200180"],
                       ["油污隔离套装_YBD-C900-K40", "油污隔离套装_YBD-C900-K110"]),
]
SPECIAL = ("止回阀", None, None)

CLEAN_MAIN = ["油烟机小保养", "油烟机大保养", "集成灶小保养", "集成灶大保养"]
EXCLUDE = ["保养超范围收费", "灶具保养"]
BIG = ["油烟机大保养", "油烟机大保养升级包"]
SMOKE_SITES = ("安装", "改装", "油烟机预埋烟管")

# 工单表列（2026-09-18：平台导出把「办事处名称」由 L 挪到 K、「工程师编号」由 AE 挪到 AF，
# 故改为按表头名解析 wo_cols()，下列常量仅作兜底）
WO_OFFICE, WO_PROD, WO_SITE, WO_TYPE, WO_DATE = 10, 73, 57, 65, 41
WO_TECH, WO_TCODE, WO_NET, WO_NO = 55, 31, 79, 53

# key -> (表头名, 兜底列号)
WO_HEADERS = {
    "office": ("办事处名称", WO_OFFICE), "prod": ("产品组", WO_PROD),
    "site": ("服务项目", WO_SITE), "type": ("工单类型", WO_TYPE),
    "date": ("服务完成时间", WO_DATE), "tech": ("服务工程师", WO_TECH),
    "tcode": ("工程师编号", WO_TCODE), "net": ("服务网点", WO_NET),
    "no": ("工单编号", WO_NO),
}
# 自购配件明细列

# CSM配件列
PJ_OFFICE, PJ_NET, PJ_TECH, PJ_TCODE = 2, 3, 5, 6
PJ_NAME, PJ_CODE, PJ_CLASS, PJ_QTY, PJ_AMT, PJ_DATE = 12, 13, 17, 18, 27, 36
# CSM服务项目列
SP_OFFICE, SP_NET, SP_TECH, SP_TCODE = 1, 3, 5, 6
SP_DATE, SP_ITEM, SP_CLASS, SP_AMT = 7, 13, 16, 24
# 工单自购配件明细列（注意：本表无「办事处」列，15列是「大区部」，
# 办事处需由 技师编号/服务网点 反查 CSM配件 或 工单表 得到）
ZG_ALGO, ZG_NET, ZG_TECH, ZG_TCODE = 15, 5, 17, 28
ZG_NAME, ZG_CODE, ZG_QTY, ZG_DATE = 13, 23, 19, 10
ZG_HEADERS = {
    "net": ("服务网点", ZG_NET), "algo": ("大区部", ZG_ALGO), "tech": ("服务工程师", ZG_TECH),
    "tcode": ("工程师编号", ZG_TCODE), "name": ("配件名称", ZG_NAME),
    "code": ("物料编码", ZG_CODE), "qty": ("配件数量", ZG_QTY), "date": ("录入完成时间", ZG_DATE),
}


def _cols(ws, header_row, spec, tag=""):
    """按表头名解析列号；找不到表头则回退兜底常量并告警（平台导出列序偶有变动）"""
    hdr = next(ws.iter_rows(min_row=header_row, max_row=header_row, values_only=True), None) or ()
    out = {}
    for k, (name, fallback) in spec.items():
        idx = next((i for i, v in enumerate(hdr)
                    if v is not None and str(v).strip() == name), None)
        if idx is None:
            idx = fallback
            print("[warn] %s 未找到表头 %r → 回退列号 %d" % (tag, name, fallback))
        out[k] = idx
    return out



def _rows(ws, header_row):
    """按表头行号返回数据行（跳过表头）"""
    return ws.iter_rows(min_row=header_row + 1, values_only=True)


def _uniq(*counters):
    """按「首次出现顺序」合并多个 Counter 的键（与参考看板的稳定排序一致）"""
    seen = OrderedDict()
    for c in counters:
        for k in c:
            seen.setdefault(k, None)
    return list(seen)


# ================= 1. 工单 / 服务量 =================
def build_workorder(excel, cutoff, use_cutoff):
    wb = openpyxl.load_workbook(excel, read_only=True, data_only=True)
    ws = wb["工单"]
    C = _cols(ws, 1, WO_HEADERS, tag="工单表")      # 按表头名解析列号（导出列序会变）
    office_net = defaultdict(lambda: defaultdict(lambda: defaultdict(lambda: [None, 0])))
    o_ord, n_ord = [], {}
    office_cnt = Counter()
    all_no = set()
    pg, wt, si = Counter(), Counter(), Counter()
    yj_cnt = Counter()
    yj_total = 0
    for r in _rows(ws, 1):
        if not r or len(r) <= C["no"]:
            continue
        d = _date(r[C["date"]])
        if d is None:
            continue
        if use_cutoff and d > cutoff:
            continue
        off, net = _s(r[C["office"]]), _s(r[C["net"]])
        tech, code = _s(r[C["tech"]]), _s(r[C["tcode"]])
        no = _s(r[C["no"]])
        if off not in o_ord:
            o_ord.append(off)
        nm = n_ord.setdefault(off, [])
        if net not in nm:
            nm.append(net)
        office_cnt[off] += 1
        all_no.add(no)
        b = office_net[off][net]
        if not b:
            b = office_net[off][net] = {"cnt": 0, "techs": OrderedDict()}
        b["cnt"] += 1
        t = b["techs"].get((tech, code))
        if t is None:
            t = b["techs"][(tech, code)] = {"name": tech, "code": code, "cnt": 0, "zh": 0, "yj": 0}
        t["cnt"] += 1
        pg[_s(r[C["prod"]])] += 1
        wt[_s(r[C["type"]])] += 1
        si[_s(r[C["site"]])] += 1
        if _s(r[C["prod"]]) == "吸油烟机" and _s(r[C["site"]]) in SMOKE_SITES:
            t["yj"] += 1
            yj_cnt[off] += 1
            yj_total += 1

    # 止回阀数量（CSM配件：销售分类=止回阀 且 上缴金额>0 → 数量求和）
    ws2 = wb["CSM配件"]
    zh_total = 0.0
    zh_office = Counter()
    for r in _rows(ws2, 2):
        if not r or len(r) <= PJ_DATE:
            continue
        if _s(r[PJ_CLASS]) != "止回阀":
            continue
        amt = r[PJ_AMT]
        if amt is None or not isinstance(amt, (int, float)) or float(amt) <= 0:
            continue
        d = _date(r[PJ_DATE])
        if use_cutoff and (d is None or d > cutoff):
            continue
        q = _num(r[PJ_QTY])
        off, net, tech, code = _s(r[PJ_OFFICE]), _s(r[PJ_NET]), _s(r[PJ_TECH]), _s(r[PJ_TCODE])
        zh_total += q
        zh_office[off] += q
        b = office_net.get(off, {}).get(net)
        if b:
            t = b["techs"].get((tech, code))
            if t is not None:
                t["zh"] += q

    offices = []
    for off in sorted(o_ord, key=lambda x: -office_cnt[x]):
        nets = []
        for net in sorted(n_ord[off], key=lambda x: -office_net[off][x]["cnt"]):
            b = office_net[off][net]
            techs = [{"name": v["name"], "code": v["code"], "cnt": v["cnt"],
                      "zh": _iq(v["zh"]), "yj": v["yj"]} for v in b["techs"].values()]
            techs.sort(key=lambda x: -x["cnt"])      # 同数按首次出现顺序（与参考一致）
            o = b["techs"]
            nets.append({"name": net, "cnt": b["cnt"],
                         "zh": _iq(sum(v["zh"] for v in o.values())),
                         "yj": sum(v["yj"] for v in o.values()), "techs": techs})
        offices.append({"name": off, "cnt": office_cnt[off],
                        "zh": _iq(zh_office[off]),
                        "yj": yj_cnt[off], "nets": nets})
    techs_seen = set()
    nets_seen = set()
    for off in office_net:
        for net, b in office_net[off].items():
            nets_seen.add(net)
            for k in b["techs"]:
                techs_seen.add(k)
    _end = svc_start(cutoff.month, cutoff.year) + datetime.timedelta(days=30)
    meta = {
        "来源": "工单表(第6个sheet) + CSM配件表(第3个sheet)",
        "周期": "%s~%s" % (svc_start(cutoff.month, cutoff.year).strftime("%Y-%m-%d"),
                          _end.strftime("%Y-%m-%d")),
        "截止": _end.strftime("%Y-%m-%d"),          # 参考看板标签固定为服务月末
        "时间口径": "服务完成时间",
        "总工单数": office_cnt and sum(office_cnt.values()),
        "办事处数": len(offices), "网点数": len(nets_seen), "工程师数": len(techs_seen),
        "维度": "办事处→网点→服务工程师（工单编号去重计数）",
        "产品组数": len(pg), "工单类型数": len(wt), "服务项目数": len(si),
        "止回阀总数": int(zh_total) if float(zh_total).is_integer() else round(zh_total, 2),
        "烟机安装量": yj_total,
        "防火阀占比": round(zh_total / yj_total * 100, 2) if yj_total else 0.0,
    }
    def lst(c):
        return [{"name": k, "cnt": v} for k, v in sorted(c.items(), key=lambda kv: -kv[1])]
    return {"meta": meta, "office": offices, "pgroup": lst(pg), "wotype": lst(wt), "sitem": lst(si)}


# ================= 2. 增值产品（valueadded）=================
def build_valueadded(excel, cutoff, svc=None):
    svc = svc or svc_start(cutoff.month, cutoff.year)
    wb = openpyxl.load_workbook(excel, read_only=True, data_only=True)
    ws = wb["CSM配件"]
    code_map = {}
    for name, codes, _ in VA_PRODUCTS:
        for c in codes:
            code_map[str(c)] = name
    agg = {n: {"total": 0.0, "office": Counter(), "net": Counter(), "tech": Counter(),
               "codes": set()} for n, _, _ in VA_PRODUCTS}
    agg["止回阀"] = {"total": 0.0, "office": Counter(), "net": Counter(), "tech": Counter(),
                   "codes": set()}
    for r in _rows(ws, 2):
        if not r or len(r) <= PJ_DATE:
            continue
        d = _date(r[PJ_DATE])
        if d is None or d > cutoff or d < svc:
            continue
        off, net, tech, code = _s(r[PJ_OFFICE]), _s(r[PJ_NET]), _s(r[PJ_TECH]), _s(r[PJ_TCODE])
        tk = "%s（%s）" % (tech, code) if code else tech
        if _s(r[PJ_CLASS]) == "止回阀":
            amt = r[PJ_AMT]
            if amt is None or not isinstance(amt, (int, float)) or float(amt) <= 0:
                continue
            q = _num(r[PJ_QTY])
            a = agg["止回阀"]
            a["total"] += q
            if off: a["office"][off] += q
            if net: a["net"][net] += q
            if tk: a["tech"][tk] += q
            continue
        nm = code_map.get(_s(r[PJ_CODE]))
        if nm:
            a = agg[nm]
            a["total"] += 1
            a["codes"].add(_s(r[PJ_CODE]))
            if off: a["office"][off] += 1
            if net: a["net"][net] += 1
            if tk: a["tech"][tk] += 1

    products, checks = [], []
    for name, codes, pnames in VA_PRODUCTS:
        a = agg[name]
        found_codes = [c for c in codes if str(c) in a["codes"]]
        missing = [c for c in codes if str(c) not in a["codes"]]
        if a["total"] > 0 and 0 < len(missing) < len(codes):
            note = "部分编码未匹配（编码 %s 未匹配），已按匹配到的编码计入" % "、".join(missing)
            level = "info"
        else:
            note, level = "", "ok"
        products.append({"name": name, "code": codes, "pname": pnames, "special": False,
                         "total": _iq(a["total"]), "office": dict(a["office"]),
                         "net": dict(a["net"]), "tech": dict(a["tech"]),
                         "verify": note or None})
        checks.append({"name": name, "code": codes, "pname": pnames, "found": a["total"] > 0,
                       "matched": found_codes, "note": note, "level": level})
    a = agg["止回阀"]
    ZHV = "特殊口径：R列销售分类=止回阀 计数（与防火阀占比同口径）"
    _o, _n, _t = Counter(), Counter(), Counter()
    for k, v in a["office"].items():
        _o[k] = _iq(v)
    for k, v in a["net"].items():
        _n[k] = _iq(v)
    for k, v in a["tech"].items():
        _t[k] = _iq(v)
    products.append({"name": "止回阀", "code": None, "pname": None, "special": True,
                     "total": _iq(a["total"]), "office": dict(_o),
                     "net": dict(_n), "tech": dict(_t), "verify": ZHV})
    checks.append({"name": "止回阀", "code": None, "pname": None, "found": True,
                   "matched": ["销售分类=止回阀"],
                   "note": "特殊口径：CSM配件表 R列销售分类=止回阀 且 上缴金额(AB列)>0 的行计数（与防火阀占比模块同口径，剔除金额为 0/空/非数字 的止回阀）",
                   "level": "special"})
    total = sum(p["total"] for p in products)
    meta = {"截止日": cutoff.strftime("%Y-%m-%d"), "月份": cutoff.strftime("%Y-%m"),
            "服务月": "%s ~ %s" % (svc.strftime("%Y-%m-%d"), cutoff.strftime("%Y-%m-%d")),
            "来源表": "CSM配件",
            "口径": "普通产品按 配件编码=产品编码 计数（服务月内）；止回阀按 销售分类=止回阀 且 上缴金额(AB列)>0 计数（与防火阀占比模块同口径，剔除金额为 0/空/非数字 的止回阀）",
            "产品项数": len(products),
            "有销量项数": sum(1 for p in products if p["total"] > 0),
            "全区总计数": int(total) if float(total).is_integer() else round(total, 2),
            "校验": checks}
    return {"meta": meta, "products": products}


# ================= 3. 增值配件（valueparts）=================
def build_valueparts(excel, cutoff, svc=None):
    svc = svc or svc_start(cutoff.month, cutoff.year)
    wb = openpyxl.load_workbook(excel, read_only=True, data_only=True)
    names = {n for n, _, _ in VA_PRODUCTS}
    code2name, pname2name = {}, {}
    for n, codes, pns in VA_PRODUCTS:
        for c in codes:
            code2name[str(c)] = n
        for pn in pns:
            pname2name[str(pn)] = n

    win = {"昨日": (cutoff, cutoff),
           "近3日": (cutoff - datetime.timedelta(days=2), cutoff),
           "近7日": (cutoff - datetime.timedelta(days=6), cutoff),
           "当月": (svc, cutoff)}

    def blank():
        return {"yest": 0.0, "last3": 0.0, "last7": 0.0, "month": 0.0}

    prod = OrderedDict()
    for n, codes, pns in VA_PRODUCTS:
        prod[n] = {"special": False, "code": codes, "pname": pns, "buckets": {}, "eng": {}}
    prod["止回阀"] = {"special": True, "code": [], "pname": ["止回阀"], "buckets": {}, "eng": {}}
    enc_pn = OrderedDict()

    def add(name, engkey, office, d, q):
        p = prod[name]
        b = p["buckets"].setdefault(office, blank())
        e = p["eng"].setdefault(engkey, {"name": engkey, "office": office, **blank()})
        for lbl, (a, z) in win.items():
            key = {"昨日": "yest", "近3日": "last3", "近7日": "last7", "当月": "month"}[lbl]
            if a <= d <= z:
                b[key] += q
                e[key] += q

    # 技师/网点 → 办事处 反查表（自购配件表自身没有办事处列）
    tech_office, net_office = {}, {}
    for r in _rows(wb["CSM配件"], 2):
        if not r or len(r) <= PJ_DATE:
            continue
        off = _s(r[PJ_OFFICE])
        if not off:
            continue
        code, tech, net = _s(r[PJ_TCODE]), _s(r[PJ_TECH]), _s(r[PJ_NET])
        if code:
            tech_office.setdefault(code, off)
        if tech:
            tech_office.setdefault("·" + tech, off)
        if net:
            net_office.setdefault(net, off)

    # CSM配件
    ws = wb["CSM配件"]
    for r in _rows(ws, 2):
        if not r or len(r) <= PJ_DATE:
            continue
        d = _date(r[PJ_DATE])
        if d is None or d > cutoff or d < svc:
            continue
        off, net, tech, code = _s(r[PJ_OFFICE]), _s(r[PJ_NET]), _s(r[PJ_TECH]), _s(r[PJ_TCODE])
        tk = "%s（%s）" % (tech, code) if code else tech
        if _s(r[PJ_CLASS]) == "止回阀":
            amt = r[PJ_AMT]
            if amt is None or not isinstance(amt, (int, float)) or float(amt) <= 0:
                continue
            add("止回阀", tk, off, d, _num(r[PJ_QTY]))
            continue
        pn_raw = _s(r[PJ_NAME])
        if pn_raw in pname2name:
            enc_pn.setdefault(pn_raw, None)      # 配件名首次出现顺序
        nm = code2name.get(_s(r[PJ_CODE])) or pname2name.get(pn_raw)
        if nm:
            add(nm, tk, off, d, _num(r[PJ_QTY]))

    # 工单自购配件明细
    ws = wb["工单自购配件明细"]
    Z = _cols(ws, 1, ZG_HEADERS, tag="自购配件表")   # 按表头名解析列号
    for r in _rows(ws, 1):
        if not r or len(r) <= Z["date"]:
            continue
        d = _date(r[Z["date"]])
        if d is None or d > cutoff or d < svc:
            continue
        net, tech, code = _s(r[Z["net"]]), _s(r[Z["tech"]]), _s(r[Z["tcode"]])
        off = (tech_office.get(code) or tech_office.get("·" + tech)
               or net_office.get(net) or "")      # 反查办事处，勿用本表「大区部」列
        tk = "%s（%s）" % (tech, code) if code else tech
        pn_raw = _s(r[Z["name"]])
        if pn_raw in pname2name:
            enc_pn.setdefault(pn_raw, None)
        nm = code2name.get(_s(r[Z["code"]])) or pname2name.get(pn_raw)
        if nm:
            add(nm, tk, off, d, _num(r[Z["qty"]]))

    def _ord_pns(pns):
        """配件名清单按源表首次出现顺序排列（与参考看板一致）"""
        idx = {p: i for i, p in enumerate(enc_pn)}
        return sorted(pns, key=lambda x: idx.get(x, len(enc_pn) + 1))

    products, checks = [], []
    for name, _, pns in VA_PRODUCTS:
        p = prod[name]
        agents = []
        for e in p["eng"].values():
            if e["month"] or e["yest"] or e["last3"] or e["last7"]:
                agents.append({"name": e["name"], "office": e["office"],
                               "yest": e["yest"], "last3": e["last3"],
                               "last7": e["last7"], "month": e["month"]})
        agents.sort(key=lambda x: -x["month"])      # 同数按首次出现顺序（与参考一致）
        tot = {"yest": 0.0, "last3": 0.0, "last7": 0.0, "month": 0.0}
        for b in p["buckets"].values():
            for k in tot:
                tot[k] += b[k]
        products.append({"name": name, "special": False, "pname": _ord_pns(pns), "code": p["code"],
                         "qty_yest": tot["yest"], "qty_last3": tot["last3"],
                         "qty_last7": tot["last7"], "qty_month": tot["month"],
                         "engineers": agents})
        checks.append({"name": name, "code": p["code"], "pname": _ord_pns(pns),
                       "found": tot["month"] > 0,
                       "note": "" if tot["month"] > 0 else "已确认本月无销售（CSM配件表/工单自购配件明细 均无匹配记录），计为 0"})
    p = prod["止回阀"]
    tot = {"yest": 0.0, "last3": 0.0, "last7": 0.0, "month": 0.0}
    for b in p["buckets"].values():
        for k in tot:
            tot[k] += b[k]
    agents = sorted([{"name": e["name"], "office": e["office"], "yest": e["yest"],
                      "last3": e["last3"], "last7": e["last7"], "month": e["month"]}
                     for e in p["eng"].values() if e["month"] or e["yest"] or e["last3"] or e["last7"]],
                    key=lambda x: -x["month"])
    products.append({"name": "止回阀", "special": True, "pname": ["止回阀"], "code": [],
                     "qty_yest": tot["yest"], "qty_last3": tot["last3"],
                     "qty_last7": tot["last7"], "qty_month": tot["month"],
                     "engineers": agents})
    checks.append({"name": "止回阀", "code": [], "pname": ["止回阀"], "found": True,
                   "note": "特殊口径：CSM配件表 销售分类=止回阀 且 上缴金额>0（与防火阀占比模块同口径，剔除金额为 0/空/非数字 的止回阀；自购配件表无此品类）"})
    products.sort(key=lambda x: (-x["qty_month"], -x["qty_last7"], x["name"]))
    meta = {"截止日": cutoff.strftime("%Y-%m-%d"),
            "服务月": "%s ~ %s" % (svc.strftime("%Y-%m-%d"), cutoff.strftime("%Y-%m-%d")),
            "口径": "普通产品：CSM配件表(配件名称/配件编码 任一匹配) + 工单自购配件明细(物料名称/物料编码 任一匹配)，数量列求和；止回阀：CSM配件表 销售分类=止回阀 且 上缴金额(AB列)>0，数量列求和（与防火阀占比模块同口径，剔除金额为 0/空/非数字 的止回阀）。按服务工程师聚合。",
            "来源表": "CSM配件 + 工单自购配件明细",
            "时间窗口": {"昨日": win["昨日"][0].strftime("%Y-%m-%d"),
                        "近3日": "%s ~ %s" % (win["近3日"][0].strftime("%Y-%m-%d"), win["近3日"][1].strftime("%Y-%m-%d")),
                        "近7日": "%s ~ %s" % (win["近7日"][0].strftime("%Y-%m-%d"), win["近7日"][1].strftime("%Y-%m-%d")),
                        "当月": "%s ~ %s" % (svc.strftime("%Y-%m-%d"), cutoff.strftime("%Y-%m-%d"))},
            "产品项数": len(products),
            "有销量项数": sum(1 for x in products if x["qty_month"] > 0),
            "全区总计数_昨日": round(sum(x["qty_yest"] for x in products), 2),
            "全区总计数_近3日": round(sum(x["qty_last3"] for x in products), 2),
            "全区总计数_近7日": round(sum(x["qty_last7"] for x in products), 2),
            "全区总计数_当月": round(sum(x["qty_month"] for x in products), 2),
            "校验": checks}
    return {"meta": meta, "products": products}


# ================= 4. 大保养（四块版：office/net/eng）=================
def build_bigcare4(excel, cutoff, use_cutoff):
    wb = openpyxl.load_workbook(excel, read_only=True, data_only=True)
    ws = wb["CSM服务项目"]
    o_clean, o_excl, o_big = Counter(), Counter(), Counter()
    n_clean, n_excl, n_big = Counter(), Counter(), Counter()
    e_clean, e_excl, e_big = Counter(), Counter(), Counter()
    e_meta, n_office = {}, {}
    t_clean = t_excl = t_big = 0
    enc = OrderedDict()            # 剔除项/大保养项：按源表首次出现顺序（与参考看板一致）
    for r in _rows(ws, 2):
        if not r or len(r) <= SP_AMT:
            continue
        if _s(r[SP_CLASS]) != "清洗保养":
            continue
        if use_cutoff:
            d = _date(r[SP_DATE])
            if d is None or d > cutoff:
                continue
        off, net = _s(r[SP_OFFICE]), _s(r[SP_NET])
        tech, code = _s(r[SP_TECH]), _s(r[SP_TCODE])
        item = _s(r[SP_ITEM])
        if item in EXCLUDE or item in BIG:
            enc.setdefault(item, None)
        ek = (tech, code)
        o_clean[off] += 1; n_clean[net] += 1; e_clean[ek] += 1; t_clean += 1
        if net and off:
            n_office[net] = off
        if ek not in e_meta:
            e_meta[ek] = (tech, code, off, net)
        if item in EXCLUDE:
            o_excl[off] += 1; n_excl[net] += 1; e_excl[ek] += 1; t_excl += 1
        if item in BIG:
            o_big[off] += 1; n_big[net] += 1; e_big[ek] += 1; t_big += 1

    offices = _uniq(o_clean, o_excl, o_big)
    oa = []
    for o in offices:
        den = o_clean[o] - o_excl[o]
        big = o_big[o]
        oa.append({"name": o, "big": big, "den": den, "pct": round(big / den * 100, 2) if den else 0.0})
    oa.sort(key=lambda x: -x["den"])
    td = sum(x["den"] for x in oa)
    oa.append({"name": "[合计]全区", "big": t_big, "den": td,
               "pct": round(t_big / td * 100, 2) if td else 0.0})

    na = []
    for n in _uniq(n_clean, n_excl, n_big):
        den = n_clean[n] - n_excl[n]
        big = n_big[n]
        na.append({"name": n, "office": n_office.get(n, ""), "big": big, "den": den,
                   "pct": round(big / den * 100, 2) if den else 0.0})
    na.sort(key=lambda x: (-x["big"], -x["den"]))
    gp = round(t_big / td * 100, 2) if td else 0.0
    na.append({"name": "[合计]全部网点", "office": "—", "big": t_big, "den": td, "pct": gp})

    ea = []
    for ek in _uniq(e_clean, e_excl, e_big):
        den = e_clean[ek] - e_excl[ek]
        big = e_big[ek]
        nm, cd, off, net = e_meta.get(ek, (ek[0], ek[1], "", ""))
        ea.append({"name": nm, "id": cd, "office": off, "net": net, "big": big, "den": den,
                   "pct": round(big / den * 100, 2) if den else 0.0})
    ea.sort(key=lambda x: (-x["big"], -x["den"]))
    ea.append({"name": "[合计]全部工程师", "id": "—", "office": "—", "net": "—",
               "big": t_big, "den": td, "pct": gp})

    excl_o = [x for x in enc if x in EXCLUDE] + [x for x in EXCLUDE if x not in enc]
    big_o = [x for x in enc if x in BIG] + [x for x in BIG if x not in enc]
    end = svc_start(cutoff.month, cutoff.year) + datetime.timedelta(days=30)
    meta = {"截止日": (cutoff if use_cutoff else end).strftime("%Y-%m-%d"),
            "来源表": "CSM服务项目",
            "周期": "%s~%s" % (svc_start(cutoff.month, cutoff.year).strftime("%Y-%m-%d"), end.strftime("%Y-%m-%d")),
            "清洗保养总单": t_clean, "剔除项": list(EXCLUDE), "大保养项": list(BIG),
            "分母": td, "大保养总数": t_big,
            "全区占比": round(t_big / td * 100, 2) if td else 0.0}
    # 口径差异：MDATA_*_B 用常量序，MONTHLY_FOUR.bigcare 用「源表首次出现序」
    enc_lists = {"剔除项": excl_o, "大保养项": big_o}
    return ({"meta": meta, "office": oa, "net": na, "eng": ea, "enc": enc_lists},
            t_clean, t_excl, t_big)


# ================= 5. 延保明细（EXTEND_TIME_DATA）=================
def build_extend_time(excel, cutoff, svc=None):
    svc = svc or svc_start(cutoff.month, cutoff.year)
    wb = openpyxl.load_workbook(excel, read_only=True, data_only=True)
    ws = wb["CSM服务项目"]
    recs = []
    for r in _rows(ws, 2):
        if not r or len(r) <= SP_AMT:
            continue
        if _s(r[SP_CLASS]) != "延保服务":
            continue
        d = _date(r[SP_DATE])
        if d is None or d > cutoff or d < svc:
            continue
        it = _s(r[SP_ITEM])
        yr = None
        if "一年" in it or "1年" in it:
            yr = "y1"
        elif "两年" in it or "二年" in it or "2年" in it:
            yr = "y2"
        elif "三年" in it or "3年" in it:
            yr = "y3"
        if yr is None:
            continue
        recs.append({"d": d.strftime("%Y-%m-%d"), "o": _s(r[SP_OFFICE]),
                     "t": _s(r[SP_TECH]), "c": "", "yr": yr, "amt": _num(r[SP_AMT])})
    recs.sort(key=lambda x: x["d"])      # 按日期升序（同日保留源表行序，与参考一致）
    return recs


# ================= 6. 清洗保养明细（CLEANING_DATA）=================
def build_cleaning(excel, cutoff, svc=None):
    svc = svc or svc_start(cutoff.month, cutoff.year)
    wb = openpyxl.load_workbook(excel, read_only=True, data_only=True)
    ws = wb["CSM服务项目"]
    recs = []
    for r in _rows(ws, 2):
        if not r or len(r) <= SP_AMT:
            continue
        if _s(r[SP_CLASS]) != "清洗保养":
            continue
        it = _s(r[SP_ITEM])
        if it not in CLEAN_MAIN:
            continue
        d = _date(r[SP_DATE])
        if d is None or d > cutoff or d < svc:
            continue
        recs.append({"o": _s(r[SP_OFFICE]), "n": _s(r[SP_NET]), "e": _s(r[SP_TECH]),
                     "ei": _s(r[SP_TCODE]), "d": d.strftime("%Y-%m-%d"), "it": it,
                     "amt": _num(r[SP_AMT])})
    return recs          # 保持源表行序（与参考看板一致）


# ================= 列序自检 =================
# (工作表, 表头行, {列号: 期望表头名}) —— 平台导出列序偶有变动，改错列会静默出错数据
HEADER_SPEC = [
    ("CSM配件", 2, {PJ_OFFICE: "办事处", PJ_NET: "服务网点", PJ_TECH: "服务工程师",
                    PJ_TCODE: "服务工程师编码", PJ_NAME: "配件名称", PJ_CODE: "配件编码",
                    PJ_CLASS: "销售分类", PJ_QTY: "数量", PJ_AMT: "上缴金额", PJ_DATE: "录入完成时间"}),
    ("CSM服务项目", 2, {SP_OFFICE: "办事处", SP_NET: "服务网点", SP_TECH: "服务工程师",
                        SP_TCODE: "服务工程师编号", SP_DATE: "录入完成时间",
                        SP_ITEM: "服务收费项目", SP_CLASS: "销售分类", SP_AMT: "上缴金额"}),
    ("服务产品收入统计", 2, {1: "服务中心", 2: "办事处", 3: "服务网点", 6: "服务工程师",
                             7: "服务工程师编码", 8: "数据来源", 24: "合计"}),
    ("WMS网点买断配件明细", 2, {3: "CSM网点名称", 8: "实收总金额（上缴金额）",
                                10: "销售分类", 11: "发货时间"}),
]


def verify_columns(excel):
    """核对各表「固定列号」是否仍对得上表头名，返回不一致清单（打印告警）"""
    wb = openpyxl.load_workbook(excel, read_only=True, data_only=True)
    bad = []
    for sheet, hrow, spec in HEADER_SPEC:
        if sheet not in wb.sheetnames:
            bad.append((sheet, -1, "表存在", "缺失"))
            continue
        ws = wb[sheet]
        hdr = list(ws.iter_rows(min_row=hrow, max_row=hrow, values_only=True))
        hdr = hdr[0] if hdr else ()
        for idx, want in spec.items():
            got = hdr[idx] if idx < len(hdr) and hdr[idx] is not None else ""
            if str(got).strip() != want:
                bad.append((sheet, idx, want, str(got).strip()))
    for sheet, idx, want, got in bad:
        print("[warn] 列序变动: %s 第%d列 期望「%s」实际「%s」" % (sheet, idx, want, got))
    if not bad:
        print("[OK] 列序自检通过（CSM配件/CSM服务项目/服务产品收入统计/WMS买断）")
    return bad
