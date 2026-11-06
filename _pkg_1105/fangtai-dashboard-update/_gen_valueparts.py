# -*- coding: utf-8 -*-
"""生成 增值配件明细 模块的 VALUE_PARTS_DATA。

数据来源：
  1) 增值产品填写模板.xlsx  -> 产品目录(名称/产品名称/产品编码/是否止回阀)
  2) CSM配件表              -> M列[12]配件名称 / N列[13]配件编码 / S列[18]数量 / R列[17]销售分类 / AB列[27]上缴金额 / AK列[36]录入完成时间 / F列[5]服务工程师 / G列[6]工程师编码
  3) 工单自购配件明细        -> AB列[27]物料名称 / X列[23]物料编码 / T列[19]配件数量 / K列[10]录入完成时间 / R列[17]服务工程师 / AC列[28]工程师编号

普通产品：CSM配件表(配件名称/配件编码 任一匹配) + 工单自购配件明细(物料名称/物料编码 任一匹配)，数量列求和。
止回阀(特殊)：CSM配件表 销售分类=止回阀 且 上缴金额(AB列)>0，数量列求和（与防火阀占比模块同口径，剔除金额为 0/空/非数字 的止回阀；自购配件表无此品类）。

时间窗口（截止日冻结 2026-08-24）：
  昨日  = 2026-08-24
  近3日 = 2026-08-22 ~ 2026-08-24
  当月  = 服务月 2026-07-28 ~ 2026-08-24

按 服务工程师(姓名+编码) 聚合三窗口数量。
"""
import argparse, json, re, os
from collections import defaultdict, OrderedDict
from datetime import datetime, date, timedelta

TEMPLATE = "增值产品填写模板.xlsx"
EXCEL = "_新8月数据.xlsx"
OUT = "_valueparts_data.json"
STORE = "_valueparts_daily.json"   # 按日累积存储（跨服务月，供 昨日/近3/近7 窗口回看上月末）

CUT = date(2026, 8, 24)          # 冻结截止日
SVC_LO = date(2026, 7, 28)       # 服务月起点
LAST3_LO = date(2026, 8, 22)   # 近3日起点（含昨日）
LAST7_LO = date(2026, 8, 18)   # 近7日起点（含昨日，共7天）

# 用户 2026-08-26 确认：以下产品本月确无销售（非编码错误），计为 0
CONFIRMED_ZERO = {"洗碗液", "排烟管弯头组件"}

def norm_code(s):
    s = str(s).strip() if s is not None else ""
    if s.isdigit():
        return re.sub(r"0+$", "", s) or "0"
    return s

def to_date(v):
    if isinstance(v, datetime):
        return v.date()
    if isinstance(v, date):
        return v
    return None

def fnum(v):
    try:
        return float(v) if v is not None else 0.0
    except Exception:
        return 0.0

def amt_gt0(v):
    """止回阀过滤口径：上缴金额(AB列, 0-based 27) 为有效数字且 > 0 才计入；0/空/非数字剔除。"""
    try:
        f = float(v)
    except (TypeError, ValueError):
        return False
    return f > 0

def read_template():
    import openpyxl
    wb = openpyxl.load_workbook(TEMPLATE, data_only=True)
    ws = wb.active
    merged = OrderedDict()  # name -> {pnames:set, codes:set, special}
    for r in range(4, ws.max_row + 1):
        name = ws.cell(r, 2).value
        pname = ws.cell(r, 3).value
        code = ws.cell(r, 4).value
        if name is None or str(name).strip() == "":
            continue
        name = str(name).strip()
        pname = str(pname).strip() if pname else ""
        code = str(code).strip() if code else ""
        special = (name == "止回阀")
        if name not in merged:
            merged[name] = {"pnames": set(), "codes": set(), "special": special}
        if pname:
            merged[name]["pnames"].add(pname)
        if code:
            merged[name]["codes"].add(code)
    return merged

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--excel", default=EXCEL)
    ap.add_argument("--date", default="2026-08-24")
    ap.add_argument("--svc-lo", help="服务月起始 YYYY-MM-DD（显式覆盖，优先）")
    ap.add_argument("--cut", help="截止日 YYYY-MM-DD（显式覆盖，优先于 --date）")
    ap.add_argument("--month", default="2026-09", help="当前服务月键（写入按日累积存储用），如 2026-09")
    ap.add_argument("--out", default=OUT)
    args = ap.parse_args()

    global CUT, SVC_LO, LAST3_LO, LAST7_LO
    if args.cut:
        CUT = datetime.strptime(args.cut, "%Y-%m-%d").date()
    else:
        CUT = datetime.strptime(args.date, "%Y-%m-%d").date()
    if args.svc_lo:
        SVC_LO = datetime.strptime(args.svc_lo, "%Y-%m-%d").date()
    LAST3_LO = CUT - timedelta(days=2)
    LAST7_LO = CUT - timedelta(days=6)
    # 当月 = 服务月起点 ~ 截止日（用户确认：当月=服务月，日期到哪天算哪天）

    import openpyxl
    prods = read_template()

    wb = openpyxl.load_workbook(args.excel, read_only=True, data_only=True)

    # ---- CSM配件表 (表头第2行, 数据第3行起) ----
    ws1 = wb["CSM配件"]
    rows1 = list(ws1.iter_rows(min_row=3, values_only=True))
    # ---- 工单自购配件明细 (表头第1行, 数据第2行起) ----
    ws2 = wb["工单自购配件明细"]
    rows2 = list(ws2.iter_rows(min_row=2, values_only=True))

    # 预建产品匹配键
    # 普通产品: 名称集合 + 编码集合(norm)
    normal = []   # list of dict
    zh_name = None
    for name, info in prods.items():
        keys = set()
        for n in info["pnames"]:
            keys.add(n)
        keys.add(name)
        code_norms = set(norm_code(c) for c in info["codes"] if c)
        if info["special"]:
            zh_name = name
            continue
        normal.append({"name": name, "keys": keys, "code_norms": code_norms,
                       "pnames": list(info["pnames"]), "codes": list(info["codes"])})

    def key_of(ename, ecode):
        return (ename or "（未知）") + (("（" + str(ecode) + "）") if ecode else "")

    # ---- 按日累积： (日期, 工程师key, 办事处, 产品, 是否止回阀) -> 数量 ----
    daily = defaultdict(float)
    def add_rec(pname, special, ename, ecode, qty, d, office):
        if d is None: return
        if not (SVC_LO <= d <= CUT):   # 只存本服务月区间记录，避免跨月重复计数
            return
        daily[(d.isoformat(), key_of(ename, ecode), office or "", pname, 1 if special else 0)] += qty

    # ---- 办事处映射：用 CSM配件表[2]办事处 建立 (工程师,编码)->办事处 ----
    eng_office = {}
    eng_office_name = {}
    # ---- 走 CSM配件表 ----
    for r in rows1:
        name_col = r[12]      # 配件名称
        code_col = r[13]      # 配件编码
        qty = fnum(r[18])     # 数量
        sale_cls = r[17]      # 销售分类
        off = r[27]           # 上缴金额
        d = to_date(r[36])    # 录入完成时间
        ename = r[5]; ecode = r[6]
        key = key_of(ename, ecode)
        if r[2]:
            eng_office[key] = r[2]
            if ename: eng_office_name[ename] = r[2]
        # 止回阀特殊（与防火阀占比模块同口径：销售分类=止回阀 且 上缴金额>0，剔除金额为 0/空/非数字 的止回阀）
        if zh_name is not None and sale_cls == "止回阀" and amt_gt0(off):
            add_rec(zh_name, True, ename, ecode, qty, d, r[2] or "")
            continue
        # 普通: 名称或编码匹配
        nm = str(name_col).strip() if name_col is not None else ""
        cn = norm_code(code_col)
        for item in normal:
            if nm in item["keys"] or (cn and cn in item["code_norms"]):
                add_rec(item["name"], False, ename, ecode, qty, d, r[2] or "")
                break

    # ---- 走 工单自购配件明细（无办事处列，用 CSM 映射兜底） ----
    for r in rows2:
        mname = r[27]         # 物料名称 (AB)
        mcode = r[23]         # 物料编码 (X)
        qty = fnum(r[19])     # 配件数量 (T)
        d = to_date(r[10])    # 录入完成时间 (K)
        ename = r[17]; ecode = r[28]
        key = key_of(ename, ecode)
        office = eng_office.get(key) or eng_office_name.get(ename) or ""
        nm = str(mname).strip() if mname is not None else ""
        cn = norm_code(mcode)
        for item in normal:
            if nm in item["keys"] or (cn and cn in item["code_norms"]):
                add_rec(item["name"], False, ename, ecode, qty, d, office)
                break

    # ---- 合并按日累积存储（保留其它服务月，供跨月窗口） ----
    store = {"months": {}}
    if os.path.exists(STORE):
        try:
            store = json.load(open(STORE, encoding="utf-8"))
            if not isinstance(store.get("months"), dict):
                store = {"months": {}}
        except Exception as e:
            print("  [warn] 读取按日存储 %s 失败：%s" % (STORE, e))
            store = {"months": {}}
    cur_records = [{"d":k[0], "k":k[1], "o":k[2], "p":k[3], "s":k[4], "q":round(v, 2)} for k, v in daily.items()]
    store["months"][args.month] = cur_records
    with open(STORE, "w", encoding="utf-8") as f:
        json.dump(store, f, ensure_ascii=False, separators=(",", ":"))
    all_recs = []
    for mv in store["months"].values():
        if isinstance(mv, list): all_recs.extend(mv)

    # ---- 从「全量按日记录」按窗口聚合（跨服务月：月初近3/近7 可回看上月末） ----
    WINS = {"yest": (CUT, CUT), "last3": (CUT - timedelta(days=2), CUT),
            "last7": (CUT - timedelta(days=6), CUT), "month": (SVC_LO, CUT)}
    agg = OrderedDict()
    for item in normal:
        agg[item["name"]] = {"eng": defaultdict(lambda: {"yest":0.0,"last3":0.0,"last7":0.0,"month":0.0,"office":""}),
                              "yest":0.0,"last3":0.0,"last7":0.0,"month":0.0}
    zh_agg = {"eng": defaultdict(lambda: {"yest":0.0,"last3":0.0,"last7":0.0,"month":0.0,"office":""}),
              "yest":0.0,"last3":0.0,"last7":0.0,"month":0.0}
    for rec in all_recs:
        st = zh_agg if rec.get("s") else agg.get(rec.get("p"))
        if st is None: continue
        try:
            rd = date.fromisoformat(rec["d"])
        except Exception:
            continue
        q = float(rec.get("q") or 0)
        for w, (lo, hi) in WINS.items():
            if lo <= rd <= hi:
                e = st["eng"][rec["k"]]
                if rec.get("o"): e["office"] = rec["o"]
                e[w] += q; st[w] += q

    # ---- 组装输出 ----
    out_products = []
    verify = []
    for item in normal:
        st = agg[item["name"]]
        engineers = []
        for k, v in st["eng"].items():
            engineers.append({"name": k, "office": v.get("office", ""), "yest": round(v["yest"], 2),
                               "last3": round(v["last3"], 2), "last7": round(v["last7"], 2), "month": round(v["month"], 2)})
        engineers.sort(key=lambda x: (-x["month"], -x["last3"], -x["yest"]))
        out_products.append({
            "name": item["name"], "special": False,
            "pname": item["pnames"], "code": item["codes"],
            "qty_yest": round(st["yest"], 2), "qty_last3": round(st["last3"], 2), "qty_last7": round(st["last7"], 2),
            "qty_month": round(st["month"], 2), "engineers": engineers,
        })
        # 校验
        note = ""
        if st["month"] == 0:
            if item["name"] in CONFIRMED_ZERO:
                note = "已确认本月无销售（CSM配件表/工单自购配件明细 均无匹配记录），计为 0"
            else:
                note = "服务月内 CSM配件表(按名称/编码) 与 工单自购配件明细(按物料名称/编码) 均未匹配到记录，计为 0"
        verify.append({"name": item["name"], "code": item["codes"], "pname": item["pnames"],
                       "found": st["month"] > 0, "note": note})

    # 止回阀
    if zh_name is not None:
        st = zh_agg
        engineers = []
        for k, v in st["eng"].items():
            engineers.append({"name": k, "office": v.get("office", ""), "yest": round(v["yest"], 2),
                               "last3": round(v["last3"], 2), "last7": round(v["last7"], 2), "month": round(v["month"], 2)})
        engineers.sort(key=lambda x: (-x["month"], -x["last3"], -x["yest"]))
        out_products.append({
            "name": zh_name, "special": True, "pname": ["止回阀"], "code": [],
            "qty_yest": round(st["yest"], 2), "qty_last3": round(st["last3"], 2), "qty_last7": round(st["last7"], 2),
            "qty_month": round(st["month"], 2), "engineers": engineers,
        })
        verify.append({"name": zh_name, "code": [], "pname": ["止回阀"], "found": True,
                       "note": "特殊口径：CSM配件表 销售分类=止回阀 且 上缴金额>0（与防火阀占比模块同口径，剔除金额为 0/空/非数字 的止回阀；自购配件表无此品类）"})

    # 排序：当月数量降序
    out_products.sort(key=lambda p: (-p["qty_month"], p["name"]))

    tot_yest = sum(p["qty_yest"] for p in out_products)
    tot_last3 = sum(p["qty_last3"] for p in out_products)
    tot_last7 = sum(p["qty_last7"] for p in out_products)
    tot_month = sum(p["qty_month"] for p in out_products)
    sold = sum(1 for p in out_products if p["qty_month"] > 0)

    data = {
        "meta": {
            "截止日": CUT.strftime("%Y-%m-%d"),
            "服务月": (SVC_LO.strftime("%Y-%m-%d") + " ~ " + CUT.strftime("%Y-%m-%d")),
            "数据口径": "CSM配件表 + 工单自购配件明细",
            "口径": "数据口径：CSM配件表 + 工单自购配件明细。普通产品：CSM配件表(配件名称/配件编码 任一匹配) + 工单自购配件明细(物料名称/物料编码 任一匹配)，数量列求和；止回阀：CSM配件表 销售分类=止回阀 且 上缴金额(AB列)>0，数量列求和（与防火阀占比模块同口径，剔除金额为 0/空/非数字 的止回阀）。按服务工程师聚合。",
            "来源表": "CSM配件 + 工单自购配件明细",
            "时间窗口": {
                "昨日": CUT.strftime("%Y-%m-%d"),
                "近3日": (LAST3_LO.strftime("%Y-%m-%d") + " ~ " + CUT.strftime("%Y-%m-%d")),
                "近7日": (LAST7_LO.strftime("%Y-%m-%d") + " ~ " + CUT.strftime("%Y-%m-%d")),
                "当月": (SVC_LO.strftime("%Y-%m-%d") + " ~ " + CUT.strftime("%Y-%m-%d")),
            },
            "产品项数": len(out_products),
            "有销量项数": sold,
            "全区总计数_昨日": round(tot_yest, 2),
            "全区总计数_近3日": round(tot_last3, 2),
            "全区总计数_近7日": round(tot_last7, 2),
            "全区总计数_当月": round(tot_month, 2),
            "校验": verify,
        },
        "products": out_products,
    }
    with open(args.out, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=0)
    print("生成完成 ->", args.out)
    print("  产品项数:", len(out_products), " 有销量:", sold)
    print("  全区总计数 昨日=%.0f 近3日=%.0f 近7日=%.0f 当月=%.0f" % (tot_yest, tot_last3, tot_last7, tot_month))
    for p in out_products:
        print("   %-12s 当月=%-6.0f 近3日=%-6.0f 昨日=%-5.0f 工程师数=%d%s" % (
            p["name"], p["qty_month"], p["qty_last3"], p["qty_yest"],
            len(p["engineers"]), " [特殊]" if p["special"] else ""))
    print("  校验:")
    for v in verify:
        if v["note"]:
            print("    - %s: %s" % (v["name"], v["note"]))

if __name__ == "__main__":
    main()
