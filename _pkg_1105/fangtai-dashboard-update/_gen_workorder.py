# -*- coding: utf-8 -*-
"""工单量统计生成器：读 Excel 第6个 sheet『工单』，按当前服务月周期过滤，
   按 办事处→网点→服务工程师 三级维度统计工单数量（工单编号去重计数），
   并提供 产品组 / 工单类型 / 服务项目 三个附加维度的计数排名。

   新增「防火阀占比」：
     止回阀数量 = CSM配件表(第3个sheet, 表头在第2行) 中 销售分类(R列)=='止回阀' 且 上缴金额(AB列)>0 的 数量(S列) 之和；
                 仅统计工单编号落在服务月周期内（与工单表同源、周期一致）；
                 金额为 0 / 空 / 非数字的止回阀 一律剔除（与增值模块同口径）。
     烟机安装量 = 工单表中 产品组=='吸油烟机' 且 服务项目 in ('安装','改装','油烟机预埋烟管') 的工单量。
     防火阀占比 = 止回阀数量 / 烟机安装量 (百分比，小数点后两位)，可按 办事处/网点/工程师 三维度查看。

   用法: python _gen_workorder.py <月报.xlsx> --date YYYY-MM-DD --out _workorder.json
"""
import sys, os, json, argparse, datetime, collections
for s in (sys.stdout, sys.stderr):
    try: s.reconfigure(encoding="utf-8", errors="replace")
    except Exception: pass
import openpyxl

# 本地优先：脚本同目录若有 build_dashboard.py 就用它（家里电脑没有 WorkBuddy 技能目录也能跑）；
# 找不到再回退到办公室技能目录。
_HERE = os.path.dirname(os.path.abspath(__file__))
SKILL = _HERE if os.path.exists(os.path.join(_HERE, "build_dashboard.py")) \
        else r"C:/Users/ABC/.workbuddy/skills/fangtai-service-achievement-dashboard/scripts"
sys.path.insert(0, SKILL)
from build_dashboard import service_period


def service_month_of(d):
    """截止日 → 服务月 (年,月)。日 >= 28 归入下一服务月，与 _gen_month.py 口径一致。
    2026-09-28 → (2026, 10)；2026-09-27 → (2026, 9)。
    """
    if d.day >= 28:
        return (d.year + 1, 1) if d.month == 12 else (d.year, d.month + 1)
    return d.year, d.month

# 工单表列索引（0-based）：★优先按表头名称解析★，找不到时才回退历史固定列号。
# 背景（2026-09-19）：源表「工单」sheet 在 F–AF 区间发生列增删重排（原 F 列「上门鉴定费」被移除，
#   新增「办事处ID/办事处名称」等），导致 L 列由「办事处名称」变成「渠道来源」、
#   工程师编号由 30 变 31。写死列号会让 6 个办事处全变「（未知办事处）」。
#   故改为按表头名称定位，源表列序再变也不会错位。
COL_FALLBACK = {        # 历史列号（2026-08 版表；仅作名称缺失时的兜底）
    "办事处名称": 11,   # 旧 L
    "服务网点":   79,   # CB
    "服务工程师": 55,   # BD
    "工程师编号": 30,   # 旧 AE（新表已是 31）
    "工单编号":   53,
    "服务完成时间": 41, # AP
    "产品组":     73,   # BV
    "工单类型":   65,   # BN
    "服务项目":   57,   # BF
}
# 兼容旧常量名（其它脚本/诊断可能引用）
COL_OFFICE = COL_FALLBACK["办事处名称"]
COL_NET    = COL_FALLBACK["服务网点"]
COL_ENG    = COL_FALLBACK["服务工程师"]
COL_ENGID  = COL_FALLBACK["工程师编号"]
COL_WOID   = COL_FALLBACK["工单编号"]
COL_TIME   = COL_FALLBACK["服务完成时间"]
COL_PGROUP = COL_FALLBACK["产品组"]
COL_WOTYPE = COL_FALLBACK["工单类型"]
COL_SITEM  = COL_FALLBACK["服务项目"]


def resolve_cols(ws):
    """按表头名称解析所需列的下标；缺名时回退 COL_FALLBACK。返回 (colmap, header_names)"""
    try:
        hdr = next(ws.iter_rows(min_row=1, max_row=1, values_only=True), ())
    except StopIteration:
        hdr = ()
    names = [str(v).strip() if v is not None else "" for v in hdr]
    col, missing = {}, []
    for key, fb in COL_FALLBACK.items():
        if key in names:
            col[key] = names.index(key)
        else:
            col[key] = fb
            missing.append(key)
    if missing:
        print("   [警告] 表头未找到列 %s，已回退历史列号" % "、".join(missing))
    return col, names

# 烟机安装量口径
PGROUP_HOOD = '吸油烟机'          # 产品组=烟机
SITEM_INSTALL = ('安装', '改装', '油烟机预埋烟管')  # 服务项目=安装/改装/油烟机预埋烟管

# CSM配件表（第3个sheet），表头在第2行（第1行为空）
CSM_SHEET = "CSM配件"
CSM_HEADER_ROW = 2                 # 1-based
CSM_COL_WOID   = 7                 # H  工单编号 (0-based 7)
CSM_COL_SALES  = 17                # R  销售分类 (0-based 17)
CSM_COL_QTY    = 18                # S  数量     (0-based 18)
CSM_COL_AMT    = 27                 # AB 上缴金额 (0-based 27)
SALES_CHECKVALVE = '止回阀'         # 销售分类=止回阀

def to_date(v):
    if isinstance(v, datetime.datetime): return v.date()
    if isinstance(v, datetime.date): return v
    return None

def tally(d, key):
    if key not in d: d[key] = 0
    d[key] += 1

def amt_gt0(v):
    """止回阀过滤口径：上缴金额 为有效数字且 > 0 才计入；0/空/非数字一律剔除。"""
    try:
        f = float(v)
    except (TypeError, ValueError):
        return False
    return f > 0

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("excel")
    ap.add_argument("--date", help="数据截止日 YYYY-MM-DD（决定服务月周期 y/m；与 --ps/--pe 二选一）")
    ap.add_argument("--ps", help="服务月起始 YYYY-MM-DD（显式覆盖，优先于 --date）")
    ap.add_argument("--pe", help="服务月结束 YYYY-MM-DD（显式覆盖，优先于 --date）")
    ap.add_argument("--out", default="_workorder.json")
    ap.add_argument("--cut", help="统计截止日 YYYY-MM-DD（仅决定 meta['截止'] 展示值；不给则取 --date，或退回服务月结束日）")
    a = ap.parse_args()
    if a.ps and a.pe:
        ps = datetime.datetime.strptime(a.ps, "%Y-%m-%d").date()
        pe = datetime.datetime.strptime(a.pe, "%Y-%m-%d").date()
        cut = datetime.datetime.strptime(a.cut, "%Y-%m-%d").date() if a.cut else pe
    else:
        y, m, d = map(int, a.date.split("-"))
        as_of = datetime.date(y, m, d)
        sy, sm = service_month_of(as_of)          # 截止日 → 服务月（月末 28 日后归下一个月）
        ps, pe = service_period(sy, sm)           # 2026-10 -> 2026-09-28 ~ 2026-10-27
        cut = as_of                               # 数据实际覆盖到哪天（≠ 周期结束日，如服务月首周）
        print("   截止日 %s → 服务月 %04d-%02d  周期 %s ~ %s" % (as_of, sy, sm, ps, pe))
    wb = openpyxl.load_workbook(a.excel, read_only=True, data_only=True)

    # ---------- 第一遍：工单表 ----------
    ws = wb["工单"]
    col, _hnames = resolve_cols(ws)
    print("   列映射(按表头解析): " + " | ".join("%s=#%d" % (k, v) for k, v in col.items()))
    c_off, c_net, c_eng = col["办事处名称"], col["服务网点"], col["服务工程师"]
    c_engid, c_woid, c_time = col["工程师编号"], col["工单编号"], col["服务完成时间"]
    c_pg, c_wt, c_si = col["产品组"], col["工单类型"], col["服务项目"]
    max_idx = max(col.values())
    offices = collections.OrderedDict()
    pgroup = collections.OrderedDict()
    wotype = collections.OrderedDict()
    sitem = collections.OrderedDict()
    total = 0
    woset = set()
    valid = {}          # woid -> (office, net, eng, engid)
    tech_node = {}      # (office,net,eng,engid) -> tech dict
    total_yj = 0
    for r in ws.iter_rows(min_row=2, values_only=True):
        if len(r) <= max_idx:
            continue
        woid = r[c_woid]
        if not woid:
            continue
        wkey = str(woid)
        if wkey in woset:        # 工单编号去重
            continue
        dt = to_date(r[c_time])
        if dt is None:
            continue
        if not (ps <= dt <= pe):  # 仅统计服务月周期内的工单
            continue
        woset.add(wkey)
        office = str(r[c_off]).strip() if r[c_off] else "（未知办事处）"
        net = str(r[c_net]).strip() if r[c_net] else "（未知网点）"
        eng = str(r[c_eng]).strip() if r[c_eng] else "（未知工程师）"
        engid = str(r[c_engid]).strip() if r[c_engid] else ""
        valid[wkey] = (office, net, eng, engid)

        if office not in offices:
            offices[office] = {"cnt": 0, "zh": 0, "yj": 0, "nets": collections.OrderedDict()}
        o = offices[office]
        o["cnt"] += 1
        if net not in o["nets"]:
            o["nets"][net] = {"cnt": 0, "zh": 0, "yj": 0, "techs": collections.OrderedDict()}
        n = o["nets"][net]
        n["cnt"] += 1
        tk = (eng, engid)
        if tk not in n["techs"]:
            n["techs"][tk] = {"name": eng, "code": engid, "cnt": 0, "zh": 0, "yj": 0}
        t = n["techs"][tk]
        t["cnt"] += 1
        tech_node[(office, net, eng, engid)] = t

        # 烟机安装量
        pg = str(r[c_pg]).strip() if r[c_pg] else ""
        wt = str(r[c_wt]).strip() if r[c_wt] else ""
        si = str(r[c_si]).strip() if r[c_si] else ""
        if pg == PGROUP_HOOD and si in SITEM_INSTALL:
            o["yj"] += 1; n["yj"] += 1; t["yj"] += 1; total_yj += 1
        # 附加维度（每张工单记一次）
        if not pg: pg = "（未填产品组）"
        if not wt: wt = "（未填工单类型）"
        if not si: si = "（未填服务项目）"
        tally(pgroup, pg); tally(wotype, wt); tally(sitem, si)
        total += 1

    # ---------- 第二遍：CSM配件表（止回阀数量，按工单编号 join 周期）----------
    total_zh = 0
    try:
        csm = wb[CSM_SHEET]
        for r in csm.iter_rows(min_row=CSM_HEADER_ROW + 1, values_only=True):
            if len(r) <= CSM_COL_AMT:
                continue
            sales = r[CSM_COL_SALES]
            if str(sales).strip() != SALES_CHECKVALVE:
                continue
            amt = r[CSM_COL_AMT]
            if not amt_gt0(amt):   # 上缴金额=0/空/非数字 的止回阀剔除
                continue
            woid = r[CSM_COL_WOID]
            if not woid:
                continue
            wkey = str(woid)
            if wkey not in valid:   # 仅统计服务月周期内工单对应的止回阀
                continue
            qty = r[CSM_COL_QTY]
            try:
                qty = float(qty)
            except (TypeError, ValueError):
                qty = 0
            if qty <= 0:
                continue
            if float(qty).is_integer():
                qty = int(qty)
            office, net, eng, engid = valid[wkey]
            t = tech_node.get((office, net, eng, engid))
            if not t:
                continue
            t["zh"] += qty
            total_zh += qty
    finally:
        wb.close()

    # ---------- 汇总：把 tech 的 zh 上卷到 net / office（先 net 后 office）----------
    for o in offices.values():
        for n in o["nets"].values():
            n["zh"] = sum(t["zh"] for t in n["techs"].values())
        o["zh"] = sum(n["zh"] for n in o["nets"].values())

    # ---------- 输出结构 ----------
    out_offices = []
    for oname, o in offices.items():
        nets = []
        for nname, n in o["nets"].items():
            techs = [{"name": t["name"], "code": t["code"], "cnt": t["cnt"],
                      "zh": t["zh"], "yj": t["yj"]}
                     for tk, t in n["techs"].items()]
            techs.sort(key=lambda x: -x["cnt"])
            nets.append({"name": nname, "cnt": n["cnt"], "zh": n["zh"], "yj": n["yj"], "techs": techs})
        nets.sort(key=lambda x: -x["cnt"])
        out_offices.append({"name": oname, "cnt": o["cnt"], "zh": o["zh"], "yj": o["yj"], "nets": nets})
    out_offices.sort(key=lambda x: -x["cnt"])
    def rank(d):
        return sorted(({"name": k, "cnt": v} for k, v in d.items()), key=lambda x: -x["cnt"])
    out_pgroup = rank(pgroup)
    out_wotype = rank(wotype)
    out_sitem = rank(sitem)
    net_total = sum(len(o["nets"]) for o in offices.values())
    eng_total = sum(len(n["techs"]) for o in offices.values() for n in o["nets"].values())
    overall_ratio = (total_zh / total_yj * 100) if total_yj else 0.0
    meta = {
        "来源": "工单表(第6个sheet) + CSM配件表(第3个sheet)",
        "周期": "%s~%s" % (ps.isoformat(), pe.isoformat()),
        "截止": cut.isoformat(),
        "时间口径": "服务完成时间",
        "总工单数": total,
        "办事处数": len(out_offices),
        "网点数": net_total,
        "工程师数": eng_total,
        "维度": "办事处→网点→服务工程师（工单编号去重计数）",
        "产品组数": len(out_pgroup),
        "工单类型数": len(out_wotype),
        "服务项目数": len(out_sitem),
        "止回阀总数": total_zh,
        "烟机安装量": total_yj,
        "防火阀占比": round(overall_ratio, 2),
    }
    out = {"meta": meta, "office": out_offices,
           "pgroup": out_pgroup, "wotype": out_wotype, "sitem": out_sitem}
    json.dump(out, open(a.out, "w", encoding="utf-8"), ensure_ascii=False, separators=(",", ":"))
    print("✅ 工单统计: 周期%s~%s 总工单=%d 办事处=%d 网点=%d 工程师=%d"
          % (ps.isoformat(), pe.isoformat(), total, len(out_offices), net_total, eng_total))
    print("   防火阀占比: 止回阀=%d / 烟机安装量=%d = %.2f%%"
          % (total_zh, total_yj, overall_ratio))
    print("   各办事处工单数: " + " | ".join("%s=%d" % (o["name"], o["cnt"]) for o in out_offices))
    print("   各办防火阀占比: " + " | ".join("%s=%.2f%%" % (o["name"], (o["zh"]/o["yj"]*100 if o["yj"] else 0)) for o in out_offices))
    print("→ 写出 %s" % a.out)

if __name__ == "__main__":
    main()
