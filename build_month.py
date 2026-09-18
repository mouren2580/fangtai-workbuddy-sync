# -*- coding: utf-8 -*-
"""按参考看板逻辑，用新 Excel 重建 2026-09 全部数据块并注入。

  python build_month.py check 2026-09-12     # 校验：全部块与参考对照
  python build_month.py build 2026-09-14     # 构建：注入并写出
"""
import os, re, sys, json, datetime
from collections import OrderedDict
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import gen_v2_lib as G
import gen_month as GM

HERE = os.path.dirname(os.path.abspath(__file__))
# 底板 = 用户最新的参考看板（每次更新时用新收到的版本覆盖本文件）
REF = os.path.join(HERE, "dashboard_ref.html")
EXCEL = os.path.join(HERE, "2026年9月西北服务产品(3).xlsx")
OUT = os.path.join(HERE, "dashboard_offline.html")

# 需要「保留历史月」的多月份容器块
MONTH_BLOCKS = ["MONTHLY_FOUR", "EXTEND_TIME_DATA", "CLEANING_DATA"]
MDATA_RE = re.compile(r'MDATA_(20\d\d)_(\d\d)_([DTBWE])\s*=')

BS = chr(92); Q = chr(34)


def extract(s, var):
    m = re.search(r'(?m)^(?:const|let|var)\s+' + re.escape(var) + r'\s*=\s*', s)
    if not m:
        return None
    i = m.end()
    if s[i] not in '{[':
        j = s.find(';', i); return s[i:j].strip()
    depth = 0; instr = False; esc = False; j = i
    while j < len(s):
        c = s[j]
        if instr:
            if esc: esc = False
            elif c == BS: esc = True
            elif c == Q: instr = False
        else:
            if c == Q: instr = True
            elif c in '{[': depth += 1
            elif c in '}]':
                depth -= 1
                if depth == 0: break
        j += 1
    return s[i:j + 1]


def replace_value(html, var, obj):
    m = re.search(r'(?m)^(const|let)(\s+)' + re.escape(var) + r'\s*=\s*', html)
    if not m:
        raise SystemExit("未找到声明: " + var)
    i = m.end()
    if html[i] not in '{[':
        j = html.find(';', i)
        return html[:i] + json.dumps(obj, ensure_ascii=False, separators=(",", ":")) + html[j:]
    depth = 0; instr = False; esc = False; j = i
    while j < len(html):
        c = html[j]
        if instr:
            if esc: esc = False
            elif c == BS: esc = True
            elif c == Q: instr = False
        else:
            if c == Q: instr = True
            elif c in '{[': depth += 1
            elif c in '}]':
                depth -= 1
                if depth == 0: break
        j += 1
    return html[:i] + json.dumps(obj, ensure_ascii=False, separators=(",", ":")) + html[j + 1:]


def make_extend(rec_et, cut):
    """由已按截止日过滤的延保明细聚合出 E 块（保证与 EXTEND_TIME_DATA 完全同口径）"""
    techs, offices, total = OrderedDict(), set(), {
        "y1": {"qty": 0, "amt": 0.0}, "y2": {"qty": 0, "amt": 0.0}, "y3": {"qty": 0, "amt": 0.0}}
    for r in rec_et:
        key = (r["t"], r["c"])
        t = techs.get(key)
        if t is None:
            t = techs[key] = {"id": r["t"], "name": r["t"], "office": r["o"],
                              "y1": {"qty": 0, "amt": 0.0}, "y2": {"qty": 0, "amt": 0.0},
                              "y3": {"qty": 0, "amt": 0.0}}
        if r["o"]:
            offices.add(r["o"])
            if not t["office"]:
                t["office"] = r["o"]
        t[r["yr"]]["qty"] += 1
        t[r["yr"]]["amt"] += r["amt"]
        total[r["yr"]]["qty"] += 1
        total[r["yr"]]["amt"] += r["amt"]
    lst = list(techs.values())
    for t in lst:
        t["tqty"] = t["y1"]["qty"] + t["y2"]["qty"] + t["y3"]["qty"]
        # 逐年金额保留浮点原值（不四舍五入），与参考看板一致
        t["tamt"] = round(t["y1"]["amt"] + t["y2"]["amt"] + t["y3"]["amt"], 2)
    for k in ("y1", "y2", "y3"):
        total[k]["amt"] = total[k]["amt"]
    lst.sort(key=lambda x: -x["tamt"])
    return {"meta": {"来源": "CSM服务项目", "筛选": "销售分类=延保服务", "金额口径": "上缴金额",
                     "年限区分": "服务收费项目列(含一年/两年/三年)",
                     "截止": cut.strftime("%Y-%m-%d") + "（同主数据）",
                     "工程师数": len(lst)},
            "offices": sorted(offices), "techs": lst, "total": total}


def make_blocks(cut):
    MONTH = cut.strftime("%Y-%m")
    svc = G.svc_start(cut.month, cut.year)
    end = svc + datetime.timedelta(days=30)

    G.verify_columns(EXCEL)          # 列序自检：平台导出偶会调整列顺序
    wo = G.build_workorder(EXCEL, cut, True)
    va = G.build_valueadded(EXCEL, cut)
    vp = G.build_valueparts(EXCEL, cut)
    bc4, t_clean, t_excl, t_big = G.build_bigcare4(EXCEL, cut, True)
    rec_et = G.build_extend_time(EXCEL, cut)
    rec_cl = G.build_cleaning(EXCEL, cut)

    # --- MDATA 系列（复用 gen_month 的 D/T/W/E；B 用 bc4 保证结构一致）---
    GM.CUTOFF = cut
    GM.MONTH = MONTH
    GM.EXCEL = EXCEL
    D = GM.build_main(EXCEL)
    D["meta"]["_stamp"] = cut.strftime("%Y-%m-%d")
    T, _ = GM.build_tech(EXCEL)
    W = GM.gen_weekly(EXCEL, end, cut)
    W["meta"] = {"sample": False,
                 "note": "真实周数据：清洗保养/延保服务取自CSM服务项目(录入完成时间,上缴金额)；"
                         "清洁耗材取自CSM配件(录入完成时间,上缴金额)+WMS网点买断配件明细(发货时间,实收/上缴金额)。"
                         "网点维度为源表维修商/经销部名。周期为服务月 %s–%s。"
                         % (svc.strftime("%Y-%m-%d"), end.strftime("%Y-%m-%d"))}
    E = make_extend(rec_et, cut)
    # B：MDATA 版（无「周期」字段，截止日=cutoff）
    B = {"meta": {k: bc4["meta"][k] for k in ["截止日", "来源表", "清洗保养总单", "剔除项",
                                             "大保养项", "分母", "大保养总数", "全区占比"]},
         "office": bc4["office"], "net": bc4["net"], "eng": bc4["eng"]}
    B["meta"]["截止日"] = cut.strftime("%Y-%m-%d")

    # --- MONTHLY_FOUR 2026-09（bigcare 版截止日=服务月末，含「周期」）---
    bc_mf = dict(bc4)
    bc_mf.pop("enc", None)
    bc_mf["meta"] = dict(bc4["meta"])
    bc_mf["meta"]["截止日"] = end.strftime("%Y-%m-%d")
    for k in ("剔除项", "大保养项"):        # MF 版用源表首现序（与 MDATA_*_B 的常量序不同）
        bc_mf["meta"][k] = bc4["enc"][k]

    # 月份内：workorder 的「截止」标签沿用服务月末（与参考一致）
    mf_month = {"workorder": wo, "valueadded": va, "valueparts": vp, "bigcare": bc_mf}

    # --- EXTEND_TIME_DATA / CLEANING_DATA ---
    et = {"months": {"2026-09": rec_et, "2026-08": None}}
    cl_data = {"months": {MONTH: rec_cl},
               "meta": {"来源表": "CSM服务项目", "筛选": "销售分类=清洗保养",
                        "生成": datetime.datetime.now().strftime("%Y-%m-%d %H:%M"),
                        "总记录": len(rec_cl), "主项": list(G.CLEAN_MAIN)}}
    return {"D": D, "T": T, "B": B, "W": W, "E": E,
            "mf_month": mf_month, "bc4": bc4,
            "rec_et": rec_et, "rec_cl": rec_cl,
            "cut": cut, "svc": svc, "end": end}


def summary(bl, tag):
    D, T, B, W, E = bl["D"], bl["T"], bl["B"], bl["W"], bl["E"]
    print("[%s] 截止=%s  周期=%s ~ %s" % (tag, bl["cut"], bl["svc"], bl["end"]))
    print("  主数据D      : 网点=%d  合计=¥%s  time.日=%s" %
          (len(D["网点"]), format(round(sum(n["合计"] for n in D["网点"]), 2), ","), D["time"]["日"]))
    print("  工程师T      : %d 人  工单消耗=¥%s  报表总合计=¥%s" %
          (T["count"], format(T["total"], ","), format(T["meta"]["报表总合计"], ",")))
    m = B["meta"]
    print("  大保养B      : 总单=%d 分母=%d 大保养=%d 占比=%.2f%%" %
          (m["清洗保养总单"], m["分母"], m["大保养总数"], m["全区占比"]))
    print("  周数据W      : %d 周  %s" % (len(W["weeks"]), [x["label"] for x in W["weeks"]]))
    t = E["total"]
    print("  延保E        : %d 人  y1=%d单/¥%s y2=%d单/¥%s y3=%d单/¥%s" %
          (E["meta"]["工程师数"], t["y1"]["qty"], t["y1"]["amt"], t["y2"]["qty"],
           t["y2"]["amt"], t["y3"]["qty"], t["y3"]["amt"]))
    w = bl["mf_month"]["workorder"]
    print("  工单workorder: 总工单=%d 网点=%d 工程师=%d 止回阀=%s 烟机安装=%s 占比=%s%%" %
          (w["meta"]["总工单数"], w["meta"]["网点数"], w["meta"]["工程师数"],
           w["meta"]["止回阀总数"], w["meta"]["烟机安装量"], w["meta"]["防火阀占比"]))
    va = bl["mf_month"]["valueadded"]
    print("  增值产品     : 全区总计数=%s  有销量=%s/12" %
          (va["meta"]["全区总计数"], va["meta"]["有销量项数"]))
    vp = bl["mf_month"]["valueparts"]
    print("  增值配件     : 昨日=%s 近3日=%s 近7日=%s 当月=%s" %
          (vp["meta"]["全区总计数_昨日"], vp["meta"]["全区总计数_近3日"],
           vp["meta"]["全区总计数_近7日"], vp["meta"]["全区总计数_当月"]))
    print("  延保明细     : %d 条  ¥%s" % (len(bl["rec_et"]), round(sum(r["amt"] for r in bl["rec_et"]), 2)))
    print("  清洁保养明细 : %d 条" % len(bl["rec_cl"]))


def check(cut):
    ref = open(REF, encoding="utf-8").read()
    bl = make_blocks(cut)
    summary(bl, "check")
    print("\n===== 与参考看板逐块对照 =====")
    rf = json.loads(extract(ref, "MONTHLY_FOUR"))
    ok = True
    for var in ["MDATA_2026_09_D", "MDATA_2026_09_T", "MDATA_2026_09_B", "MDATA_2026_09_W", "MDATA_2026_09_E"]:
        mine = bl[var[10:]] if var[10:] in bl else None
        r = json.loads(extract(ref, var))
        mm = {"D": bl["D"], "T": bl["T"], "B": bl["B"], "W": bl["W"], "E": bl["E"]}[var[-1]]
        same = json.dumps(r, sort_keys=True) == json.dumps(mm, sort_keys=True)
        ok &= same
        print("  %-20s %s" % (var, "IDENTICAL" if same else "DIFFERS"))
        if not same:
            rk, mk = set(r.keys()), set(mm.keys())
            if rk != mk:
                print("      keys ref=", sorted(rk), " mine=", sorted(mk))
            for k in rk & mk:
                if json.dumps(r[k], sort_keys=True) != json.dumps(mm[k], sort_keys=True):
                    print("      [%s] ref=%s" % (k, json.dumps(r[k], ensure_ascii=False)[:180]))
                    print("      [%s] min=%s" % (k, json.dumps(mm[k], ensure_ascii=False)[:180]))
    for mod in ["workorder", "valueadded", "valueparts", "bigcare"]:
        r = rf["months"]["2026-09"][mod]
        m = bl["mf_month"][mod]
        same = json.dumps(r, sort_keys=True) == json.dumps(m, sort_keys=True)
        ok &= same
        print("  MONTHLY_FOUR.%-11s %s" % (mod, "IDENTICAL" if same else "DIFFERS"))
        if not same:
            for k in set(r.keys()) & set(m.keys()):
                if json.dumps(r[k], sort_keys=True) != json.dumps(m[k], sort_keys=True):
                    print("      [%s] ref=%s" % (k, json.dumps(r[k], ensure_ascii=False)[:200]))
                    print("      [%s] min=%s" % (k, json.dumps(m[k], ensure_ascii=False)[:200]))
    ret = json.loads(extract(ref, "EXTEND_TIME_DATA"))
    same = json.dumps(ret["months"]["2026-09"], sort_keys=True) == json.dumps(bl["rec_et"], sort_keys=True)
    ok &= same
    print("  EXTEND_TIME_DATA 2026-09  %s" % ("IDENTICAL" if same else "DIFFERS"))
    rcl = json.loads(extract(ref, "CLEANING_DATA"))
    same = json.dumps(rcl["months"]["2026-09"], sort_keys=True) == json.dumps(bl["rec_cl"], sort_keys=True)
    ok &= same
    print("  CLEANING_DATA 2026-09     %s" % ("IDENTICAL" if same else "DIFFERS"))
    print("\n>>> 总体:", "全部一致 ✅" if ok else "存在差异 ⚠")
    return ok


def merge_retained(html, prev):
    """把上一版中存在、而新底版中缺失的历史月份块补回来，避免换底版时丢历史月。"""
    if not prev:
        return html, []
    kept = []
    # 1) 多月份容器：补缺月
    for var in MONTH_BLOCKS:
        cur = json.loads(extract(html, var))
        old = json.loads(extract(prev, var))
        for mk, mv in (old.get("months") or {}).items():
            if mk not in (cur.get("months") or {}):
                cur.setdefault("months", {})[mk] = mv
                kept.append("%s.months[%s]" % (var, mk))
        html = replace_value(html, var, cur)
    # 2) 顶层 MDATA_YYYY_MM_* 声明：补缺块
    cur_names = set(re.findall(r'(?m)^(?:const|let)\s+(MDATA_20\d\d_\d\d_[DTBWE])\s*=', html))
    for m in re.finditer(r'(?m)^(?:const|let)\s+(MDATA_20\d\d_\d\d_[DTBWE])\s*=\s*', prev):
        name = m.group(1)
        if name in cur_names:
            continue
        val = extract(prev, name)
        anchor = re.search(r'(?m)^(?:const|let)\s+MDATA_20\d\d_\d\d_[DTBWE]\s*=', html)
        if anchor:
            html = html[:anchor.start()] + "const %s = %s;\n" % (name, val) + html[anchor.start():]
            cur_names.add(name)
            kept.append(name)
    return html, kept


def build(cut):
    if not os.path.exists(REF):
        raise SystemExit("参考看板不存在: " + REF)
    html = open(REF, encoding="utf-8").read()
    prev = open(OUT, encoding="utf-8").read() if os.path.exists(OUT) else None
    html, kept = merge_retained(html, prev)
    if kept:
        print("[保留历史月] " + ", ".join(kept))
    bl = make_blocks(cut)
    summary(bl, "build")

    for var, key in [("MDATA_2026_09_D", "D"), ("MDATA_2026_09_T", "T"),
                     ("MDATA_2026_09_B", "B"), ("MDATA_2026_09_W", "W"),
                     ("MDATA_2026_09_E", "E")]:
        html = replace_value(html, var, bl[key])
    # MONTHLY_FOUR：仅替换 2026-09 分支，保留 1-8 月与 year
    mf = json.loads(extract(html, "MONTHLY_FOUR"))
    mf["months"]["2026-09"] = bl["mf_month"]
    html = replace_value(html, "MONTHLY_FOUR", mf)
    # EXTEND_TIME_DATA
    et = json.loads(extract(html, "EXTEND_TIME_DATA"))
    et["months"]["2026-09"] = bl["rec_et"]
    html = replace_value(html, "EXTEND_TIME_DATA", et)
    # CLEANING_DATA
    cl = json.loads(extract(html, "CLEANING_DATA"))
    cl["months"]["2026-09"] = bl["rec_cl"]
    cl["meta"]["生成"] = datetime.datetime.now().strftime("%Y-%m-%d %H:%M")
    cl["meta"]["总记录"] = len(bl["rec_cl"])
    html = replace_value(html, "CLEANING_DATA", cl)

    build_id = cut.strftime("%Y%m%d").replace("2026", "2026")  # placeholder
    build_id = datetime.datetime.now().strftime("%Y%m%d-%H%M")
    html = re.sub(r'var BUILD = "[^"]+";', 'var BUILD = "%s";' % build_id, html, count=1)

    open(OUT, "w", encoding="utf-8").write(html)
    vj = {"v": build_id, "cut": cut.strftime("%Y-%m-%d"),
          "t": datetime.datetime.now().strftime("%Y-%m-%dT%H:%M:%S")}
    open(os.path.join(HERE, "version.json"), "w", encoding="utf-8").write(
        json.dumps(vj, ensure_ascii=False))

    # 写后校验：每个块可 JSON 解析
    chk = open(OUT, encoding="utf-8").read()
    for var in ["MDATA_2026_09_D", "MDATA_2026_09_T", "MDATA_2026_09_B", "MDATA_2026_09_W",
                "MDATA_2026_09_E", "MONTHLY_FOUR", "EXTEND_TIME_DATA", "CLEANING_DATA"]:
        json.loads(extract(chk, var))
    print("\n[OK] 已写入 %s   BUILD=%s  截止=%s" % (OUT, build_id, cut))
    return build_id


if __name__ == "__main__":
    mode = sys.argv[1] if len(sys.argv) > 1 else "check"
    ds = sys.argv[2] if len(sys.argv) > 2 else "2026-09-12"
    y, m, d = [int(x) for x in ds.split("-")]
    if mode == "check":
        check(datetime.date(y, m, d))
    else:
        build(datetime.date(y, m, d))
