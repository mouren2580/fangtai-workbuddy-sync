#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
增值产品数量 数据生成器
================================
读 增值产品填写模板.xlsx 的产品目录（名称 / 产品名称 / 产品编码），
对 _新8月数据.xlsx 的 CSM配件表 计数，输出 _valueadded_data.json。

计数口径：
  普通产品：CSM配件表 配件编码 == 产品编码（含"尾随零归一化"匹配，处理如
            112500020018 → 1125000200180 这类截断），按 服务月 录入完成时间 过滤。
  止回阀  ：CSM配件表 销售分类(R列)=止回阀 且 上缴金额(AB列)>0 的行计数（与防火阀占比模块同口径，剔除金额为 0/空/非数字 的止回阀）。

按 办事处 / 服务网点 / 服务工程师 三维聚合，便于看板下钻。
"""
import openpyxl, json, argparse, re
from collections import defaultdict, OrderedDict
from datetime import datetime, date

EXCEL   = "_新8月数据.xlsx"
TEMPLATE= "增值产品填写模板.xlsx"
OUT     = "_valueadded_data.json"

CUT_LO  = date(2026, 7, 28)   # 服务月起始（8月 = 7.28–8.27）

# 用户 2026-08-26 确认：以下产品本月确无销售（非编码错误），计为 0 且不标红
CONFIRMED_ZERO = {"洗碗液", "排烟管弯头组件"}

def norm_code(s):
    """归一化编码：去空格；尾随零剥离（避免 112500020018 vs 1125000200180 误判）"""
    if s is None: return ""
    s = str(s).strip()
    if not s: return ""
    t = s.rstrip("0")
    return t if t else s

def in_period(d, cut_hi):
    if not isinstance(d, datetime): return False
    dd = d.date()
    return CUT_LO <= dd <= cut_hi

def amt_gt0(v):
    """止回阀过滤口径：上缴金额(AB列, 0-based 27) 为有效数字且 > 0 才计入；0/空/非数字剔除。"""
    try:
        f = float(v)
    except (TypeError, ValueError):
        return False
    return f > 0

def read_template():
    """按"名称"分组：同一显示名（如油杯垫 K40/K110）合并为一个产品目录项。"""
    wb = openpyxl.load_workbook(TEMPLATE, data_only=True)
    ws = wb.active
    groups = OrderedDict()
    for r in range(4, ws.max_row + 1):
        name = ws.cell(r, 2).value
        if name is None or str(name).strip() == "":
            continue
        name = str(name).strip()
        pname = ws.cell(r, 3).value
        code = ws.cell(r, 4).value
        pname = str(pname).strip() if pname else ""
        code = str(code).strip() if code else ""
        groups.setdefault(name, []).append({"pname": pname, "code": code})
    return groups

def main():
    global CUT_LO
    ap = argparse.ArgumentParser()
    ap.add_argument("--date", default="2026-08-24", help="统计截止日(冻结值)，与 --lo/--hi 二选一")
    ap.add_argument("--lo", help="服务月起始 YYYY-MM-DD（显式覆盖，优先）")
    ap.add_argument("--hi", help="服务月结束/截止 YYYY-MM-DD（显式覆盖，优先）")
    ap.add_argument("--excel", default=EXCEL)
    ap.add_argument("--template", default=TEMPLATE)
    ap.add_argument("--out", default=OUT)
    args = ap.parse_args()
    if args.lo and args.hi:
        CUT_LO = datetime.strptime(args.lo, "%Y-%m-%d").date()
        cut_hi = datetime.strptime(args.hi, "%Y-%m-%d").date()
    else:
        cut_hi = datetime.strptime(args.date, "%Y-%m-%d").date()

    groups = read_template()

    wb = openpyxl.load_workbook(args.excel, read_only=True, data_only=True)
    ws = wb["CSM配件"]
    rows = list(ws.iter_rows(min_row=3, values_only=True))

    # 预建 actual-code → [row] 索引（仅服务月内的行）
    code_index = defaultdict(list)
    zhihui_rows = []
    for r in rows:
        lc = r[36]  # 录入完成时间
        if not in_period(lc, cut_hi):
            continue
        code_index[str(r[13]).strip() if r[13] is not None else ""].append(r)
        if r[17] == "止回阀" and amt_gt0(r[27]):
            zhihui_rows.append(r)

    # 校验信息
    verify = []
    out_products = []

    # 普通产品计数（按名称合并多编码）
    for name, items in groups.items():
        if name == "止回阀":
            continue
        matched_rows = []
        matched_actual = []
        code_notes = []
        for it in items:
            tcode = it["code"]
            ncode = norm_code(tcode)
            if tcode and tcode in code_index:
                matched_rows.extend(code_index[tcode]); matched_actual.append(tcode)
            else:
                hit = False
                for ac, rs in code_index.items():
                    if ac and norm_code(ac) == ncode:
                        matched_rows.extend(rs); matched_actual.append(ac); hit = True
                if not hit:
                    # 注意：模板里登记了这个编码，只是本月 CSM配件表里没有对应记录。
                    # 措辞必须说清是"本月无记录"，不能说"未匹配"，否则会被误读成目录缺编码。
                    code_notes.append(tcode or "（空）")
        found = len(matched_rows) > 0
        office = defaultdict(int); net = defaultdict(int); tech = defaultdict(int)
        for r in matched_rows:
            office[r[2] or "（未知）"] += 1
            net[r[3] or "（未知）"] += 1
            tech[(r[5] or "（未知）") + (("（" + str(r[6]) + "）") if r[6] else "")] += 1
        total = len(matched_rows)
        note = ""
        level = "ok"
        if name in CONFIRMED_ZERO and not found:
            # 用户确认本月确无销售（非编码错误）：干净显示 0，不报警
            level = "ok"
            note = ""
            code_notes = []
        if code_notes:
            if found:
                note = "部分编码服务月内无记录（" + "、".join(code_notes) + "），已按有记录的编码计入"
                level = "info"
            else:
                note = "编码 " + "、".join(code_notes) + " 服务月内无记录，计为 0"
                level = "error"
        verify.append({"name": name, "code": [it["code"] for it in items],
                       "pname": [it["pname"] for it in items],
                       "found": found, "matched": sorted(set(matched_actual)),
                       "note": note, "level": level})
        out_products.append({
            "name": name,
            "code": [it["code"] for it in items],
            "pname": [it["pname"] for it in items],
            "special": False,
            "total": total,
            "office": dict(office), "net": dict(net), "tech": dict(tech),
            "verify": (note or None),
        })

    # 止回阀 特殊计数
    office = defaultdict(int); net = defaultdict(int); tech = defaultdict(int)
    for r in zhihui_rows:
        office[r[2] or "（未知）"] += 1
        net[r[3] or "（未知）"] += 1
        tech[(r[5] or "（未知）") + (("（" + str(r[6]) + "）") if r[6] else "")] += 1
    zh_total = len(zhihui_rows)
    verify.append({"name": "止回阀", "code": None, "pname": None, "found": True,
                   "matched": ["销售分类=止回阀"],
                   "note": "特殊口径：CSM配件表 R列销售分类=止回阀 且 上缴金额(AB列)>0 的行计数（与防火阀占比模块同口径，剔除金额为 0/空/非数字 的止回阀）",
                   "level": "special"})
    out_products.append({
        "name": "止回阀", "code": None, "pname": None, "special": True,
        "total": zh_total,
        "office": dict(office), "net": dict(net), "tech": dict(tech),
        "verify": "特殊口径：R列销售分类=止回阀 计数（与防火阀占比同口径）",
    })

    grand = sum(p["total"] for p in out_products)
    sold = sum(1 for p in out_products if p["total"] > 0)

    data = {
        "meta": {
            "截止日": cut_hi.isoformat(),
            "月份": cut_hi.strftime("%Y-%m"),
            "服务月": "%s ~ %s" % (CUT_LO.isoformat(), cut_hi.isoformat()),
            "来源表": "CSM配件",
            "数据口径": "CSM配件表",
            "口径": "数据口径：CSM配件表。普通产品按 配件编码=产品编码 计数（服务月内）；止回阀按 销售分类=止回阀 且 上缴金额(AB列)>0 计数（与防火阀占比模块同口径，剔除金额为 0/空/非数字 的止回阀）",
            "产品项数": len(out_products),
            "有销量项数": sold,
            "全区总计数": grand,
            "校验": verify,
        },
        "products": out_products,
    }
    with open(args.out, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=0)
    print("生成完成 →", args.out)
    print("  产品项数:", len(out_products), " 有销量:", sold, " 全区总计数:", grand)
    print("  止回阀:", zh_total)
    for p in out_products:
        flag = "  [i] 见校验明细" if p["verify"] and not p["special"] else ""
        print("   %-10s 计数=%-5d%s" % (p["name"], p["total"], flag))
    print("  校验明细:")
    for v in verify:
        if v["note"]:
            print("    - %s: %s" % (v["name"], v["note"]))

if __name__ == "__main__":
    main()
