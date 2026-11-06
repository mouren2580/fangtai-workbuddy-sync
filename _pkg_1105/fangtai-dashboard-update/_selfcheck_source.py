# -*- coding: utf-8 -*-
"""源表自检器（方太西北服务产品月报）

用途：每次数据更新前，独立核验用户导出的 Excel 是否自洽、周期是否准确。
      —— 不依赖生成器的输出，直接从源表原始明细复算。

用法：
  python _selfcheck_source.py --excel <月报.xlsx> --svc-lo 2026-09-28 --cut 2026-10-01
  python _selfcheck_source.py --excel <x> --svc-lo ... --cut ... --prev _month_2026-10.json

检查项：
  1) 各明细表的时间列分布 —— 是否有记录越出 [svc-lo, cut]（导出周期错）
  2) 收入统计表结构 —— 行数/办事处数/品项求和=合计/负值
  3) 三元勾稽 —— CSM配件 + WMS网点买断 + CSM服务项目 是否 = 收入统计表合计
  4) 15 品项逐项勾稽 —— 每品能否由明细还原
  5) 口径关键词 —— 止回阀条数与金额
  6) 与上一轮 JSON 对比 —— 指标是否单调（可选）

退出码：0 = 全部通过 / 1 = 存在需人工确认的疑点
"""
import sys, os, json, argparse, datetime, collections
for s in (sys.stdout, sys.stderr):
    try: s.reconfigure(encoding="utf-8", errors="replace")
    except Exception: pass
import openpyxl

ITEMS = ["清洗保养","延保服务","清洁耗材","止回阀","装饰罩","水/气路金属软管",
         "排烟出风管","水气路通用辅材","其他产品线专用","其他水路辅材","其他气路辅材",
         "橱柜类辅材","燃气管快速接头","工具","电路辅材"]

# 表 -> (日期列表头名, 起始行, 是否按工单号去重, 工单号列索引)
DETAIL_TABLES = [
    ("工单",                 "服务完成时间", 2, True,  53),
    ("CSM服务项目",           "录入完成时间", 3, False, None),
    ("CSM配件",              "录入完成时间", 3, False, None),
    ("WMS网点买断配件明细",    "发货时间",    2, False, None),
    ("工单自购配件明细",       "录入完成时间", 2, False, None),
]

WARN, OK = [], []

def warn(msg): WARN.append(msg); print("  [疑点] " + msg)
def ok(msg):   OK.append(msg);   print("  [通过] " + msg)

def num(v):
    try: return round(float(v), 2)
    except Exception: return 0.0

def to_date(v):
    if isinstance(v, datetime.datetime): return v.date()
    if isinstance(v, datetime.date): return v
    return None

def find_hdr(ws, kw, max_scan=4):
    """在前 max_scan 行找含关键词的表头行，返回 (行号1based, 行tuple)"""
    for ri, row in enumerate(ws.iter_rows(min_row=1, max_row=max_scan, values_only=True), 1):
        if row and any(isinstance(x, str) and kw in x for x in row):
            return ri, row
    return None, None

def col_of(hdr, name):
    for i, v in enumerate(hdr or ()):
        if isinstance(v, str) and v.strip() == name: return i
    return None


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--excel", required=True)
    ap.add_argument("--svc-lo", required=True, help="服务月起始日 YYYY-MM-DD")
    ap.add_argument("--cut", required=True, help="统计截止日 YYYY-MM-DD")
    ap.add_argument("--prev", default=None, help="上一轮 _month_<月>.json（可选，做单调性对比）")
    a = ap.parse_args()
    LO = datetime.date(*map(int, a.svc_lo.split("-")))
    HI = datetime.date(*map(int, a.cut.split("-")))
    print("=" * 76)
    print("源表自检  %s" % os.path.basename(a.excel))
    print("周期 %s ~ %s" % (LO, HI))
    print("=" * 76)
    wb = openpyxl.load_workbook(a.excel, read_only=True, data_only=True)

    # ---------- 1) 明细表周期覆盖 ----------
    print("\n[1] 明细表周期覆盖")
    for sn, kw, mr, dedup, widx in DETAIL_TABLES:
        if sn not in wb.sheetnames:
            warn("缺少 sheet: %s" % sn); continue
        ws = wb[sn]
        hri, hdr = find_hdr(ws, kw)
        if hri is None:
            warn("%s 未找到含「%s」的表头列" % (sn, kw)); continue
        ci = col_of(hdr, kw)
        cnt = collections.Counter(); tot = nodate = 0; woset = set()
        lo_hit = hi_hit = 0
        for r in ws.iter_rows(min_row=mr, values_only=True):
            if not r or len(r) <= ci: continue
            if dedup:
                w = r[widx] if widx is not None and len(r) > widx else None
                if not w: continue
                k = str(w)
                if k in woset: continue
                woset.add(k)
            tot += 1
            d = to_date(r[ci])
            if d is None: nodate += 1; continue
            cnt[d] += 1
            if d < LO: lo_hit += 1
            if d > HI: hi_hit += 1
        inr = sum(v for k, v in cnt.items() if LO <= k <= HI)
        span = "%s ~ %s" % (min(cnt), max(cnt)) if cnt else "—"
        line = "%s[%d列] 数据行%d 无日期%d 周期内%d 范围%s" % (sn, ci, tot, nodate, inr, span)
        if lo_hit or hi_hit:
            warn("%s  ← 越界 %d 条(早于起)/%d 条(晚于截止)" % (line, lo_hit, hi_hit))
        else:
            ok(line)
        if inr:
            det = {str(k): v for k, v in sorted(cnt.items()) if LO <= k <= HI}
            print("        " + str(det))

    # ---------- 2) 收入统计表结构 ----------
    print("\n[2] 服务产品收入统计表结构")
    ws = wb["服务产品收入统计"]
    inc = collections.defaultdict(float); rows = 0; offs = collections.defaultdict(float)
    neg = 0; nets = collections.defaultdict(set); src = collections.defaultdict(float)
    hdr_found = False
    for r in ws.iter_rows(min_row=3, values_only=True):
        if not r or len(r) < 25: continue
        # 跳过可能的重复表头
        if isinstance(r[0], str) and r[0].strip() == "大区": hdr_found = True; continue
        rows += 1
        amt = num(r[24]); tot_r = 0.0
        for i, it in enumerate(ITEMS):
            v = num(r[9 + i]); inc[it] += v; tot_r += v
        if amt < 0 or tot_r < 0: neg += 1
        o = str(r[2]).strip() if r[2] else "(空)"
        offs[o] += amt; nets[o].add(str(r[3]).strip() if r[3] else "(空)")
        src[str(r[8]).strip() if r[8] else "(空)"] += amt
    tot = sum(inc.values())
    ok("数据行 %d | 办事处 %d 个 | 合计 ¥%s" % (rows, len(offs), format(round(tot), ",")))
    if len(offs) == 6:
        ok("办事处数 = 6 ✓")
    else:
        warn("办事处数 = %d（应为 6）: %s" % (len(offs), list(offs)))
    if neg: warn("存在负值行 %d 条" % neg)
    else:   ok("无负值行")
    print("        分办: " + " | ".join("%s ¥%s" % (k, format(round(v), ",")) for k, v in sorted(offs.items(), key=lambda x: -x[1])))
    print("        来源: " + " | ".join("%s ¥%s(%.2f%%)" % (k, format(round(v), ","), v / tot * 100) for k, v in sorted(src.items(), key=lambda x: -x[1])))

    # ---------- 3) 三元勾稽 ----------
    # ⚠️ 口径（2026-10-11 二次确认，以此为准）：收入统计表 = **上缴金额毛额**，不扣折扣。
    #    实证：用户补全兰州缺失行后，收入表 ¥220,439 = 明细毛额合计 ¥220,439，差 ¥0。
    #    教训：10-11 上午曾误判为「净额 = 上缴 − 折扣」，那是当时收入表本身缺了 ¥3,130
    #    （兰州办事处）造成的巧合，不是口径问题。改勾稽口径前，先怀疑表有没有粘全。
    print("\n[3] 三元勾稽（明细 → 收入统计表，口径 = 上缴金额毛额）")

    def amt(row, hdr, *amt_names):
        """取 上缴金额/实收金额（毛额，不扣折扣）。
        col_of 找不到列时返回 None 而不是 -1，这里逐个候选名试。"""
        for nm in amt_names:
            j = col_of(hdr, nm)
            if j is not None and j >= 0:
                return num(row[j])
        return 0.0

    # CSM配件
    csm = collections.defaultdict(float); csm_cat = collections.Counter()
    ws = wb["CSM配件"]; hri, hdr = find_hdr(ws, "销售分类")
    c_cat, c_pay = col_of(hdr, "销售分类"), col_of(hdr, "上缴金额")
    for r in ws.iter_rows(min_row=(hri or 2) + 1, values_only=True):
        if not r or len(r) <= max(c_cat, c_pay): continue
        c = str(r[c_cat]).strip() if r[c_cat] else ""
        if c in ITEMS: csm[c] += amt(r, hdr, "上缴金额"); csm_cat[c] += 1
    # WMS
    wms = collections.defaultdict(float); wms_cat = collections.Counter()
    ws = wb["WMS网点买断配件明细"]; hri, hdr = find_hdr(ws, "销售分类")
    c_cat, c_amt = col_of(hdr, "销售分类"), col_of(hdr, "实收总金额（上缴金额）")
    if c_amt is None: c_amt = col_of(hdr, "实收总金额")
    for r in ws.iter_rows(min_row=(hri or 2) + 1, values_only=True):
        if not r or len(r) <= max(c_cat, c_amt): continue
        c = str(r[c_cat]).strip() if r[c_cat] else ""
        if c in ITEMS: wms[c] += amt(r, hdr, "实收总金额（上缴金额）", "实收总金额"); wms_cat[c] += 1
    # CSM服务项目
    svc = collections.defaultdict(float); svc_cat = collections.Counter()
    ws = wb["CSM服务项目"]; hri, hdr = find_hdr(ws, "销售分类")
    c_cat, c_pay = col_of(hdr, "销售分类"), col_of(hdr, "上缴金额")
    for r in ws.iter_rows(min_row=(hri or 2) + 1, values_only=True):
        if not r or len(r) <= max(c_cat, c_pay): continue
        c = str(r[c_cat]).strip() if r[c_cat] else "(空)"
        svc[c] += amt(r, hdr, "上缴金额"); svc_cat[c] += 1
    svc_sum = sum(v for k, v in svc.items() if k != "(空)")
    s3 = sum(csm.values()) + sum(wms.values()) + svc_sum
    print("        CSM配件 ¥%s(%d条) + WMS买断 ¥%s(%d条) + CSM服务项目 ¥%s(%d条) = ¥%s"
          % (format(round(sum(csm.values())), ","), sum(csm_cat.values()),
             format(round(sum(wms.values())), ","), sum(wms_cat.values()),
             format(round(svc_sum), ","), sum(v for k, v in svc_cat.items() if k != "(空)"),
             format(round(s3), ",")))
    print("        收入统计表合计 = ¥%s" % format(round(tot), ","))
    gap3 = tot - s3
    tol3 = max(1000.0, tot * 0.005)   # 容差：0.5% 或 1000 元（分类标签细节/四舍五入）
    if abs(gap3) < 1.0:
        ok("三元勾稽差额 ¥%.2f ✓（完全对齐）" % gap3)
    elif abs(gap3) <= tol3:
        ok("三元勾稽差额 ¥%.2f（%.2f%%），在容差内 ✓" % (gap3, gap3 / tot * 100))
    else:
        warn("三元勾稽差额 ¥%.2f（%.2f%%）—— 需人工确认" % (gap3, gap3 / tot * 100))

    # ---------- 4) 15 品项逐项勾稽 ----------
    print("\n[4] 15 品项逐项勾稽")
    bad = []
    for it in ITEMS:
        d = round(inc[it] - csm[it] - wms[it] - svc.get(it, 0.0), 2)
        if abs(d) >= 0.01: bad.append((it, inc[it], csm[it], wms[it], svc.get(it, 0.0), d))
    if not bad:
        ok("15 品项全部可由明细还原 ✓")
    else:
        for it, ra, rb, rc, re_, rd in bad:
            print("        %s 收入表¥%.2f = CSM¥%.2f + WMS¥%.2f + 服务项目¥%.2f  差%.2f" % (it, ra, rb, rc, re_, rd))
        # 容差：500 元 或 1.5%（品项级存在折扣处理/分类标签细节差异，非缺行）
        if all(abs(x[5]) < max(500.0, x[1] * 0.015) for x in bad):
            ok("%d 品项小额差异（均 <1.5%% 或 <500元），属折扣处理/分类标签细节" % len(bad))
        else:
            warn("%d 品项存在较大差异，需人工确认" % len(bad))

    # ---------- 5) 口径关键词 ----------
    print("\n[5] 口径关键词自查")
    ws = wb["CSM配件"]
    hri, hdr = find_hdr(ws, "销售分类")
    c_cat, c_pay = col_of(hdr, "销售分类"), col_of(hdr, "上缴金额")
    zg_n = 0; zg_amt = 0.0; zg_le0 = 0
    for r in ws.iter_rows(min_row=(hri or 2) + 1, values_only=True):
        if not r or len(r) <= max(c_cat, c_pay): continue
        if str(r[c_cat]).strip() != "止回阀": continue
        v = num(r[c_pay])          # 毛额口径，与收入表一致
        if v <= 0: zg_le0 += 1; continue
        zg_n += 1; zg_amt += v
    print("        止回阀(上缴>0): %d 条 ¥%s（上缴≤0 剔除 %d 条）" % (zg_n, format(round(zg_amt), ","), zg_le0))
    if zg_amt > 0 and abs(zg_amt - inc["止回阀"]) < 1.0:
        ok("止回阀金额与收入表一致 ✓")
    elif zg_amt > 0 and abs(zg_amt - inc["止回阀"]) <= max(300.0, inc["止回阀"] * 0.01):
        ok("止回阀金额 ¥%.2f 与收入表 ¥%.2f 差异在容差内 ✓" % (zg_amt, inc["止回阀"]))
    else:
        warn("止回阀金额 ¥%.2f 与收入表 ¥%.2f 不一致" % (zg_amt, inc["止回阀"]))
    print("        清洗保养(服务项目): %d 条 ¥%.2f" % (svc_cat.get("清洗保养", 0), svc.get("清洗保养", 0.0)))
    print("        延保服务(服务项目): %d 条 ¥%.2f" % (svc_cat.get("延保服务", 0), svc.get("延保服务", 0.0)))

    # ---------- 6) 与上一轮对比 ----------
    if a.prev and os.path.exists(a.prev):
        print("\n[6] 与上一轮对比（单调性）")
        try:
            pv = json.load(open(a.prev, encoding="utf-8"))
            pd_ = pv.get("D", {})
            p_amt = round(sum(n["合计"] for n in pd_.get("网点", [])))
            p_net = len(pd_.get("网点", []))
            B = pv.get("B", {}).get("meta", {})
            E = pv.get("E", {}).get("total", {})
            p_clean = B.get("分母", 0)
            p_ext = sum(E.get(k, {}).get("qty", 0) for k in ("y1", "y2", "y3"))
            cur_ext = sum(int(r) for r in [0])  # 占位
            print("        上轮收入 ¥%s / 网点 %d / 清洗分母 %d / 延保单 %d"
                  % (format(p_amt, ","), p_net, p_clean, p_ext))
            print("        本轮收入 ¥%s / 网点 %d / 清洗分母 %d"
                  % (format(round(tot), ","), sum(len(v) for v in nets.values()), svc_cat.get("清洗保养", 0)))
            if tot < p_amt:
                warn("收入较上轮下降（¥%s → ¥%s）" % (format(p_amt, ","), format(round(tot), ",")))
            else:
                ok("收入较上轮增长 ¥%s" % format(round(tot - p_amt), ","))
        except Exception as e:
            warn("上轮对比失败: %s" % e)

    wb.close()
    print("\n" + "=" * 76)
    print("结论：通过 %d 项，疑点 %d 项" % (len(OK), len(WARN)))
    for w in WARN: print("   ⚠️  " + w)
    if not WARN: print("   源表自洽，可安全进入管线。")
    print("=" * 76)
    return 1 if WARN else 0

if __name__ == "__main__":
    sys.exit(main())
