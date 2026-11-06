# -*- coding: utf-8 -*-
"""单均收入生成器（口径① + 口径⑤）· 三维度：办事处 / 服务网点 / 服务工程师

口径① 单均服务产值     = 工单表「实收费用」合计 ÷ 上门工单数（全部工单 − 远程指导）
口径⑤ 单均服务产品收入 = 服务产品收入 ÷ 出单工单数（按工单编号去重）

⑤ 分子（2026-10-17 用户定稿）＝ CSM服务项目 + CSM配件 + **工单自购配件明细**
  · 三张表都带「工单编号」，逐笔按工单编号归属到办事处 / 网点 / 工程师，**不做任何分摊**
  · 工单自购配件含「网点买断」在工单上的实际消耗/收费，买断因此天然落到具体的人
  · ⚠️ 这是「工单收费口径」，与「服务产品收入统计表」的网点上缴口径不同，两者不可直接对账
  · 分母＝三表去重后的工单编号数（同一张工单多笔只算一单）

输出 _unit_revenue.json —— {meta:{latest,...}, months:{"YYYY-MM":{meta,total,office,net,eng,items}}}
注意：「远程指导」在工单表的「服务项目」列，不在「工单类型」列。
"""
import sys, io, os, json, collections, datetime, argparse
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")
from openpyxl import load_workbook

ap = argparse.ArgumentParser()
ap.add_argument("--excel", required=True)
ap.add_argument("--svc-lo", required=True)
ap.add_argument("--cut", required=True)
ap.add_argument("--month", required=True)
ap.add_argument("--out", default="_unit_revenue.json")
a = ap.parse_args()

LO = datetime.date(*map(int, a.svc_lo.split("-")))
CUT = datetime.date(*map(int, a.cut.split("-")))
REMOTE = "远程指导"

wb = load_workbook(a.excel, data_only=True, read_only=True)


def find_hdr(ws, key):
    for i, row in enumerate(ws.iter_rows(max_row=8, values_only=True), start=1):
        vals = [str(v).strip() if v is not None else "" for v in row]
        if key in vals:
            return i, vals
    return None, None


def idx(hdr, name):
    try:
        return hdr.index(name)
    except ValueError:
        return None


def num(v):
    try:
        return float(v or 0)
    except Exception:
        return 0.0


def to_date(v):
    if isinstance(v, datetime.datetime):
        return v.date()
    if isinstance(v, datetime.date):
        return v
    if isinstance(v, str):
        s = v.strip()[:10].replace("/", "-").replace(".", "-")
        try:
            y, m, d = s.split("-")
            return datetime.date(int(y), int(m), int(d))
        except Exception:
            return None
    return None


def pick_sheet(*keys, exclude=()):
    for sn in wb.sheetnames:
        if any(k in sn for k in keys) and not any(e in sn for e in exclude):
            return sn
    raise SystemExit("找不到 sheet: %s" % (keys,))


def blank():
    return {"cnt": 0, "remote": 0, "recv": 0.0, "paidCnt": 0}


# ---------- 1) 工单表：上门工单数 + 实收费用 + 工单档案（编号 → 办事处/网点/工程师） ----------
sh = pick_sheet("工单", exclude=("自购",))
ws = wb[sh]
hri, hdr = find_hdr(ws, "办事处名称")
i_t, i_no = idx(hdr, "服务完成时间"), idx(hdr, "工单编号")
i_off, i_net = idx(hdr, "办事处名称"), idx(hdr, "服务网点")
i_eng, i_eid = idx(hdr, "服务工程师"), idx(hdr, "工程师编号")
i_sitem = idx(hdr, "服务项目")
# ① 分子：用户 2026-10-17 15:18 定稿 = 工单表「实收费用」
# （曾短暂试过「技师应缴金额」，全月仅差 1 张商用净水机换芯单 ¥540，用户判定可忽略，仍用实收费用）
i_recv = idx(hdr, "实收费用")
AMT_COL = "实收费用"
if i_recv is None:
    i_recv = idx(hdr, "技师应缴金额")
    AMT_COL = "技师应缴金额"

off, net, eng = collections.defaultdict(blank), collections.defaultdict(blank), collections.defaultdict(blank)
net_office, eng_meta = {}, {}
order_meta = {}          # 工单编号 -> {office, net, ekey}（全量建库，不按周期过滤）
n_all = n_remote = n_paid = 0
recv_all = 0.0
seen = set()
for r in ws.iter_rows(min_row=(hri or 2) + 1, values_only=True):
    if not r or len(r) <= max(i_t, i_no, i_off, i_sitem, i_recv):
        continue
    no = str(r[i_no]).strip() if r[i_no] else ""
    o = str(r[i_off]).strip() if r[i_off] else "(空)"
    nn = str(r[i_net]).strip() if i_net is not None and r[i_net] else "(空)"
    en = str(r[i_eng]).strip() if i_eng is not None and r[i_eng] else "(空)"
    eid = str(r[i_eid]).strip() if i_eid is not None and r[i_eid] else ""
    ekey = eid or (en + "@" + nn)
    if no and no not in order_meta:
        order_meta[no] = {"office": o, "net": nn, "ekey": ekey}
    d = to_date(r[i_t])
    if d is None or not (LO <= d <= CUT):
        continue
    if no:
        if no in seen:
            continue
        seen.add(no)
    net_office.setdefault(nn, o)
    if ekey not in eng_meta:
        eng_meta[ekey] = {"name": en or "(空)", "id": eid, "office": o, "net": nn}
    sitem = str(r[i_sitem]).strip() if r[i_sitem] else ""
    v = num(r[i_recv])
    n_all += 1
    recv_all += v
    targets = [(off[o]), (net[nn]), (eng[ekey])]
    if sitem == REMOTE:
        n_remote += 1
        for t in targets:
            t["remote"] += 1
        continue
    n_paid += 1 if v > 0 else 0
    for t in targets:
        t["cnt"] += 1
        t["recv"] += v
        if v > 0:
            t["paidCnt"] += 1

n_net = n_all - n_remote

# ---------- 2) ⑤ 服务产品收入（工单归属口径）：CSM服务项目 + CSM配件 + 工单自购配件 ----------
ITEMS15 = ["清洗保养", "延保服务", "止回阀", "清洁耗材", "装饰罩", "排烟出风管",
           "其他产品线专用", "水/气路金属软管", "水气路通用辅材", "其他水路辅材",
           "其他气路辅材", "橱柜类辅材", "燃气管快速接头", "工具", "电路辅材"]

rev_off = collections.defaultdict(float)
rev_net = collections.defaultdict(float)
rev_eng = collections.defaultdict(float)
ord_off = collections.defaultdict(set)      # 出单工单（去重）—— 办事处
ord_net = collections.defaultdict(set)      # 网点
ord_eng = collections.defaultdict(set)      # 工程师
all_orders = set()                          # 全区出单工单
rev_total = 0.0
src_sum = {}
unmatched = collections.Counter()


def add_amount(no, amt, t_office, t_net, t_ekey):
    """把一笔金额归属到三个维度；有工单编号优先按工单档案定位"""
    om = order_meta.get(no) if no else None
    if om:
        o, nn, ek = om["office"], om["net"], om["ekey"]
    else:
        o, nn, ek = t_office, t_net, t_ekey
        unmatched["有单号未匹配" if no else "无单号"] += 1
    rev_off[o] += amt
    rev_net[nn] += amt
    rev_eng[ek] += amt
    if no:
        ord_off[o].add(no)
        ord_net[nn].add(no)
        ord_eng[ek].add(no)
        all_orders.add(no)
    if ek not in eng_meta:
        eng_meta[ek] = {"name": ek.split("@")[0], "id": "" if "@" in ek else ek,
                        "office": o, "net": nn}


# 2a) CSM配件 / CSM服务项目：销售分类∈15品项，上缴金额>0
for sheet, eid_names in ((pick_sheet("CSM配件"), ("服务工程师编码", "服务工程师编号")),
                         (pick_sheet("CSM服务项目"), ("服务工程师编号", "服务工程师编码"))):
    w = wb[sheet]
    h2, hd = find_hdr(w, "销售分类")
    if h2 is None:
        continue
    c_cat, c_amt = idx(hd, "销售分类"), idx(hd, "上缴金额")
    c_no = idx(hd, "工单编号")
    c_off, c_net = idx(hd, "办事处"), idx(hd, "服务网点")
    c_eng = idx(hd, "服务工程师")
    c_eid = None
    for nm in eid_names:
        j = idx(hd, nm)
        if j is not None:
            c_eid = j
            break
    if c_cat is None or c_amt is None:
        continue
    sub = 0.0
    for r in w.iter_rows(min_row=h2 + 1, values_only=True):
        if not r or len(r) <= max(c_cat, c_amt):
            continue
        c = str(r[c_cat]).strip() if r[c_cat] else ""
        if c not in ITEMS15:
            continue
        v = num(r[c_amt])
        if v <= 0:
            continue
        no = str(r[c_no]).strip() if c_no is not None and r[c_no] else ""
        o = str(r[c_off]).strip() if c_off is not None and r[c_off] else "(空)"
        nn = str(r[c_net]).strip() if c_net is not None and r[c_net] else "(空)"
        en = str(r[c_eng]).strip() if c_eng is not None and r[c_eng] else "(空)"
        eid = str(r[c_eid]).strip() if c_eid is not None and r[c_eid] else ""
        add_amount(no, v, o, nn, eid or (en + "@" + nn))
        sub += v
    src_sum[sheet] = sub
    rev_total += sub

# 2b) 工单自购配件明细：全额计入（含「网点买断」配件），该表无销售分类
ws2 = wb[pick_sheet("工单自购配件", exclude=("CSM",))]
hd2 = [str(x).strip() if x else "" for x in next(ws2.iter_rows(max_row=1, values_only=True))]
c_t, c_amt = idx(hd2, "录入完成时间"), idx(hd2, "自购配件收费金额")
c_no, c_net = idx(hd2, "工单编号"), idx(hd2, "服务网点")
c_eng, c_cp = idx(hd2, "服务工程师"), idx(hd2, "配件名称")
i_eid2 = 28 if len(hd2) > 28 else None      # 「工程师编号」列无表头，固定第 29 列
sub = 0.0
bo_buyout = 0.0
for r in ws2.iter_rows(min_row=2, values_only=True):
    if not r or c_t is None or c_amt is None:
        continue
    d = to_date(r[c_t])
    if d is None or not (LO <= d <= CUT):
        continue
    v = num(r[c_amt])
    if v <= 0:
        continue
    no = str(r[c_no]).strip() if c_no is not None and r[c_no] else ""
    nn = str(r[c_net]).strip() if c_net is not None and r[c_net] else "(空)"
    en = str(r[c_eng]).strip() if c_eng is not None and r[c_eng] else "(空)"
    eid = str(r[i_eid2]).strip() if i_eid2 and len(r) > i_eid2 and r[i_eid2] else ""
    om = order_meta.get(no) if no else None
    o = om["office"] if om else net_office.get(nn, "(空)")
    add_amount(no, v, o, nn, eid or (en + "@" + nn))
    sub += v
    if c_cp is not None and r[c_cp] and "买断" in str(r[c_cp]):
        bo_buyout += v
src_sum["工单自购配件明细"] = sub
rev_total += sub

# ---------- 3) 服务产品收入统计表（上缴口径）+ 品项（结构参考） ----------
ws3 = wb[pick_sheet("服务产品收入统计")]
inc_total = 0.0
for r in ws3.iter_rows(min_row=3, values_only=True):
    if not r or len(r) < 25:
        continue
    if isinstance(r[0], str) and r[0].strip() == "大区":
        continue
    inc_total += num(r[24])

cnt_i, amt_i = collections.Counter(), collections.defaultdict(float)
for sheet, amts in ((pick_sheet("CSM配件"), ("上缴金额",)),
                    (pick_sheet("CSM服务项目"), ("上缴金额",)),
                    (pick_sheet("WMS"), ("实收总金额（上缴金额）", "实收总金额"))):
    w = wb[sheet]
    h2, hd = find_hdr(w, "销售分类")
    if h2 is None:
        continue
    c_cat = idx(hd, "销售分类")
    c_amt = None
    for nm in amts:
        j = idx(hd, nm)
        if j is not None:
            c_amt = j
            break
    if c_cat is None or c_amt is None:
        continue
    for r in w.iter_rows(min_row=h2 + 1, values_only=True):
        if not r or len(r) <= max(c_cat, c_amt):
            continue
        c = str(r[c_cat]).strip() if r[c_cat] else ""
        if c not in ITEMS15:
            continue
        v = num(r[c_amt])
        if v <= 0:
            continue
        cnt_i[c] += 1
        amt_i[c] += v

# ---------- 4) 组装 ----------
def pack(d, rev_map, ord_map, extra=None):
    out = []
    for k, v in d.items():
        if v["cnt"] <= 0:
            continue
        rev = rev_map.get(k, 0.0)
        oc = len(ord_map.get(k, ()))
        row = {"cnt": v["cnt"], "remote": v["remote"], "recv": round(v["recv"], 2),
               "rev": round(rev, 2), "paidCnt": v["paidCnt"], "ordCnt": oc,
               "a": round(v["recv"] / v["cnt"], 2),
               "b": round(rev / oc, 2) if oc else 0.0,
               "bUp": round(rev / v["cnt"], 2)}
        if extra:
            row.update(extra(k))
        out.append(row)
    out.sort(key=lambda x: (-x["b"], -x["cnt"]))
    return out


office = pack(off, rev_off, ord_off, lambda k: {"name": k})
net = pack(net, rev_net, ord_net, lambda k: {"name": k, "office": net_office.get(k, "")})
eng = pack(eng, rev_eng, ord_eng,
           lambda k: {"name": eng_meta.get(k, {}).get("name", k),
                      "id": eng_meta.get(k, {}).get("id", ""),
                      "office": eng_meta.get(k, {}).get("office", ""),
                      "net": eng_meta.get(k, {}).get("net", "")})

items = []
for it in ITEMS15:
    if cnt_i[it]:
        items.append({"name": it, "cnt": cnt_i[it], "rev": round(amt_i[it], 2),
                      "unit": round(amt_i[it] / cnt_i[it], 2)})
items.sort(key=lambda x: -x["rev"])

out = {"meta": {}, "months": {}}
if os.path.exists(a.out):
    try:
        out["months"] = json.load(open(a.out, encoding="utf-8")).get("months", {})
    except Exception as e:
        print("⚠️ 旧文件读取失败，将重建：%s" % e)

n_ord = len(all_orders)
out["months"][a.month] = {
    "meta": {
        "周期": "%s~%s" % (LO, CUT),
        "截止": str(CUT),
        "工单全量": n_all,
        "远程指导剔除": n_remote,
        "上门工单": n_net,
        "有收费工单": n_paid,
        "出单工单": n_ord,
        "出单率": round(n_ord / n_net * 100, 2) if n_net else 0.0,
        "口径": ("①单均服务产值=工单表「实收费用」÷上门工单（全部工单−远程指导）。"
                 "⑤单均服务产品收入=服务产品收入÷出单工单（按工单编号去重）；"
                 "分子=CSM服务项目表+CSM配件表+工单自购配件明细（含网点买断在工单上的消耗），"
                 "三表均按工单编号逐笔归属到办事处/网点/工程师，不做分摊。"
                 "⚠️⑤为工单收费口径，与「服务产品收入统计表」的网点上缴口径不同，不可直接对账。"),
        "来源拆分": {k: round(v, 2) for k, v in src_sum.items()},
        "买断配件": round(bo_buyout, 2),
        "未匹配": dict(unmatched),
        "上缴口径合计": round(inc_total, 2),
        "①金额列": AMT_COL,
    },
    "total": {"recv": round(recv_all, 2), "rev": round(rev_total, 2),
              "a": round(recv_all / n_net, 2) if n_net else 0.0,
              "b": round(rev_total / n_ord, 2) if n_ord else 0.0,
              "bUp": round(rev_total / n_net, 2) if n_net else 0.0,
              "ordCnt": n_ord, "ordRate": round(n_ord / n_net * 100, 2) if n_net else 0.0,
              "bkRate": round(n_paid / n_net * 100, 2) if n_net else 0.0},
    "office": office,
    "net": net,
    "eng": eng,
    "items": items,
}
out["meta"] = {"latest": a.month, "截止": str(CUT),
               "来源": "工单表 + CSM服务项目表 + CSM配件表 + 工单自购配件明细"}
json.dump(out, open(a.out, "w", encoding="utf-8"), ensure_ascii=False, indent=1)

t = out["months"][a.month]["total"]
print("✔ 单均收入 %s：工单 %d（剔远程 %d）→ 上门 %d 单；%s ¥%.2f"
      % (a.month, n_all, n_remote, n_net, AMT_COL, t["recv"]))
print("   ⑤分子 ¥%.2f = %s" % (rev_total, " + ".join("%s ¥%.2f" % (k, v) for k, v in src_sum.items())))
print("   出单工单（去重）%d，出单率 %.2f%%" % (n_ord, t["ordRate"]))
print("   ①¥%.2f（÷上门）  ⑤¥%.2f（÷出单）  ⑤¥%.2f（÷上门，参考）" % (t["a"], t["b"], t["bUp"]))
print("   维度：办事处 %d / 网点 %d / 工程师 %d / 品项 %d → %s"
      % (len(office), len(net), len(eng), len(items), a.out))
print("   网点买断配件（工单口径）¥%.2f；未匹配：%s" % (bo_buyout, dict(unmatched) or "无"))
wb.close()
