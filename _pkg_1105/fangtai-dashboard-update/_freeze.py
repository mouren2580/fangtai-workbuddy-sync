# -*- coding: utf-8 -*-
"""服务月数据「锁死」工具 —— 快照 / 还原 / 校验。

为什么需要：管线里有三个环节会「重算历史月份」而不是只算当月——
  _gen_unit.py            整文件覆盖 _unit_revenue.json（只写当月 → 旧月会消失）
  _gen_cleaning.py        扫整张 Excel 按服务月重新分组（旧月被重算）
  _gen_valueparts.py      只替换当月键（旧月安全，但仍做保护）
_merge_month.py / _gen_extend_time_months.py 只替换当月键，本身安全。
所以「锁死」= 把某月各数据块快照到 _frozen/<月>/，之后每次跑管线先还原，旧月永远不动。

用法：
  python _freeze.py --save 2026-10        # 锁死（快照当前工作文件里的该月数据）
  python _freeze.py --save 2026-11 --only lastyear   # 只锁「去年同月 3+1」（dashboard.html 内嵌）
  python _freeze.py --restore            # 还原所有已锁死月份（run_update.py 自动调用）
  python _freeze.py --verify             # 校验所有已锁死月份 vs 当前工作文件
  python _freeze.py --verify 2026-10     # 只校验某月
  python _freeze.py --list               # 列出已锁死月份
  python _freeze.py --unlock 2026-10 --yes   # 解锁（删除快照）
"""
import argparse
import datetime
import hashlib
import json
import os
import re
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
os.chdir(HERE)
FROZEN = "_frozen"
HTML = "dashboard.html"
LY_RE = re.compile(r"const EMBED_LASTYEAR_BY_MONTH = (\{.*?\});\n", re.S)

# 月度四块键名
MF_KEYS = ("workorder", "valueadded", "valueparts", "bigcare")


def read_lastyear(mk):
    """从 dashboard.html 取该月的「去年同月 3+1」块（去掉 _stamp，锁死的是数值本身）。"""
    try:
        h = open(HTML, encoding="utf-8").read()
        m = LY_RE.search(h)
        if not m:
            return None
        blk = json.loads(m.group(1)).get(mk)
        if isinstance(blk, dict):
            blk = {k: v for k, v in blk.items() if k != "_stamp"}
        return blk
    except Exception:
        return None


def write_lastyear(mk, obj):
    """把该月块写回 dashboard.html，并补一个新的 _stamp（防老端本地缓存压制）。"""
    h = open(HTML, encoding="utf-8").read()
    m = LY_RE.search(h)
    if not m:
        return False
    cur = json.loads(m.group(1))
    blk = dict(obj)
    blk["_stamp"] = datetime.datetime.now().astimezone().isoformat(timespec="seconds")
    cur[mk] = blk
    payload = ("const EMBED_LASTYEAR_BY_MONTH = "
               + json.dumps(cur, ensure_ascii=False, separators=(",", ":")) + ";\n")
    open(HTML, "w", encoding="utf-8").write(h[:m.start()] + payload + h[m.end():])
    return True



def jload(p, default=None):
    try:
        with open(p, encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return default


def jdump(p, obj):
    with open(p, "w", encoding="utf-8") as f:
        json.dump(obj, f, ensure_ascii=False, separators=(",", ":"))


def sha(obj):
    s = json.dumps(obj, ensure_ascii=False, sort_keys=True, separators=(",", ":"))
    return hashlib.sha1(s.encode("utf-8")).hexdigest()


def month_dir(mk):
    return os.path.join(FROZEN, mk)


def collect(mk):
    """从当前工作文件里摘出该月需要锁死的数据块。缺什么记 None。"""
    st = {}

    st["month_block"] = jload("_month_%s.json" % mk)

    mf = {}
    for k, suf in zip(MF_KEYS, ("wo", "va", "vp", "bc")):
        mf[k] = jload("_mf_%s_%s.json" % (mk, suf))
    st["mf"] = mf

    ext = jload("_extend_time.json", {}) or {}
    st["extend"] = (ext.get("months") or {}).get(mk)

    cl = jload("_cleaning_data.json", {}) or {}
    st["cleaning"] = (cl.get("months") or {}).get(mk)

    vpd = jload("_valueparts_daily.json", {}) or {}
    st["vp_daily"] = (vpd.get("months") or {}).get(mk)

    ur = jload("_unit_revenue.json", {}) or {}
    st["unit"] = (ur.get("months") or {}).get(mk)

    # 已合并进 _monthly_four.json 的块（可能比 _mf_* 更权威）
    m4 = jload("_monthly_four.json", {}) or {}
    st["merged"] = (m4.get("months") or {}).get(mk)

    # 去年同月 3+1（内嵌在 dashboard.html，管线不会重算，但换机 / 换 HTML 会丢）
    st["lastyear"] = read_lastyear(mk)

    return st


def summary_of(st):
    """给 meta.json 用的关键数字，便于日后核对。"""
    s = {}
    try:
        wo = (st.get("mf") or {}).get("workorder", {}).get("meta", {})
        va = (st.get("mf") or {}).get("valueadded", {}).get("meta", {})
        vp = (st.get("mf") or {}).get("valueparts", {}).get("meta", {})
        bc = (st.get("mf") or {}).get("bigcare", {}).get("meta", {})
        s["工单"] = wo.get("总工单数")
        s["网点"] = wo.get("网点数")
        s["工程师"] = wo.get("工程师数")
        s["防火阀占比"] = wo.get("防火阀占比")
        s["增值产品"] = va.get("全区总计数")
        s["增值配件当月"] = vp.get("全区总计数_当月")
        s["大保养占比"] = bc.get("全区占比")
        ut = st.get("unit") or {}
        tot = ut.get("total") or ut
        s["①单均服务产值"] = tot.get("a")
        s["⑤单均服务产品收入"] = tot.get("b")
        s["出单工单"] = (ut.get("meta") or {}).get("出单工单")
        s["截止"] = (ut.get("meta") or {}).get("截止") or wo.get("截止")
    except Exception as e:
        s["error"] = str(e)
    return s


def save(mk, excel="", only=""):
    if only == "lastyear":
        # 只锁「去年同月 3+1」：合并进已有快照，不动其它块
        d = month_dir(mk)
        old = jload(os.path.join(d, "state.json"), {}) or {}
        ly = read_lastyear(mk)
        if ly is None:
            print("❌ %s 的 EMBED_LASTYEAR_BY_MONTH 里没有该月数据，无从锁死" % mk)
            return 1
        old["lastyear"] = ly
        os.makedirs(d, exist_ok=True)
        jdump(os.path.join(d, "state.json"), old)
        meta = jload(os.path.join(d, "meta.json"), {}) or {}
        meta.setdefault("月份", mk)
        meta["去年同比锁死时间"] = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        meta.setdefault("摘要", {})
        meta["摘要"]["去年同比小计"] = {
            k: round(sum(v.get(k, 0) for v in ly.values() if isinstance(v, dict)), 2)
            for k in ("清洗保养", "延保服务", "清洁耗材", "止回阀")
        }
        meta.setdefault("块校验", {})["lastyear"] = sha(ly)
        jdump(os.path.join(d, "meta.json"), meta)
        print("🔒 已锁死 %s 的去年同月 3+1 → %s" % (mk, d))
        print("   小计：%s" % json.dumps(meta["摘要"]["去年同比小计"], ensure_ascii=False))
        return 0

    st = collect(mk)
    empty = [k for k in ("month_block", "extend", "cleaning", "unit") if st.get(k) is None]
    if st.get("mf") and all(v is None for v in st["mf"].values()):
        empty.append("mf")
    if empty:
        print("❌ %s 缺少数据块：%s —— 请先跑一次更新再锁死" % (mk, ", ".join(empty)))
        return 1
    d = month_dir(mk)
    os.makedirs(d, exist_ok=True)
    jdump(os.path.join(d, "state.json"), st)
    meta = {
        "月份": mk,
        "锁死时间": datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "源表": excel,
        "摘要": summary_of(st),
        "块校验": {k: sha(v) for k, v in st.items()},
    }
    jdump(os.path.join(d, "meta.json"), meta)
    print("🔒 已锁死 %s → %s" % (mk, d))
    print("   摘要：%s" % json.dumps(meta["摘要"], ensure_ascii=False))
    return 0


def restore_one(mk, quiet=False):
    d = month_dir(mk)
    st = jload(os.path.join(d, "state.json"))
    if not st:
        print("⚠️  %s 无快照，跳过" % mk)
        return False
    changed = []

    def write_if_diff(path, obj):
        if obj is None:
            return
        old = jload(path)
        if old is not None and sha(old) == sha(obj):
            return
        jdump(path, obj)
        changed.append(os.path.basename(path))

    write_if_diff("_month_%s.json" % mk, st.get("month_block"))
    for k, suf in zip(MF_KEYS, ("wo", "va", "vp", "bc")):
        write_if_diff("_mf_%s_%s.json" % (mk, suf), (st.get("mf") or {}).get(k))

    # 月份键嵌在容器里的，逐键替换
    def patch_container(path, obj, label):
        if obj is None:
            return
        c = jload(path)
        if c is None:
            return
        cur = (c.get("months") or {}).get(mk)
        if cur is not None and sha(cur) == sha(obj):
            return
        c.setdefault("months", {})[mk] = obj
        jdump(path, c)
        changed.append("%s[%s]" % (label, mk))

    patch_container("_monthly_four.json", st.get("merged") or {
        k: (st.get("mf") or {}).get(k) for k in MF_KEYS
    }, "_monthly_four")
    patch_container("_extend_time.json", st.get("extend"), "_extend_time")
    patch_container("_cleaning_data.json", st.get("cleaning"), "_cleaning_data")
    patch_container("_valueparts_daily.json", st.get("vp_daily"), "_valueparts_daily")
    patch_container("_unit_revenue.json", st.get("unit"), "_unit_revenue")
    # 兜底：meta.latest 若指向不存在的月份（生成器换了新月但被还原），改回最大月份
    ur = jload("_unit_revenue.json")
    if ur:
        mset = set((ur.get("months") or {}).keys())
        latest = (ur.get("meta") or {}).get("latest")
        if mset and latest not in mset:
            ur.setdefault("meta", {})["latest"] = max(mset)
            jdump("_unit_revenue.json", ur)
            changed.append("_unit_revenue.meta.latest")

    if not quiet:
        if changed:
            print("🔒 还原 %s：%s" % (mk, ", ".join(changed)))
        else:
            print("🔒 %s 已是最新（无需还原）" % mk)
    return True


def restore_all(quiet=False):
    if not os.path.isdir(FROZEN):
        if not quiet:
            print("（无已锁死月份）")
        return 0
    mks = sorted(x for x in os.listdir(FROZEN) if os.path.isdir(month_dir(x)))
    if not mks:
        if not quiet:
            print("（无已锁死月份）")
        return 0
    n = 0
    for mk in mks:
        if restore_one(mk, quiet=quiet):
            n += 1
    if n and not quiet:
        print("🔒 已锁死月份：%s" % ", ".join(mks))
    return n


def verify_one(mk):
    d = month_dir(mk)
    meta = jload(os.path.join(d, "meta.json"), {}) or {}
    st = jload(os.path.join(d, "state.json"))
    if not st:
        print("⚠️  %s 无快照" % mk)
        return 1
    cur = collect(mk)
    diffs = []
    for k in st:
        if sha(st[k]) != sha(cur.get(k)):
            diffs.append(k)
    if not diffs:
        print("✔ %s 未变动（%s 锁死）" % (mk, meta.get("锁死时间") or meta.get("去年同比锁死时间", "?")))
        return 0
    print("✘ %s 有 %d 个数据块被改动：%s" % (mk, len(diffs), ", ".join(diffs)))
    print("   跑 `python _freeze.py --restore` 可还原")
    return 1


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--save", metavar="YYYY-MM")
    ap.add_argument("--restore", action="store_true")
    ap.add_argument("--verify", nargs="?", const="__all__", metavar="YYYY-MM")
    ap.add_argument("--list", action="store_true")
    ap.add_argument("--unlock", metavar="YYYY-MM")
    ap.add_argument("--yes", action="store_true")
    ap.add_argument("--excel", default="")
    ap.add_argument("--only", default="", choices=["", "lastyear"],
                    help="lastyear = 只锁「去年同月 3+1」")
    ap.add_argument("--quiet", action="store_true")
    a = ap.parse_args()

    if a.save:
        return save(a.save, a.excel, a.only)
    if a.restore:
        restore_all(quiet=a.quiet)
        return 0
    if a.verify:
        mks = ([a.verify] if a.verify != "__all__"
               else sorted(x for x in os.listdir(FROZEN) if os.path.isdir(month_dir(x))) if os.path.isdir(FROZEN) else [])
        if not mks:
            print("（无已锁死月份）")
            return 0
        return max(verify_one(mk) for mk in mks)
    if a.list:
        if not os.path.isdir(FROZEN):
            print("（无已锁死月份）")
            return 0
        mks = sorted(x for x in os.listdir(FROZEN) if os.path.isdir(month_dir(x)))
        for mk in mks:
            meta = jload(os.path.join(month_dir(mk), "meta.json"), {}) or {}
            _t = meta.get("锁死时间") or meta.get("去年同比锁死时间", "?")
            _c = (meta.get("摘要") or {}).get("截止")
            _extra = "" if meta.get("锁死时间") else "  （仅去年同比）"
            print("%s  锁死于 %s  截止 %s%s" % (mk, _t, _c, _extra))
        if not mks:
            print("（无已锁死月份）")
        return 0
    if a.unlock:
        d = month_dir(a.unlock)
        if not os.path.isdir(d):
            print("⚠️  %s 未锁死" % a.unlock)
            return 0
        if not a.yes:
            print("要解锁 %s 请加 --yes" % a.unlock)
            return 1
        import shutil
        shutil.rmtree(d)
        print("🔓 已解锁 %s" % a.unlock)
        return 0
    ap.print_help()
    return 0


if __name__ == "__main__":
    sys.exit(main())
