# -*- coding: utf-8 -*-
"""方太西北服务产品达成看板 —— 一键更新（自检 → 生成 → 合并 → 注入 → 发布）

在本文件夹内执行：
    python run_update.py --excel "2026年10月西北服务产品.xlsx" --cut 2026-10-04
    python run_update.py --excel "..." --cut 2026-10-04 --no-publish   # 只更新本地，不发布

参数：
    --excel        月报 Excel 路径（放在本文件夹内可只写文件名）
    --cut          统计截止日 = 拿到 Excel 那天的前一天（铁律），格式 YYYY-MM-DD
    --month        服务月键，如 2026-10；默认由 --cut 自动推导
    --no-publish   只更新本地 dashboard.html，不推送到 GitHub
    --skip-selfcheck  跳过源表自检（不建议用）
"""
import argparse
import os
import re
import shutil
import subprocess
import sys
import time
from datetime import date

HERE = os.path.dirname(os.path.abspath(__file__))
os.chdir(HERE)
sys.path.insert(0, HERE)
for _s in (sys.stdout, sys.stderr):
    try:
        _s.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass

from build_dashboard import service_period, service_month_of  # noqa: E402

PY = sys.executable
ENV = dict(os.environ, PYTHONIOENCODING="utf-8")
STEP = [0]


def head(msg):
    STEP[0] += 1
    print("\n" + "=" * 68)
    print("[%d] %s" % (STEP[0], msg))
    print("=" * 68)


def run(args, desc):
    """跑一个子脚本；失败即中止整个流程。"""
    t0 = time.time()
    r = subprocess.run([PY] + args, env=ENV, cwd=HERE)
    if r.returncode != 0:
        print("\n❌ 失败：%s（退出码 %d）" % (desc, r.returncode))
        print("   已中止，dashboard.html 未被发布；排查上面的报错后重跑即可。")
        sys.exit(1)
    print("✔ %s（%.1fs）" % (desc, time.time() - t0))


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--excel", required=True, help="月报 Excel 路径")
    ap.add_argument("--cut", required=True, help="统计截止日 YYYY-MM-DD（= 拿到表的当天 −1 天）")
    ap.add_argument("--month", default=None, help="服务月键，如 2026-10；默认自动推导")
    ap.add_argument("--no-publish", action="store_true", help="只更新本地，不推送 GitHub")
    ap.add_argument("--skip-selfcheck", action="store_true", help="跳过源表自检")
    ap.add_argument("--ignore-freeze", action="store_true",
                    help="跳过已锁死月份还原（仅当确实要修改已锁死月份时才用）")
    a = ap.parse_args()

    excel = a.excel
    if not os.path.isabs(excel):
        excel = os.path.join(HERE, excel)
    if not os.path.exists(excel):
        print("❌ 找不到 Excel：%s" % excel)
        sys.exit(1)

    cut = date.fromisoformat(a.cut)
    if a.month:
        y, m = (int(x) for x in a.month.split("-"))
    else:
        y, m = service_month_of(cut)
    mk = "%04d-%02d" % (y, m)
    ps, pe = service_period(y, m)
    lo, hi, pe_s = ps.isoformat(), cut.isoformat(), pe.isoformat()

    print("=" * 68)
    print("方太看板 · 一键更新")
    print("  源表     : %s" % excel)
    print("  统计截止 : %s   （D−1 铁律：拿到表那天的前一天）" % hi)
    print("  服务月   : %s   周期 %s ~ %s" % (mk, lo, pe_s))
    print("  解释器   : %s" % PY)
    print("=" * 68)

    # ---- 0. 环境 ----
    head("环境检查")
    try:
        import openpyxl  # noqa: F401
        print("✔ openpyxl %s" % openpyxl.__version__)
    except ImportError:
        print("❌ 缺少 openpyxl，请先执行：  pip install openpyxl")
        sys.exit(1)
    print("✔ 脚本同目录 build_dashboard.py：", "有" if os.path.exists(os.path.join(HERE, "build_dashboard.py")) else "无")

    # ---- 1. 源表自检 ----
    if a.skip_selfcheck:
        print("\n[跳过] 源表自检")
    else:
        head("源表自检（不通过会中止）")
        run(["_selfcheck_source.py", "--excel", excel, "--svc-lo", lo, "--cut", hi],
            "源表自检")

    # ---- 2. 备份 ----
    head("备份当前看板与状态文件")
    ts = time.strftime("%Y%m%d-%H%M%S")
    shutil.copyfile("dashboard.html", "_bak_dashboard_%s.html" % ts)
    for f in ("_monthly_four.json", "_extend_time.json", "_cleaning_data.json", "_valueparts_daily.json", "_unit_revenue.json"):
        if os.path.exists(f):
            shutil.copyfile(f, "_bak_%s_%s" % (f, ts))
    print("✔ 备份前缀 _bak_*_%s（后悔了可以直接改回）" % ts)

    # ---- 3. 生成 ----
    head("生成各维度数据")
    run(["_gen_month.py", excel, "--date", hi, "--out", "_month_%s.json" % mk], "月度五块矩阵")
    run(["_gen_workorder.py", excel, "--date", hi, "--out", "_mf_%s_wo.json" % mk], "工单 / 防火阀")
    run(["_gen_valueadded.py", "--lo", lo, "--hi", hi, "--excel", excel,
         "--out", "_mf_%s_va.json" % mk], "增值产品")
    run(["_gen_valueparts.py", "--svc-lo", lo, "--cut", hi, "--excel", excel,
         "--month", mk, "--out", "_mf_%s_vp.json" % mk], "增值配件明细")
    run(["_gen_bigcare.py", "--excel", excel, "--ps", lo, "--pe", pe_s, "--cut", hi,
         "--out", "_mf_%s_bc.json" % mk], "大保养")
    run(["_gen_extend_time_months.py", "--excel", excel, "--svc-lo", lo,
         "--cut", hi, "--month", mk], "延保年限")
    run(["_gen_cleaning.py", "--excel", excel], "清洗保养")
    run(["_gen_unit.py", "--excel", excel, "--svc-lo", lo, "--cut", hi,
         "--month", mk, "--out", "_unit_revenue.json"], "单均收入")

    # ---- 3.5 还原已锁死月份（_gen_unit.py / _gen_cleaning.py 会重算历史月份）----
    if a.ignore_freeze:
        print("\n[--ignore-freeze] 跳过已锁死月份还原")
    else:
        head("还原已锁死月份（锁死的数据不允许被重算）")
        run(["_freeze.py", "--restore"], "还原已锁死月份")

    # ---- 4. 合并 + 注入 ----
    head("合并并注入 dashboard.html")
    run(["_merge_month.py", "--month", mk], "合并月度四块")
    run(["_reinject_monthlyfour.py"], "注入 MONTHLY_FOUR")
    run(["_inject_month.py", "--default-month", mk], "注入 MONTH_DATA + 默认月")
    run(["_reinject_extendtime.py"], "注入 EXTEND_TIME_DATA")
    run(["_reinject_cleaning.py"], "注入 CLEANING_DATA")
    run(["_reinject_unit.py"], "注入 UNIT_REVENUE_DATA")

    # ---- 5. 校验 ----
    head("校验")
    h = open("dashboard.html", encoding="utf-8").read()
    hit = len(re.findall(r'"统计截止日":"%s"' % hi, h))
    cur = re.search(r'CURRENT_MONTH\s*=\s*"([0-9-]+)"', h)
    print("  统计截止日命中 %d 处 → %s" % (hit, "✔" if hit else "✘ 没找到，数据可能没注入"))
    print("  默认打开月份 = %s → %s" % (cur.group(1) if cur else "?", "✔" if cur and cur.group(1) == mk else "✘"))
    blocks = re.findall(r"<script(?![^>]*src=)[^>]*>(.*?)</script>", h, re.S)
    import tempfile
    bad = 0
    node_missing = False
    for i, b in enumerate(blocks):
        if not b.strip():
            continue
        p = os.path.join(tempfile.gettempdir(), "_syn_%d.js" % i)
        open(p, "w", encoding="utf-8").write(b)
        try:
            r = subprocess.run(["node", "--check", p], capture_output=True, text=True)
            if r.returncode != 0:
                bad += 1
                print("  ✘ 内联脚本块 %d 语法错误" % i)
        except FileNotFoundError:
            node_missing = True
        finally:
            if os.path.exists(p):
                os.remove(p)
    if node_missing:
        print("  内联脚本语法检查：未装 Node.js，已跳过（不影响数据）")
    elif bad == 0:
        print("  内联脚本语法检查：✔")
    # 锁死月份校验（防止历史月份被悄悄改动）
    fz = subprocess.run([PY, "_freeze.py", "--verify"], env=ENV, cwd=HERE,
                        capture_output=True, text=True)
    for ln in (fz.stdout or "").strip().splitlines():
        print("  " + ln)
    if fz.returncode != 0:
        print("\n❌ 锁死月份被改动，已停止，未发布。先跑 python _freeze.py --restore 还原。")
        sys.exit(1)
    if hit == 0 or bad:
        print("\n❌ 校验未通过，已停止，未发布。")
        sys.exit(1)

    # ---- 6. 发布 ----
    if a.no_publish:
        print("\n[--no-publish] 已跳过发布。浏览器打开本文件夹的 dashboard.html 检查效果。")
    else:
        head("发布到 GitHub Pages")
        run(["_bump_build.py"], "更新 BUILD 号 / version.json")
        import hashlib
        shutil.copyfile("dashboard.html", os.path.join("publish_gh", "index.html"))
        open(os.path.join("publish_gh", ".nojekyll"), "w").write("")
        sa = hashlib.sha1(open("dashboard.html", "rb").read()).hexdigest()
        sb = hashlib.sha1(open(os.path.join("publish_gh", "index.html"), "rb").read()).hexdigest()
        if sa != sb:
            print("❌ 发布副本与源文件不一致，已停止")
            sys.exit(1)
        print("✔ sha1 一致 %s" % sa[:10])

        token = None
        if os.path.exists(".gh_token"):
            token = open(".gh_token", encoding="utf-8").read().strip()
        if not token:
            print("❌ 找不到令牌：请在本文件夹新建 .gh_token 文件，把 GitHub 令牌粘进去（单独一行）")
            print("   本地数据已更新完成，补好令牌后重跑即可发布；或改用 --no-publish。")
            sys.exit(1)
        run(["_gh_push.py", "mouren2580", "fangtai-dashboard", "publish_gh",
             "--token", token, "--branch", "main"], "推送 GitHub")

    # ---- 7. 汇总 ----
    head("完成")
    try:
        import json
        mf = json.load(open("_monthly_four.json", encoding="utf-8"))
        blk = mf["months"][mk]
        wo = blk["workorder"]["meta"]
        va = blk["valueadded"]["meta"]
        print("  服务月 %s（截止 %s）" % (mk, hi))
        print("  工单 %s 单 | 网点 %s 个 | 工程师 %s 人" % (wo.get("总工单数"), wo.get("网点数"), wo.get("工程师数")))
        print("  防火阀占比 %.2f%%" % wo.get("防火阀占比", 0))
        print("  增值产品 %s 件 | 增值配件（当月）%s 件" % (va.get("全区总计数"), blk["valueparts"]["meta"].get("全区总计数_当月")))
        print("  大保养占比 %.2f%%" % blk["bigcare"]["meta"].get("全区占比", 0))
    except Exception as e:
        print("  （汇总读取失败，不影响数据：%s）" % e)
    print("\n浏览器打开 dashboard.html 可先本地核对；线上按 Ctrl+Shift+R 硬刷新。")
    print("注意：应用分享链接（app.workbuddy.host）只能在办公室的 WorkBuddy 里更新，家里更新后记得回去补一次。")


if __name__ == "__main__":
    main()
