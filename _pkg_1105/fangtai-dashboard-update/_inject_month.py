# -*- coding: utf-8 -*-
"""把各月 _month_YYYY-MM.json 注入 dashboard.html，构建多月份结构（月份下拉）。
   - 顶层 5 个数据块(DASHBOARD_DATA/TECH_DATA/BIG_CARE_DATA/WEEKLY_DATA/EXTEND_WARRANTY_DATA)
     保持为「基准月」(取顶层 DASHBOARD_DATA.meta.月份，当前=2026-08)，仅 const→let 以便切换。
   - 其余月份作为 MDATA_<Y>_<M>_* 常量注入，并构建 MONTH_DATA 包装 + CURRENT_MONTH + applyMonth。
   - 用 /*__MONTHS_START__*/.../*__MONTHS_END__*/ 锚点包裹，可重复运行（幂等刷新）。
   用法: python _inject_month.py
"""
import sys, os, re, json, glob, argparse
for s in (sys.stdout, sys.stderr):
    try: s.reconfigure(encoding="utf-8", errors="replace")
    except Exception: pass

ap = argparse.ArgumentParser()
ap.add_argument("--default-month", default="2026-09", help="初始显示月份(默认2026-09=当前服务月)")
args = ap.parse_args()

HTML = "dashboard.html"

def jd(o): return json.dumps(o, ensure_ascii=False, separators=(",", ":"))

html = open(HTML, encoding="utf-8").read()

# 1) const → let（幂等）
for var in ["DASHBOARD_DATA", "BIG_CARE_DATA", "TECH_DATA", "WEEKLY_DATA", "EXTEND_WARRANTY_DATA"]:
    html = re.sub(r"^const %s = " % var, "let %s = " % var, html, flags=re.M)
html = re.sub(r"^const ITEMS = DASHBOARD_DATA\.meta\.品项;", "let ITEMS = DASHBOARD_DATA.meta.品项;", html, flags=re.M)
html = re.sub(r"^(const|let) TIME = DASHBOARD_DATA\.time;", "let TIME = DASHBOARD_DATA.time;", html, flags=re.M)

# 2) 收集各月 JSON
files = sorted(glob.glob("_month_*.json"))
months = {}
for f in files:
    mm = re.search(r"_month_((\d{4})-(\d{2}))\.json", f)
    if not mm: continue
    months[mm.group(1)] = json.load(open(f, encoding="utf-8"))

# 基准月 = 顶层 DASHBOARD_DATA.meta.月份
m = re.search(r"let DASHBOARD_DATA = (\{.*?\});", html, re.S)
if not m:
    print("❌ 未找到顶层 DASHBOARD_DATA"); sys.exit(1)
base_month = json.loads(m.group(1))["meta"]["月份"]
default_month = args.default_month
print("基准月 =", base_month, "| 默认显示月 =", default_month, "| 发现月份:", sorted(months))

# 2.5) 月份回退：SAVED_MONTH 跟随默认月；单月回退兜底改用 CURRENT_MONTH（原先硬编码会过期）
n1 = len(re.findall(r"let SAVED_MONTH = '[^']*';", html))
html = re.sub(r"let SAVED_MONTH = '[^']*';", "let SAVED_MONTH = '%s';" % default_month, html)
n2 = len(re.findall(r"MONTH_DATA\[SAVED_MONTH\] \? SAVED_MONTH : '[^']*'", html))
html = re.sub(r"MONTH_DATA\[SAVED_MONTH\] \? SAVED_MONTH : '[^']*'",
              "MONTH_DATA[SAVED_MONTH] ? SAVED_MONTH : CURRENT_MONTH", html)
print("✔ 回退月 SAVED_MONTH → %s（%d 处）；单月回退兜底改用 CURRENT_MONTH（%d 处）" % (default_month, n1, n2))

# 3) 构建注入块
parts = []
for mk in sorted(months):
    if mk == base_month:
        continue
    y, mo = mk.split("-")
    pre = "MDATA_%s_%s" % (y, mo)
    D = months[mk]["D"]; T = months[mk]["T"]; B = months[mk]["B"]; W = months[mk]["W"]; E = months[mk]["E"]
    parts.append("const %s_D = %s;" % (pre, jd(D)))
    parts.append("const %s_T = %s;" % (pre, jd(T)))
    parts.append("const %s_B = %s;" % (pre, jd(B)))
    parts.append("const %s_W = %s;" % (pre, jd(W)))
    parts.append("const %s_E = %s;" % (pre, jd(E)))

entries = ['  "%s": {D:DASHBOARD_DATA, T:TECH_DATA, B:BIG_CARE_DATA, W:WEEKLY_DATA, E:EXTEND_WARRANTY_DATA}' % base_month]
for mk in sorted(months):
    if mk == base_month:
        continue
    y, mo = mk.split("-")
    pre = "MDATA_%s_%s" % (y, mo)
    entries.append('  "%s": {D:%s_D, T:%s_T, B:%s_B, W:%s_W, E:%s_E}' % (mk, pre, pre, pre, pre, pre))
month_data = "const MONTH_DATA = {\n" + ",\n".join(entries) + "\n};"

apply_fn = (
    'let CURRENT_MONTH = "%s";\n' % default_month +
    'function applyMonth(m){\n'
    '  if(!MONTH_DATA[m]) return;\n'
    '  CURRENT_MONTH = m;\n'
    '  const M = MONTH_DATA[m];\n'
    '  DASHBOARD_DATA = M.D; WEEKLY_DATA = M.W; BIG_CARE_DATA = M.B; EXTEND_WARRANTY_DATA = M.E; TECH_DATA = M.T;\n'
    '  TIME = DASHBOARD_DATA.time;\n'
    '  NETS = DASHBOARD_DATA.网点;\n'
    '  ITEMS = (DASHBOARD_DATA.meta && DASHBOARD_DATA.meta.品项) || ITEMS;\n'
    "  const s=(DASHBOARD_DATA.meta&&DASHBOARD_DATA.meta.统计截止日)||'';\n"
    "  const d=s?new Date(s+'T00:00:00'):new Date();\n"
    '  computeTime(isNaN(d.getTime())?new Date():d);\n'
    '  recomputeTime(); reinitOrder(); ensureTargets();\n'
    "  const sel=document.getElementById('monthSel'); if(sel) sel.value=m;\n"
    "  // 同步 工单/防火阀/增值产品/增值配件/大保养 四块到当前月（与顶部 monthSel 统一由 applyMonth 驱动；2026-08 用基准块，其余月用 MONTHLY_FOUR）\n"
    "  if(typeof mfApply==='function'){ try{ mfApply(CURRENT_MONTH==='2026-08'?'base':CURRENT_MONTH); }catch(e){ console.error('mfApply in applyMonth', e); } }\n"
    '  renderAll();\n'
    '  if (typeof renderTech === "function") renderTech();  // 切月后刷新服务工程师明细(读全局最新 TECH_DATA)；放模板里避免被重注入覆盖\n'
    '}\n'
)

block = "/*__MONTHS_START__*/\n" + "\n".join(parts) + "\n" + month_data + "\n" + apply_fn + "/*__MONTHS_END__*/"

# 4) 插入/替换
if "/*__MONTHS_START__*/" in html:
    html = re.sub(r"/\*__MONTHS_START__\*/.*?/\*__MONTHS_END__\*/", block, html, flags=re.S)
    print("✔ 已就地刷新 MONTH_DATA 块")
else:
    html = re.sub(r"(?:const|let) TIME = DASHBOARD_DATA\.time;", block + r"\n\g<0>", html, count=1)
    print("✔ 已插入 MONTH_DATA 块（首次）")

open(HTML, "w", encoding="utf-8").write(html)

# 校验
print("MONTH_DATA 出现次数:", html.count("const MONTH_DATA ="))
print("applyMonth 出现次数:", html.count("function applyMonth("))
print("let DASHBOARD_DATA:", "let DASHBOARD_DATA =" in html, "| 残留 const DASHBOARD_DATA:", "const DASHBOARD_DATA =" in html)
