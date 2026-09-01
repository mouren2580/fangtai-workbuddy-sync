# -*- coding: utf-8 -*-
import base64, re

SRC = r"D:\WorkBuddy\dashboard_offline.html"
XLSX = r"D:\WorkBuddy\2026年9月西北服务产品.xlsx"
html0 = open(SRC, encoding="utf-8").read()
html = html0

# 1. xlsxData 块替换成 9月 Excel
with open(XLSX, "rb") as f:
    data = f.read()
b64 = base64.b64encode(data).decode("ascii")
pat = re.compile(r'(<script[^>]*id="xlsxData"[^>]*>)(.*?)(</script>)', re.S)
m = pat.search(html)
assert m, "xlsxData block not found"
html = pat.sub(lambda mm: mm.group(1)+b64+mm.group(3), html, count=1)
print("xlsxData replaced, new b64 len =", len(b64))

# 2. 423行 meta：月份 / 统计截止日（双引号字面量）
assert '"月份":"2026-08"' in html
html = html.replace('"月份":"2026-08"', '"月份":"2026-09"')
assert '"统计截止日":"2026-08-20"' in html
html = html.replace('"统计截止日":"2026-08-20"', '"统计截止日":"2026-08-31"')

# 3. monthBadge 初始值
assert 'id="monthBadge">2026-08<' in html
html = html.replace('id="monthBadge">2026-08<', 'id="monthBadge">2026-09<')

# 4. CURRENT_MONTH 默认
assert 'let CURRENT_MONTH = "2026-08";' in html
html = html.replace('let CURRENT_MONTH = "2026-08";', 'let CURRENT_MONTH = "2026-09";')

# 5. SAVED_MONTH 默认
assert "let SAVED_MONTH = '2026-08';" in html
html = html.replace("let SAVED_MONTH = '2026-08';", "let SAVED_MONTH = '2026-09';")

# 6. 年度视图退出回退月
assert "MONTH_DATA[SAVED_MONTH] ? SAVED_MONTH : '2026-08'" in html
html = html.replace("MONTH_DATA[SAVED_MONTH] ? SAVED_MONTH : '2026-08'", "MONTH_DATA[SAVED_MONTH] ? SAVED_MONTH : '2026-09'")

# 7. rebuildFromWorkbook 截止日块
old_g = '''  /* 【铁律】截止日固定为通知更新日(2026-08-21)前一天=2026-08-20，不随导入数据自动滚动 */
  DEFAULT_ASOF = new Date('2026-08-20T00:00:00');
  DASHBOARD_DATA.meta.统计截止日 = '2026-08-20';'''
new_g = '''  /* 【铁律】截止日固定为通知更新日(2026-09-01)前一天=2026-08-31，不随导入数据自动滚动 */
  DEFAULT_ASOF = new Date('2026-08-31T00:00:00');
  DASHBOARD_DATA.meta.统计截止日 = '2026-08-31';'''
assert old_g in html, "rebuildFromWorkbook block not matched"
html = html.replace(old_g, new_g)

# 8. 离线加载截止日块
old_h = '''        // 【铁律】截止日固定 2026-08-20（通知更新日 8-21 前一天），不随数据自动滚动
        DASHBOARD_DATA.meta._stamp = new Date().toISOString();
        DASHBOARD_DATA.meta.统计截止日 = '2026-08-20';
        DEFAULT_ASOF = new Date('2026-08-20T00:00:00');
        computeTime(DEFAULT_ASOF);
        updateAsOfBadge(); syncPeriodUI(); renderAll();
        document.title = "方太西北服务看板（离线·数据更新 " + '2026-08-20' + "）";
        console.log("[offline] 内嵌 Excel 已加载", res, "统计截止日", '2026-08-20');'''
new_h = '''        // 【铁律】截止日固定 2026-08-31（通知更新日 9-01 前一天），不随数据自动滚动
        DASHBOARD_DATA.meta._stamp = new Date().toISOString();
        DASHBOARD_DATA.meta.统计截止日 = '2026-08-31';
        DEFAULT_ASOF = new Date('2026-08-31T00:00:00');
        computeTime(DEFAULT_ASOF);
        updateAsOfBadge(); syncPeriodUI(); renderAll();
        document.title = "方太西北服务看板（离线·数据更新 " + '2026-08-31' + "）";
        console.log("[offline] 内嵌 Excel 已加载", res, "统计截止日", '2026-08-31');'''
assert old_h in html, "offline block not matched"
html = html.replace(old_h, new_h)

# 9. MONTH_DATA 键 2026-08 -> 2026-09（DASHBOARD_DATA 变量运行时为9月数据，引用一致）
old_key = '"2026-08": {D:DASHBOARD_DATA, T:TECH_DATA, B:BIG_CARE_DATA, W:WEEKLY_DATA, E:EXTEND_WARRANTY_DATA},'
new_key = '"2026-09": {D:DASHBOARD_DATA, T:TECH_DATA, B:BIG_CARE_DATA, W:WEEKLY_DATA, E:EXTEND_WARRANTY_DATA},'
assert old_key in html, "MONTH_DATA key not matched"
html = html.replace(old_key, new_key)

# 10. EMBED_TARGETS_BY_MONTH 加 2026-09（沿用8月目标）
lines = html.split('\n')
for i, line in enumerate(lines):
    if line.startswith('const EMBED_TARGETS_BY_MONTH = {'):
        mm = re.search(r'"2026-08":\s*(\{.*\})\s*;', line, re.S)
        assert mm, "2026-08 target not found in EMBED_TARGETS_BY_MONTH"
        grp = mm.group(1)
        lines[i] = line[:mm.start()] + '"2026-08": ' + grp + ', "2026-09": ' + grp + ';'
        print("EMBED_TARGETS_BY_MONTH: added 2026-09 (reuse 8月目标)")
        break
html = '\n'.join(lines)

# 校验：不应残留 2026-08-20
assert '2026-08-20' not in html, "残留 2026-08-20 未清除"
# 9月应出现
assert html.count('2026-09') >= 5

open(SRC, "w", encoding="utf-8").write(html)
print("ALL REPLACEMENTS DONE")
print("2026-09 occurrences:", html.count('2026-09'))
print("2026-08-20 occurrences:", html.count('2026-08-20'))
