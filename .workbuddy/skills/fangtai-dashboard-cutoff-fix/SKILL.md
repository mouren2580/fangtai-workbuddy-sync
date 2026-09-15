---
name: fangtai-dashboard-cutoff-fix
description: 排查并修复方太西北服务产品看板统计截止日被业务单号字符串误解析的问题
agent_created: true
---

# 触发条件
- 用户反馈看板「截止」/「数据覆盖至」日期与实际 Excel 数据最大日期不符
- 看板剩余天数、进度百分比因此不对

# 核心根因
`dashboard_offline.html` 的 `deriveCutoffFromWorkbook()` 曾对字符串做 `new Date()` 宽松解析。业务单号（如 `CK20260802000012`、`FGDG20260802005133`）内含 `20260820` 子串，会被误解析为 `2026-08-20`，导致统计截止日变大。

# 调试步骤
1. 用 `check_xlsx_dates.py` / `find_0820.py` 扫描 Excel 真实日期列最大值，确认业务数据实际截止日。
2. 用 `find_0820_strings.py` 扫描所有可能被 JS `new Date()` 解析为错误日期的字符串（单号、YYYYMMDD 字符串等）。
3. 检查 `dashboard_offline.html` 的 `deriveCutoffFromWorkbook()`：
   - 只应接受 `v instanceof Date`（SheetJS 已识别的日期格式单元格）
   - 不应再对字符串/数字做宽松解析

# 修复步骤
1. 修改 `deriveCutoffFromWorkbook()`：遇到非 Date 值直接 `continue`。
2. 用 `reembed_xlsx.py` 把最新 xlsx 重嵌进 `dashboard_offline.html` 和 `deploy_cs/index.html`（如 xlsx 未变则 base64 不变）。
3. 发布与同步：**统一按下述 skill 的四端 SOP 执行**（2026-09-15 已更新，勿再用旧的 `workbuddy_cloudstudio_deploy` 或 PAT 推 Pages）：
   → `D:\WorkBuddy\.workbuddy\skills\fangtai-dashboard-publish\SKILL.md`
   - GitHub Pages：`python push_gh_pages.py`（SSH，需沙箱旁路）
   - CloudStudio：`workbuddy_sites_deploy(directory="D:\WorkBuddy\deploy_cs")`（需用户当轮确认；报「预留域名未绑定」是假失败）
   - 备份仓：`git push origin main`（Gitee）+ `git push github main`（均需沙箱旁路）
   - 注意 `dashboard_offline.html` 的改动要同步到 `dashboard.html` 与 `dashboard_ref.html`（底板）

# 验证
- CloudStudio 地址：https://0717bc4b30824b8d8a407555473b321e.app.workbuddy.link
- GitHub Pages：https://mouren2580.github.io/fangtai-dashboard/
- 刷新后 badge 应显示正确的「截止 YYYY-MM-DD」和「数据覆盖至 YYYY-MM-DD」
