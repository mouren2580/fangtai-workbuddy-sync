# 项目长期记忆（方太西北服务产品月报 / D:\WorkBuddy）

> 操作步骤 SOP 见技能：`D:\WorkBuddy\.workbuddy\skills\fangtai-dashboard-publish\SKILL.md`。本文件只记「事实、口径、坑、当前状态」。

## 一、仓库与发布（4 端）
- **Pages（同事主链接，优先保证）**：`https://mouren2580.github.io/fangtai-dashboard/`，源 `git@github.com:mouren2580/fangtai-dashboard.git`，用 `push_gh_pages.py`（**前台跑，不能后台**，SSH 会挂死）。
- **备份仓双推**：Gitee `origin`（私有）＋ GitHub `github`（仍 Public）。认证全走 SSH `~/.ssh/id_ed25519_github` / `_gitee`，**PAT 已失效**。任何读 `~/.ssh` 的 git 操作都要 `dangerouslyDisableSandbox:true`。
- **CloudStudio**：目录 `deploy_cs`，旧链 `https://0717bc4b30824b8d8a407555473b321e.app.workbuddy.link`。报「预留域名未绑定…本次发布已停止」= **假失败**，curl `version.json` 验证即可。
- **家里 host**：`fangtai-service-dashboard.app.workbuddy.host` —— 家里机器建的，办公室**覆盖不了**，只能用户在家重发。
- ⚠️ Gitee 单文件 100MB；`git add -A` 前删掉临时大文件（如 `_old_*.xlsx`）。

## 二、看板架构（单文件 HTML）
- 数据块：顶层 `MDATA_YYYY_MM_{D,T,B,W,E}`；`BASE_MONTH="2026-08"` 锚点（8 月用基准块，故**无 `MDATA_2026_08_*` 属正常**）；`CURRENT_MONTH` 控制默认月。
- 全局块：`MONTHLY_FOUR.months[月].{workorder,valueadded,valueparts,bigcare}` + `.year`（固定累计快照，当前仍「1–10 月/截止 10-04」，不可重算）、`EXTEND_TIME_DATA`、`CLEANING_DATA`、`VALUE_ADDED/PARTS_DATA`、`WORKORDER_DATA`。
- `MONTH_DATA` 映射（月份 → 五个 MDATA 常量）**必须含当月**：月份下拉框选项由 `Object.keys(MONTH_DATA).sort()` 生成，`applyMonth` 也靠它取数。
- **4 份副本**：`dashboard_offline.html`（源）/ `dashboard.html` / `deploy_cs/index.html` 三者 md5 必须一致；`dashboard_ref.html` 是 build 的**底板**——**改样式/JS 必须同步写它**，但**数据更新不要写它**（留作 check 对照，其数据通常比线上旧）。
- 缓存穿透：页面 `fetch('version.json')` 按版本号去重（`localStorage._cbTried`，非 sessionStorage 计数）+ 红色提示条兜底。

## 三、工具链与口径
- `build_month.py check <cut>`（与底板对照）/ `build <cut>`（只换当月分支，保留历史月 + 注入 version.json/BUILD）。开头自动跑 `G.verify_columns()` 列序自检。
- `gen_v2_lib.py`（Excel 解析：workorder/valueadded/valueparts/bigcare4/extend_time/cleaning）、`gen_month.py`（D/T/W）。
- **环境**：必须用 venv python `C:/Users/40973/.workbuddy/binaries/python/envs/default/Scripts/python.exe`（系统 python 无 openpyxl）。Bash 需带 PATH 前缀（PortableGit usr/bin + mingw64/bin）。
- **截止日铁律**：截止日 = 更新日前一天；服务月 = 上月 28 日 → 当月 27 日（`service_month()`，**gen_v2_lib 与 gen_month 各一份，改一处必须同步**）。
- 除 D/T 外全部按截止日过滤；**D/T 取「服务产品收入统计」全量快照**，换 Excel 必变，属正常。
- 校验脚本：`scan_dates2.py`/`scan_head.py`（扫各表最大日期）、`verify_build.py`（历史月+功能回归）、`chk_w_off.py`（周数据归属）、`dump_head.py`（表结构诊断）、`_check_diff915.py`（语义比对，`prev` 必须指向 `_prev.html`）。

### 口径坑（血泪清单）
- **平台会挪列序**：工单/自购表已改按表头名解析（`_cols` + `WO_HEADERS`/`ZG_HEADERS`），其余靠 `verify_columns` 自检。判据：**报 3–8 条错位 = 挪列，可兜底；缺关键业务列（如配件名称）= 这张表导错了，必须重导**。
- `工单自购配件明细` 无办事处列（col15 是大区部）；`CSM配件` col0 也是大区不是办事处 → 技师归属须反查。
- W 块清洁耗材曾用 CSM配件 col0（大区）当办事处 → 假办事处「西北大区部」。已按表头定位 + 「必须以办事处结尾」兜底。
- 明细表跨 Excel 行序不可复现，只校验「多重集一致」；`valueparts.code[]`、`bigcare.meta.剔除项` 是首现序，顺序变化属预期。
- E 块必须由已过滤延保明细聚合；`EXTEND_TIME_DATA` 记录 `c` 恒为空串。
- 百分比统一 `toFixed(2)`；KPI 环形字号 13；不改 CSS 宽度、`hsl()`、条形图 `Math.round`。

## 四、当前状态（2026-10-30 10:40，BUILD **20261030-1040**，截止 **2026-10-29**，**首个 11 月构建**）
- Excel `2026年11月西北服务产品.xlsx`（md5 `b71ed871…`，10:27 导出，**仅 1.0MB**）——换月后首份导出，**只含 10-28~10-29 两天**（工单 1,564 / CSM服务项目 453 / CSM配件 598 / 自购明细 378 / 服务产品收入统计 163 行）。工作表结构与 10 月完全一致（含 `WMS网点买断配件明细`）。
- 服务月 **2026-11 = 10.28–11.27（31 天）**，`CURRENT_MONTH=2026-11`，数据块 **45 → 50 个**（新增 `MDATA_2026_11_{D,T,B,W,E}`）。
- **2026-11 核心数（2 天）**：工单 **1,564**（=Excel 行数）/ 止回阀 103 / 烟机安装 536（**19.22%**）/ 增值产品 160（8/12 有销量）/ 增值配件当月 179（昨日 85、近7日 179）/ 大保养 **0÷52 = 0.00%**（总单 56）/ 延保 **9 条 ¥3,251.8**（y3 9 单、6 人）/ 清洗 **52 单**（免费 9 / 收费 43 ¥4,159）/ D **¥39,803.2（68 网点）**/ T **157 人 ¥26,543.8**。
- 达成：11 月目标 **¥544,500**（⚠️ 比 10 月的 495,000 高 10%，是 EMBED_TARGETS_BY_MONTH 里既有的月目标，不是我改的），达成率 **7.31%**、时间进度 6.45%、三大爆品 7.33%；周数据仅第 1 周有值 ¥12,176.8。
- 四端：Pages `5d7c887` ✅、旧 .link ✅（均 `20261030-1040`）、三副本 md5 `621c9e4f…` ✅、备份仓 `cfa66f7` 双推 ✅；家里 host `20261029-1446`（截止 10-28，慢一天）。
- ⚠️ **10 月报表只到 10-20，缺 10/21–10/27 七天**：11 月这份导出从 10/28 起，不含 10 月补数；要补全 10 月需再导一份 9.28–10.27 的「10 月」数据。
- ⚠️ **年度累计 `MONTHLY_FOUR.year` 仍停在「1–10 月、截止 10-04」**（总工单 180,378），这是固定快照、无法由单月 Excel 重算；11 月视图不影响，年度视图会偏旧。
- ⚠️ **11 月的「去年同期」未填**：`LASTYEAR_BY_MONTH["2026-11"]` 由 `seedLastYear()` 自动补空，同比列会显示 0/空，需在「设置去年同期」补 2025-11 的数。

### 本次新增认知（2026-10-30 跨月构建）
- 🐞 **跨月 build 会把上个月退回底板旧快照（已修）**：`build_month.merge_retained` 旧逻辑只「补缺月」——月份已存在就保留 REF 的。REF 是 10-16 底板，其 2026-10 只有 **4,629 单/10-04**，而线上是 16,475 单/10-20。10 月内的构建因当月会被替换而看不出来，**第一次跨到新月份就静默丢数**。已改为 **prev（上一版 build）优先覆盖**，REF 只补 prev 没有的月份/块；`verify_build.py` 增加「历史月逐月比对 _prev.html，差异必须为 0」。
- 🐞 **新月份根本没有 MDATA 块**：`build()` 用 `replace_value` 只会替换已存在的声明，`MDATA_2026_11_*` 不存在会直接 `SystemExit`。已加「有则替换、无则插到首个 MDATA 声明前」+ `ensure_month_data()` 往 `MONTH_DATA` 补一行（下拉框选项由 `Object.keys(MONTH_DATA).sort()` 生成，补上即自动多出该月）。
- ✅ **跨月后必须做浏览器实测**：静态校验查不出运行时问题。用本地 Edge 无头渲染即可（`msedge.exe --headless=new --disable-gpu --dump-dom file:///D:/WorkBuddy/dashboard_offline.html`），再在 DOM 里核对月徽标/周徽标/工单数是否为新月份值。


### 本次新增认知（2026-10-21）
- 🐞 **「网点买断」单据会被平台回溯撤回**：用 git 里的旧 Excel 逐行对比证明，CSM配件 窗口内 5,017 行**逐行完全一致**，但自购配件表少了 244 行（零新增、全是旧行消失，集中在 10-09~10-15），其中 35 行「网点买断」命中增值耗材 → valueparts 当月比上一版少 35 件。**不是代码 bug，是源数据撤回**；这类波动以后还会出现，历史月快照无法重算。
- 🐞 周数据 `m2o`（网点→办事处）原只取自 CSM服务项目，少数网点在该表办事处字段为空/填成「XX服务中心」→ 归成「—」（本次 1 个网点 ¥99）。已加 **CSM配件 预建映射**兜底（`gen_month.gen_weekly`）。

## 五、家里 vs 办公室两套实现
- 家里：`run_update.py` + `_gen_*.py` + 4 个累积 json（**1–9 月历史状态，漏拷会让年度累计变小**）；办公室：`build_month.py` + `gen_v2_lib.py` + `gen_month.py`。数据块兼容但代码独立，会各自漂移。办公室已支持任意月份。
