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
- 全局块：`MONTHLY_FOUR.months[月].{workorder,valueadded,valueparts,bigcare}` + `.year`（**2026-11-06 起为活的年度累计：203,951 单 / 截止 11-04**，由家里套件算出）、`EXTEND_TIME_DATA`、`CLEANING_DATA`、`VALUE_ADDED/PARTS_DATA`、`WORKORDER_DATA`、`UNIT_REVENUE_DATA`（单均收入，家里套件独有）。
- `MONTH_DATA` 映射（月份 → 五个 MDATA 常量）**必须含当月**：月份下拉框选项由 `Object.keys(MONTH_DATA).sort()` 生成，`applyMonth` 也靠它取数。
- **4 份副本**：`dashboard_offline.html`（源）/ `dashboard.html` / `deploy_cs/index.html` 三者 md5 必须一致；`dashboard_ref.html` 是 `build_month.py` 的**底板**——**改样式/JS 必须同步写它**。2026-11-06 起 ref 已同步为新结构（含 `UNIT_REVENUE_DATA`、10 月整月），旧 ref（5.00MB / 无单均收入）存为 `_ref_old_1016.html` 后删除，git 历史可回溯。
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

## 四、当前状态（2026-11-06 10:28，BUILD **20261106-1028**，截止 **2026-11-05**，家里工具链产物）
- **来源**：家里更新包 `_pkg_1105/` 里的 `run_update.py`。用户**不用再发包**——直接拿新 Excel 在该目录重跑 `run_update.py --excel <xlsx> --cut <D-1> --no-publish`，累积 json / 锁死月都在目录里，跑完用 `_adopt_home_dashboard.py <BUILD> <cut>` 接入工作区（换缓存穿透段 + 写 3 副本 + version.json）。
- Excel `2026年11月西北服务产品.xlsx`（**md5 `a6e691f5…`，4.18MB，5 张明细表统一到 11-05**；工单 7,013 / CSM服务项目 2,152 / CSM配件 2,788 / 自购明细 1,685 / WMS买断 84）。
- 服务月 **2026-11 = 10.28–11.27（31 天）**，`CURRENT_MONTH=2026-11`，50 个 MDATA 块，`MONTH_DATA` 含 2026-01~2026-11。
- **2026-11 核心数（截止 11-05，第 2 周/共 5 周，时间进度 29.03%）**：工单 **7,013** / 网点 94 / 工程师 229 / 防火阀（止回阀）占比 **18.96%** / 增值产品 **683** / 增值配件当月 **806** / 大保养 **1.66%** / 延保 **64 条 ¥23,930.50** / 清洗 **241 单 ¥19,320**（大小保养合计 249，剔除项口径差）/ 单月收入 **¥170,950（87 网点）** / 达成率 **31.40%**。
- **年度累计（`MONTHLY_FOUR.year`，活算）**：工单 **204,717**、防火阀 17.92%、增值产品 **17,330**、配件 **20,067**、大保养 1.81%（164/9,377），截止 11-05。
- **10 月锁死**：`_frozen/2026-10`（21,955 单 / 到 10-27），每次跑更新自动还原，不被重算。
- 四端：Pages commit `80d9dd2` ✅、旧 `.link` ✅（均 `20261106-1028`）、三副本 md5 `f6aae1cc…` 一致 ✅、备份仓 `1ba9994` 双推 ✅；家里 host 仍 `20261105-1033`（办公室覆盖不了）。

### 本次新增认知（2026-11-06 接入家里更新包 / 11-05 增量）
- **接包流程**：解压 → 读 `先看这个.txt` + `_frozen/*/meta.json` + `publish_gh/version.json` → 扫 Excel 各表最大日期（`_selfcheck_source.py` 已含此检查）→ 包目录跑 `run_update.py --cut <截止> --no-publish` → `_cmp_blocks.py` 与上一版逐块比对（应只有当月块 + MONTHLY_FOUR + CLEANING_DATA + EXTEND_TIME_DATA + UNIT_REVENUE_DATA 变化，**历史月必须 0 差异**）→ `_adopt_home_dashboard.py <BUILD> <cut>` → `_chk_js.py` + Edge 无头 → 发布。
- 🐞 **最大的坑：家里 `dashboard.html` 的缓存穿透是旧实现**（`_cbN` 计数）。必须整段替换为办公室修复版（`_adopt_home_dashboard.py` 内置断言 `_cbTried` 存在、`_cbN` 消失）。
- 🐞 **`_merge_month.py` 漏改 `year.bigcare.meta` 文案**：workorder/valueadded/valueparts 的「截止日/服务月」都按 `cut` 动态修正了，只有 bigcare 沿用 `_gen_monthly_four.py` 里写死的「2026-08-26 / 2025-12-28~2026-08-26」，导致年度视图大保养卡显示「截止 08-26」而数值已在变。**已补 `yfix['bigcare']`**，重跑 `_merge_month.py --month 2026-11` + `_reinject_monthlyfour.py` 即修（幂等，year 数值不变）。
- ✅ **核对方法**：`_cmp_blocks.py` 抽顶层声明逐块 md5；`_chk_1105.py` 打印两月关键数；`_chk_js.py` 逐块 `node --check`；Edge `--headless=new --dump-dom` 查 KPI 卡（本版 **31 张**）、月徽标、周徽标、达成率、年度累计。
- ⚠️ 旧 `.link` 发布工具继续报「预留域名未绑定」假失败（本次是新后缀 `-37944`），`curl .../version.json` 已是新版即成功。
- `.gitignore` 已补 `_pkg_*/*/_bak_*`、`_pkg_*/*/_pre_*`、`_dom*.txt`、`~$*.xlsx`，避免接包备份把仓库撑胖。

## 五、家里 vs 办公室两套实现（**2026-11-06 起以家里那套为准**）
- 家里套件（包内 `fangtai-dashboard-update/`，已留存 `_pkg_1105/`）：`run_update.py` 一键 = 源表自检 → 生成（`_gen_month/workorder/valueadded/valueparts/bigcare/extend_time_months/cleaning/unit`）→ **还原 `_frozen/` 锁死月** → 合并注入（`_merge_month` / `_reinject_*`）→ 校验（截止日命中、默认月、内联 JS 语法、锁死月 `--verify`）→ 发布（需 `.gh_token`，办公室走 SSH 不用它）。累积状态 5 个 `_*.json` + `_month_YYYY-MM.json`（**一个都不能少，漏拷年度累计会变小**）。
- 办公室旧实现 `build_month.py + gen_v2_lib.py + gen_month.py`：**已落后**——没有 `_frozen` 锁死、没有 `UNIT_REVENUE_DATA`、年度累计是死快照、10 月数据需自己拼。**新数据不要再直接跑 `build_month.py build`**，否则会把单均收入模块和锁死机制丢掉（历史月虽能靠 `_prev` 续命，但会埋雷）。要用它就先开一个新月份话题确认。
- 家里套件落后于办公室的唯一一处：**缓存穿透仍是旧的 sessionStorage `_cbN` 计数版**（微信里次数耗光就永久失效 → 手机停在旧月份）。所以每次接家里包都必须把这段换回办公室的「按版本号去重 + 红色提示条」版本（见技能）。
- 家里只能更 Pages；`deploy_cs` 与应用链接必须在办公室发布。
