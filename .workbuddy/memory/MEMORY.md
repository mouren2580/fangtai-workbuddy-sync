# 项目长期记忆（方太西北服务产品月报 / WorkBuddy 工作区）

> 详细操作步骤已固化为技能：`D:\WorkBuddy\.workbuddy\skills\fangtai-dashboard-publish\SKILL.md`（改数据/改样式/发布四端 SOP）。本文件只记「事实、口径、坑、当前状态」。

## 一、仓库与推送（3 个远端）
- **备份仓**：Gitee `git@gitee.com:xbftsjkb/fangtai-workbuddy-sync.git`（远程名 `origin`，私有一份）＋ GitHub `git@github.com:mouren2580/fangtai-workbuddy-sync.git`（远程名 `github`，**仍 Public，建议用户改 Private**）。
- **Pages 站**：`https://mouren2580.github.io/fangtai-dashboard/`，源 `git@github.com:mouren2580/fangtai-dashboard.git`（根目录 `main`）。**这是分享给同事的主链接，每次更新必须同步**。
- **认证**：GitHub 全部走 SSH `~/.ssh/id_ed25519_github`（同一把钥匙用于 Pages 仓与备份仓）；Gitee 用 `~/.ssh/id_ed25519_gitee`。**PAT 已全部失效，别再折腾 PAT。**
- **任何读 `~/.ssh` 的 git 操作都要 `dangerouslyDisableSandbox: true`**；`push_gh_pages.py` 不能后台跑（SSH 会挂死），必须前台。
- ⚠️ Gitee 单文件 100MB 硬限制；`git add -A` 前确认无大中间产物。`.gitignore` 已含 `models/`、`*.bin`、`2026-09-12-20-12-51/`、`_prev*.html`、`_check_*.py`、`_bak_*`。
- 流程：改数据 → `git push origin main` + `git push github main` → `python push_gh_pages.py`。

## 二、看板架构（单文件 HTML）
- 多月份：顶层 `MDATA_YYYY_MM_{D,T,B,W,E}`（D 主数据/T 工程师/B 大保养/W 周数据/E 延保）；`BASE_MONTH="2026-08"` 锚点（8 月用基准块 `DASHBOARD_DATA/TECH_DATA/BIG_CARE_DATA/WEEKLY_DATA/EXTEND_WARRANTY_DATA`，故**无 `MDATA_2026_08_*` 属正常**）；`CURRENT_MONTH` 默认 `2026-09`。
- 另有全局块：`MONTHLY_FOUR.months[月].{workorder,valueadded,valueparts,bigcare}`、`.year`（1–8 月累计，**当月 Excel 无历史数据、不可重算，保持原样**）、`EXTEND_TIME_DATA.months[月]`（延保明细）、`CLEANING_DATA.months[月]`（清洗保养明细）、`VALUE_ADDED_DATA`/`VALUE_PARTS_DATA`/`WORKORDER_DATA`。
- 缓存穿透：页面 `fetch('version.json')` 比 BUILD；用户按 **Ctrl+Shift+R** 硬刷新。`?newPanel=true` 之类查询参数无效。
- **同一份看板 4 个副本**：`dashboard_offline.html`（源）/ `dashboard.html`（离线交付）/ `deploy_cs/index.html`（CloudStudio）三者 md5 应一致；`dashboard_ref.html` 是 `build_month.py` 的**底板**（存用户参考版数据），**改样式/JS 也要写进它，否则下次 build 会被覆盖回旧样式**；但**数据更新不要写它**。

## 三、数据工具链
- **入口 `build_month.py`**（顶部 `EXCEL` 指向当月 Excel、`REF` → `dashboard_ref.html`）：
  - `python build_month.py check <截止日>`：按该截止日重算并与底板逐块对照（底板数据是旧截止日，故 DIFFERS 属正常，看 summary 数字）；
  - `python build_month.py build <截止日>`：只替换 `2026-09` 分支、保留 1–8 月，注入 `dashboard_offline.html` + `version.json` + `var BUILD`。`merge_retained()` 会补回底板缺的历史月。
  - `make_blocks()` 开头会跑 `G.verify_columns(EXCEL)` **列序自检**。
- **解析库 `gen_v2_lib.py`**：`build_workorder/build_valueadded/build_valueparts/build_bigcare4/build_extend_time/build_cleaning`；`gen_month.py` 保留供 D（服务产品收入统计）/T（工程师）/W（周数据）。
- **环境**：解析 Excel **必须用 venv python** `C:/Users/40973/.workbuddy/binaries/python/envs/default/Scripts/python.exe`（系统 3.13.12 无 openpyxl）。Bash 里所有命令要带 PATH 前缀（PortableGit usr/bin + mingw64/bin），否则 `dirname`/`git` 都找不到。
- **截止日铁律**：`截止日 = 更新日前一天`；服务周期 = 上月 28 日 → 当月 27 日（9 月 = 8.28–9.27，31 天，时间进度 = 已过天数 ÷ 31）。
- **口径**：除 D/T 外全部按截止日过滤；**D/T 取「服务产品收入统计」报表全量（不过滤日期）**，故换 Excel 快照必变，`check` 对不上属正常。

### 🐞 列序坑（2026-09-18 踩，已修）
平台导出**会调整列顺序**。本次「工单」表把 `办事处名称` 由 L→K、`工程师编号` 由 AE→AF，硬编码列号导致：办事处名读空、技师编码读成「已支付」、办事处数只剩 2 —— **数量有、归属全错且不报错**。
- 已改为 **按表头名解析**：`gen_v2_lib._cols(ws, header_row, spec, tag)` + `WO_HEADERS`（工单表）/`ZG_HEADERS`（工单自购配件明细），找不到表头才回退旧常量并打印 `[warn]`。
- 另有 `verify_columns(excel)` 核对 CSM配件/CSM服务项目/服务产品收入统计/WMS买断 的固定列号是否仍对得上表头名，每次 build 自动跑。
- **后验手段**：把上一版 `dashboard_offline.html` 存成 `_prev*.html`，`build` 后用语义比对脚本（列表按递归多重集比，区分「仅顺序不同」与「真实差异」）验证历史截止日重算是否与上一版一致。

### 其他口径坑
- `工单自购配件明细` **没有办事处列**：col15 是「大区部」、col12 服务中心、col5 网点 → 技师归属必须用「技师编号→办事处」反查（源 `CSM配件` col2/col5/col6 或 `工单表` col30/col11）。`CSM配件` col0 也是「大区」不是办事处。
- 明细表按源表行序，**跨 Excel 文件排列不可复现**，只校验「多重集一致」。
- `valueadded.products[]` 有 `verify`，`valueparts.products[]` 没有；`valueparts` 配件名 `pname` 按源表首现序。
- 大保养：`清洗保养总单`=销售分类(col16)；`剔除项`/`大保养项`=服务收费项目(col13)，三者独立可重叠；分母 = 总单 − 剔除。`MDATA_*_B.meta` 用常量序，`MONTHLY_FOUR.bigcare.meta` 用源表首现序。
- `EXTEND_TIME_DATA` 记录 `c` 恒为空串、E 块 `techs[].id` = 姓名（E 必须由已过滤的延保明细聚合，`gen_extend` 不过滤日期不能用）；E 块逐年金额不四舍五入（保留浮点尾差），只有 `tamt` 舍入。
- 合计行固定文案 `[合计]全区`/`[合计]全部网点`/`[合计]全部工程师`。

## 四、发布与显示规范
- **CloudStudio 链接**：`https://0717bc4b30824b8d8a407555473b321e.app.workbuddy.link`（目录 `deploy_cs`），用 `workbuddy_sites_deploy`（旧名 `workbuddy_cloudstudio_deploy`）。
  - ⚠️ **必须先问用户是否覆盖线上**（工具强约束：对已有线上链接的目录不会静默覆盖）。
  - ⚠️ 返回「预留域名未绑定…本次发布已停止」= **假失败**（内容其实已发布，平台正从 `.link` 迁 `.host`）→ 直接 `curl <链接>/version.json` 看 `v`/`cut` 确认，别重试。
  - ⚠️ 不要传 `updateExistingApp:true`（本工作区会报 no existing app to update）。
- **显示规范**：所有百分比两位小数（`const pct = x => (x*100).toFixed(2) + '%'`，另需改同比/合计/品项分布/大保养/自购耗材/年度面板/技师面板的 `.toFixed(1)`）；**不要改** CSS 宽度、`hsl()`、条形图 `Math.round`。KPI 环形图文字用 `font-size="10.5"`。自检 `grep -n "toFixed(1)"` 只应剩条形图一处。

### ⚠️ 用户发来 dashboard(x).html 时必须先 diff，不能直接覆盖
用户会把自己另存/改过的看板发回来，里面可能有**本地没有的新功能**。做法：把长行截断到 150 字符后用 `difflib.SequenceMatcher` 比对骨架，逐块看差异再决定。
- 2026-09-27 那次用户版比本地多 14KB，含：①**「各品项占比环形图」模块**（`.iPieScope`/`#itemPie`/`#itemScopeSeg`/`PIE_COLORS`/`pieColorOf()`/`renderItemPie()`/`renderItemScopeSeg()`，可切全区/办事处、点击下钻）；②条形图改按品项配色；③技师面板 `pct` 改 `toFixed(2)`（我 9-15 漏改的）；④「锁定规则 2026-09-21」注释；⑤KPI 环形字号用 **13**（非 10.5，以用户版为准）。
- 处理：把用户版同时覆盖 `dashboard_offline.html` 与 `dashboard_ref.html`（作新底板），再 build 新数据。

### 语义比对脚本的已知局限
`_check_diff915.py` 对「dict 列表」用**字符串**多重集比，浮点尾差（如 `6251.700000000001` vs `…002`）会被误报成真实差异；`bigcare.meta.剔除项/大保养项`、`valueparts.code[]` 是**源表首现序**，跨 Excel 顺序必变。这两种都属预期，别当 bug 查。

## 五、当前状态（2026-09-27 09:00）
- BUILD `20260927-0859`，**截止 `2026-09-26`**（周期 8.28–9.27，时间进度第 30 天），四端（Pages `a9c37ff` / CloudStudio / 工作区 / 备份仓 `983e7a7`）同步。
- 核心数字：工单 **21,357**（= Excel 行数，自洽）/ 止回阀 **1,336**（烟机安装 7,162、占比 **18.65%**）/ 大保养 **16÷704＝2.27%**（总单 718）/ 延保明细 **185 条 ¥69,049.60** / 清洗保养明细 **703 条** / 增值产品 2,120 / 增值配件当月 2,877 / D 报表全量 **¥515,160.26**（95 网点）/ T **222 人 ¥423,895.00**。
- 数据源 Excel 已切到 `2026年9月西北服务产品(5).xlsx`（(3) 已被用户移走）；底板 `dashboard_ref.html` = 用户 9-26 版（含品项环形图）。
- 历史月 1–8 月 + 8 月清洗 788 / 延保 139 + 40 个 MDATA 块完整 ✅；新功能 9 项抽样保留 ✅；真实 `toFixed(1)` 残留 0 处。
