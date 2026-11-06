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

## 四、当前状态（2026-11-06 10:03，BUILD **20261106-1003**，截止 **2026-11-04**，**已切换到家里的工具链产物**）
- **来源变更**：本版不再由 `build_month.py` 生成，而是接入用户家里电脑的更新包（`方太看板_家里电脑更新包_20261105.zip`），用包里的 `run_update.py` 在办公室重跑复现（cut 11-04，--no-publish），逐块比对与家里原版**数值完全一致**（仅 3 处差异：`CLEANING_DATA.meta.生成` 时间戳、`MDATA_2026_11_D.meta._stamp`、`MONTHLY_FOUR` 里 `pname` 首次出现顺序）。
- Excel `2026年11月西北服务产品.xlsx`（**md5 `10c96870…`，4.35MB，4 张表统一到 11-04**；工单 6,247 / CSM服务项目 1,890 / CSM配件 2,354 记录 / 自购明细 1,479）。注意工作区原来那份 1.0MB（`b71ed871…`，只到 10-29）已被覆盖为新版。
- 服务月 **2026-11 = 10.28–11.27（31 天）**，`CURRENT_MONTH=2026-11`，**50 个 MDATA 块**、`MONTH_DATA` 含 2026-01~2026-11。
- **2026-11 核心数（截止 11-04）**：工单 **6,247**（=Excel 行数）/ 网点 94 / 工程师 227 / 止回阀占比 **18.69%** / 增值产品 **609** / 增值配件当月 **729** / 大保养 **1.37%** / 延保 **52 条 ¥19,579.50**（全 y3、26 人）/ 清洗 **219 单 ¥17,113**（免费 34 / 收费 185；油烟机小保养 216 + 大保养 3）/ 单月实际总收入 **¥146,575**（86 网点）/ 达成率 **26.92%**、时间进度 **25.81%**、三大爆品 **32.53%**（目标 ¥544,500）。
- ✅ **本次接入补齐了之前三个缺口**（家里包自带累积状态与锁死月）：
  1. **10 月整月补全**：`21,955 单 / 99 网点 / 243 工程师 / 止回阀 18.69% / 增值产品 2,254 / 配件 2,520 / 大保养 1.09%`，截止 **10-27**（之前办公室只到 10-20、16,475 单）。`_frozen/2026-10` 已锁死，跑更新时自动还原、不会被重算。
  2. **年度累计活了**：`MONTHLY_FOUR.year` 现在是 **203,951 单、截止 2026-11-04**（旧版是固定快照 180,378 / 截止 10-04）。
  3. **11 月去年同期已填**：`EMBED_LASTYEAR_BY_MONTH['2026-11']` = 清洗保养 74,730 / 延保 10,053.1 / 清洁耗材 21,862 / 止回阀 126,260（`_frozen/2026-11` 锁死，stamp 2026-10-29）。
- **新增模块 `UNIT_REVENUE_DATA`（单均收入）**：办公室旧实现没有；`months['2026-10'/'2026-11']`，含「工单全量 / 远程指导剔除 / 上门工单 / 有收费工单 / 出单工单」等口径。
- 四端：Pages commit `facc616` ✅、旧 .link ✅（均 `20261106-1003`）、三副本 + 底板 md5 `582fe1f7…` 一致 ✅、备份仓 `08d4908` 双推 ✅；家里 host `20261105-1033`（同一份数据，仅 BUILD 标签旧）。
- ⚠️ 新 CloudStudio 域名 `fangtai-nw-dashboard-98984.app.workbuddy.host` 返回「链接已失效」——该 app 从未创建成功，**对外仍只发 Pages 与旧 `.link` 两条**。
- 家里包已解压留存 `D:\WorkBuddy\_pkg_1105\fangtai-dashboard-update\`（含全部累积 json、`_frozen/`、`数据源参考/`），**已随备份仓提交**，勿删。

### 本次新增认知（2026-11-06 接入家里更新包）
- **接包流程**：解压 → 先读 `先看这个.txt` + `_frozen/*/meta.json` + `publish_gh/version.json`（确认包里截止日/锁死月份）→ 扫 `数据源参考/` 那份 Excel 各表最大日期 → 在包目录跑 `run_update.py --excel 数据源参考/xxx.xlsx --cut <截止> --no-publish` → 与包内原始 `dashboard.html` 逐块比对（`_cmp_blocks.py`，只该有时间戳/首现序差异）→ 拷成 `dashboard_offline.html` 并**补回缓存穿透修复** → 改 BUILD → 同步 3 副本 + `version.json` → 发布。
- 🐞 **最大的坑：家里 `dashboard.html` 的缓存穿透是旧实现**。直接拿来发布会让 10-16 修好的「手机停在旧月份」问题复发（`_cbN` 计数耗尽即永久失效）。必须把 `/*__CACHEBUST_START__*/ … /*__CACHEBUST_END__*/` 整段替换为办公室修复版，并同步改 `var BUILD`（脚本 `_adopt_home_dashboard.py` 就是干这个的）。
- ✅ **核对方法**：`_cmp_blocks.py` 抽出两份 HTML 的顶层声明逐块 md5 比对；正常结果应只有 `CLEANING_DATA.meta.生成`、`MDATA_*_D.meta._stamp`、`MONTHLY_FOUR` 的 `pname` 首现序 3 处不同。
- ✅ **无头渲染必做**：`msedge --headless=new --dump-dom`，看 KPI 卡数量（本版 30 张）、月徽标、周徽标、达成率；DOM 里的 `NaN`/`undefined` 基本都来自内置 SheetJS，不是数据问题。
- ⚠️ 旧 `.link` 发布工具继续报「预留域名未绑定」假失败；`fangtai-nw-dashboard-98984` 这个新域名是死链，别发给同事。


- 🐞 **跨月 build 会把上个月退回底板旧快照（已修）**：`build_month.merge_retained` 旧逻辑只「补缺月」——月份已存在就保留 REF 的。REF 是 10-16 底板，其 2026-10 只有 **4,629 单/10-04**，而线上是 16,475 单/10-20。10 月内的构建因当月会被替换而看不出来，**第一次跨到新月份就静默丢数**。已改为 **prev（上一版 build）优先覆盖**，REF 只补 prev 没有的月份/块；`verify_build.py` 增加「历史月逐月比对 _prev.html，差异必须为 0」。
- 🐞 **新月份根本没有 MDATA 块**：`build()` 用 `replace_value` 只会替换已存在的声明，`MDATA_2026_11_*` 不存在会直接 `SystemExit`。已加「有则替换、无则插到首个 MDATA 声明前」+ `ensure_month_data()` 往 `MONTH_DATA` 补一行（下拉框选项由 `Object.keys(MONTH_DATA).sort()` 生成，补上即自动多出该月）。
- ✅ **跨月后必须做浏览器实测**：静态校验查不出运行时问题。用本地 Edge 无头渲染即可（`msedge.exe --headless=new --disable-gpu --dump-dom file:///D:/WorkBuddy/dashboard_offline.html`），再在 DOM 里核对月徽标/周徽标/工单数是否为新月份值。


### 本次新增认知（2026-10-21）
- 🐞 **「网点买断」单据会被平台回溯撤回**：用 git 里的旧 Excel 逐行对比证明，CSM配件 窗口内 5,017 行**逐行完全一致**，但自购配件表少了 244 行（零新增、全是旧行消失，集中在 10-09~10-15），其中 35 行「网点买断」命中增值耗材 → valueparts 当月比上一版少 35 件。**不是代码 bug，是源数据撤回**；这类波动以后还会出现，历史月快照无法重算。
- 🐞 周数据 `m2o`（网点→办事处）原只取自 CSM服务项目，少数网点在该表办事处字段为空/填成「XX服务中心」→ 归成「—」（本次 1 个网点 ¥99）。已加 **CSM配件 预建映射**兜底（`gen_month.gen_weekly`）。

## 五、家里 vs 办公室两套实现（**2026-11-06 起以家里那套为准**）
- 家里套件（包内 `fangtai-dashboard-update/`，已留存 `_pkg_1105/`）：`run_update.py` 一键 = 源表自检 → 生成（`_gen_month/workorder/valueadded/valueparts/bigcare/extend_time_months/cleaning/unit`）→ **还原 `_frozen/` 锁死月** → 合并注入（`_merge_month` / `_reinject_*`）→ 校验（截止日命中、默认月、内联 JS 语法、锁死月 `--verify`）→ 发布（需 `.gh_token`，办公室走 SSH 不用它）。累积状态 5 个 `_*.json` + `_month_YYYY-MM.json`（**一个都不能少，漏拷年度累计会变小**）。
- 办公室旧实现 `build_month.py + gen_v2_lib.py + gen_month.py`：**已落后**——没有 `_frozen` 锁死、没有 `UNIT_REVENUE_DATA`、年度累计是死快照、10 月数据需自己拼。**新数据不要再直接跑 `build_month.py build`**，否则会把单均收入模块和锁死机制丢掉（历史月虽能靠 `_prev` 续命，但会埋雷）。要用它就先开一个新月份话题确认。
- 家里套件落后于办公室的唯一一处：**缓存穿透仍是旧的 sessionStorage `_cbN` 计数版**（微信里次数耗光就永久失效 → 手机停在旧月份）。所以每次接家里包都必须把这段换回办公室的「按版本号去重 + 红色提示条」版本（见技能）。
- 家里只能更 Pages；`deploy_cs` 与应用链接必须在办公室发布。
