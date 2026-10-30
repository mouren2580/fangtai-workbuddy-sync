---
name: fangtai-dashboard-publish
description: 方太西北服务产品看板做任意改动（数据/样式/格式）后，同步到本地 4 份副本 + GitHub Pages + CloudStudio + 两个备份仓的四端发布 SOP。改看板样式、百分比格式、文案、BUILD 版本号时使用。
agent_created: true
---

# 触发条件
- 用户要求改方太看板的显示/样式/文案（如「百分比显示两位小数」「某处字号」「图例文案」）
- 用户要求把看板更新到线上（GitHub Pages / CloudStudio）
- 用户发来新 Excel 需要重建数据（另见 build_month.py 流程，见文末）

# 环境准备（每条 Bash 都要带）
```bash
export PATH="/c/Users/40973/.workbuddy/binaries/PortableGit/versions/1.2.0/usr/bin:/c/Users/40973/.workbuddy/binaries/PortableGit/versions/1.2.0/mingw64/bin:/c/WINDOWS/system32:$PATH"
export PYTHONIOENCODING=utf-8
PY="C:/Users/40973/.workbuddy/binaries/python/envs/default/Scripts/python.exe"   # ← 必须用这个 venv
```
裸 `bash` 里 `dirname`/`cd` 会报错，`python`/`git` 也不在 PATH，必须显式加前缀。
⚠️ **脚本路径别写 `/d/WorkBuddy/x.py`**：会被解析成 `D:\d\WorkBuddy\x.py` 报 `can't open file`。用 `D:/WorkBuddy/x.py`，或先 `cd /d/WorkBuddy` 再写相对路径。
⚠️ **不要用** `binaries/python/versions/3.13.12/python.exe`（系统级那个）——它**没装 openpyxl**，跑 `build_month.py` 会
`ModuleNotFoundError: No module named 'openpyxl'`。解析 Excel 一律用上面 venv 里的 python（openpyxl 3.1.5）。
⚠️ **git 操作路径必须写 Windows 形式**（`C:/Users/...` 或相对当前目录）：`git -C "/c/Users/..."` 会报
`fatal: cannot change to '...': No such file or directory`（git.exe 不认 MSYS `/c/` 路径）。bash 的 `cp`/`ls` 两种都认。

# 关键认知：同一份看板有 4 个副本，必须一起改
| 文件 | 用途 |
|---|---|
| `D:\WorkBuddy\dashboard_offline.html` | **权威源**。改这里，然后复制到下面两个 |
| `D:\WorkBuddy\dashboard.html` | 给单位离线使用的副本 |
| `D:\WorkBuddy\deploy_cs\index.html` | CloudStudio 部署目录内容 |
| `D:\WorkBuddy\dashboard_ref.html` | **build_month.py 的底板**。样式/JS 类改动必须也写进它，否则下次 `python build_month.py build <截止日>` 重建时改动被覆盖 |

三份 HTML 改完应完全一致：
```bash
md5sum dashboard_offline.html dashboard.html deploy_cs/index.html   # 三者相同
```
`dashboard_ref.html` 的 BUILD 值与数据块与它们不同（它是参考底板），只要求承载了同一份样式/JS 改动。

# 版本号 / 缓存穿透
改完内容后，4 份 HTML 里的 `var BUILD = "YYYYMMDD-HHMM";`（第 15 行）都要改成新时间戳，并更新：
```bash
C:/.../python.exe -c "
import json,datetime;print(json.dumps({'v':'YYYYMMDD-HHMM','cut':'2026-09-14','t':datetime.datetime.now().strftime('%Y-%m-%dT%H:%M:%S')},ensure_ascii=False))"
```
写入 `version.json` 和 `deploy_cs/version.json`（页面 `fetch('version.json')` 比对 BUILD 做缓存穿透，字段名是 `v`/`cut`/`t`，**不是** `version`/`cutoff`）。

### 🐞 缓存穿透失效：手机一直显示几个月前的旧页面（2026-10-16 修）
症状：用户反馈「手机打开链接显示到 8 月」。先确认**服务端没问题**（`curl <链接> | grep -o 'var BUILD = "[^"]*"\|let CURRENT_MONTH = "[^"]*"'` 均为最新），
那必然是手机端缓存的旧 HTML 没被换掉。
- 根因：旧实现用 `sessionStorage` 计数兜底防死循环——`var n=parseInt(sessionStorage.getItem('_cbN')||'0',10); if(n>=2) return;`。
  **微信/QQ 内置浏览器的 sessionStorage 常长期不清除**，次数耗尽后 `return` 直接退出，缓存穿透**永久失效**，之后每次数据更新手机都看不到。
- 修复：改成**按版本号去重**（`localStorage['_cbTried']`）——每个新版本只自动重载一次；同一版本已尝试过仍不一致才提示人工刷新，既不死循环也不卡死。
  另加了红色提示条兜底：`发现新版 xxx（当前显示 yyy）→ 点此刷新`。
- 这段代码在 `/*__CACHEBUST_START__*/ … /*__CACHEBUST_END__*/` 块里，**改它必须 4 份 HTML 一起改（含 `dashboard_ref.html` 底板）**，
  否则下次 `build_month.py build` 会把底板里的旧实现覆盖回来。各份文件的 `var BUILD` 值不同，Edit 时 old_string 要带各自的 BUILD。
- 改完自检：`grep -o 'var BUILD = "[^"]*"' dashboard_offline.html dashboard.html deploy_cs/index.html dashboard_ref.html` 四份应一致；
  JS 语法可用 node 校验：提取该 script 段到临时 `.js` 再 `node --check`（注意用 `src.rindex('<script>', 0, i)` 限定在前缀内，
  写成 `rindex('<script>', i)` 会搜到 CACHEBUST_END 之后的 script，切出空串）。

# 百分比显示规范（用户要求：全部两位小数）
- 全局格式化函数：`const pct = x => (x*100).toFixed(2) + '%';`（一行改动覆盖达成率/时间进度/进度偏差/占比等大多数显示位）
- 其余用 `.toFixed(1)+'%'` 直出的点（看过板版本可能行号漂移）：
  - 同比列 `(yoy*100)`、合计行 `(yoySum*100)`、品项分布 `(x.r*100)`
  - 大保养明细 `r.pct.toFixed(2)`、自购耗材占比 `(p.total/grand*100)`
  - 年度目标面板 `const pctTxt = ...(r*100).toFixed(2)`
  - 技师面板「占筛选合计」`(t.eff/sum*100).toFixed(2)`，空值兜底 `'0.00'`
  - 已符合的：防火阀占比（`ratio.toFixed(2)`）等
- **不要改**：CSS 宽度/高度（`style="width:${w}%"`）、`hsl(...%)` 颜色、`Math.round(x*100)` 用于条形图宽度的地方
- 改完必须自检残留：`grep -n "toFixed(1)" dashboard_offline.html` → 只应剩条形图宽度那一处
- ⚠️ KPI 卡右上角的**环形图**（`viewBox="0 0 56 56"`，`.kpi .ring` 56px）内文字从 `58.1%` 变 `58.06%` 会撞到环体，需把该 `<text>` 的 `font-size="13"` 调成 `font-size="10.5"`

# 发布三步
1. **GitHub Pages（同事主用链接，优先保证）**
   ```bash
   cd /d/WorkBuddy && "$PY" push_gh_pages.py       # 走 SSH，无需 PAT
   ```
   - 脚本自动：维护常驻克隆 → 覆盖 `index.html`+`version.json` → commit → push
   - **必须加 `dangerouslyDisableSandbox: true`**（要读 `~/.ssh`，沙箱会拦）
   - ⚠️ **千万不要用 `run_in_background` 跑它**：后台环境下 SSH 会被挂起，实测卡 25 分钟零输出、克隆目录里文件根本没被覆盖。
     必须**前台**执行（长超时 300000ms）。若已卡住，用 TaskStop 杀掉后走下面的手动流程。
   - 复核：`"$PY" push_gh_pages.py --status` 或 `curl -s https://mouren2580.github.io/fangtai-dashboard/version.json`
   - **手动 fallback（脚本卡死时用，旁路 + 前台）**：
     ```bash
     CL="C:/Users/40973/.workbuddy/tmp/fangtai-dashboard"     # 必须 Windows 路径形式
     export GIT_SSH_COMMAND="ssh -o StrictHostKeyChecking=accept-new -o ConnectTimeout=20"
     git -C "$CL" fetch --depth 1 origin main && git -C "$CL" reset --hard origin/main
     cp deploy_cs/index.html "$CL/index.html"; cp deploy_cs/version.json "$CL/version.json"
     git -C "$CL" add -A
     git -C "$CL" -c user.name=mouren2580 -c user.email=409737410@qq.com commit -q -m "看板数据更新至 <cut>"
     git -C "$CL" push origin main
     ```
     ⚠️ 只覆盖 `index.html` 与 `version.json`，仓库里的 `drainage/`、`sync-kit/`、`sync.sh`、`index.orig.html`、`.nojekyll` 一律不碰。
2. **CloudStudio**：`workbuddy_sites_deploy(directory="D:\\WorkBuddy\\deploy_cs", language="static", appName="方太西北服务产品月报看板", domainPrefix="fangtai-nw-dashboard", userAskedToPublish=true)`
   - **必须用户在本轮明确同意**才能覆盖线上链接（工具强约束，先问一句）
   - ⚠️ 报 `应用预留域名 ... 未绑定到本次发布环境` 是**假失败**：内容通常已发布。用 `curl -s https://0717bc4b30824b8d8a407555473b321e.app.workbuddy.link/version.json` 验证，返回新 BUILD 即成功，不要把报错原样转述给用户当失败
   - 传 `updateExistingApp:true` 可能因「工作区无现有 app 记录」报错，直接去掉该参数重试
   - ⚠️ **新链接 `fangtai-service-dashboard.app.workbuddy.host` 办公室这边覆盖不了**：它是家里那台机器建的 app，这边用同前缀部署只会新建带随机后缀的域名（`…-66565.app.workbuddy.host`），原链接不动。要更新只能在家里那台机器再发一次；或让同事看 Pages / 旧 `.link`。
3. **两个备份仓**
   ```bash
   export GIT_SSH_COMMAND="ssh -o StrictHostKeyChecking=accept-new"
   git add -A && git commit -m "..." && git push origin main      # Gitee
   git push github main                                          # GitHub
   ```
   - 同样需要 `dangerouslyDisableSandbox: true`
   - `git add -A` 前先 `git status --porcelain` 确认没有大文件（Gitee 单文件 100MB 硬限制；`.gitignore` 已挡 `models/`、`*.bin`、`_bak_*`）
   - 注意 `core.autocrlf=true`：用 Python 以 `io.open(...,newline='')` 读写会把工作区文件 CRLF→LF，文件字节数变小但 **git diff 只显示真实内容改动**，属正常

# 收尾验证（两端线上都查一次）
```bash
curl -s "https://mouren2580.github.io/fangtai-dashboard/version.json"
curl -s "https://0717bc4b30824b8d8a407555473b321e.app.workbuddy.link/version.json"
```
两边 `v` 应等于新 BUILD；必要时再抓页面印证具体改动（如 `grep -o "const pct = x => (x\*100)\.toFixed(2)"`）。GitHub Pages 构建需 1–3 分钟，刚推完立刻查可能是旧版。

# 相关：新 Excel 数据更新流程（数据类改动走这条）
**截止日铁律**：**截止日 = 更新日前一天**（如 10-06 更新 → 2026-10-05）。
**服务周期**：上月 28 日 → 当月 27 日。⚠️ **天数不固定**：9 月 8.28–9.27 = 31 天，但 **10 月 9.28–10.27 = 30 天**（取决于上月的天数）。
→ 已统一为 `service_month(cutoff)`（`gen_v2_lib.py` 与 `gen_month.py` **各一份，改一处必须同步另一处**），规则「每月 28 日及以后归下一个服务月；终点 = 当月 27 日；跨年 12-28 → 次年 1 月服务月」。
→ 🐞 旧代码 `end = svc + timedelta(days=30)` 只对 31 天服务月成立，10 月是第一个 30 天服务月才暴露（算出 9/28–10/28）。若再看到周期多一天，就是这里。

1. **先确认真实截止日**：别直接信"更新日−1"，扫一下 Excel 各表日期列最大值（工单 col41 / CSM服务项目 col7 / CSM配件 col36 / 工单自购配件明细 col10，0 基索引）。
   ⚠️ **表头行位置不统一（2026-10-08 踩到）**：`工单`/`工单自购配件明细` 表头在第 1 行（row index 0），但
   **`CSM服务项目`/`CSM配件`/`服务产品收入统计` 表头在第 2 行（row index 1，第 1 行整行空）**。
   → 扫日期脚本若直接拿第 1 行当表头，会拿到一行全空字符串，报"未找到日期列"。做法：**自动定位表头行**（第一个非空单元格数 ≥8 的行）再按名字找日期列。现成脚本 `scan_dates2.py`（第1行表头）+ `scan_head.py`（自动定位表头行）。
   ⚠️ **各表截止日可能不一致（2026-10-08 踩到）**：那次 `工单`/`CSM配件`/`工单自购配件明细` 已到 10-07，而 **`CSM服务项目` 停在 10-05**（该表未刷新，行数与前一次完全相同）。后果是**大保养 B / 清洗保养明细 / 延保明细 三块停在 10-05**，其余模块到 10-07。
   → 处理：**按主表（工单）的最大日期定截止日**，但必须**如实告知用户哪几块滞后**，别让人以为全量都到最新。
   → 判断某表是否刷新：比对该表总行数与上次扫描是否相同（本次 CSM服务项目 1,690 = 1,690）。
   ⚠️ **列序坑（2026-09-18 踩过）**：平台导出**会调整列顺序**。上次「工单」表把 `办事处名称` L(11)→K(10)、`工程师编号` AE(30)→AF(31)，硬编码列号导致办事处名读空、技师编码读成「已支付」，**数量对、归属全错且不报错**。
   → 现已改为按表头名解析（`gen_v2_lib._cols` + `WO_HEADERS`/`ZG_HEADERS`），`build_month.make_blocks()` 开头还会跑 `G.verify_columns(EXCEL)` 自检其余固定列号。看 `[OK] 列序自检通过` 才算过。
   - ⚠️ **「列序变动」≠「表导出错了」（2026-10-16 踩到，必看）**：那次用户导出的 **`CSM配件` 表整张变成了与 `CSM服务项目` 完全相同的 38 列**，
     表里**压根没有「配件名称/配件编码/数量」**。症状是 `verify_columns` 一口气报 **8 条**列序变动（办事处/配件名称/配件编码/销售分类/数量/上缴金额/录入完成时间…全错）。
     → 判据：**报表 3–8 条列错位 = 平台挪了列（可自动兜底）；连报 8 条以上且缺关键业务列 = 这张表导错了，必须让用户重新导出**，别硬套兜底列号跑出一份看着正常、实则全错的看板。
     → `CSM配件` 正常表头应为：`大区/服务中心/办事处/服务网点/网点类型/服务工程师/服务工程师编码/工单编号/工单类型/产品线/服务项目/保内外/配件名称(12)/配件编码(13)/…/销售分类(17)/数量(18)/上缴金额(27)/录入完成时间(36)`。
       若它的第 7 列是「录入完成时间」、第 13 列是「服务收费项目」→ 导错了。
     → 同理，若 `CSM配件` 与 `CSM服务项目` 表头逐字相同，也可能是导出时选错了报表模板。
2. **留一份上一版做对照**（关键，别忘）：
   ```bash
   cp dashboard_offline.html _prev.html        # build 会覆盖 dashboard_offline.html
   ```
   `.gitignore` 已忽略 `_prev*.html`。
3. `"$PY" build_month.py check <上一次的截止日>` —— 用新 Excel 按旧截止日重算，**与上一次 build 的 summary 数字对比**（工单数、止回阀、延保条数、清洗条数、大保养）。
   一致 ⇒ 历史数据无修订、解析器正常；不一致 ⇒ 逐项查是补录还是解析异常。
   ⚠️ check 的"与参考看板逐块对照"是拿 `dashboard_ref.html` 比的，而 REF 往往停在更早的截止日，**满屏 DIFFERS 属预期**，不能当异常。
   ✅ **更严格的做法：语义比对**（比 summary 数字可靠得多）——`_check_diff915.py` 把重算结果与上一版逐叶子比对，列表按**递归多重集**比较：
   - **先改对脚本参数**：`CUT = <上一次的截止日>`，`prev = _prev.html`（**同截止日的上一版 build**）。
     🐞 **别指向 `dashboard_ref.html`**（底板截止日更早）→ 会满屏「真实差异」，白排查一轮。
     🐞 **改脚本后必须 `grep` 确认落盘**：用 python 字符串 replace 改脚本时，若只 `print` 一段硬编码文字，会掩盖替换失败（2026-10-08 踩到）。
   - 只报「真实差异」，把「明细行序不同」单独归类（明细表按源表行序，跨 Excel 文件本就不可复现，不算错）。
   - 预期结论：`B`/`W`/`E`/`workorder`/`valueadded`/`valueparts`/`bigcare`/`EXTEND_TIME_DATA`/`CLEANING_DATA` **全部一致**；
     **只有 `D`/`T` 有真实差异且属正常**（二者取「报表全量」、不过滤日期，换快照必变，表现为网点数/工程师人数/合计金额增长）。
   - 注：脚本对「标量数组」（如 `剔除项`/`大保养项`/`code[]`）按索引比，会把**首现序**差异误报为真实差异——这两类属预期。
   - ⚠️ **`valueparts` 也可能有真实差异，且不一定是我方 bug**（2026-10-21 实测 −35 件）：平台会把「网点买断」单据
     **回溯撤回**（工单自购配件明细表旧行整批消失、零新增）。判据与排查手段见下文「🐞 网点买断单据被平台撤回」。
4. `"$PY" build_month.py build <新截止日>` → 输出 summary + 新 BUILD（写入 `dashboard_offline.html` 与 `version.json`）。
   **自洽校验**：工单总数应 = Excel 工单表有效日期总行数（扫日期列即可，10 月实测 5,214 = 5,214 ✅）。
   ⚠️ 若打印 `[warn] 未找到 var CURRENT_MONTH 声明`：页面里是 `let CURRENT_MONTH`（非 `var`），脚本正则已改为 `(?:var|let)`，改完就不会再出现；出现说明页面默认月份没跟着跨月，要修。
5. **历史月回归校验**（换底板/换 Excel 后必做，防历史月丢失）：用 `build_month.extract()` 逐月比 `MONTHLY_FOUR.months`+`year`、`EXTEND_TIME_DATA.months`、`CLEANING_DATA.months`，
   并核对 `MDATA_*` 声明数量（**10 月 45 个、11 月起 50 个；每多一个月 +5**）与基准块（8 月清洗 788 条、延保 139 条；9 月清洗 728、延保 192）。校验完删除 `_prev.html`。
   再抽样确认新功能还在：`.iPieScope`/`id="itemPie"`/`PIE_COLORS`/`renderItemPie`/`锁定规则 2026-09-21`/`toFixed(2) + '%'`。
6. 同步 4 份副本 + version.json（见上文），再走发布三步。
   ⚠️ **`dashboard_ref.html` 不要同步成新版**：它是下次 check 的对照底板，只在用户发来新版 dashboard 时才覆盖。

### 🐞 周数据 W 的办事处列错位（2026-10-08 修，影响金额归属）
`gen_month.gen_weekly()` 的清洁耗材取自 `CSM配件`，旧代码 `off = g(r, 0)` —— 但**该表 col0 是「大区」，办事处在 col2**
（表头序：大区 / 服务中心 / 办事处 / 服务网点）。后果：周数据多出一个**假办事处「西北大区部」**，吸走清洁耗材金额
（10 月两周约 ¥7,428，占两周总额约 12%），各真办事处达成被稀释，且全程不报错。
- 已按表头名定位「办事处」列（找不到回退 col2），并加「必须以『办事处』结尾，否则回退网点→办事处映射或『—』」兜底；
  同时现在允许用 CSM配件 的网点→办事处映射补 WMS 买断行归属（旧注释写"不能用"，正是因为 col0 取错，现在可靠了）。
- **验证**：`chk_w_off.py` 打印 W 的 office 异常键 + net 异常网点数，应全为 0，且**周合计金额与修复前一致**（只改归属、总额守恒）。
- ⚠️ 历史月（1–9 月）W 块是当时快照，**仍带旧的错误归属**，当月 Excel 无法重算；切到旧月仍可能看到「西北大区部」。
- 通用教训：**跨表取「办事处」必须先确认该表 col0 是啥**——同名含义在不同表里列号不同，`CSM服务项目` col1=办事处、`CSM配件` col2=办事处、`服务产品收入统计` col2=办事处，三者 col0 都是「大区」。
   ⚠️ **备份仓提交后别只看 `git commit` 的回显**：沙箱外视图可能滞后，显示 `nothing to commit` 而实际已提交成功。以 `git rev-parse --short HEAD` / `git log --oneline -1` 为准；若确实 `nothing to commit` 且 HEAD 未变，等几秒重跑 `git status` 再提交。

### 🐞 网点买断单据被平台撤回 → valueparts 波动（2026-10-21 确认，非 bug）
同一截止日重算时 `MONTHLY_FOUR.valueparts` 出现真实差异（当月 1414→1379，−35 件；昨日/近3日/近7日同步变小），
排查后确认是**平台侧回溯撤回「网点买断」单据**，不是解析问题。

**排查手段（强烈推荐，一步定性）：用 git 里的旧 Excel 逐行对比。**
```bash
git log --oneline -3 -- "2026年10月西北服务产品.xlsx"      # xlsx 已被 git 跟踪
git show <上次commit>:"2026年10月西北服务产品.xlsx" > _old.xlsx
```
然后按 `(技师, 技师编码, 物料编码, 配件名, 数量, 日期)` 建多重集，对两张表分别求差集：
- CSM配件：本次 **5,017 行逐行完全一致（0 差异）** ⇒ 解析器正常，直接排除列错位/口径问题
- 工单自购配件明细：旧 2,943 → 新 2,699，**旧有 244 行消失、零新增**（集中在 10-09~10-15）
- 消失行里命中增值产品的正好是「网点买断」35 行/35 件（漂洗剂13+洗碗盐10+洗碗粉10+清洁膏2），**与 −35 分毫不差**

结论：**源数据撤回，可放心继续 build**，但要如实告诉用户「增值配件里的网点买断部分会被平台回溯，历史快照无法重算，
数字可能比上一版略小」。⚠️ 用完记得删掉 `_old.xlsx`，别提交进仓库。

### 🐞 周数据网点→办事处映射缺兜底（2026-10-21 修）
`gen_month.gen_weekly()` 的 `m2o`（网点→办事处）原本只从 `CSM服务项目` 取。少数网点在那张表的办事处字段
为空或误填成「XX服务中心」（如「库尔勒泽玉…」填成「乌鲁木齐服务中心」），结果归属成「—」（本次 1 个网点 ¥99）。
- 已加 **CSM配件 预建映射**（该表办事处列可靠）：在 `m2o` 定义后、CSM服务项目 循环前先铺一层底；
  并把 CSM服务项目 的 `off` 改为「不以『办事处』结尾就用 `m2o` 兜底」。
- 验证：`chk_w_off.py` 的 net 异常网点数应为 0，且**周合计金额不变**（只改归属）。

### 🐞 跨月构建：新月份没有数据块 / 上个月被底板旧快照覆盖（2026-10-30 修，两个必看坑）
第一次跑到「库里还没有的月份」时会连踩两个坑，且**都不会报错**（第二个甚至数据看起来正常）。

**坑 1：新月份根本没有 `MDATA_*` 块**
`build()` 用 `replace_value()` 只会替换**已存在**的声明，`MDATA_2026_11_D` 不存在就直接 `SystemExit: 未找到声明`。
- 已改为「有则替换、无则插到首个 MDATA 声明之前」（必须在前，否则 `MONTH_DATA` 引用时未定义）。
- 还要往 `MONTH_DATA` 补一行映射——`ensure_month_data()` 会做。**下拉框选项是 `Object.keys(MONTH_DATA).sort()` 动态生成的**，补上就自动多出该月，不用改 HTML。

**坑 2：上个月被静默退回底板的旧快照（最危险）**
`merge_retained()` 旧逻辑只「补缺月」：月份已存在就保留 `dashboard_ref.html` 的。而 REF 往往停在更早截止日
（实测 REF 的 2026-10 只有 **4,629 单/10-04**，线上是 **16,475 单/10-20**）。
- 在**当月内部**更新时看不出来（当月稍后会被新数据覆盖），**一旦跨到新月份，上个月就整块退回旧快照**，页面无声丢数。
- 已改为 **prev（上一版 build）优先覆盖**，REF 只补 prev 没有的月份/块；`MONTHLY_FOUR.year` 也取 prev 的。
- **回归必检**：`verify_build.py` 会逐月比对 `_prev.html`，`MONTHLY_FOUR`/`EXTEND_TIME_DATA`/`CLEANING_DATA`/全部 MDATA 块
  **差异必须为 0**，有差异说明又退回去了。

**跨月后必做浏览器实测**（静态校验查不出运行时问题）：
```bash
"/c/Program Files (x86)/Microsoft/Edge/Application/msedge.exe" --headless=new --disable-gpu \
  --virtual-time-budget=8000 --dump-dom "file:///D:/WorkBuddy/dashboard_offline.html" > _dom.html 2> _dom.err
```
在 `_dom.html` 里核对：`id="monthBadge"` 是新月、`id="weekBadge"` 是「第1周/共N周」、KPI 里的工单数/达成率不为 0/NaN；
`_dom.err` 里出现 `Uncaught` 才算真报错（Chromium 的 QQBrowser/importer 之类 ERROR 可忽略）。**用完删除 `_dom.html`**（几 MB，别提交）。

**换月时另外要提醒用户的两件事**（源数据限制，不是算错）：
- 新月份那份 Excel 只含「本月服务月」的数据（如 11 月导出只有 10-28 起的 2 天），**上个月最后几天的补数不在里面**。
  若上次更新停在上月中旬（如 10 月只到 10-20），就要让用户**另外导一份上月（9.28–10.27）**，否则上月永远缺那几天。
- `MONTHLY_FOUR.year`（年度累计）是固定快照，**无法由单月 Excel 重算**；`LASTYEAR_BY_MONTH[新月份]` 会由 `seedLastYear()` 补空，
  同比列显示 0/空，需在「设置去年同期」里补去年同月数据。
