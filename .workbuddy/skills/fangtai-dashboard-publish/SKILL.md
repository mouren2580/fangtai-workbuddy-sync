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
**截止日铁律**：**截止日 = 更新日前一天**（如 9-16 更新 → 2026-09-15）。服务周期 上月28日→当月27日（9 月 = 8.28–9.27，共 31 天）。

1. **先确认真实截止日**：别直接信"更新日−1"，扫一下 Excel 各表日期列最大值（工单 col41 / CSM服务项目 col7 / CSM配件 col36 / 工单自购配件明细 col10，0 基索引），两者应吻合。
   ⚠️ **列序坑（2026-09-18 踩过）**：平台导出**会调整列顺序**。上次「工单」表把 `办事处名称` L(11)→K(10)、`工程师编号` AE(30)→AF(31)，硬编码列号导致办事处名读空、技师编码读成「已支付」，**数量对、归属全错且不报错**。
   → 现已改为按表头名解析（`gen_v2_lib._cols` + `WO_HEADERS`/`ZG_HEADERS`），`build_month.make_blocks()` 开头还会跑 `G.verify_columns(EXCEL)` 自检其余固定列号。
   → **仍要看 summary 的数字是否合理**（办事处数应为 6、网点 ~98、工程师 ~238），异常就先去核对表头。
2. **留一份上一版做对照**（关键，别忘）：
   ```bash
   cp dashboard_offline.html _prev.html        # build 会覆盖 dashboard_offline.html
   ```
   `.gitignore` 已忽略 `_prev*.html`。
3. `"$PY" build_month.py check <上一次的截止日>` —— 用新 Excel 按旧截止日重算，**与上一次 build 的 summary 数字对比**（工单数、止回阀、延保条数、清洗条数、大保养）。
   一致 ⇒ 历史数据无修订、解析器正常；不一致 ⇒ 逐项查是补录还是解析异常。
   ⚠️ check 的"与参考看板逐块对照"是拿 `dashboard_ref.html` 比的，而 REF 往往停在更早的截止日，**满屏 DIFFERS 属预期**，不能当异常。
   ✅ **更严格的做法：语义比对**（比 summary 数字可靠得多）——把重算结果与 `_prev.html` 逐叶子比对，列表按**递归多重集**比较：
   - 只报「真实差异」，把「明细行序不同」单独归类（明细表按源表行序，跨 Excel 文件本就不可复现，不算错）；
   - 预期结论：`workorder`/`valueadded`/`W` 完全一致；`valueparts`/`bigcare`/`B`/`E`/`EXTEND_TIME_DATA`/`CLEANING_DATA` 仅顺序不同；**只有 `D`/`T` 有真实差异且属正常**（二者取「报表全量」、不过滤日期，换快照必变）。
4. `"$PY" build_month.py build <新截止日>` → 输出 summary + 新 BUILD（写入 `dashboard_offline.html` 与 `version.json`）。
   **自洽校验**：工单总数增量应 ≈ 新增天数的工单行数之和（如 9-15→9-17 的 682+755=1437）。
5. **历史月回归校验**（换底板/换 Excel 后必做，防历史月丢失）：用 `build_month.extract()` 逐月比 `MONTHLY_FOUR.months`(1–8月)+`year`、`EXTEND_TIME_DATA.months`、`CLEANING_DATA.months`，
   并核对 `MDATA_*` 声明数量（当前 40 个）与 8 月基准块（8 月清洗 788 条、延保 139 条）。校验完删除 `_prev.html`。
6. 同步 4 份副本 + version.json（见上文），再走发布三步。
