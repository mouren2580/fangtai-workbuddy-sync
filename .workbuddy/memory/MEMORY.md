# 项目长期记忆（方太西北服务产品月报 / WorkBuddy 工作区）

## 跨电脑任务同步约定（家里 ↔ 单位）
- **仓库**：Gitee 私有 `https://gitee.com/xbftsjkb/fangtai-workbuddy-sync.git`，分支 `main`。
- **理由**：GitHub 国内访问超时，统一用 Gitee。
- **同步范围**：`D:\WorkBuddy` 全工作区（Excel 月报、看板、`.workbuddy/` 记忆、配置）。`.codebuddy/` 为机器本地设置不纳入。
- **本机（家里）认证**：SSH 密钥 `~/.ssh/id_ed25519_gitee`，Gitee 公钥标题 `home-workbuddy`。
- **单位电脑接入**：clone → 生成自有 SSH 密钥 → 加入同一 Gitee 账号 → 双向 pull/push。
- **日常流程**：改后 `git add -A && git commit -m "..." && git push`；另一台先 `git pull`。
- **本机推送的权限坑（2026-08-19 确认）**：`git push` 走 SSH 需读 `~/.ssh`，被**沙箱拦截**，必须加 `dangerouslyDisableSandbox:true` 且**用户需放行 bypass 弹窗**（用户曾连续拒绝导致推不出去）。另：本机 Git 凭据助手是 Sogou 的 `helper-selector`，在沙箱里写 `AppData\LocalLow\SogouPY` 会失败、且无 Gitee 凭据缓存，故 HTTPS 推送也不可用；`manager-core` 未安装。
- **可靠推送方案**：(1) 用户放行 bypass 后 `git push origin main`；(2) 用户提供 Gitee PAT，用 `git push https://<user>:<token>@gitee.com/xbftsjkb/fangtai-workbuddy-sync.git main`（绕开 SSH 与 Sogou 助手）；(3) 用户手动在终端 `git push origin main`。
- ⚠️ **Gitee 单文件 100MB 硬限制（2026-09-15 实测踩坑）**：`git add -A` 曾误提交 `2026-09-12-20-12-51/models/faster-whisper-small/model.bin`（**462MB**，whisper 语音模型），Gitee 直接 `pre-receive hook declined / exceeds quota 100MB` 拒收整个 push。
  - **教训**：本工作区做「全量 add」前务必确认没有大体积中间产物（模型、音视频、缓存）。
  - **已加 `.gitignore`**：`2026-09-12-20-12-51/`、`models/`、`*.bin`、`_bak_*`、`_probe_*.py`、`_check_*.py`、`_diff_*.py`。
  - **补救套路**：若大文件只出现在**最新一次未成功推送**的提交里 → `git rm -r --cached <path>` + `git commit --amend`，历史即不再含该 blob（不必 filter-branch）。若已推送过，才需要 `git filter-repo` 重写历史。

## GitHub 双备份远程（2026-08-19 新增）
- **GitHub 仓库**：`https://github.com/mouren2580/fangtai-workbuddy-sync.git`（当前 Public，建议改 Private），分支 `main`，远程名 `github`。
- **为何加**：与 Gitee 互为双保险。
- **本机（家里）推 GitHub**：用户本机/浏览器直连 GitHub 会超时，**只能由 AI 在（非沙箱）环境用 SSH 推**；走 SSH 后不需要 PAT。
- **推送到 GitHub 必须 `dangerouslyDisableSandbox: true`**（SSH 要读 `~/.ssh`，沙箱会拦；HTTPS 443 也曾被拦）。
- **✅ 推送方式＝SSH（2026-09-15 起，已弃用 PAT 模板）**：`github` 远程已改为 `git@github.com:mouren2580/fangtai-workbuddy-sync.git`，直接 **`git push github main`** 即可。对应的 `~/.ssh/config` 已配 `Host github.com → IdentityFile ~/.ssh/id_ed25519_github`；首次需 `StrictHostKeyChecking=accept-new`。
  - 历史遗留（仅供理解，勿再用）：旧方案是 `git -c url."https://x-access-token:<PAT>@github.com/".insteadOf="https://github.com/" push github main`，依赖 PAT；PAT 已失效。
  - **2026-09-15 状态：GitHub 备份仓已补齐至 `cb2d99e`（与 Gitee `origin/main` 同步）✅**，不再落后。
  - ⚠️ 该仓**仍是 Public**，内含方太业务数据/Excel/看板 —— 建议用户在 GitHub 网页改成 Private（AI 无权限代改）。
- **SSH 公钥（2026-09-15 确认已注册并可用）**：`~/.ssh/id_ed25519_github`，公钥 `ssh-ed25519 AAAAC3NzaC1lZDI1NTE5AAAAIKZzj9JlxinRnj3dOlYstcG4FqIUFCPIAiTNLNUqPZmk home-workbuddy-github`。实测 `ssh -T git@github.com` → `Hi mouren2580! You've successfully authenticated`。**同一把钥匙同时用于 `fangtai-dashboard`（Pages 仓）与 `fangtai-workbuddy-sync`（备份仓）**。
- **更新顺序**：改完先 `git push origin main`（Gitee 备份仓），再 `python push_gh_pages.py` 推 GitHub Pages（SSH，见下文）。
- **✅ PAT 已弃用，改走 SSH（2026-09-15 结论，别再折腾 PAT 了）**：classic PAT `ghp_...Lp1` 在 9-11、9-15 两度实测 `401 Bad credentials`（已失效/被撤销）。**但不需要新 PAT**——`~/.ssh/id_ed25519_github` 早已注册到 `mouren2580` 账号，`ssh -T git@github.com` 能通过认证，**直接 SSH 推 GitHub Pages 即可**（详见「GitHub Pages 改走 SSH 推送」条目）。
  - 本机**没有任何 GitHub 凭据**：无环境变量、无 `.git-credentials`、Windows 凭据管理器也无 GitHub 条目、git credential.helper = `<no helper>` → 一旦 SSH 也不可用，就只剩"用户自己推"一条路。
  - 若哪天还想回到 PAT：用户须**在 GitHub 网页自行创建**（AI 不能代建），建议 fine-grained、Resource owner=`mouren2580`、仅授权 `fangtai-dashboard`、Contents=Read and write、设合理过期；填给 `_gh_push.py` 或写进 `~/.workbuddy/.github_pat`。

## 网页看板（fangtai-dashboard）机制
- **多月份架构（2026-09 起）**：`dashboard_offline.html` 用 `MONTH_DATA` 对象按月份键（`2026-01`…`2026-09`）存各月数据，每块 = `MDATA_YYYY_MM_{D,T,B,W,E}`（`D`主数据/`T`工程师/`B`大保养/`W`周数据/`E`延保）；`BASE_MONTH="2026-08"` 锚点不被重建覆盖；`CURRENT_MONTH` 默认 `2026-09`。页面 `fetch('version.json')` 做缓存穿透（BUILD 标记比对）。
- **数据生成工具（2026-09-15 换代，务必用新工具）**：`build_dashboard.py` 与早期 `gen_month.py` 单跑法**已过时**。改为**「以用户最新参考看板为底板 + 只替换当月数据块 + 保留历史月」**：
  - **入口 `build_month.py`**：
    - `python build_month.py check <截止日>` → 用新 Excel 按该截止日重算，与底板（`dashboard_ref.html`）**逐块对照**，验证解析器；`build` 模式才写文件。
    - `python build_month.py build <截止日>` → 重建并注入 `dashboard_offline.html` + 写 `version.json` + 改 `var BUILD`。**只替换 2026-09 分支**（1–8 月与 `MONTHLY_FOUR.year` 原样保留）。
    - **`REF` 指向工作区 `dashboard_ref.html`**（= 用户最新那份参考看板；用户每次发来新版就覆盖它）。
    - **`merge_retained()`（2026-09-15 新增，重要）**：换底板时自动把**旧版有、新底板没有的月份/块补回来**（`MONTHLY_FOUR`/`EXTEND_TIME_DATA`/`CLEANING_DATA` 缺月、以及顶层 `MDATA_*` 缺块）。**背景**：曾因底板换成一份缺 `CLEANING_DATA.months["2026-08"]` 的参考版，导致 8 月 788 条清洗明细被丢掉。
  - **解析库 `gen_v2_lib.py`**：`build_workorder / build_valueadded / build_valueparts / build_bigcare4 / build_extend_time / build_cleaning`。
  - `gen_month.py` 保留，仅供 D（服务产品收入统计）/ T（工程师）/ W（周数据）复用。
  - **✅ 验证结论（最重要口径）**：**除 D/T 外所有模块都按「统计截止日 = 更新日 − 1」过滤**（不是「整服务月不过滤」）。
    ⚠️ **D/T 例外**：`服务产品收入统计` 取「**报表全量**」（不过滤日期，meta 里写明 `截止: 报表全量`），所以换 Excel 快照就会变，`check` 时与参考版对不上属**正常**。
  - **✅ 2026-09-15 二次校验**：以用户 9-14 版（BUILD `20260914-0913`，截止 9-13）为基准反算 → `workorder`/`valueadded`/`valueparts`(含技师归属)/`MONTHLY_FOUR.bigcare`/`CLEANING_DATA`/`MDATA_*_W` **完全一致**。
- **模块清单（参考版逻辑）**：`MDATA_YYYY_MM_{D,T,B,W,E}` + `MONTHLY_FOUR.months[月].{workorder,valueadded,valueparts,bigcare}` + `EXTEND_TIME_DATA.months[月]`（延保明细）+ `CLEANING_DATA.months[月]`（清洗保养明细）。`MONTHLY_FOUR.year` = 1–8 月累计，**当月 Excel 只含本月，无法重算，保持原样**。
- **`MDATA_2026_08_*` 不存在是正常的**：8 月用基准块 `DASHBOARD_DATA/TECH_DATA/BIG_CARE_DATA/WEEKLY_DATA/EXTEND_WARRANTY_DATA`。
- **易踩的口径坑**：① 大保养 `net`/`eng` 同值排序用**首次出现顺序**（非名称），合计行固定文案 `[合计]全区`/`[合计]全部网点`/`[合计]全部工程师`；② meta 标签不统一：`workorder.meta.截止` 与 `MONTHLY_FOUR.bigcare.meta.截止日` = **服务月末（9-27）**，而 `MDATA_*_B.meta.截止日` = **截止日**且无「周期」字段；③ 延保 `EXTEND_TIME_DATA` 记录 `c` 恒为空串、E 块 `techs[].id` = **姓名**，E 必须由已过滤的延保明细聚合（`gen_extend` 不过滤日期，不能用）；④ `valueadded.products[]` **有** `verify`，`valueparts.products[]` **没有**；⑤ 明细表按源表行序，跨 Excel 文件排列不可复现，**只校验「多重集一致」**即可；⑥ **`剔除项`/`大保养项` 两处顺序不同**：`MDATA_*_B.meta` 用**常量序** `["保养超范围收费","灶具保养"]`/`["油烟机大保养升级包","油烟机大保养"]`，`MONTHLY_FOUR.bigcare.meta` 用**源表首现序**（`build_bigcare4` 额外返回 `enc` 清单，只给 MF 分支套用）；⑦ `valueparts` 的配件名 `pname` 也按**源表首现序**；⑧ E 块逐年金额**不四舍五入**（保留浮点尾差，如 `42616.90000000001`），只有 `tamt` 舍入。
- 🐞 **已修复的重大口径 bug：`工单自购配件明细` 没有办事处列！** 该表 col15 = **「大区部」**（不是办事处）、col12 = 服务中心、col5 = 服务网点。
  - 症状：只在自购配件表出现的耗材（洗碗粉/洗碗盐/漂洗剂/厨小护/蒸烤清洁剂/锅支架清洁膏）技师归属全被标成「西北大区部」——**数量对、归属错**。
  - 正确做法：建 **技师编号→办事处** 反查表（源：`CSM配件` col2=办事处 + col5=工程师 + col6=工程师编码；`工单表` col30=工程师编码 + col11=办事处名称 亦可，两者 0 一对多），自购行取 `tech_office[编码] → tech_office[姓名] → net_office[网点]`。代码里 `ZG_OFFICE` 已改名 `ZG_ALGO` 防误用。
  - 同类坑：`CSM配件` col0 也是「大区」不是办事处（周数据归集曾踩，见 2026-09-15 日志）。
- **如何更新看板（AI 流程）**：① 把新 Excel 路径写进 `build_month.py` 的 `EXCEL` 常量（`gen_v2_lib.py` 无路径常量，由入口传入）→ 先 `check <参考版截止日>` 确认无异常 → 再 `build <更新日−1>`；② 复制 `dashboard_offline.html`+`version.json` 到 `deploy_cs/` → 用 `workbuddy_sites_deploy`（旧名 `workbuddy_cloudstudio_deploy`）部署；③ `git push origin main`（Gitee，沙箱旁路）；④ 另存一份 `dashboard.html` 供单位离线使用。
- **⚠️ 线上部署需用户当轮确认**：`workbuddy_sites_deploy` 对方「已有线上链接的目录」会拒绝静默覆盖，返回 `sites_deploy_needs_confirmation`，会要求先问用户一句「改动已完成，需要我同步更新到线上分享链接吗？（线上现有内容会被覆盖）」——**这是工具强约束，即使工作区既有「自动部署」约定也必须先问**。
- **截止日口径（铁律）**：`截止日 = 更新日前一天`。如 9-11 更新 → 锁 **2026-09-10**；9-15 更新 → 锁 **2026-09-14**；服务周期 = 上月28日→当月27日（9月=8.28–9.27，31天，时间进度=已过天数÷31）。
- **最新状态（2026-09-15 09:20）**：工作区 BUILD `20260915-0904`，截止 `2026-09-14`。**三端已全部同步 ✅**。D 合计 **¥306,777.34**（89 网点）/ 工单 **13,136**（工程师 212 人、工单消耗 ¥261,641.40）/ 防火阀占比 **19.02%**（止回阀 833 ÷ 烟机安装 4,379）/ 大保养 **2.67%**（12÷449，总单 459）/ 延保 **113 单 ¥42,974.90** / 清洗保养明细 448 条 / 增值产品 1,271 / 增值配件当月 1,774。三模块「止回阀」互校一致 = 833。
- **三端版本现状（2026-09-15 09:20 实测，全部一致 ✅）**：GitHub Pages `20260915-0904`（Pages 本体与本地**逐块一致 10/10**）｜CloudStudio `20260915-0904`｜工作区 `20260915-0904`，截止均 `2026-09-14`。
- **🎉 GitHub Pages 改走 SSH 推送（2026-09-15 破解，彻底摆脱 PAT）**：地址 `https://mouren2580.github.io/fangtai-dashboard/`（源 `fangtai-dashboard` 仓 `main` / 根目录，用户**主要分享给同事看的就是这条**）。
  - classic PAT `ghp_...Lp1` 已失效（API `401`），但 **`~/.ssh/id_ed25519_github` 早已注册到 mouren2580 账号**——`ssh -T git@github.com` 返回 `Hi mouren2580! You've successfully authenticated`。所以**用 SSH 直接推**，不再需要 PAT。
  - **主用脚本 `push_gh_pages.py`**（工作区根）：维护常驻克隆 `~/.workbuddy/tmp/fangtai-dashboard`，`fetch + reset --hard` → 用 `deploy_cs/` 里的 `index.html`+`version.json` 覆盖 → commit → `git push origin main`；`--status` 只比对本地上线版本、不推送。
  - ⚠️ **必须非沙箱运行**（要读 `~/.ssh`）；`~/.ssh/config` 已配 `Host github.com → IdentityFile ~/.ssh/id_ed25519_github`。
  - ⚠️ 该仓除 `index.html` 外还有 `drainage/`、`sync-kit/`、`sync.sh`、`index.orig.html`、`.nojekyll`——**只覆盖 index.html 与 version.json，其它一律不碰**。
  - GitHub Pages 重建约 **1–3 分钟**，用 `python push_gh_pages.py --status` 复核（线上 `version.json` 的 `v`/`cut`）。
  - 备用：API 版脚本 `_gh_push.py`（需有效 PAT，目前不可用，已支持 `--check` 自检）。
- **CloudStudio / 线上分享链接**：`https://0717bc4b30824b8d8a407555473b321e.app.workbuddy.link`（目录 `D:\WorkBuddy\deploy_cs`）。**2026-09-16 13:00 已更新至 BUILD `20260916-1259`（截止 9-15）✅**，旧 `.link` 链接实测已生效。
- **⚠️ 部署工具的「预留域名未绑定」报错＝假失败（2026-09-15 实测）**：`workbuddy_sites_deploy` 可能返回
  `应用预留域名 fangtai-dashboard.app.workbuddy.host 未绑定到本次发布环境，本次发布已停止。`
  **但内容其实已经上传并 release 成功**（`artifactRelease` 返回 201，`releaseUrl` 为 sandbox 的 `.host` 域名；
  报错只发生在最后的预留域名绑定校验——平台正从 `.link` 迁到 `.host`）。
  → **不要当成失败重试**，直接 `curl -s "<链接>/version.json"` 看 `v`/`cut` 是否已变即可确认。
  → 详见 `~/.workbuddy/logs/sites/sites-deploy-YYYYMMDD.log`（有完整分阶段日志）。
  → 另：`updateExistingApp:true` 在本工作区会报 `no existing app to update`（旧部署记录与新工具不兼容），**去掉该参数**即可正常走完流程。
- **📐 显示规范：所有百分比一律两位小数（用户 2026-09-15 要求）**：全局改 `const pct = x => (x*100).toFixed(2) + '%'` 即可覆盖绝大多数显示位；另需改同比/合计/品项分布/大保养/自购耗材/年度目标面板/技师面板的 `.toFixed(1)`。**不要改** CSS 宽度、`hsl()` 颜色、条形图宽度用的 `Math.round(x*100)`。KPI 卡环形图（`.kpi .ring` 56px）内文字要配 `font-size="10.5"` 才不撞环体。自检：`grep -n "toFixed(1)" dashboard_offline.html` 只应剩条形图宽度一处。
- **⚠️ 同一份看板有 4 个副本，改样式/JS 必须一起改**：`dashboard_offline.html`（源）、`dashboard.html`（离线交付）、`deploy_cs/index.html`（CloudStudio）——三者 md5 应一致；另 `dashboard_ref.html` 是 `build_month.py` 的**底板**，不写进去下次 `build_month.py build` 重建就丢改动。
- **🆕 发布 SOP 已固化为 skill**：`D:\WorkBuddy\.workbuddy\skills\fangtai-dashboard-publish\SKILL.md`（改看板 → 改 4 副本 + BUILD/version.json → push_gh_pages.py → 部署 CloudStudio → 双备份仓推送 → curl 验收）。改看板前先读它，可省掉 PATH/沙箱/假失败等一堆试错。
- **最新状态（2026-09-16 13:00）**：BUILD `20260916-1259`，截止 `2026-09-15`。四端一致（Pages `e48903f` / CloudStudio / 工作区 / 备份仓 `23730a0`）。核心数字：工单 **13,886**（+750）/ 止回阀 **884**（烟机安装 4,663、占比 18.96%）/ 大保养 **13÷465＝2.80%**（总单 475）/ 延保明细 **115 条 ¥43,661.90** / 清洗保养 **464 条** / 增值产品 1,349 / 增值配件当月 1,868 / D 报表全量 **¥321,560.84**（91 网点，不过滤日期）/ T 工程师 214 人 ¥274,581.90。历史月 1–8 月 + 8 月清洗 788 条、延保 139 条均完整保留 ✅
- **🔧 环境要点（2026-09-16 确认）**：解析 Excel 的 `build_month.py` **必须用 venv python** `C:/Users/40973/.workbuddy/binaries/python/envs/default/Scripts/python.exe`（系统级 3.13.12 **无 openpyxl**）；`push_gh_pages.py` **不能后台运行**（SSH 会挂起，实测卡 25 分钟），必须前台+沙箱旁路，卡住就手动推 `git -C "C:/Users/.../tmp/fangtai-dashboard"`（git 不认 `/c/` 路径形式）。
- **`?newPanel=true` 之类查询参数无效**：页面里没有任何 `newPanel` 代码，纯缓存穿透/来源标记，可忽略。判断"线上是哪一版"就看 `version.json` 的 `v`/`cut` 与页面 `var BUILD`。
- **历史背景（已过时，仅供参考）**：早期为内嵌 xlsx / 按钮上传机制，2026-08-20 改造、2026-09-11 用户改"在线版"硬编码 JS。
