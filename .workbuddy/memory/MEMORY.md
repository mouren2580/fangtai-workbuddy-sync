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

## GitHub 双备份远程（2026-08-19 新增）
- **GitHub 仓库**：`https://github.com/mouren2580/fangtai-workbuddy-sync.git`（当前 Public，建议改 Private），分支 `main`，远程名 `github`。
- **为何加**：与 Gitee 互为双保险。
- **本机（家里）推 GitHub 限制**：用户本机/浏览器连不通 GitHub（ERR_CONNECTION_TIMED_OUT）；只能由 AI 在沙箱用用户提供的 PAT 推送。
- **沙箱推送 GitHub 的关键坑（2026-08-20 实测确认）**：`git push` 到 GitHub **必须加 `dangerouslyDisableSandbox:true`**！沙箱默认拦截 git 的出站 443（报错 `Failed to connect to github.com port 443`），而 `curl github.com` 能通是因为它会被系统**自动旁路沙箱**——但 git 不会自动旁路。第一次失败、加旁路后 `8740542..54a4233 main->main` 成功。Gitee 的 SSH push 同理需旁路（另需 `StrictHostKeyChecking=accept-new` 跳过 known_hosts）。
- **推送命令模板（2026-09-01 修正）**：`git -c url."https://x-access-token:<PAT>@github.com/".insteadOf="https://github.com/" push github main`（前面加沙箱旁路参数）。⚠️ **旧模板 `https://<PAT>@github.com/` 的坑**：它只把 PAT 放在**用户名位**、密码位为空，依赖 Windows 凭据缓存；缓存一旦失效（或新克隆的仓库无缓存）就会弹窗要密码并失败。必须用 `x-access-token:<PAT>`（或 `<PAT>:<PAT>`）把 PAT 明确放在**密码位**，才能稳定免交互推送。克隆独立仓库（如 fangtai-dashboard）后也可用 `git remote set-url origin https://x-access-token:<PAT>@github.com/...` 直接写死。
- **SSH 公钥备份**（未启用，因改用 PAT）：`~/.ssh/id_ed25519_github` 已生成，公钥 `ssh-ed25519 AAAAC3NzaC1lZDI1NTE5AAAAIKZzj9JlxinRnj3dOlYstcG4FqIUFCPIAiTNLNUqPZmk home-workbuddy-github`。
- **更新顺序**：改完先 `git push origin main`（Gitee），再按需喊 AI 用 PAT 推 GitHub。
- **PAT 安全策略（2026-08-20 用户决定）**：原用 **classic PAT**（`ghp_...Lp1`，scope=`repo`）。用户曾要求"撤销并替换成受限 PAT"后说"先不动了"。
  - ⚠️ **2026-09-11 实测：该 classic PAT 已失效**——调 GitHub API 返回 `401 Bad credentials`（可能被撤销/提前过期）。AI 当前**无法推送 GitHub Pages**，需用户提供新的有效 PAT。
  - ⚠️ **AI 能力边界**：① 创建 fine-grained PAT 必须由用户在 GitHub 网页生成，AI 不能代建；② 撤销 classic PAT 需用户在网页手动撤销。
  - **推荐替换**：Fine-grained PAT，Resource owner=`mouren2580`，Repository access=仅 `fangtai-dashboard`（Pages 站）+ 可选 `fangtai-workbuddy-sync`，Contents=Read and write，合理过期。
  - **切换顺序**：用户提供新 PAT → AI 改 `_gh_push.py`/insteadOf 里的 token 并验证 push → 用户再网页撤销旧 PAT。

## 网页看板（fangtai-dashboard）机制
- **多月份架构（2026-09 起）**：`dashboard_offline.html` 用 `MONTH_DATA` 对象按月份键（`2026-01`…`2026-09`）存各月数据，每块 = `MDATA_YYYY_MM_{D,T,B,W,E}`（`D`主数据/`T`工程师/`B`大保养/`W`周数据/`E`延保）；`BASE_MONTH="2026-08"` 锚点不被重建覆盖；`CURRENT_MONTH` 默认 `2026-09`。页面 `fetch('version.json')` 做缓存穿透（BUILD 标记比对）。
- **数据生成工具**：`build_dashboard.py`（原 `build()` 主数据+大保养生成器）**已丢失**（git/磁盘均无）。工作区现用 **`gen_month.py`**（自包含）重建某月五块并注入 HTML：
  - 复用 `sync-kit/sync.py` 的 `build_tech`/`gen_extend`/`gen_weekly`（T/E/W）。
  - `build_main`（D）：`服务产品收入统计` 按 (办事处,服务中心,服务网点) 聚合 15 品项+合计+来源，并补 `time` 字段（服务月进度）。
  - `build_bigcare`（B）：`清洗保养总单`=**销售分类(col16)=清洗保养**；`剔除项`/`大保养项`=**服务收费项目(col13)**，与清洗保养总单**可重叠**（独立 if）；`分母=总单-剔除`。
  - `gen_weekly` 周起点**向后**取服务月所在周周一（如 9 月从 8/24 起），脚手架整月 5 周，数据按截止日过滤。
  - 顶部 `EXCEL` 路径需改成新 Excel；`CUTOFF`=更新日前一天（铁律）。运行：`python gen_month.py`（需 openpyxl，已装于 venv）。
- **如何更新看板（AI 流程）**：① 改 `gen_month.py` 的 `EXCEL` 路径 → 跑它重建并注入 `dashboard_offline.html` + 写 `version.json`（BUILD 号）+ 改 `var BUILD`；② 复制 `dashboard_offline.html`+`version.json` 到 `deploy_cs/` → `workbuddy_cloudstudio_deploy`（同 URL）；③ `git push origin main`（Gitee，沙箱旁路）；④ GitHub Pages：`_gh_push.py`（需有效 PAT，读 `GITHUB_PAT` 环境变量）。
- **截止日口径（铁律）**：`截止日 = 更新日前一天`。如 9-11 更新 → 锁 **2026-09-10**；服务周期 = 上月28日→当月27日（9月=8.28–9.27，31天）。
- **GitHub Pages 地址**：`https://mouren2580.github.io/fangtai-dashboard/`（源 `fangtai-dashboard` 仓 `main`/根目录）。**2026-09-11 状态：仍停留在用户 9-11 BUILD（commit 41896a0e08，KPI 与本次一致），因 classic PAT 失效 AI 未能推送最新版。**
- **CloudStudio 地址**：`https://0717bc4b30824b8d8a407555473b321e.app.workbuddy.link`（2026-09-11 已部署最新版 ✅）。
- **历史背景（已过时，仅供参考）**：早期为内嵌 xlsx / 按钮上传机制，2026-08-20 改造、2026-09-11 用户改"在线版"硬编码 JS。
