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
- **推送命令模板**（沙箱内、旁路，token 仅本次内存，勿写入 .git/config）：`git -c url."https://<PAT>@github.com/".insteadOf="https://github.com/" push github main`（前面加沙箱旁路参数）。
- **SSH 公钥备份**（未启用，因改用 PAT）：`~/.ssh/id_ed25519_github` 已生成，公钥 `ssh-ed25519 AAAAC3NzaC1lZDI1NTE5AAAAIKZzj9JlxinRnj3dOlYstcG4FqIUFCPIAiTNLNUqPZmk home-workbuddy-github`。
- **更新顺序**：改完先 `git push origin main`（Gitee），再按需喊 AI 用 PAT 推 GitHub。

## 网页看板（fangtai-dashboard）机制
- 地址：`https://mouren2580.github.io/fangtai-dashboard/`，GitHub Pages 静态站（源 `main`/根目录，legacy build，public）。
- **现状（2026-08-20 改造后）**：`index.html` 已改为**内嵌最新 xlsx 自动加载**版——基于原看板代码 + 内嵌 `dashboard_offline.html` 数据，打开即显示最新数据，且保留「同步Excel」按钮可手动覆盖。原版备份在仓库 `index.orig.html`。
- **如何更新该看板（AI 在沙箱执行）**：浅克隆（`--depth 1`）`mouren2580/fangtai-dashboard` → 用最新 `dashboard_offline.html` 覆盖 `index.html` → `git config user.email/user.name` 设身份 → commit → push（沙箱旁路 + PAT `insteadOf`）。推送后 Pages 自动重建（~1-3 分钟）。`dashboard_offline.html` 本身由工作区 xlsx 重嵌生成（见 2026-08-20 日志）。
- **截止日**：由 `dashboard_offline.html` 的"从 Excel 过账日期取最大值"逻辑自动推断，不再硬编码。
- **其它可达看板**：① CloudStudio 离线地址 `https://0717bc4b30824b8d8a407555473b321e.app.workbuddy.link`（内嵌 xlsx，需 AI 重部署更新）；② 家里本地 `dashboard/local.html` + `python -m http.server 8090`（实时读工作区 xlsx，未启用）。
- **历史背景（已过时，仅供参考）**：改造前该看板纯前端、仅按钮上传、数据硬编码 8/17 快照、AI 无法远程更新；2026-08-20 已通过仓库 push 改造解决。
