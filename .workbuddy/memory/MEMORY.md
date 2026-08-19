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
- **本机（家里）推 GitHub 限制**：家用网络连不通 GitHub，本机也无 GitHub 凭据；只能由 AI 在沙箱（能通 GitHub）用用户提供的 PAT 推送，或用户开代理后自推。
- **推送命令模板**（沙箱内，token 仅本次内存，勿写入 .git/config）：`git -c url."https://<PAT>@github.com/".insteadOf="https://github.com/" push -u github main`。
- **SSH 公钥备份**（未启用，因改用 PAT）：`~/.ssh/id_ed25519_github` 已生成，公钥 `ssh-ed25519 AAAAC3NzaC1lZDI1NTE5AAAAIKZzj9JlxinRnj3dOlYstcG4FqIUFCPIAiTNLNUqPZmk home-workbuddy-github`。
- **更新顺序**：改完先 `git push origin main`（Gitee），再按需喊 AI 用 PAT 推 GitHub。

## 网页看板（fangtai-dashboard）机制
- 地址：`https://mouren2580.github.io/fangtai-dashboard/?newPanel=true`，GitHub Pages 静态站，纯前端 SPA。
- **数据来源**：仅通过浏览器内 🔄 同步Excel 按钮 → FileReader 读取用户选择的 `2026年8月西北服务产品.xlsx`（用 SheetJS/XLSX 解析）。**无任何远程 fetch / 后端接口**，无法从外部推送数据。
- 目标/设置/截止日等存在 `localStorage`，换电脑或清缓存需重设（或同步Excel自动取前一天）。
- 结论：看板刷新必须人工在浏览器里点「同步Excel」选文件。AI 无法 remotely 更新它。数据源就是工作区里的 xlsx（更新后推 Gitee 即可两边共用）。
- 若想「真正自动同步」，需改造看板源码让其从某个可公开访问的 URL（如 Gitee raw）拉 xlsx；但这要改 GitHub 上的看板仓库（本机访问 GitHub 受限），属于后续可选项。

### 看板的「数据更新日期 / 截止日期」机制（2026-08-17 逆向确认）
- 数据里带 `"统计截止日":"YYYY-MM-DD"` 和 `"_stamp":"...Z"`（导入时间戳）字段；内嵌默认数据是每月一条快照（截止日 01-27…08-15 等）。
- `asOfBtn` 点击弹窗原文：**「截止日期已固定为数据自带『统计截止日』，不随打开时间自动变化。如需更新，请通知我重新导入数据（截止日=通知更新日−1）」**。
- 即：这两个日期是**数据派生**的，不是自由输入框；更新办法是重新导入（同步）Excel。AI 无法在浏览器里改这两个值。
- 当前内嵌默认截止日仍是 2026-08-15；用户同步 8/16 版 Excel 后应刷新为 8/16（或 导入日−1）。若同步后仍是旧值，说明看板未从用户 Excel 取该字段 → 需看板作者 mouren2580 重导，或本地自建可设日期的版本。
