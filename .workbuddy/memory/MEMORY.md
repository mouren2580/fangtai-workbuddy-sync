# 项目长期记忆（方太西北服务产品月报 / WorkBuddy 工作区）

## 跨电脑任务同步约定（家里 ↔ 单位）
- **仓库**：Gitee 私有 `https://gitee.com/xbftsjkb/fangtai-workbuddy-sync.git`，分支 `main`。
- **理由**：GitHub 国内访问超时，统一用 Gitee。
- **同步范围**：`D:\WorkBuddy` 全工作区（Excel 月报、看板、`.workbuddy/` 记忆、配置）。`.codebuddy/` 为机器本地设置不纳入。
- **本机（家里）认证**：SSH 密钥 `~/.ssh/id_ed25519_gitee`，Gitee 公钥标题 `home-workbuddy`。
- **单位电脑接入**：clone → 生成自有 SSH 密钥 → 加入同一 Gitee 账号 → 双向 pull/push。
- **日常流程**：改后 `git add -A && git commit -m "..." && git push`；另一台先 `git pull`。

## 网页看板（fangtai-dashboard）机制
- 地址：`https://mouren2580.github.io/fangtai-dashboard/?newPanel=true`，GitHub Pages 静态站，纯前端 SPA。
- **数据来源**：仅通过浏览器内 🔄 同步Excel 按钮 → FileReader 读取用户选择的 `2026年8月西北服务产品.xlsx`（用 SheetJS/XLSX 解析）。**无任何远程 fetch / 后端接口**，无法从外部推送数据。
- 目标/设置/截止日等存在 `localStorage`，换电脑或清缓存需重设（或同步Excel自动取前一天）。
- 结论：看板刷新必须人工在浏览器里点「同步Excel」选文件。AI 无法 remotely 更新它。数据源就是工作区里的 xlsx（更新后推 Gitee 即可两边共用）。
- 若想「真正自动同步」，需改造看板源码让其从某个可公开访问的 URL（如 Gitee raw）拉 xlsx；但这要改 GitHub 上的看板仓库（本机访问 GitHub 受限），属于后续可选项。
