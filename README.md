# WorkBuddy 工作区 · 双机同步仓库

本仓库用于在家里电脑与单位电脑之间同步方太西北大区服务部的月报工作，
包括：Excel 数据、看板 HTML、工作区记忆（`.workbuddy/memory`）。

## 同步方式
- 远端：GitHub 私有仓库
- 家里 / 单位两台电脑均 `clone` 同一仓库
- 改完执行 `git pull` → `git add -A` → `git commit` → `git push`
- 另一台电脑 `git pull` 即可拿到更新

## 注意事项
- `.codebuddy/`（机器本地设置）已被 .gitignore 排除，不进入同步
- 大文件（如 xlsx）直接纳入版本管理即可，无需 Git LFS
- 提交前请确认 git 的用户名/邮箱已设置（见 `git config user.name/email`）
