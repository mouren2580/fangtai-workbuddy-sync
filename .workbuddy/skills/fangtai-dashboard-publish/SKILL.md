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
PY=C:/Users/40973/.workbuddy/binaries/python/versions/3.13.12/python.exe
```
裸 `bash` 里 `dirname`/`cd` 会报错，`python`/`git` 也不在 PATH，必须显式加前缀。

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
   cd /d/WorkBuddy && C:/.../python.exe push_gh_pages.py       # 走 SSH，无需 PAT
   ```
   - 脚本自动：维护常驻克隆 → 覆盖 `index.html`+`version.json` → commit → push
   - **必须加 `dangerouslyDisableSandbox: true`**（要读 `~/.ssh`，沙箱会拦）
   - 复核：`python push_gh_pages.py --status` 或 `curl -s https://mouren2580.github.io/fangtai-dashboard/version.json`
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

# 相关：新 Excel 数据更新流程
`build_month.py` 顶部 `EXCEL` 指向新表 → `python build_month.py check <参考版截止日>` 验证解析器逐项吻合 → `python build_month.py build <更新日前一天>` → 然后走上文发布三步。
截止日铁律：**截止日 = 更新日前一天**（如 9-15 更新 → 2026-09-14）。服务周期 上月28日→当月27日（9 月 = 8.28–9.27，共 31 天）。
