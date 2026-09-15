# -*- coding: utf-8 -*-
"""把工作区最新看板同步到 GitHub Pages（mouren2580/fangtai-dashboard）。

**不需要 PAT** —— 走 SSH（~/.ssh/id_ed25519_github，已注册到 mouren2580 账号）。
本机 2026-09-15 实测 `ssh -T git@github.com` => "Hi mouren2580! You've successfully
authenticated"，故彻底绕过已失效的 classic PAT。

流程：维护一份常驻克隆 → 覆盖 index.html / version.json → commit → push。
只动这两个文件，**不碰** 仓库里的 drainage/、sync-kit/、sync.sh、index.orig.html。

用法：
  python push_gh_pages.py            # 同步并推送
  python push_gh_pages.py --status   # 只看远端当前版本，不推送

⚠️ 沙箱限制：本脚本会读 ~/.ssh，需以「非沙箱」方式运行（dangerouslyDisableSandbox）。
"""
import os, sys, json, shutil, subprocess, urllib.request

HERE = os.path.dirname(os.path.abspath(__file__))
DEPLOY = os.path.join(HERE, "deploy_cs")
CLONE = os.path.join(os.path.expanduser("~"), ".workbuddy", "tmp", "fangtai-dashboard")
REPO = "git@github.com:mouren2580/fangtai-dashboard.git"
# 工作区 deploy_cs/ 里的这两个文件会被原样推到仓库根目录
SRC = ["index.html", "version.json"]
AUTHOR_NAME, AUTHOR_EMAIL = "mouren2580", "409737410@qq.com"
PAGES_URL = "https://mouren2580.github.io/fangtai-dashboard/"


def run(cmd, check=True):
    r = subprocess.run(cmd, capture_output=True, text=True,
                       encoding="utf-8", errors="replace")
    if check and r.returncode != 0:
        print("  命令失败:", " ".join(cmd))
        print("  stdout:", (r.stdout or "").strip()[:500])
        print("  stderr:", (r.stderr or "").strip()[:500])
        sys.exit(1)
    return r


def local_version():
    try:
        return json.load(open(os.path.join(DEPLOY, "version.json"), encoding="utf-8"))
    except Exception:
        return {}


def remote_version():
    try:
        with urllib.request.urlopen(PAGES_URL + "version.json", timeout=25) as r:
            return json.loads(r.read().decode("utf-8"))
    except Exception as e:
        return {"_error": repr(e)[:120]}


def main():
    v = local_version()
    print("本地版本:", json.dumps(v, ensure_ascii=False))
    if "--status" in sys.argv:
        print("线上版本:", json.dumps(remote_version(), ensure_ascii=False))
        return

    # 1) 准备常驻克隆
    os.makedirs(os.path.dirname(CLONE), exist_ok=True)
    if not os.path.isdir(os.path.join(CLONE, ".git")):
        print("克隆仓库…")
        run(["git", "clone", "--depth", "1", REPO, CLONE])
    else:
        print("拉取远端最新…")
        run(["git", "-C", CLONE, "fetch", "--depth", "1", "origin", "main"])
        run(["git", "-C", CLONE, "reset", "--hard", "origin/main"])

    # 2) 覆盖两个文件
    for fn in SRC:
        src = os.path.join(DEPLOY, fn)
        if not os.path.exists(src):
            sys.exit("缺少源文件: " + src)
        shutil.copyfile(src, os.path.join(CLONE, fn))
        print("  覆盖 %-14s %d bytes" % (fn, os.path.getsize(src)))

    # 3) 有没有实际变化
    st = run(["git", "-C", CLONE, "status", "--porcelain"]).stdout.strip()
    if not st:
        print("远端内容与本地一致，无需推送。")
        return

    # 4) 提交并推送
    msg = "看板更新 BUILD %s（数据截止 %s）" % (v.get("v", "?"), v.get("cut", "?"))
    run(["git", "-C", CLONE, "add", "-A"])
    run(["git", "-C", CLONE, "-c", "user.name=" + AUTHOR_NAME,
         "-c", "user.email=" + AUTHOR_EMAIL, "commit", "-q", "-m", msg])
    r = run(["git", "-C", CLONE, "push", "origin", "main"])
    print((r.stderr or "").strip()[-300:])
    print("已推送:", msg)
    sha = run(["git", "-C", CLONE, "rev-parse", "--short", "HEAD"]).stdout.strip()
    print("commit:", sha)
    print("commit 页: https://github.com/mouren2580/fangtai-dashboard/commit/" + sha)
    print("Pages 构建通常需 1–3 分钟，稍后用 --status 复核。")


if __name__ == "__main__":
    main()
