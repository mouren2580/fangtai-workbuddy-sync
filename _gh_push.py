# -*- coding: utf-8 -*-
"""GitHub Pages 同步：把工作区最新看板推到 mouren2580/fangtai-dashboard（main / 根目录）。

同时更新 index.html 与 version.json（页面靠 version.json 做缓存穿透）。

PAT 读取优先级：
  1) 命令行第 1 个参数        python _gh_push.py ghp_xxx
  2) 环境变量                 GITHUB_PAT=ghp_xxx python _gh_push.py
  3) 本地文件                 ~/.workbuddy/.github_pat （只写一行 token，最省事，一次配置长期可用）
注意：token 只放在用户目录，**不要**放进 D:\\WorkBuddy（那是 git 工作区会被同步走）。

用法：
  python _gh_push.py                 # 用文件/环境变量里的 PAT
  python _gh_push.py <PAT>           # 临时指定
  python _gh_push.py --check         # 只查 PAT 是否有效，不推送
"""
import os, sys, json, base64, time, urllib.request, urllib.error, ssl

API = "https://api.github.com"
OWNER, REPO, BRANCH = "mouren2580", "fangtai-dashboard", "main"
HERE = os.path.dirname(os.path.abspath(__file__))
DEPLOY = os.path.join(HERE, "deploy_cs")
PAT_FILE = os.path.join(os.path.expanduser("~"), ".workbuddy", ".github_pat")
FILES = ["index.html", "version.json"]


def pick_token():
    args = [a for a in sys.argv[1:] if not a.startswith("--")]
    if args and args[0].strip():
        return args[0].strip()
    t = (os.environ.get("GITHUB_PAT") or "").strip()
    if t:
        return t
    if os.path.exists(PAT_FILE):
        t = open(PAT_FILE, encoding="utf-8").read().strip()
        if t:
            return t
    return None


TOKEN = pick_token()
CHECK_ONLY = "--check" in sys.argv

if not TOKEN:
    sys.exit("未找到 PAT。请用以下任一方式提供：\n"
             "  1) python _gh_push.py <PAT>\n"
             "  2) set GITHUB_PAT=<PAT> && python _gh_push.py\n"
             "  3) 把 PAT 写进 %s （一行，推荐）" % PAT_FILE)


def req(method, url, payload=None, retries=6, quiet=False):
    data = json.dumps(payload).encode("utf-8") if payload is not None else None
    last = None
    for i in range(retries):
        r = urllib.request.Request(url, data=data, method=method)
        r.add_header("Authorization", "token " + TOKEN)
        r.add_header("Accept", "application/vnd.github+json")
        r.add_header("Content-Type", "application/json")
        r.add_header("User-Agent", "wb-sync")
        try:
            ctx = ssl.create_default_context()
            with urllib.request.urlopen(r, timeout=180, context=ctx) as resp:
                body = resp.read().decode("utf-8")
                return resp.status, (json.loads(body) if body.strip() else {})
        except urllib.error.HTTPError as e:
            body = e.read().decode("utf-8", "replace")
            try:
                j = json.loads(body)
            except Exception:
                j = {"message": body[:300]}
            # 4xx（除 403/429）直接返回，不重试
            if 400 <= e.code < 500 and e.code not in (403, 429):
                return e.code, j
            last = (e.code, j)
        except Exception as e:
            last = (0, {"message": repr(e)[:200]})
        if not quiet:
            wait = 3 * (i + 1)
            print("  retry %d/%d after %ds: %s" % (i + 1, retries, wait,
                                                   str(last[1].get("message", ""))[:80]))
            time.sleep(wait)
    return last


def build_message():
    try:
        v = json.load(open(os.path.join(DEPLOY, "version.json"), encoding="utf-8"))
        return "看板更新 BUILD %s（数据截止 %s）" % (v.get("v", "?"), v.get("cut", "?"))
    except Exception:
        return "看板更新"


def main():
    base = "%s/repos/%s/%s" % (API, OWNER, REPO)

    # 0) 先验 PAT
    code, j = req("GET", base, retries=2)
    if code != 200:
        sys.exit("PAT 无效或无权访问 %s/%s：HTTP %s %s" % (OWNER, REPO, code,
                                                       j.get("message", "")))
    print("PAT 有效，仓库可访问：%s（默认分支 %s）" % (j.get("full_name"), j.get("default_branch")))
    if CHECK_ONLY:
        return

    # 1) 当前 main 指向
    code, j = req("GET", base + "/git/ref/heads/" + BRANCH, retries=3)
    if code != 200 or "object" not in j:
        sys.exit("GET ref 失败 %s %s" % (code, j))
    parent = j["object"]["sha"]
    print("当前 main = %s" % parent[:10])

    # 2) 父 commit 的 tree
    code, jc = req("GET", base + "/git/commits/" + parent, retries=3)
    if code != 200:
        sys.exit("GET commit 失败 %s" % code)
    base_tree = jc["tree"]["sha"]

    # 3) 逐个建 blob
    tree_items = []
    for fn in FILES:
        p = os.path.join(DEPLOY, fn)
        if not os.path.exists(p):
            sys.exit("缺少文件：%s" % p)
        raw = open(p, "rb").read()
        code, jb = req("POST", base + "/git/blobs",
                       {"content": base64.b64encode(raw).decode("ascii"),
                        "encoding": "base64"})
        if code not in (200, 201):
            sys.exit("BLOB(%s) 失败 %s %s" % (fn, code, jb))
        tree_items.append({"path": fn, "mode": "100644", "type": "blob", "sha": jb["sha"]})
        print("  blob %-14s %d bytes -> %s" % (fn, len(raw), jb["sha"][:10]))

    # 4) tree
    code, jt = req("POST", base + "/git/trees",
                   {"base_tree": base_tree, "tree": tree_items})
    if code not in (200, 201):
        sys.exit("TREE 失败 %s %s" % (code, jt))

    # 5) commit
    msg = build_message()
    code, jc2 = req("POST", base + "/git/commits",
                    {"message": msg, "tree": jt["sha"], "parents": [parent]})
    if code not in (200, 201):
        sys.exit("COMMIT 失败 %s %s" % (code, jc2))
    new_sha = jc2["sha"]

    # 6) 移动分支
    code, jr = req("PATCH", base + "/git/refs/heads/" + BRANCH,
                   {"sha": new_sha, "force": False})
    if code not in (200, 201):
        sys.exit("REF 失败 %s %s" % (code, jr))

    print("OK  GitHub Pages 已更新 -> commit %s" % new_sha[:10])
    print("    commit message: %s" % msg)
    print("    https://github.com/%s/%s/commit/%s" % (OWNER, REPO, new_sha))


if __name__ == "__main__":
    main()
