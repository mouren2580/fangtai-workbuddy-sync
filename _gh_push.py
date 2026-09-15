# -*- coding: utf-8 -*-
"""最小 GitHub Pages 推送：用 Git Data API 把 index.html 更新到 mouren2580/fangtai-dashboard main。
PAT 从环境变量 GITHUB_PAT 读取，不在文件中硬编码。"""
import os, sys, json, base64, time, urllib.request, urllib.error, ssl

API = "https://api.github.com"
OWNER, REPO, BRANCH = "mouren2580", "fangtai-dashboard", "main"
SRC = r"D:\WorkBuddy\deploy_cs\index.html"
TOKEN = os.environ.get("GITHUB_PAT")
if not TOKEN:
    sys.exit("未设置 GITHUB_PAT")
MSG = "同步2026-09看板数据(截止2026-09-10) BUILD 20260911-1702"


def req(method, url, payload=None, retries=6):
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
            with urllib.request.urlopen(r, timeout=120, context=ctx) as resp:
                body = resp.read().decode("utf-8")
                return resp.status, (json.loads(body) if body.strip() else {})
        except urllib.error.HTTPError as e:
            body = e.read().decode("utf-8", "replace")
            try:
                j = json.loads(body)
            except Exception:
                j = {"message": body[:300]}
            if 400 <= e.code < 500 and e.code not in (403, 429):
                return e.code, j
            last = (e.code, j)
        except Exception as e:
            last = (0, {"message": repr(e)[:200]})
        wait = 3 * (i + 1)
        print("  retry %d/%d after %ds: %s" % (i + 1, retries, wait, str(last[1].get("message", ""))[:80]))
        time.sleep(wait)
    return last


def main():
    base = "%s/repos/%s/%s" % (API, OWNER, REPO)
    raw = open(SRC, "rb").read()
    # 1) 当前 main
    code, j = req("GET", base + "/git/ref/heads/" + BRANCH, retries=3)
    if code != 200 or "object" not in j:
        sys.exit("GET ref 失败 %s %s" % (code, j))
    parent = j["object"]["sha"]
    # 2) 父 commit 的 tree
    code, jc = req("GET", base + "/git/commits/" + parent, retries=3)
    if code != 200:
        sys.exit("GET commit 失败 %s" % code)
    base_tree = jc["tree"]["sha"]
    # 3) blob
    code, jb = req("POST", base + "/git/blobs",
                   {"content": base64.b64encode(raw).decode("ascii"), "encoding": "base64"})
    if code not in (200, 201):
        sys.exit("BLOB 失败 %s %s" % (code, jb))
    blob_sha = jb["sha"]
    # 4) tree
    code, jt = req("POST", base + "/git/trees",
                   {"base_tree": base_tree, "tree": [{"path": "index.html", "mode": "100644",
                                                      "type": "blob", "sha": blob_sha}]})
    if code not in (200, 201):
        sys.exit("TREE 失败 %s %s" % (code, jt))
    # 5) commit
    code, jc2 = req("POST", base + "/git/commits",
                    {"message": MSG, "tree": jt["sha"], "parents": [parent]})
    if code not in (200, 201):
        sys.exit("COMMIT 失败 %s %s" % (code, jc2))
    new_sha = jc2["sha"]
    # 6) ref
    code, jr = req("PATCH", base + "/git/refs/heads/" + BRANCH, {"sha": new_sha, "force": False})
    if code not in (200, 201):
        sys.exit("REF 失败 %s %s" % (code, jr))
    print("OK GitHub Pages 已更新 -> commit %s" % new_sha[:10])


if __name__ == "__main__":
    main()
