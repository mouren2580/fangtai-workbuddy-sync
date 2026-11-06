# -*- coding: utf-8 -*-
"""
通过 GitHub Git Data API 推送静态站点（绕开 github.com 主站，只用 api.github.com）。
适用场景：git push 因主站 443 不通失败，但 api.github.com 可达。

用法:
  python _gh_push.py <owner> <repo> <本地目录> --token <ghp_xxx> [--branch main] [--message "msg"]
"""
import argparse
import base64
import json
import os
import ssl
import sys
import time
import urllib.error
import urllib.request

API = "https://api.github.com"


def req(method, url, token, payload=None, retries=4):
    data = json.dumps(payload).encode("utf-8") if payload is not None else None
    last = None
    for i in range(retries):
        r = urllib.request.Request(url, data=data, method=method)
        r.add_header("Authorization", "token " + token)
        r.add_header("Accept", "application/vnd.github+json")
        r.add_header("Content-Type", "application/json")
        r.add_header("User-Agent", "wb-gh-push")
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
                j = {"message": body[:400]}
            # 4xx 通常无需重试
            if 400 <= e.code < 500 and e.code not in (403, 429):
                return e.code, j
            last = (e.code, j)
        except Exception as e:                      # 网络类错误
            last = (0, {"message": repr(e)})
        wait = 2 * (i + 1)
        print("    retry %d/%d after %ds  (%s)" % (i + 1, retries, wait, last[1].get("message", "")[:80]))
        time.sleep(wait)
    return last


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("owner")
    ap.add_argument("repo")
    ap.add_argument("directory")
    ap.add_argument("--token", required=True)
    ap.add_argument("--branch", default="main")
    ap.add_argument("--message", default="update dashboard")
    a = ap.parse_args()

    base = "%s/repos/%s/%s" % (API, a.owner, a.repo)

    # 1) 收集文件
    files = []
    for root, _dirs, names in os.walk(a.directory):
        if ".git" in root.replace("\\", "/").split("/"):
            continue
        for n in names:
            p = os.path.join(root, n)
            rel = os.path.relpath(p, a.directory).replace("\\", "/")
            files.append((rel, p))
    files.sort()
    print("files: %d" % len(files))

    # 2) 上传 blob
    tree = []
    for rel, p in files:
        raw = open(p, "rb").read()
        code, j = req("POST", base + "/git/blobs", a.token,
                      {"content": base64.b64encode(raw).decode("ascii"), "encoding": "base64"})
        if code not in (200, 201):
            print("  BLOB FAIL %s -> %s %s" % (rel, code, j.get("message")))
            sys.exit(1)
        print("  blob %-16s %8d B  sha=%s" % (rel, len(raw), j["sha"][:10]))
        tree.append({"path": rel, "mode": "100644", "type": "blob", "sha": j["sha"]})

    # 3) 取当前分支 head（存在则作为 parent）
    parents = []
    base_tree = None
    code, j = req("GET", base + "/git/ref/heads/" + a.branch, a.token, retries=2)
    if code == 200 and "object" in j:
        parents = [j["object"]["sha"]]
        c2, jc = req("GET", base + "/git/commits/" + parents[0], a.token, retries=2)
        if c2 == 200:
            base_tree = jc["tree"]["sha"]
        print("  parent commit = %s" % parents[0][:10])
    else:
        print("  no existing branch -> initial commit")

    # 4) 建 tree
    payload = {"tree": tree}
    if base_tree:
        payload["base_tree"] = base_tree
    code, j = req("POST", base + "/git/trees", a.token, payload)
    if code not in (200, 201):
        print("  TREE FAIL %s %s" % (code, j.get("message")))
        sys.exit(1)
    tree_sha = j["sha"]
    print("  tree = %s" % tree_sha[:10])

    # 5) 建 commit
    code, j = req("POST", base + "/git/commits", a.token,
                  {"message": a.message, "tree": tree_sha, "parents": parents})
    if code not in (200, 201):
        print("  COMMIT FAIL %s %s" % (code, j.get("message")))
        sys.exit(1)
    commit_sha = j["sha"]
    print("  commit = %s" % commit_sha[:10])

    # 6) 更新/创建 ref
    if parents:
        code, j = req("PATCH", base + "/git/refs/heads/" + a.branch, a.token,
                      {"sha": commit_sha, "force": False})
    else:
        code, j = req("POST", base + "/git/refs", a.token,
                      {"ref": "refs/heads/" + a.branch, "sha": commit_sha})
    if code not in (200, 201):
        print("  REF FAIL %s %s" % (code, j.get("message")))
        sys.exit(1)
    print("OK  %s@%s -> %s" % (a.branch, commit_sha[:10], "https://github.com/%s/%s" % (a.owner, a.repo)))


if __name__ == "__main__":
    main()
