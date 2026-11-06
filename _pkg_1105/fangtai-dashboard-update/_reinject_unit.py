# -*- coding: utf-8 -*-
"""把 _unit_revenue.json 注入 dashboard.html 的
/*__UNIT_START__*/ ... /*__UNIT_END__*/ 锚点（仅替换 UNIT_REVENUE_DATA 行）。幂等。"""
import json, re, sys

HTML = "dashboard.html"
SRC = "_unit_revenue.json"


def main():
    data = json.load(open(SRC, encoding="utf-8"))
    payload = json.dumps(data, ensure_ascii=False, separators=(",", ":"))
    payload = payload.replace("</", "<\\/")   # 防 </script> 提前闭合
    line = "var UNIT_REVENUE_DATA = %s;   // 单均收入(口径①服务产值/⑤服务产品)，_reinject_unit.py 注入" % payload

    h = open(HTML, encoding="utf-8").read()
    if "/*__UNIT_START__*/" not in h or "/*__UNIT_END__*/" not in h:
        print("❌ 找不到 UNIT 锚点"); sys.exit(1)
    pat = re.compile(r"var UNIT_REVENUE_DATA = .*?;\s*(//[^\n]*)?\n")
    if not pat.search(h):
        print("❌ 锚点内未找到 UNIT_REVENUE_DATA 声明行"); sys.exit(1)
    h2, n = pat.subn(line + "\n", h, count=1)
    open(HTML, "w", encoding="utf-8").write(h2)
    print("✅ 已注入 UNIT_REVENUE_DATA（替换 %d 处，个月份: %s）" % (n, ",".join(sorted(data.get("months", {})))))
    m = re.search(r"var UNIT_REVENUE_DATA = (\{.*?\});\s*//", h2, re.S)
    if not m:
        print("❌ 注入后校验失败：找不到声明"); sys.exit(1)
    json.loads(m.group(1))
    print("✅ 注入内容 JSON 合法")


if __name__ == "__main__":
    main()
