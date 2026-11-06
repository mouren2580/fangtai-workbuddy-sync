# -*- coding: utf-8 -*-
"""把 _cleaning_data.json(清洗保养原始记录: 按服务月) 注入 dashboard.html 的
/*__CLEANING_START__*/ ... /*__CLEANING_END__*/ 锚点，仅替换 CLEANING_DATA 声明行。幂等。"""
import json, re, sys

HTML = "dashboard.html"
SRC = "_cleaning_data.json"

def main():
    data = json.load(open(SRC, encoding="utf-8"))
    payload = json.dumps(data, ensure_ascii=False, separators=(",", ":"))
    payload = payload.replace("</", "<\\/")  # 防 </script> 提前闭合
    line = "let CLEANING_DATA = %s;   // 清洗保养原始记录(按服务月)，_reinject_cleaning.py 注入" % payload

    h = open(HTML, encoding="utf-8").read()
    if "/*__CLEANING_START__*/" not in h or "/*__CLEANING_END__*/" not in h:
        print("ERR 找不到 CLEANING 锚点"); sys.exit(1)
    pat = re.compile(r"let CLEANING_DATA = .*?;(\s*//.*)?\n")
    if not pat.search(h):
        print("ERR 锚点内未找到 CLEANING_DATA 声明行"); sys.exit(1)
    h2 = pat.sub(line + "\n", h, count=1)
    if h2 == h:
        print("WARN 未发生变化（可能已注入）")
    open(HTML, "w", encoding="utf-8").write(h2)
    # 校验
    m = re.search(r"let CLEANING_DATA = (\{.*?\});", h2, re.S)
    json.loads(m.group(1))
    n_months = len(data.get("months", {}))
    print("OK 已注入 CLEANING_DATA（服务月数:%d, 总记录:%d）" % (n_months, data.get("meta", {}).get("总记录", 0)))
    print("OK 注入内容 JSON 合法")

if __name__ == "__main__":
    main()
