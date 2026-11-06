# -*- coding: utf-8 -*-
"""把 _extend_time.json(延保时间窗口: 昨日/近3日/近7日) 注入 dashboard.html 的
/*__EXTENDTIME_START__*/ ... /*__EXTENDTIME_END__*/ 锚点，仅替换 EXTEND_TIME_DATA 行，
保留 EXTEND_YEAR_DATA / extWin 的声明。幂等。"""
import json, re, sys

HTML = "dashboard.html"
SRC = "_extend_time.json"

def main():
    data = json.load(open(SRC, encoding="utf-8"))
    payload = json.dumps(data, ensure_ascii=False, separators=(",", ":"))
    payload = payload.replace("</", "<\\/")  # 防 </script> 提前闭合
    line = "let EXTEND_TIME_DATA = %s;   // 延保时间窗口(昨日/近3日/近7日)，_reinject_extendtime.py 注入" % payload

    h = open(HTML, encoding="utf-8").read()
    if "/*__EXTENDTIME_START__*/" not in h or "/*__EXTENDTIME_END__*/" not in h:
        print("❌ 找不到 EXTENDTIME 锚点"); sys.exit(1)
    # 仅替换锚点内的 EXTEND_TIME_DATA 声明行（保留 EXTEND_YEAR_DATA / extWin）
    pat = re.compile(r"let EXTEND_TIME_DATA = .*?;(\s*//.*)?\n")
    if not pat.search(h):
        print("❌ 锚点内未找到 EXTEND_TIME_DATA 声明行"); sys.exit(1)
    h2 = pat.sub(line + "\n", h, count=1)
    if h2 == h:
        print("⚠ 未发生变化（可能已注入）")
    open(HTML, "w", encoding="utf-8").write(h2)
    print("✅ 已注入 EXTEND_TIME_DATA（窗口: %s）" % ", ".join(data.get("windows", {}).keys()))
    # 校验
    m = re.search(r"let EXTEND_TIME_DATA = (\{.*?\});", h2, re.S)
    json.loads(m.group(1))
    print("✅ 注入内容 JSON 合法")

if __name__ == "__main__":
    main()
