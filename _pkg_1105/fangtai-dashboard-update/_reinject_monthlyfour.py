# -*- coding: utf-8 -*-
"""把 _monthly_four.json 注入 dashboard.html 的 MONTHLY_FOUR 锚点。
重注入只替换数据块，不触碰切换逻辑（逻辑在锚点之外，由 _gen_monthly_four 流程的
dashboard.html 静态编辑提供）。幂等：锚点存在则整体替换，不存在则在 VALUEPARTS_END 后创建。
"""
import json, os, sys

HERE = os.path.dirname(os.path.abspath(__file__))
HTML = os.path.join(HERE, "dashboard.html")
SRC = os.path.join(HERE, "_monthly_four.json")
START = "/*__MONTHLYFOUR_START__*/"
END = "/*__MONTHLYFOUR_END__*/"
ANCHOR_AFTER = "/*__VALUEPARTS_END__*/"

def main():
    if not os.path.exists(SRC):
        sys.stderr.write("缺少 _monthly_four.json，请先运行 _gen_monthly_four.py\n")
        raise SystemExit(1)
    data = json.load(open(SRC, encoding='utf-8'))
    js = json.dumps(data, ensure_ascii=False, separators=(',', ':'))
    # 防止数据中的 </script> 提前闭合脚本块
    js = js.replace('</', '<\\/')
    block = "%s\nlet MONTHLY_FOUR = %s;\n%s" % (START, js, END)

    html = open(HTML, encoding='utf-8').read()
    if START in html and END in html:
        i = html.index(START); j = html.index(END) + len(END)
        html = html[:i] + block + html[j:]
        print("已整体替换 MONTHLY_FOUR 锚点")
    elif ANCHOR_AFTER in html:
        html = html.replace(ANCHOR_AFTER, ANCHOR_AFTER + "\n" + block + "\n", 1)
        print("已新建 MONTHLY_FOUR 锚点（置于 %s 之后）" % ANCHOR_AFTER)
    else:
        sys.stderr.write("找不到注入锚点 %s\n" % ANCHOR_AFTER)
        raise SystemExit(1)

    open(HTML, 'w', encoding='utf-8').write(html)
    sz = os.path.getsize(HTML)
    print("dashboard.html 现 %.2f MB" % (sz / 1048576.0))
    # 校验 JSON 可解析
    try:
        json.loads(js.replace('<\\/', '</'))
        print("✅ 数据 JSON 自洽可解析")
    except Exception as e:
        print("⚠️ JSON 解析异常:", e)

if __name__ == '__main__':
    main()
