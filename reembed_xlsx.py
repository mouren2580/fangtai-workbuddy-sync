import base64, re, os

SRC = r"D:\WorkBuddy\dashboard_offline.html"
XLSX = r"D:\WorkBuddy\2026年8月西北服务产品.xlsx"
DEPLOY = r"D:\WorkBuddy\deploy_cs\index.html"

html = open(SRC, encoding="utf-8").read()
data = open(XLSX, "rb").read()
b64 = base64.b64encode(data).decode("ascii")

pat = re.compile(r'<script[^>]*id="xlsxData"[^>]*>.*?</script>', re.S)
m = pat.search(html)
if not m:
    raise SystemExit("xlsxData block not found")

def repl(mm):
    block = mm.group(0)
    end_of_open = block.index(">") + 1
    return block[:end_of_open] + b64 + "</script>"

new_html = pat.sub(repl, html, count=1)
if new_html == html:
    raise SystemExit("NO CHANGE")

os.makedirs(os.path.dirname(DEPLOY), exist_ok=True)
open(SRC, "w", encoding="utf-8").write(new_html)
open(DEPLOY, "w", encoding="utf-8").write(new_html)
print("OK reembed. b64 length =", len(b64), "file size =", len(new_html))
