# -*- coding: utf-8 -*-
"""每次部署前 bump 缓存穿透版本号 BUILD。
同时更新两处，二者必须一致：
  1) dashboard.html 内 /*__CACHEBUST_START__*/…END*/ 块的 `var BUILD = "<tag>"`
  2) publish_gh/version.json 的 `v` 字段（页面启动时拉取比对，判断是否已过期）
用法: python _bump_build.py [--tag 20260902-1030]
"""
import re, sys, os, json, datetime, argparse
for s in (sys.stdout, sys.stderr):
    try: s.reconfigure(encoding='utf-8')
    except Exception: pass

ap = argparse.ArgumentParser()
ap.add_argument('--tag', default=None, help='自定义版本串，默认取当前时间 YYYYMMDD-HHMM')
ap.add_argument('--html', default='dashboard.html')
ap.add_argument('--json', default=os.path.join('publish_gh', 'version.json'))
a = ap.parse_args()

tag = a.tag or datetime.datetime.now().strftime('%Y%m%d-%H%M')

# 1) HTML 内 BUILD
h = open(a.html, encoding='utf-8').read()
m = re.search(r'/\*__CACHEBUST_START__\*/.*?/\*__CACHEBUST_END__\*/', h, re.S)
if not m:
    print('❌ 未找到 CACHEBUST 锚点'); sys.exit(1)
blk = m.group(0)
old = re.search(r'var BUILD = "([^"]*)"', blk)
if not old:
    print('❌ 块内未找到 var BUILD'); sys.exit(1)
new_blk = blk.replace('var BUILD = "%s"' % old.group(1), 'var BUILD = "%s"' % tag)
h = h[:m.start()] + new_blk + h[m.end():]
open(a.html, 'w', encoding='utf-8').write(h)
print('✅ HTML BUILD: %s -> %s' % (old.group(1), tag))

# 2) version.json（cut 取「默认显示月」的统计截止日，便于用户一眼核对）
cut = ''
cm = re.search(r'let CURRENT_MONTH = "([\d-]+)"', h)
if cm:
    mk = cm.group(1).replace('-', '_')
    mm = re.search(r'MDATA_%s_D = \{"meta":\{.*?"统计截止日":"([^"]*)"' % mk, h, re.S)
    if mm:
        cut = mm.group(1)
os.makedirs(os.path.dirname(a.json) or '.', exist_ok=True)
data = {'v': tag, 'cut': cut, 't': datetime.datetime.now().isoformat(timespec='seconds')}
with open(a.json, 'w', encoding='utf-8') as f:
    json.dump(data, f, ensure_ascii=False, separators=(',', ':'))
print('✅ version.json:', json.dumps(data, ensure_ascii=False))
