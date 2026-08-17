import re, zipfile

# ---- 1) 校验新 Excel 含目标工作表 ----
z = zipfile.ZipFile("2026年8月西北服务产品.xlsx")
wbxml = z.read('xl/workbook.xml').decode('utf-8')
names = re.findall(r'<sheet [^>]*name="([^"]+)"', wbxml)
print("工作表:", names)
assert "服务产品收入统计" in names, "缺少 服务产品收入统计 表！"

# ---- 2) 修补本地看板副本 ----
p = "dashboard/local.html"
s = open(p, encoding='utf-8').read()

# 若已注入则跳过，保证幂等
if "auto-sync" in s:
    print("已注入过，跳过")
else:
    # 截止日 2026-08-16 -> 2026-08-17 (驱动 DEFAULT_ASOF)
    if '统计截止日":"2026-08-15"' in s:
        s = s.replace('统计截止日":"2026-08-15"', '统计截止日":"2026-08-17"')
    else:
        s = s.replace('统计截止日":"2026-08-16"', '统计截止日":"2026-08-17"')
    # 数据更新时间 -> 今天
    s = s.replace('_stamp":"2026-08-16T02:20:59.306Z"', '_stamp":"2026-08-18T07:25:00.000Z"')
    s = s.replace('_stamp":"2026-08-17T07:25:00.000Z"', '_stamp":"2026-08-18T07:25:00.000Z"')

    AUTO = r'''
<script>
/* 自动同步：从同源服务器拉取最新 Excel 并重建看板（无需手动点同步Excel） */
(async function(){
  try{
    const resp = await fetch('/2026年8月西北服务产品.xlsx');
    if(!resp.ok) throw new Error('HTTP '+resp.status);
    const buf = await resp.arrayBuffer();
    const wb = XLSX.read(buf, {type:'array'});
    const res = rebuildFromWorkbook(wb);
    if(res){
      DASHBOARD_DATA.meta._stamp = new Date().toISOString();
      DASHBOARD_DATA.meta.统计截止日 = '2026-08-17';
      computeTime(new Date('2026-08-17T00:00:00'));
      updateAsOfBadge();
      syncPeriodUI(); renderAll();
      console.log('[auto-sync] 已用最新 Excel 重建', res);
    } else {
      console.warn('[auto-sync] rebuild 返回空');
    }
  }catch(e){
    console.error('[auto-sync] 失败（将以嵌入数据展示）', e);
  }
})();
</script>
'''
    anchor = "</body>\n</html>"
    assert anchor in s, "未找到结尾锚点"
    s = s.replace(anchor, AUTO + anchor, 1)
    open(p, 'w', encoding='utf-8').write(s)
    print("本地看板已写入自动同步脚本")
