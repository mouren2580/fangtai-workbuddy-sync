import openpyxl, datetime

f = r"D:\WorkBuddy\2026年8月西北服务产品.xlsx"
wb = openpyxl.load_workbook(f, data_only=True)

def serial_to_date(s):
    return datetime.datetime(1899, 12, 30) + datetime.timedelta(days=s)

overall_max = None
for ws in wb.worksheets:
    hdr = None
    date_cols = []
    sheet_max = None
    for i, row in enumerate(ws.iter_rows(values_only=True)):
        if hdr is None:
            if i < 6 and row and any(isinstance(c, str) and ('日期' in c or '时间' in c) for c in row):
                hdr = row
                date_cols = [j for j, c in enumerate(row) if isinstance(c, str) and ('日期' in c or '时间' in c)]
                continue
            if i >= 6 and hdr is None:
                break
        if hdr is not None:
            for j in date_cols:
                if j >= len(row):
                    continue
                v = row[j]
                if v is None:
                    continue
                d = None
                if isinstance(v, datetime.datetime):
                    d = v
                elif isinstance(v, (int, float)):
                    if 40000 < v < 70000:
                        d = serial_to_date(v)
                if d:
                    d = d.replace(hour=0, minute=0, second=0, microsecond=0)
                    if sheet_max is None or d > sheet_max:
                        sheet_max = d
    if sheet_max:
        print(f"[{ws.title}] max={sheet_max.date()}  date_cols={date_cols}")
        if overall_max is None or sheet_max > overall_max:
            overall_max = sheet_max

print("OVERALL_MAX_DATE =", overall_max.date() if overall_max else None)
