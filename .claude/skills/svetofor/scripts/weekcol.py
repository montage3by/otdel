"""For sheets whose header has a week column for 26.10–01.11 (or 27.10–02.11 / 03.11), print each row's label with the previous-week and this-week values. With --fill, print cell fill colors instead (for Светофор)."""
import sys, re, datetime as dt
import openpyxl

WEEK = re.compile(r'(26|27)\.10\.?\s*[-–]\s*0?[12]\.11|26\.10\.?-\s*0?1\.11|^0?3\.11$')
PHONE = re.compile(r'(\+?7|8)[\s(-]*\d{3}[\s)-]*\d{3}[\s-]*\d{2}[\s-]*\d{2}')

def fmt(v):
    if v is None:
        return ''
    if isinstance(v, (dt.datetime, dt.date)):
        return v.strftime('%d.%m')
    if isinstance(v, float):
        return f'{v:.2f}'.rstrip('0').rstrip('.')
    return PHONE.sub('[тел]', str(v).replace('\n', ' / '))[:1500]

path, fill = sys.argv[1], '--fill' in sys.argv
only = [a for a in sys.argv[2:] if not a.startswith('--')]
wb = openpyxl.load_workbook(path, read_only=not fill, data_only=True)
for ws in wb.worksheets:
    if only and not any(o.lower() in ws.title.lower() for o in only):
        continue
    rows = list(ws.iter_rows(max_row=400))
    for hi, row in enumerate(rows):
        cols = [c for c in row if isinstance(c.value, str) and WEEK.search(c.value.strip())]
        if not cols:
            continue
        print(f'=== {ws.title} (header row {hi + 1})')
        for col in cols:
            ci = col.column - 1
            print(f'  -- column «{col.value.strip()}»')
            for r in rows[hi + 1:]:
                if ci >= len(r):
                    continue
                label = next((fmt(c.value) for c in r[:3] if c.value not in (None, '')), '')
                if fill:
                    f = r[ci].fill.fgColor.rgb if r[ci].fill and r[ci].fill.fill_type else None
                    if label:
                        print(f'    {label[:60]}: {f}  {fmt(r[ci].value)[:200]}')
                else:
                    cur, prev = fmt(r[ci].value), fmt(r[ci - 1].value) if ci else ''
                    if cur:
                        print(f'    {label[:80]}: {cur}   (пред.: {prev[:120]})')
        break
