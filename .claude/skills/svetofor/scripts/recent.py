"""Print rows that mention dates in the window 26.10–02.11.2026 (or date cells in it) from each sheet of an xlsx."""
import sys, re, datetime as dt
import openpyxl

START, END = dt.date(2026, 10, 26), dt.date(2026, 11, 2)
TXT = re.compile(r'\b(2[6-9]|3[01])\.10(\.(20)?26)?\b|\b0?[12]\.11(\.(20)?26)?\b|\b(2[6-9]|3[01])\.10\s*[-–]\s*0?[1-2]\.11')
PHONE = re.compile(r'(\+?7|8)[\s(-]*\d{3}[\s)-]*\d{3}[\s-]*\d{2}[\s-]*\d{2}')

def in_window(v):
    if isinstance(v, dt.datetime):
        v = v.date()
    if isinstance(v, dt.date):
        return START <= v <= END
    if isinstance(v, str):
        return bool(TXT.search(v))
    return False

def fmt(v):
    if isinstance(v, (dt.datetime, dt.date)):
        return v.strftime('%d.%m')
    if isinstance(v, float):
        return f'{v:.2f}'.rstrip('0').rstrip('.')
    s = str(v).replace('\n', ' ')
    return PHONE.sub('[тел]', s)[:160]

path = sys.argv[1]
wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
for ws in wb.worksheets:
    if any(k in ws.title.lower() for k in ('доступ', 'пароль', 'логин')):
        continue
    hits = []
    header = None
    for i, row in enumerate(ws.iter_rows(values_only=True)):
        if i > 4000:
            break
        if header is None and sum(1 for c in row if c not in (None, '')) >= 3:
            header = row
        if any(in_window(c) for c in row):
            hits.append(row)
    if hits:
        print(f'=== {ws.title} ({len(hits)} rows)')
        if header:
            print('  HDR:', ' | '.join(fmt(c) for c in header if c not in (None, ''))[:400])
        for r in hits[:40]:
            print('  ', ' | '.join(fmt(c) for c in r if c not in (None, ''))[:600])
