"""In report sheets, find cells naming the 26.10–01.11 week and print the block of rows under them (vertical report layout)."""
import sys, re, datetime as dt
import openpyxl

WEEK = re.compile(r'(26|27)\.10\.?\s*[-–]\s*0?[12]\.11')
KEYS = ('отчет', 'отчёт', 'директ', 'таргет', 'вк', 'итог', 'kpi', 'кпи', 'валидац')
PHONE = re.compile(r'(\+?7|8)[\s(-]*\d{3}[\s)-]*\d{3}[\s-]*\d{2}[\s-]*\d{2}')

def fmt(v):
    if isinstance(v, (dt.datetime, dt.date)):
        return v.strftime('%d.%m')
    if isinstance(v, float):
        return f'{v:.2f}'.rstrip('0').rstrip('.')
    return PHONE.sub('[тел]', str(v).replace('\n', ' / '))[:700]

wb = openpyxl.load_workbook(sys.argv[1], read_only=True, data_only=True)
for ws in wb.worksheets:
    if not any(k in ws.title.lower() for k in KEYS):
        continue
    rows = [r for r in ws.iter_rows(values_only=True, max_row=3000)]
    for i, r in enumerate(rows):
        hit = [j for j, c in enumerate(r) if isinstance(c, str) and WEEK.search(c)]
        if not hit:
            continue
        j = hit[0]
        # header-style (many week labels in the row) is handled by weekcol.py
        if sum(1 for c in r if isinstance(c, str) and re.search(r'\d\d\.\d\d', c)) > 3:
            continue
        print(f'=== {ws.title}: row {i + 1}')
        for rr in rows[i:i + 14]:
            cells = [fmt(c) for c in rr[max(0, j - 1):j + 12] if c not in (None, '')]
            if cells:
                print('   ', ' | '.join(cells)[:1200])
