import sys, re, openpyxl
PHONE = re.compile(r'(\+?7|8)[\s(-]*\d{3}[\s)-]*\d{3}[\s-]*\d{2}[\s-]*\d{2}')
PAT = re.compile(r'26\.10\s*[–-]\s*0?1\.11|26\.10\.2026|за неделю 26|Неделя 26|26\.10\.?\s*-\s*1\.11')
wb = openpyxl.load_workbook(sys.argv[1], read_only=True, data_only=True)
for ws in wb.worksheets:
    if any(k in ws.title.lower() for k in ('контент', 'кп', 'кп ', 'кон-план')): continue
    for row in ws.iter_rows(values_only=True, max_row=4000):
        for c in row:
            if isinstance(c, str) and len(c) > 250 and PAT.search(c):
                print(f'=== {ws.title}\n' + PHONE.sub('[тел]', c)[:3500] + '\n')
