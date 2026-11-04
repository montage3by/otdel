import sys, openpyxl
wb = openpyxl.load_workbook(sys.argv[1], data_only=True)
ws = next(w for w in wb.worksheets if '2.0' in w.title)
hdr = [c.value for c in ws[1]]
import datetime as dt
idx = [i for i, v in enumerate(hdr) if (v.strftime('%d.%m') if isinstance(v, dt.datetime) else str(v).strip()) in ('13.10', '20.10', '27.10', '03.11')]
names = {'FFFF0000': 'КРАСН', 'FFFFFF00': 'ЖЁЛТ', 'FF00FF00': 'ЗЕЛ', None: '-'}
print('cols', [hdr[i] for i in idx])
proj=''
for row in ws.iter_rows(min_row=2, max_row=200):
    label = next((str(c.value) for c in row[:2] if c.value), '')
    if not label: continue
    if label not in ('Общая картина','Внешний фактор','ОПТ','Контент','Дизайн'): proj=label
    label=proj+' / '+label
    out = []
    for i in idx:
        c = row[i]
        f = c.fill.fgColor.rgb if c.fill and c.fill.fill_type else None
        out.append(f"{names.get(f, f)}{(' «'+str(c.value)[:150].replace(chr(10),' ')+'»') if c.value else ''}")
    print(f'{label[:45]:45} | ' + ' | '.join(out))
