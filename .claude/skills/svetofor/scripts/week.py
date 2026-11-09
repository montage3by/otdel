"""Всё по одной неделе из дневника: колонки-недели, блоки-недели и длинные комментарии.

Использование: python3 -I week.py <дневник.xlsx> <ДД.ММ начала> <ДД.ММ конца> [<ДД.ММ начала-1> <ДД.ММ конца-1>]
Пример: week.py diary.xlsx 02.11 08.11
Второй период (необязательно) — если в дневнике недели идут со вторника, например 03.11 09.11.
Телефоны маскируются. Выход: === лист / тип / что нашли.
"""
import sys, re, datetime as dt
import openpyxl

PHONE = re.compile(r'(\+?7|8)[\s(-]*\d{3}[\s)-]*\d{3}[\s-]*\d{2}[\s-]*\d{2}')
SKIP_LONG = ('контент', 'кп', 'кон-план', 'контент-план')


def num(d):
    day, mon = d.split('.')
    return rf'0?{int(day)}\.{mon}'


def week_re(periods):
    parts = [rf'{num(a)}\.?(?:\d{{2,4}})?\s*[-–—]\s*{num(b)}' for a, b in periods]
    return re.compile('|'.join(parts))


def fmt(v, n=1500):
    if v is None:
        return ''
    if isinstance(v, (dt.datetime, dt.date)):
        return v.strftime('%d.%m')
    if isinstance(v, float):
        return f'{v:.2f}'.rstrip('0').rstrip('.')
    return PHONE.sub('[тел]', str(v).replace('\n', ' / '))[:n]


args = [a for a in sys.argv[1:] if not a.startswith('--')]
path = args[0]
periods = [(args[1], args[2])] + ([(args[3], args[4])] if len(args) >= 5 else [])
WEEK = week_re(periods)

wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
seen_long = set()
found = False
SKIP_SHEET = ('контент', 'кп', 'кон-план', 'контент-план', 'дзен')
for ws in wb.worksheets:
    title = ws.title.lower()
    if any(k in title for k in SKIP_SHEET) and '--content' not in sys.argv:
        continue
    printed_cols = set()
    rows = list(ws.iter_rows(values_only=True, max_row=4000))
    for i, r in enumerate(rows):
        hits = [j for j, c in enumerate(r) if isinstance(c, str) and WEEK.search(c) and len(c) < 60]
        if not hits:
            continue
        j = hits[0]
        dated = sum(1 for c in r if isinstance(c, str) and re.search(r'\d\d?\.\d\d', c) and len(c) < 60)
        found = True
        if dated > 3:
            if j in printed_cols:
                continue
            printed_cols.add(j)
            # заголовок: недели — колонки. Печатаем эту колонку и предыдущую.
            print(f'=== {ws.title} / колонка «{r[j].strip()}» (строка {i + 1})')
            for rr in rows[i + 1:i + 60]:
                if j >= len(rr):
                    continue
                lab = next((fmt(c, 70) for c in rr[:3] if c not in (None, '')), '')
                cur = fmt(rr[j])
                if cur and lab:
                    prev = fmt(rr[j - 1], 120) if j else ''
                    print(f'    {lab}: {cur}   (пред.: {prev})')
        else:
            # блок: неделя — строка, таблица под ней
            print(f'=== {ws.title} / блок «{r[j].strip()}» (строка {i + 1})')
            for rr in rows[i:i + 14]:
                cells = [fmt(c, 1200) for c in rr[max(0, j - 1):j + 12] if c not in (None, '')]
                if cells:
                    print('    ' + ' | '.join(cells)[:1800])
    for r in rows:
        for c in r:
            if isinstance(c, str) and len(c) > 250 and WEEK.search(c) and c[:200] not in seen_long:
                seen_long.add(c[:200])
                found = True
                print(f'=== {ws.title} / комментарий\n' + fmt(c, 4000) + '\n')
if not found:
    print('НЕДЕЛЯ НЕ НАЙДЕНА')
