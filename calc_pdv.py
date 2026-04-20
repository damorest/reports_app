#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Розрахунок сум ПДВ по елеваторних послугах та формування зведеного звіту.

Використання:
    python3 calc_pdv.py <input_file.xls>

Приклад:
    python3 calc_pdv.py reports/main_report_april2026.xls

Якщо input_file не вказано — береться значення за замовчуванням (DEFAULT_INPUT).
Результат зберігається поруч із вхідним файлом з суфіксом _result.xls.
"""
import sys
import os
import xlrd
import xlwt
from xlutils.copy import copy
from collections import defaultdict

DEFAULT_INPUT = '/Users/macbookpro/mhp/reports/2026-03/main_report_march2026.xls'

if len(sys.argv) > 1:
    INPUT = sys.argv[1]
    if not os.path.isabs(INPUT):
        INPUT = os.path.join(os.getcwd(), INPUT)
else:
    INPUT = DEFAULT_INPUT

base, ext = os.path.splitext(INPUT)
OUTPUT = base + '_result' + ext

# ============================================================
# Ціни ПДВ (з ПДВ − без ПДВ) з прайс-листів PDF
# Ключ: рядок-фрагмент (lowercase) назви організації
# ============================================================
PRICES = {
    'вквк': {
        'кукурудза': {'приймання': 4.46,  'очистка': 5.42, 'сушка': 23.40, 'зберігання': 0.62},
        'соняшник':  {'приймання': 5.43,  'очистка': 5.84, 'сушка': 24.70, 'зберігання': 0.70},
        'соя':       {'приймання': 5.00,  'очистка': 5.43, 'сушка': 24.29, 'зберігання': 0.62},
    },
    'катеринопільськ': {
        'кукурудза': {'приймання': 5.00,  'очистка': 5.42, 'сушка': 23.40, 'зберігання': 0.62},
        'соняшник':  {'приймання': 5.60,  'очистка': 5.84, 'сушка': 25.34, 'зберігання': 0.70},
        'соя':       {'приймання': 5.00,  'очистка': 5.43, 'сушка': 24.29, 'зберігання': 0.61},
    },
    'мзвкк': {
        'кукурудза': {'приймання': 5.00,  'очистка': 5.42, 'сушка': 23.40, 'зберігання': 0.62},
        'соняшник':  {'приймання': 5.60,  'очистка': 5.84, 'сушка': 24.70, 'зберігання': 0.70},
        'соя':       {'приймання': 5.00,  'очистка': 5.43, 'сушка': 24.29, 'зберігання': 0.61},
    },
    'андріяшів': {
        'кукурудза': {'приймання': 4.20,  'очистка': 5.40, 'сушка': 23.92, 'зберігання': 0.67},
        'соняшник':  {'приймання': 5.00,  'очистка': 5.62, 'сушка': 24.48, 'зберігання': 0.76},
        'соя':       {'приймання': 5.00,  'очистка': 5.62, 'сушка': 24.48, 'зберігання': 0.72},
    },
    'ямпільськ': {
        'кукурудза': {'приймання': 5.00,  'очистка': 5.44, 'сушка': 24.00, 'зберігання': 0.63},
        'соя':       {'приймання': 5.60,  'очистка': 5.73, 'сушка': 24.00, 'зберігання': 0.65},
    },
    'вендичанськ': {
        'кукурудза': {'приймання': 4.46,  'очистка': 5.42, 'сушка': 22.51, 'зберігання': 0.62},
        'соняшник':  {'приймання': 5.43,  'очистка': 5.84, 'сушка': 24.70, 'зберігання': 0.70},
        'соя':       {'приймання': 5.00,  'очистка': 5.43, 'сушка': 24.29, 'зберігання': 0.62},
    },
    'елеваторний': {  # Соколівський / Шпиківський / Калинівський
        'кукурудза': {'приймання': 4.46,  'очистка': 5.42, 'сушка': 21.60, 'зберігання': 0.62},
        'соняшник':  {'приймання': 5.43,  'очистка': 5.84, 'сушка': 24.68, 'зберігання': 0.70},
        'соя':       {'приймання': 5.00,  'очистка': 5.43, 'сушка': 24.29, 'зберігання': 0.62},
    },
    'воскресинц': {
        'кукурудза': {'приймання': 6.00,  'очистка': 6.00, 'сушка': 24.00, 'зберігання': 0.67},
        'соя':       {'приймання': 6.19,  'очистка': 6.29, 'сушка': 24.00, 'зберігання': 0.79},
    },
    'львівськ': {
        'кукурудза': {'приймання': 6.19,  'очистка': 6.84, 'сушка': 24.00, 'зберігання': 0.62},
        'соя':       {'приймання': 6.41,  'очистка': 6.99, 'сушка': 24.51, 'зберігання': 0.74},
    },
    'краснянськ': {
        'кукурудза': {'приймання': 6.00,  'очистка': 6.00, 'сушка': 24.00, 'зберігання': 0.63},
        'соя':       {'приймання': 6.00,  'очистка': 6.00, 'сушка': 24.51, 'зберігання': 0.76},
    },
    'новомосковськ': {
        'соняшник':  {'приймання': 5.59,  'очистка': 6.07, 'сушка': 29.40, 'зберігання': 0.72},
        'соя':       {'приймання': 5.00,  'очистка': 6.06, 'сушка': 28.00, 'зберігання': 0.71},
    },
    'яготинськ': {
        'кукурудза': {'приймання': 5.80,  'очистка': 5.58, 'сушка': 23.40, 'зберігання': 0.61},
        'соняшник':  {'приймання': 5.97,  'очистка': 6.03, 'сушка': 25.85, 'зберігання': 0.70},
        'соя':       {'приймання': 6.21,  'очистка': 6.03, 'сушка': 25.05, 'зберігання': 0.61},
    },
    'перспектив': {  # Городенківський ел-тор ф-я Перспектив
        'кукурудза': {'приймання': 6.60,  'очистка': 7.68, 'сушка': 24.00, 'зберігання': 0.70},
        'соя':       {'приймання': 7.08,  'очистка': 7.62, 'сушка': 25.34, 'зберігання': 1.05},
    },
    # Батьківські компанії → використовують прайс відповідного елеватора
    'агрокряж': {   # МХП-Агрокряж ТОВ = Вендичанський
        'кукурудза': {'приймання': 4.46,  'очистка': 5.42, 'сушка': 22.51, 'зберігання': 0.62},
        'соняшник':  {'приймання': 5.43,  'очистка': 5.84, 'сушка': 24.70, 'зберігання': 0.70},
        'соя':       {'приймання': 5.00,  'очистка': 5.43, 'сушка': 24.29, 'зберігання': 0.62},
    },
    'урожайна': {   # МХП-Урожайна Країна = Андріяшівський
        'кукурудза': {'приймання': 4.20,  'очистка': 5.40, 'сушка': 23.92, 'зберігання': 0.67},
        'соняшник':  {'приймання': 5.00,  'очистка': 5.62, 'сушка': 24.48, 'зберігання': 0.76},
        'соя':       {'приймання': 5.00,  'очистка': 5.62, 'сушка': 24.48, 'зберігання': 0.72},
    },
    'агро-с': {     # МХП-Агро-С ТОВ = Яготинський
        'кукурудза': {'приймання': 5.80,  'очистка': 5.58, 'сушка': 23.40, 'зберігання': 0.61},
        'соняшник':  {'приймання': 5.97,  'очистка': 6.03, 'сушка': 25.85, 'зберігання': 0.70},
        'соя':       {'приймання': 6.21,  'очистка': 6.03, 'сушка': 25.05, 'зберігання': 0.61},
    },
    'птахофабрика': {  # Вінницька птахофабрика ТОВ = ВКВК
        'кукурудза': {'приймання': 4.46,  'очистка': 5.42, 'сушка': 23.40, 'зберігання': 0.62},
        'соняшник':  {'приймання': 5.43,  'очистка': 5.84, 'сушка': 24.70, 'зберігання': 0.70},
        'соя':       {'приймання': 5.00,  'очистка': 5.43, 'сушка': 24.29, 'зберігання': 0.62},
    },
}

# Пріоритетний список ключів (довші/специфічніші — першими)
ORG_KEYS_ORDERED = [
    'катеринопільськ', 'мзвкк', 'андріяшів', 'ямпільськ', 'вендичанськ',
    'елеваторний', 'воскресинц', 'львівськ', 'краснянськ', 'новомосковськ',
    'яготинськ', 'перспектив', 'агрокряж', 'урожайна', 'агро-с',
    'птахофабрика', 'вквк',
]

# Для визначення "переміщень між філіями" — групування в одну юридичну сутність
ENTITY = {
    'вквк': 'ВКВК', 'птахофабрика': 'ВКВК',
    'вендичанськ': 'МХП-Агрокряж', 'агрокряж': 'МХП-Агрокряж',
    'андріяшів': 'МХП-УК', 'урожайна': 'МХП-УК',
    'яготинськ': 'МХП-Агро-С', 'агро-с': 'МХП-Агро-С',
    'катеринопільськ': 'Катеринопільський', 'мзвкк': 'МЗВКК',
    'ямпільськ': 'Ямпільський', 'елеваторний': 'Елеваторний',
    'воскресинц': 'Захід-Агро', 'львівськ': 'Захід-Агро', 'краснянськ': 'Захід-Агро',
    'новомосковськ': 'Оріль-Лідер',
    'перспектив': 'Перспектив',
}


def get_org_key(name):
    if not name or not isinstance(name, str):
        return None
    low = name.lower()
    for k in ORG_KEYS_ORDERED:
        if k in low:
            return k
    return None


def get_entity(name):
    k = get_org_key(name)
    return ENTITY.get(k) if k else None


# Ціни ранніх зернових (TE-03, ВКВК) — єдиний доступний прайс на ранні зернові.
# Застосовується для всіх організацій, де зустрічається пшениця/ячмінь/ріпак.
EARLY_GRAIN_PRICES = {
    'пшениця': {'приймання': 5.30, 'очистка': 5.52, 'сушка': 20.78, 'зберігання': 0.64},
    'ріпак':   {'приймання': 5.84, 'очистка': 6.20, 'сушка': 22.16, 'зберігання': 0.70},
}


def get_crop_key(nom):
    if not nom:
        return None
    low = str(nom).lower()
    if 'кукурудза' in low or 'сорго' in low:
        return 'кукурудза'
    if 'соняшник' in low:
        return 'соняшник'
    if 'соя' in low:
        return 'соя'
    if 'пшениця' in low or 'ячмінь' in low or 'жито' in low or 'овес' in low:
        return 'пшениця'
    if 'ріпак' in low:
        return 'ріпак'
    return None


def get_vat(org, nom, service, warn_collector=None):
    """Повертає суму ПДВ. Якщо ціна не знайдена — додає попередження у warn_collector."""
    ck = get_crop_key(nom)
    if not ck:
        if warn_collector is not None and nom and str(nom).strip():
            warn_collector.append({
                'тип': 'Невідома культура',
                'організація': str(org).strip(),
                'культура': str(nom).strip(),
                'послуга': service,
                'опис': f'Культуру "{str(nom).strip()}" не знайдено у прайс-листах. Сума ПДВ = 0.',
            })
        return 0.0
    # Ранні зернові — єдиний прайс для всіх організацій (TE-03)
    if ck in EARLY_GRAIN_PRICES:
        return EARLY_GRAIN_PRICES[ck].get(service, 0.0)
    ok = get_org_key(org)
    if not ok:
        if warn_collector is not None:
            warn_collector.append({
                'тип': 'Невідома організація',
                'організація': str(org).strip(),
                'культура': str(nom).strip(),
                'послуга': service,
                'опис': f'Організацію "{str(org).strip()}" не знайдено у прайс-листах. Сума ПДВ = 0.',
            })
        return 0.0
    price = PRICES.get(ok, {}).get(ck, {}).get(service, None)
    if price is None:
        if warn_collector is not None:
            warn_collector.append({
                'тип': 'Відсутня ціна',
                'організація': str(org).strip(),
                'культура': str(nom).strip(),
                'послуга': service,
                'опис': f'Немає ціни для {str(nom).strip()} / {service} у прайсі організації "{str(org).strip()}". Сума ПДВ = 0.',
            })
        return 0.0
    return price


import re as _re
_DATE_RE = _re.compile(r'^\d{2}\.\d{2}\.\d{4}')  # 01.03.2026...

def is_data_row(cell_type, val):
    """Рядок є рядком даних (не дата, не заголовок, не підсумок)."""
    if cell_type != xlrd.XL_CELL_TEXT:
        return False
    low = val.lower().strip()
    if not low:
        return False
    # Відкидаємо дати, записані як текст
    if _DATE_RE.match(low):
        return False
    skip = {'организация', 'итого', 'дата', 'отбор:', 'сформирован:',
            'послуги зберігання', 'номенклатура', 'отправительполучатель',
            'лабораторный', 'лабораторний', 'аналізи', 'вище норми', 'нижче норми',
            'анализ', 'артеміда', 'урожай нвф', 'околиця'}
    for s in skip:
        if low.startswith(s):
            return False
    return True


def is_valid_nom(nom):
    """Назва культури — не порожня, не занадто довга, не виглядає як назва організації."""
    if not nom or not str(nom).strip():
        return False
    s = str(nom).strip()
    if len(s) > 60:
        return False
    low = s.lower()
    org_words = ('елеватор', 'філія', ' тов', ' прат', ' фг ', ' пп ', 'вентилюв', 'поступлен')
    if any(w in low for w in org_words):
        return False
    return True


# ============================================================
# ОСНОВНА ФУНКЦІЯ ОБРОБКИ
# ============================================================
def process(input_bytes: bytes) -> tuple:
    """
    Обробляє вхідний .xls файл (як bytes) та повертає:
      (output_bytes, unique_warnings, normal_count, internal_count)
    """
    import io

    rb = xlrd.open_workbook(file_contents=input_bytes, formatting_info=True)
    wb = copy(rb)

    summary  = defaultdict(lambda: defaultdict(float))
    warnings = []

    # ----------------------------------------------------------
    # 1. ЗБЕРІГАННЯ  (sheet 0)
    # cols: org=0, kontrag=6, kultura=10, delta=19, cina=20, suma=21
    # ----------------------------------------------------------
    rs = rb.sheet_by_name('зберігання')
    ws = wb.get_sheet(0)

    if not rs.cell_value(6, 20):
        ws.write(6, 20, 'ціна ПДВ')
    if not rs.cell_value(6, 21):
        ws.write(6, 21, 'Сума ПДВ')

    for i in range(7, rs.nrows):
        ct  = rs.cell_type(i, 0)
        org = rs.cell_value(i, 0)
        if not is_data_row(ct, org):
            continue
        kontrag = rs.cell_value(i, 6)
        nom     = rs.cell_value(i, 10)
        delta   = rs.cell_value(i, 19)
        if not isinstance(delta, (int, float)) or delta == 0 or not is_valid_nom(nom):
            ws.write(i, 20, '')
            ws.write(i, 21, '')
            continue
        cina = get_vat(org, nom, 'зберігання', warnings)
        suma = round(delta * cina, 6)
        ws.write(i, 20, cina)
        ws.write(i, 21, suma)
        if cina > 0:
            summary[(str(org).strip(), str(kontrag).strip())]['зберігання'] += suma

    # ----------------------------------------------------------
    # 2. СУШКА  (sheet 1)
    # cols: org=0, kontrag=1, nom=2, delta_och=5
    #       cina_och=9, suma_och=10, cina_sus=11, suma_sus=12
    # ----------------------------------------------------------
    rs2 = rb.sheet_by_name('сушка')
    ws2 = wb.get_sheet(1)

    for i in range(20, rs2.nrows):
        ct  = rs2.cell_type(i, 0)
        org = rs2.cell_value(i, 0)
        if not is_data_row(ct, org):
            ws2.write(i, 9,  '')
            ws2.write(i, 10, '')
            ws2.write(i, 11, '')
            ws2.write(i, 12, '')
            continue
        kontrag = rs2.cell_value(i, 1)
        nom     = rs2.cell_value(i, 2)
        d_och   = rs2.cell_value(i, 5)
        d = float(d_och) / 1000.0 if isinstance(d_och, (int, float)) and d_och else 0.0
        if not d or not is_valid_nom(nom):
            ws2.write(i, 9,  '')
            ws2.write(i, 10, '')
            ws2.write(i, 11, '')
            ws2.write(i, 12, '')
            continue
        c_och = get_vat(org, nom, 'очистка', warnings)
        c_sus = get_vat(org, nom, 'сушка',   warnings)
        s_och = round(d * c_och, 6)
        s_sus = round(d * c_sus, 6)
        ws2.write(i, 9,  c_och)
        ws2.write(i, 10, s_och)
        ws2.write(i, 11, c_sus)
        ws2.write(i, 12, s_sus)
        total_sushka = s_och + s_sus
        if total_sushka > 0:
            summary[(str(org).strip(), str(kontrag).strip())]['сушка'] += total_sushka

    # ----------------------------------------------------------
    # 3. ПРИЙМАННЯ  (sheet 2)
    # cols: org=0, kontrag=5, kultura=11, fizves=20, cina=28, suma=29
    # ----------------------------------------------------------
    rs3 = rb.sheet_by_name('приймання')
    ws3 = wb.get_sheet(2)

    if not rs3.cell_value(2, 28):
        ws3.write(2, 28, 'цінаПДВ')

    for i in range(3, rs3.nrows):
        ct  = rs3.cell_type(i, 0)
        org = rs3.cell_value(i, 0)
        if not is_data_row(ct, org):
            ws3.write(i, 28, '')
            ws3.write(i, 29, '')
            continue
        kontrag = rs3.cell_value(i, 5)
        nom     = rs3.cell_value(i, 11)
        fizves  = rs3.cell_value(i, 20)
        if not isinstance(fizves, (int, float)) or fizves == 0 or not is_valid_nom(nom):
            ws3.write(i, 28, '')
            ws3.write(i, 29, '')
            continue
        cina = get_vat(org, nom, 'приймання', warnings)
        suma = round(float(fizves) * cina, 6)
        ws3.write(i, 28, cina)
        ws3.write(i, 29, suma)
        if cina > 0:
            summary[(str(org).strip(), str(kontrag).strip())]['приймання'] += suma

    # ----------------------------------------------------------
    # 4. ЗВЕДЕНИЙ ЗВІТ
    # ----------------------------------------------------------
    ws4 = wb.add_sheet('Зведений звіт')

    hdr      = xlwt.easyxf('font: bold true; borders: bottom thin')
    num      = xlwt.easyxf(num_format_str='#,##0.00')
    num_bold = xlwt.easyxf('font: bold true', num_format_str='#,##0.00')
    red_num  = xlwt.easyxf('font: colour red', num_format_str='#,##0.00')
    red_bold = xlwt.easyxf('font: bold true, colour red', num_format_str='#,##0.00')

    for c, w in enumerate([45, 45, 18, 18, 18, 18]):
        ws4.col(c).width = w * 256

    headers = ['Організація', 'Контрагент', 'Зберігання', 'Сушка/Очистка', 'Приймання', 'Разом']
    for c, h in enumerate(headers):
        ws4.write(0, c, h, hdr)

    normal_rows = []
    internal_rows = []
    for (org, kont), svc in sorted(summary.items()):
        zbr = svc.get('зберігання', 0)
        sus = svc.get('сушка', 0)
        prm = svc.get('приймання', 0)
        tot = zbr + sus + prm
        e_org  = get_entity(org)
        e_kont = get_entity(kont)
        is_int = bool(e_org and e_kont and e_org == e_kont)
        entry = (org, kont, zbr, sus, prm, tot, is_int)
        (internal_rows if is_int else normal_rows).append(entry)

    row = 1
    tot_zbr = tot_sus = tot_prm = tot_all = 0.0
    for org, kont, zbr, sus, prm, tot, _ in normal_rows:
        ws4.write(row, 0, org)
        ws4.write(row, 1, kont)
        ws4.write(row, 2, zbr, num)
        ws4.write(row, 3, sus, num)
        ws4.write(row, 4, prm, num)
        ws4.write(row, 5, tot, num)
        tot_zbr += zbr; tot_sus += sus; tot_prm += prm; tot_all += tot
        row += 1
    ws4.write(row, 0, 'РАЗОМ', num_bold)
    ws4.write(row, 2, tot_zbr, num_bold)
    ws4.write(row, 3, tot_sus, num_bold)
    ws4.write(row, 4, tot_prm, num_bold)
    ws4.write(row, 5, tot_all, num_bold)
    row += 2

    if internal_rows:
        ws4.write(row, 0, 'ПЕРЕМІЩЕННЯ МІЖ ФІЛІЯМИ (не включаються до основного звіту)', hdr)
        row += 1
        for c, h in enumerate(headers):
            ws4.write(row, c, h, hdr)
        row += 1
        int_zbr = int_sus = int_prm = int_all = 0.0
        for org, kont, zbr, sus, prm, tot, _ in internal_rows:
            ws4.write(row, 0, org,  xlwt.easyxf('font: colour grey50'))
            ws4.write(row, 1, kont, xlwt.easyxf('font: colour grey50'))
            ws4.write(row, 2, zbr, red_num)
            ws4.write(row, 3, sus, red_num)
            ws4.write(row, 4, prm, red_num)
            ws4.write(row, 5, tot, red_num)
            int_zbr += zbr; int_sus += sus; int_prm += prm; int_all += tot
            row += 1
        ws4.write(row, 0, 'РАЗОМ переміщення', red_bold)
        ws4.write(row, 2, int_zbr, red_bold)
        ws4.write(row, 3, int_sus, red_bold)
        ws4.write(row, 4, int_prm, red_bold)
        ws4.write(row, 5, int_all, red_bold)

    # ----------------------------------------------------------
    # 5. ПОПЕРЕДЖЕННЯ
    # ----------------------------------------------------------
    seen_warns = set()
    unique_warnings = []
    for w in warnings:
        key = (w['тип'], w['організація'], w['культура'], w['послуга'])
        if key not in seen_warns:
            seen_warns.add(key)
            unique_warnings.append(w)

    warn_orange = xlwt.easyxf('font: bold true; pattern: pattern solid, fore_colour orange')
    warn_red    = xlwt.easyxf('pattern: pattern solid, fore_colour light_orange')
    warn_hdr    = xlwt.easyxf('font: bold true; borders: bottom thin')
    ok_green    = xlwt.easyxf('font: bold true; pattern: pattern solid, fore_colour light_green')

    ws5 = wb.add_sheet('Попередження')
    for c, w in enumerate([25, 45, 25, 15, 65]):
        ws5.col(c).width = w * 256

    if unique_warnings:
        ws5.write(0, 0, f'⚠ Знайдено {len(unique_warnings)} позицій без ціни ПДВ — суми для них = 0', warn_orange)
        ws5.write(2, 0, 'Тип проблеми',  warn_hdr)
        ws5.write(2, 1, 'Організація',   warn_hdr)
        ws5.write(2, 2, 'Культура',      warn_hdr)
        ws5.write(2, 3, 'Послуга',       warn_hdr)
        ws5.write(2, 4, 'Що зробити',    warn_hdr)
        for r, w in enumerate(unique_warnings, start=3):
            ws5.write(r, 0, w['тип'],         warn_red)
            ws5.write(r, 1, w['організація'], warn_red)
            ws5.write(r, 2, w['культура'],    warn_red)
            ws5.write(r, 3, w['послуга'],     warn_red)
            ws5.write(r, 4, w['опис'],        warn_red)
    else:
        ws5.write(0, 0, '✓ Всі організації та культури знайдені у прайс-листах.', ok_green)

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue(), unique_warnings, len(normal_rows), len(internal_rows)


# ============================================================
# CLI-запуск
# ============================================================
if __name__ == '__main__':
    with open(INPUT, 'rb') as f:
        data = f.read()

    out_bytes, unique_warnings, n_normal, n_internal = process(data)

    with open(OUTPUT, 'wb') as f:
        f.write(out_bytes)

    if unique_warnings:
        print(f'\n⚠  ПОПЕРЕДЖЕННЯ ({len(unique_warnings)} унікальних):')
        for w in unique_warnings:
            print(f'   [{w["тип"]}] {w["організація"]} | {w["культура"]} | {w["послуга"]}')
        print('   → Відкрий вкладку "Попередження" у файлі результату.')
    else:
        print('\n✓ Попереджень немає — всі ціни знайдено.')

    print(f'\n✅ Збережено: {OUTPUT}')
    print(f'   Основних рядків: {n_normal}, переміщень: {n_internal}')
