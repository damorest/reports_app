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
        # Організації поза списком МХП (сторонні елеватори) теж рахуються
        # за цим прайсом — сигналізуємо про це, щоб рішення було свідомим.
        if warn_collector is not None and not get_org_key(org):
            warn_collector.append({
                'тип': 'Організація поза прайсом',
                'організація': str(org).strip(),
                'культура': str(nom).strip(),
                'послуга': service,
                'опис': f'Організації "{str(org).strip()}" немає у прайс-листах МХП. '
                        f'Застосовано загальний прайс ранніх зернових (TE-03). Перевірте суму.',
            })
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

def _cv(sheet, row, col, default=0):
    """Безпечне читання клітинки — повертає default якщо рядок коротший."""
    return sheet.cell_value(row, col) if len(sheet.row(row)) > col else default

def _inn(val):
    """Нормалізує ІНН: число 361879902121.0 → рядок '361879902121'. Порожнє → ''."""
    if not val:
        return ''
    s = str(val).strip()
    if s.endswith('.0'):
        s = s[:-2]
    return s if s and s != '0' else ''

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
    # Сторонні елеватори (Урожай НВФ тощо) НЕ фільтруються: їхні рядки
    # потрапляють у розрахунок і, за відсутності прайсу, дають суму 0
    # плюс попередження «Невідома організація» — щоб нічого не губилось мовчки.
    skip = {'организация', 'итого', 'дата', 'отбор:', 'сформирован:',
            'послуги зберігання', 'номенклатура', 'отправительполучатель',
            'лабораторный', 'лабораторний', 'аналізи', 'вище норми', 'нижче норми',
            'анализ'}
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
# ДИНАМІЧНИЙ ПОШУК КОЛОНОК ЗА ЗАГОЛОВКАМИ
# Вигрузка з 1С «пливе» від місяця до місяця (змінюється ширина
# об'єднаних комірок), тому індекси колонок НЕ хардкодяться,
# а щоразу шукаються за назвами у рядку заголовків.
# ============================================================

def _hnorm(v):
    """Нормалізує текст заголовка для порівняння."""
    if not isinstance(v, str):
        return ''
    return ' '.join(v.lower().split())


def _header_map(sheet, row):
    """{нормалізований_заголовок: [індекси колонок]} для вказаного рядка."""
    m = defaultdict(list)
    for c in range(sheet.ncols):
        h = _hnorm(_cv(sheet, row, c, ''))
        if h:
            m[h].append(c)
    return m


def _find_header_row(sheet, required, max_scan=20):
    """Номер рядка заголовків — перший, у якому присутні всі назви з `required`."""
    for r in range(min(max_scan, sheet.nrows)):
        m = _header_map(sheet, r)
        if all(name in m for name in required):
            return r
    raise ValueError(
        'Вкладка "%s": не знайдено рядок заголовків (очікували колонки: %s). '
        'Схоже, змінився формат вигрузки.' % (sheet.name, ', '.join(required))
    )


def _col(hmap, name, sheet_name, after=None):
    """Індекс колонки за назвою; `after` — брати перше входження правіше цієї колонки."""
    idxs = hmap.get(name, [])
    if after is not None:
        idxs = [i for i in idxs if i > after]
    if not idxs:
        raise ValueError(
            'Вкладка "%s": не знайдено колонку "%s". Схоже, змінився формат вигрузки.'
            % (sheet_name, name)
        )
    return idxs[0]


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

    summary          = defaultdict(lambda: defaultdict(float))
    pair_is_internal = {}   # (org, kontrag) -> bool на основі ІНН
    warnings         = []

    # ----------------------------------------------------------
    # 1. ЗБЕРІГАННЯ
    # Колонки визначаються за заголовками (див. _find_header_row).
    # Результат дописується у дві нові колонки в кінці вкладки.
    # ----------------------------------------------------------
    rs = rb.sheet_by_name('зберігання')
    ws = wb.get_sheet(rb.sheet_names().index('зберігання'))

    h_row     = _find_header_row(rs, ('организация', 'контрагент', 'культура',
                                      'кількість зберігання т/д'))
    hm        = _header_map(rs, h_row)
    c_org     = _col(hm, 'организация', 'зберігання')
    c_kontrag = _col(hm, 'контрагент',  'зберігання')
    c_kultura = _col(hm, 'культура',    'зберігання')
    c_inn_o   = _col(hm, 'инн', 'зберігання', after=c_org)
    c_inn_k   = _col(hm, 'инн', 'зберігання', after=c_kontrag)
    c_qty     = _col(hm, 'кількість зберігання т/д', 'зберігання')
    c_cina, c_suma = rs.ncols, rs.ncols + 1

    ws.write(h_row, c_cina, 'ціна ПДВ')
    ws.write(h_row, c_suma, 'Сума ПДВ')

    for i in range(h_row + 1, rs.nrows):
        ct  = rs.cell_type(i, c_org)
        org = rs.cell_value(i, c_org)
        if not is_data_row(ct, org):
            continue
        inn_o   = _inn(_cv(rs, i, c_inn_o,   ''))
        kontrag = _cv(rs, i, c_kontrag, '')
        nom     = _cv(rs, i, c_kultura, '')
        inn_k   = _inn(_cv(rs, i, c_inn_k,   ''))
        delta   = _cv(rs, i, c_qty)
        if not isinstance(delta, (int, float)) or delta == 0 or not is_valid_nom(nom):
            ws.write(i, c_cina, '')
            ws.write(i, c_suma, '')
            continue
        cina = get_vat(org, nom, 'зберігання', warnings)
        suma = round(delta * cina, 6)
        ws.write(i, c_cina, cina)
        ws.write(i, c_suma, suma)
        if cina > 0:
            key = (str(org).strip(), str(kontrag).strip())
            summary[key]['зберігання'] += suma
            pair_is_internal.setdefault(key, bool(inn_o and inn_k and inn_o == inn_k))

    # ----------------------------------------------------------
    # 2. СУШКА
    # Заголовки у двох рядках: h_row (Организация/Контрагент/…)
    # та h_row+1 (Количество_очистка_т% / Количество_сушка_т%).
    # ----------------------------------------------------------
    rs2 = rb.sheet_by_name('сушка')
    ws2 = wb.get_sheet(rb.sheet_names().index('сушка'))

    h_row2     = _find_header_row(rs2, ('организация', 'контрагент', 'реальная номенклатура'))
    hm2        = _header_map(rs2, h_row2)
    hm2b       = _header_map(rs2, h_row2 + 1)
    c2_org     = _col(hm2, 'организация', 'сушка')
    c2_kontrag = _col(hm2, 'контрагент',  'сушка')
    c2_nom     = _col(hm2, 'реальная номенклатура', 'сушка')
    c2_inn_o   = _col(hm2, 'инн', 'сушка', after=c2_org)
    c2_inn_k   = _col(hm2, 'инн', 'сушка', after=c2_kontrag)
    c2_och     = _col(hm2b, 'количество_очистка_т%', 'сушка')
    c2_sus     = _col(hm2b, 'количество_сушка_т%',   'сушка')
    c2_cina_och, c2_suma_och = rs2.ncols,     rs2.ncols + 1
    c2_cina_sus, c2_suma_sus = rs2.ncols + 2, rs2.ncols + 3

    ws2.write(h_row2, c2_cina_och, 'ціна ПДВ очистка')
    ws2.write(h_row2, c2_suma_och, 'Сума ПДВ очистка')
    ws2.write(h_row2, c2_cina_sus, 'ціна ПДВ сушка')
    ws2.write(h_row2, c2_suma_sus, 'Сума ПДВ сушка')

    for i in range(h_row2 + 2, rs2.nrows):
        ct  = rs2.cell_type(i, c2_org)
        org = rs2.cell_value(i, c2_org)
        if not is_data_row(ct, org):
            continue
        inn_o   = _inn(_cv(rs2, i, c2_inn_o,   ''))
        kontrag = _cv(rs2, i, c2_kontrag, '')
        nom     = _cv(rs2, i, c2_nom,     '')
        inn_k   = _inn(_cv(rs2, i, c2_inn_k,   ''))
        d_och_v = _cv(rs2, i, c2_och)
        d_sus_v = _cv(rs2, i, c2_sus)
        d_och = float(d_och_v) / 1000.0 if isinstance(d_och_v, (int, float)) and d_och_v else 0.0
        d_sus = float(d_sus_v) / 1000.0 if isinstance(d_sus_v, (int, float)) and d_sus_v else 0.0
        if not (d_och or d_sus) or not is_valid_nom(nom):
            continue
        c_och_v = get_vat(org, nom, 'очистка', warnings)
        c_sus_v = get_vat(org, nom, 'сушка',   warnings)
        s_och = round(d_och * c_och_v, 6)
        s_sus = round(d_sus * c_sus_v, 6)
        if d_och:
            ws2.write(i, c2_cina_och, c_och_v)
            ws2.write(i, c2_suma_och, s_och)
        if d_sus:
            ws2.write(i, c2_cina_sus, c_sus_v)
            ws2.write(i, c2_suma_sus, s_sus)
        total_sushka = s_och + s_sus
        if total_sushka > 0:
            key = (str(org).strip(), str(kontrag).strip())
            summary[key]['сушка'] += total_sushka
            pair_is_internal.setdefault(key, bool(inn_o and inn_k and inn_o == inn_k))

    # ----------------------------------------------------------
    # 3. ПРИЙМАННЯ
    # ----------------------------------------------------------
    rs3 = rb.sheet_by_name('приймання')
    ws3 = wb.get_sheet(rb.sheet_names().index('приймання'))

    h_row3     = _find_header_row(rs3, ('организация', 'отправительполучатель',
                                        'культура', 'физвес'))
    hm3        = _header_map(rs3, h_row3)
    c3_org     = _col(hm3, 'организация', 'приймання')
    c3_kontrag = _col(hm3, 'отправительполучатель', 'приймання')
    c3_kultura = _col(hm3, 'культура', 'приймання')
    c3_inn_o   = _col(hm3, 'инн', 'приймання', after=c3_org)
    c3_inn_k   = _col(hm3, 'инн', 'приймання', after=c3_kontrag)
    c3_fiz     = _col(hm3, 'физвес', 'приймання')
    c3_cina, c3_suma = rs3.ncols, rs3.ncols + 1

    ws3.write(h_row3, c3_cina, 'цінаПДВ')
    ws3.write(h_row3, c3_suma, 'Сума ПДВ')

    for i in range(h_row3 + 1, rs3.nrows):
        ct  = rs3.cell_type(i, c3_org)
        org = rs3.cell_value(i, c3_org)
        if not is_data_row(ct, org):
            continue
        inn_o   = _inn(_cv(rs3, i, c3_inn_o,   ''))
        kontrag = _cv(rs3, i, c3_kontrag, '')
        nom     = _cv(rs3, i, c3_kultura, '')
        inn_k   = _inn(_cv(rs3, i, c3_inn_k,   ''))
        fizves  = _cv(rs3, i, c3_fiz)
        if not isinstance(fizves, (int, float)) or fizves == 0 or not is_valid_nom(nom):
            ws3.write(i, c3_cina, '')
            ws3.write(i, c3_suma, '')
            continue
        cina = get_vat(org, nom, 'приймання', warnings)
        suma = round(float(fizves) * cina, 6)
        ws3.write(i, c3_cina, cina)
        ws3.write(i, c3_suma, suma)
        if cina > 0:
            key = (str(org).strip(), str(kontrag).strip())
            summary[key]['приймання'] += suma
            pair_is_internal.setdefault(key, bool(inn_o and inn_k and inn_o == inn_k))

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
        is_int = pair_is_internal.get((org, kont), False)
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
        ws5.write(0, 0, f'⚠ Знайдено {len(unique_warnings)} позицій, що потребують уваги '
                        f'(немає ціни → сума 0, або застосовано загальний прайс)', warn_orange)
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
