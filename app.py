import streamlit as st
import xlrd
import subprocess
from calc_pdv import process

def _git_version():
    try:
        return subprocess.check_output(
            ['git', 'rev-parse', '--short', 'HEAD'],
            stderr=subprocess.DEVNULL
        ).decode().strip()
    except Exception:
        return 'unknown'

REQUIRED_SHEETS = ('зберігання', 'сушка', 'приймання')
MAX_SIZE_MB = 50

st.set_page_config(
    page_title='Розрахунок ПДВ — МХП Елеватори',
    page_icon='🌾',
    layout='centered',
)

st.title('🌾 Розрахунок ПДВ по елеваторних послугах')
st.write('Завантажте щомісячний звіт у форматі `.xls`, натисніть **Розрахувати** — і скачайте результат.')


def validate(file_bytes: bytes, filename: str):
    """Повертає рядок з помилкою або None якщо все ок."""
    size_mb = len(file_bytes) / 1024 / 1024
    if size_mb > MAX_SIZE_MB:
        return f'Файл завеликий ({size_mb:.1f} МБ). Максимум — {MAX_SIZE_MB} МБ.'

    if filename.lower().endswith('.xlsx'):
        return (
            'Файл у форматі `.xlsx` — збережіть його як `.xls` (Файл → Зберегти як → '
            'Excel 97-2003 Workbook) і завантажте знову.'
        )

    try:
        rb = xlrd.open_workbook(file_contents=file_bytes)
    except xlrd.biffh.XLRDError:
        return 'Не вдалося відкрити файл. Переконайтесь що це файл Excel формату `.xls`.'
    except Exception:
        return 'Файл пошкоджений або має невідомий формат.'

    sheet_names = rb.sheet_names()
    missing = [s for s in REQUIRED_SHEETS if s not in sheet_names]
    if missing:
        return (
            f'У файлі відсутні обов\'язкові вкладки: **{", ".join(missing)}**.\n\n'
            f'Наявні вкладки: {", ".join(sheet_names)}.\n\n'
            'Переконайтесь що завантажено правильний файл щомісячного звіту.'
        )

    return None


uploaded = st.file_uploader('Оберіть файл звіту (.xls)', type=['xls'])

if uploaded is not None:
    st.info(f'Файл завантажено: **{uploaded.name}**')

    if st.button('Розрахувати ПДВ', type='primary'):
        file_bytes = uploaded.read()

        error = validate(file_bytes, uploaded.name)
        if error:
            st.error(error)
            st.stop()

        with st.spinner('Обробка...'):
            try:
                out_bytes, unique_warnings, n_normal, n_internal = process(file_bytes)
            except Exception as e:
                st.error(
                    f'Сталась неочікувана помилка під час обробки.\n\n'
                    f'Деталі: `{e}`\n\n'
                    'Перевірте структуру файлу або зверніться до розробника.'
                )
                st.stop()

        if n_normal == 0 and n_internal == 0:
            st.warning(
                'Файл оброблено, але не знайдено жодного рядку з даними. '
                'Можливо, структура файлу відрізняється від очікуваної.'
            )
        else:
            st.success(f'Готово! Основних рядків: **{n_normal}**, переміщень між філіями: **{n_internal}**')

        if unique_warnings:
            st.warning(
                f'⚠ Знайдено **{len(unique_warnings)}** позицій без ціни ПДВ (суми = 0). '
                'Деталі у вкладці "Попередження" файлу результату.'
            )
            with st.expander('Показати попередження'):
                for w in unique_warnings:
                    st.write(f'**[{w["тип"]}]** {w["організація"]} | {w["культура"]} | {w["послуга"]}')
        else:
            st.info('✓ Всі ціни знайдено, попереджень немає.')

        result_name = uploaded.name.replace('.xls', '_result.xls')
        st.download_button(
            label='⬇ Завантажити результат',
            data=out_bytes,
            file_name=result_name,
            mime='application/vnd.ms-excel',
        )

st.divider()
st.caption(f'МХП Елеватори · Розрахунок ПДВ · commit {_git_version()}')
