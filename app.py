import streamlit as st
from calc_pdv import process

st.set_page_config(
    page_title='Розрахунок ПДВ — МХП Елеватори',
    page_icon='🌾',
    layout='centered',
)

st.title('🌾 Розрахунок ПДВ по елеваторних послугах')
st.write('Завантажте щомісячний звіт у форматі `.xls`, натисніть **Розрахувати** — і скачайте результат.')

uploaded = st.file_uploader('Оберіть файл звіту (.xls)', type=['xls'])

if uploaded is not None:
    st.info(f'Файл завантажено: **{uploaded.name}**')

    if st.button('Розрахувати ПДВ', type='primary'):
        with st.spinner('Обробка...'):
            try:
                input_bytes = uploaded.read()
                out_bytes, unique_warnings, n_normal, n_internal = process(input_bytes)
            except Exception as e:
                st.error(f'Помилка обробки: {e}')
                st.stop()

        st.success(f'Готово! Основних рядків: **{n_normal}**, переміщень між філіями: **{n_internal}**')

        if unique_warnings:
            st.warning(f'⚠ Знайдено **{len(unique_warnings)}** позицій без ціни ПДВ (суми = 0). Деталі у вкладці "Попередження" файлу результату.')
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
