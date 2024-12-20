import streamlit as st
import pandas as pd
import io

# Заголовок застосунку
st.title("Аудит регіону")

# Інструкція для користувача
st.write("""
Цей застосунок дозволяє:
- Завантажити Excel-файл з даними про лоти (очікується таблиця з листом "Sheet1").
- Відфільтрувати дані за вказаними компаніями.
- Показати топ-20 областей за різними показниками.
- Завантажити результат у форматі XLSX з автопідбором ширини стовпців.
- **Додатково**: 
  - Створити окремий лист для кожної області з топ-20 компаніями за сумою виграних тендерів, 
    де три вказані компанії об'єднані в одну "АМЕТРІН ФК".
  - Додати колонку "2023" на основі даних з "Sheet2" з колонки "Сума лота".
  - Замінити назву колонки "Сума виграних тендерів" на "2024" у листах областей.
  - Додати колонку "динаміка" (у %) за формулою ((2024-2023)/2023)*100 з округленням до цілих та символом "%".
  - У першій строке регіональних листів показувати загальні суми за 2024 (з Sheet1) та 2023 (з Sheet2) по області.
  - Додати колонки "Доля 2024" та "Доля 2023", які рахуються відносно значень у першій строкі.
  - Колонку "динаміка" поставити перед долями.
  - Додати колонку "Прирост долі" як різницю між "Доля 2024" та "Доля 2023".
""")

uploaded_file = st.file_uploader("Завантажте Excel-файл з даними", type=["xlsx"])

original_target_companies = [
    'ТОВАРИСТВО З ОБМЕЖЕНОЮ ВІДПОВІДАЛЬНІСТЮ "ДІЯ ФАРМ"',
    'ТОВАРИСТВО З ОБМЕЖЕНОЮ ВІДПОВІДАЛЬНІСТЮ "АМЕТРІН ФК"',
    'ТОВАРИСТВО З ОБМЕЖЕНОЮ ВІДПОВІДАЛЬНІСТЮ "МОДЕРН-ФАРМ"'
]

grouped_company_name = "АМЕТРІН ФК"
companies_to_group = original_target_companies

if uploaded_file:
    try:
        data_sheet1 = pd.read_excel(uploaded_file, sheet_name='Sheet1')
        data_sheet2 = pd.read_excel(uploaded_file, sheet_name='Sheet2')

        required_columns_sheet1 = ['Сума лота', 'Переможець', 'Регіон організатора']
        missing_cols_sheet1 = [col for col in required_columns_sheet1 if col not in data_sheet1.columns]
        if missing_cols_sheet1:
            st.error(f"Відсутні необхідні стовпці в Sheet1: {', '.join(missing_cols_sheet1)}")
        else:
            st.write("Приклад завантажених даних Sheet1:", data_sheet1.head())

            # Обробка Sheet1
            data_sheet1['Сума лота'] = (
                data_sheet1['Сума лота']
                .astype(str)
                .str.replace('\u00a0', '', regex=True)
                .str.replace(',', '.', regex=True)
            )
            data_sheet1['Сума лота'] = pd.to_numeric(data_sheet1['Сума лота'], errors='coerce')

            data_sheet1['Переможець'] = (
                data_sheet1['Переможець']
                .astype(str)
                .str.split('|').str[0]
                .str.strip()
            )

            data_sheet1['Переможець'] = data_sheet1['Переможець'].apply(
                lambda x: grouped_company_name if x in companies_to_group else x
            )

            filtered_data = data_sheet1[
                (data_sheet1['Переможець'].isin([grouped_company_name])) &
                (data_sheet1['Регіон організатора'].notna()) &
                (data_sheet1['Регіон організатора'] != '-')
            ]

            total_region_summary = (
                data_sheet1
                .dropna(subset=['Регіон організатора'])
                [data_sheet1['Регіон організатора'] != '-']
                .groupby('Регіон організатора', as_index=False)
                .agg(total_sum=('Сума лота', 'sum'))
            )

            if filtered_data.empty:
                st.warning("За вказаними компаніями не знайдено жодного лота.")
            else:
                companies_region_summary = (
                    filtered_data
                    .groupby('Регіон організатора', as_index=False)
                    .agg(
                        sum_companies=('Сума лота', 'sum'),
                        count_companies=('Сума лота', 'count')
                    )
                )

                merged_summary = pd.merge(
                    total_region_summary,
                    companies_region_summary,
                    on='Регіон організатора',
                    how='inner'
                )

                # Переименовываем "Сума виграних тендерів" на "2024"
                merged_summary.rename(columns={
                    'Регіон організатора': 'Область',
                    'total_sum': '2024',
                    'sum_companies': 'Сума виграних тендерів компаній',
                    'count_companies': 'Кількість виграних тендерів'
                }, inplace=True)

                # доля = (Сума виграних тендерів компаній / 2024)*100
                merged_summary['доля'] = (merged_summary['Сума виграних тендерів компаній'] / merged_summary['2024']) * 100
                merged_summary['доля'] = merged_summary['доля'].round(2)

                merged_summary = merged_summary[['Область', 'Сума виграних тендерів компаній', 'Кількість виграних тендерів', '2024', 'доля']]

                st.subheader("Топ-20 областей")
                st.dataframe(merged_summary, use_container_width=True)

                required_columns_sheet2 = ['Переможець', 'Сума лота', 'Регіон організатора']
                missing_cols_sheet2 = [col for col in required_columns_sheet2 if col not in data_sheet2.columns]
                if missing_cols_sheet2:
                    st.error(f"Відсутні необхідні стовпці в Sheet2 для підрахунку 2023 по регіону: {', '.join(missing_cols_sheet2)}")
                else:
                    st.write("Приклад завантажених даних Sheet2:", data_sheet2.head())

                    data_sheet2['Переможець'] = (
                        data_sheet2['Переможець']
                        .astype(str)
                        .str.split('|').str[0]
                        .str.strip()
                    )
                    data_sheet2['Переможець'] = data_sheet2['Переможець'].apply(
                        lambda x: grouped_company_name if x in companies_to_group else x
                    )

                    data_sheet2['Сума лота'] = (
                        data_sheet2['Сума лота']
                        .astype(str)
                        .str.replace('\u00a0', '', regex=True)
                        .str.replace(',', '.', regex=True)
                    )
                    data_sheet2['Сума лота'] = pd.to_numeric(data_sheet2['Сума лота'], errors='coerce')

                    # Подсчет общей суммы за 2023 по регіонам
                    total_region_summary_2023 = (
                        data_sheet2
                        .dropna(subset=['Регіон організатора'])
                        [data_sheet2['Регіон організатора'] != '-']
                        .groupby('Регіон організатора', as_index=False)['Сума лота'].sum()
                    )
                    total_region_summary_2023.rename(columns={'Сума лота': 'total_sum_2023'}, inplace=True)
                    total_sum_dict_2023 = total_region_summary_2023.set_index('Регіон організатора')['total_sum_2023'].to_dict()

                    # Словник для загальної суми по області за 2024
                    total_sum_dict = merged_summary.set_index('Область')['2024'].to_dict()

                    output = io.BytesIO()
                    with pd.ExcelWriter(output, engine='openpyxl') as writer:
                        merged_summary.to_excel(writer, sheet_name='Data', index=False)
                        workbook = writer.book
                        worksheet = writer.sheets['Data']
                        for i, col in enumerate(merged_summary.columns, start=1):
                            max_length = max(
                                [len(str(val)) for val in merged_summary[col].values if val is not None] + [len(col)]
                            )
                            adjusted_width = max_length + 2
                            worksheet.column_dimensions[worksheet.cell(row=1, column=i).column_letter].width = adjusted_width

                        # Подсчёт сумм 2023 по компаниям для мапинга
                        sum_2023_by_company = data_sheet2.groupby('Переможець', as_index=False)['Сума лота'].sum()
                        sum_2023_dict = sum_2023_by_company.set_index('Переможець')['Сума лота'].to_dict()

                        for region in merged_summary['Область']:
                            region_data = data_sheet1[data_sheet1['Регіон організатора'] == region]

                            top_companies = (
                                region_data
                                .groupby('Переможець', as_index=False)
                                .agg(total_sum=('Сума лота', 'sum'))
                                .sort_values(by='total_sum', ascending=False)
                                .head(20)
                            )

                            top_companies.rename(columns={
                                'Переможець': 'Назва компанії',
                                'total_sum': '2024'
                            }, inplace=True)

                            top_companies['2023'] = top_companies['Назва компанії'].map(sum_2023_dict)

                            top_companies = top_companies[top_companies['Назва компанії'].notna() & (top_companies['Назва компанії'] != '-')]

                            # Формула динаміки: ((2024 - 2023)/2023)*100
                            def calc_dynamic(row):
                                if pd.isna(row['2023']) or row['2023'] == 0:
                                    return 0
                                return ((row['2024'] - row['2023']) / row['2023']) * 100

                            top_companies['динаміка'] = top_companies.apply(calc_dynamic, axis=1)
                            top_companies['динаміка'] = top_companies['динаміка'].round(0).astype(int).astype(str) + '%'

                            # Получаем общие суммы по региону
                            total_sum_region_2024 = total_sum_dict.get(region, 0)
                            total_sum_region_2023 = total_sum_dict_2023.get(region, 0)

                            # Формируем итоговую строку с общими суммами
                            summary_row = pd.DataFrame([{
                                'Назва компанії': 'ВСЬОГО',
                                '2024': total_sum_region_2024,
                                '2023': total_sum_region_2023,
                                'динаміка': ''
                            }])

                            # Добавляем строку сверху
                            top_companies = pd.concat([summary_row, top_companies], ignore_index=True)

                            # Рассчёт долей
                            def calc_share(val, total):
                                if total == 0 or pd.isna(val):
                                    return 0
                                return (val / total) * 100

                            total_2024 = top_companies.loc[0, '2024'] if pd.notna(top_companies.loc[0, '2024']) else 0
                            total_2023 = top_companies.loc[0, '2023'] if pd.notna(top_companies.loc[0, '2023']) else 0

                            # Считаем доли в числовом формате, чтобы потом вычислить прирост
                            dola_2024_numeric = top_companies['2024'].apply(lambda x: calc_share(x, total_2024))
                            dola_2023_numeric = top_companies['2023'].apply(lambda x: calc_share(x, total_2023))

                            # Прирост долі = Доля 2024 - Доля 2023
                            pririst_doli = dola_2024_numeric - dola_2023_numeric

                            # Преобразуем в проценты со знаком '%'
                            top_companies['Доля 2024'] = dola_2024_numeric.round(0).astype(int).astype(str) + '%'
                            top_companies['Доля 2023'] = dola_2023_numeric.round(0).astype(int).astype(str) + '%'
                            top_companies['Прирост долі'] = pririst_doli.round(0).astype(int).astype(str) + '%'

                            # Переставляем колонки:
                            # Сначала "динаміка", потом доли и прирост
                            # Итоговый порядок:
                            # Назва компанії | 2024 | 2023 | динаміка | Доля 2024 | Доля 2023 | Прирост долі
                            top_companies = top_companies[['Назва компанії', '2024', '2023', 'динаміка', 'Доля 2024', 'Доля 2023', 'Прирост долі']]

                            base_sheet_name = f"2024_{region}"
                            sheet_name = base_sheet_name[:31]

                            original_sheet_name = sheet_name
                            counter = 1
                            while sheet_name in writer.book.sheetnames:
                                suffix = f"_{counter}"
                                sheet_name = f"{original_sheet_name[:31 - len(suffix)]}{suffix}"
                                counter += 1
                                if counter > 100:
                                    st.error(f"Неможливо створити унікальне ім'я для листа області: {region}")
                                    break

                            top_companies.to_excel(writer, sheet_name=sheet_name, index=False)
                            worksheet = writer.sheets[sheet_name]
                            for i, col in enumerate(top_companies.columns, start=1):
                                max_length = max(
                                    [len(str(val)) for val in top_companies[col].values if val is not None] + [len(col)]
                                )
                                adjusted_width = max_length + 2
                                worksheet.column_dimensions[worksheet.cell(row=1, column=i).column_letter].width = adjusted_width

                    output.seek(0)
                    xlsx_data = output.read()

                    st.download_button(
                        label="Завантажити результати у форматі XLSX",
                        data=xlsx_data,
                        file_name="analysis_results.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )

    except Exception as e:
        st.error(f"Помилка обробки даних: {e}")
else:
    st.info("Будь ласка, завантажте Excel-файл для початку аналізу.")
