import pandas as pd
import numpy as np

def filter_skillspace_data(df):
    # 1. Формируем ФИО
    df['ФИО_Заголовок'] = (
        df['Фамилия'].fillna('') + ' ' + 
        df['Имя'].fillna('') + ' ' + 
        df['Отчество'].fillna('')
    ).str.strip()

    names_list = df['ФИО_Заголовок'].tolist()

    # 2. Транспонируем таблицу
    df_t = df.T.reset_index()

    # 3. Устанавливаем заголовки
    df_t.columns = ['Параметр'] + names_list

    # 4. Очистка от переносов строк
    for col in df_t.columns:
        df_t[col] = df_t[col].apply(
            lambda x: str(x).replace('\n', ' ').replace('\r', '').replace('\t', ' ').strip() 
            if pd.notnull(x) else np.nan
        )

    # 5. Фильтрация
    # Условие 1: содержит "Получено баллов"
    mask_include = df_t['Параметр'].str.contains('Получено баллов', case=False, na=False)
    # Условие 2: НЕ содержит "Статус"
    mask_exclude = ~df_t['Параметр'].str.contains('Статус', case=False, na=False)

    df_filtered = df_t[mask_include & mask_exclude].copy()

    # Удаляем техническую строку 'ФИО_Заголовок'
    df_filtered = df_filtered[df_filtered['Параметр'] != 'ФИО_Заголовок']

    # 6. Удаление пустых значений
    df_filtered[names_list] = df_filtered[names_list].replace(r'^\s*$', np.nan, regex=True)
    df_filtered[names_list] = df_filtered[names_list].replace(['nan', 'NaN', 'None'], np.nan)
    
    # Удаляем строки, где у всех пусто
    df_filtered = df_filtered.dropna(subset=names_list, how='all')

    return df_filtered