import pandas as pd
import re
import json
import os
import numpy as np

from fuzzywuzzy import fuzz

def filter_skillspace_data(df):
    # Тотальная очистка ФИО от любых случайных пробелов (концевых, двойных)
    df['ФИО_Заголовок'] = (
        df['Фамилия'].fillna('') + ' ' + 
        df['Имя'].fillna('') + ' ' + 
        df['Отчество'].fillna('')
    ).apply(lambda x: " ".join(str(x).split()))

    names_list = df['ФИО_Заголовок'].tolist()

    df_t = df.T.reset_index()

    df_t.columns = ['Параметр'] + names_list

    for col in df_t.columns:
        df_t[col] = df_t[col].apply(
            lambda x: str(x).replace('\n', ' ').replace('\r', '').replace('\t', ' ').strip() 
            if pd.notnull(x) else np.nan
        )

    mask_include = df_t['Параметр'].str.contains('Получено баллов', case=False, na=False)
    mask_exclude = ~df_t['Параметр'].str.contains('Статус', case=False, na=False)

    df_filtered = df_t[mask_include & mask_exclude].copy()

    df_filtered = df_filtered[df_filtered['Параметр'] != 'ФИО_Заголовок']

    df_filtered[names_list] = df_filtered[names_list].replace(r'^\s*$', np.nan, regex=True)
    df_filtered[names_list] = df_filtered[names_list].replace(['nan', 'NaN', 'None'], np.nan)
    
    df_filtered = df_filtered.dropna(subset=names_list, how='all')

    return df_filtered

def load_config():
    if os.path.exists('data/settings.json'):
        with open('data/settings.json', 'r', encoding='utf-8') as f:
            try:
                data = json.load(f)
                print("--- Конфигурация settings.json успешно загружена ---")
                return data
            except Exception as e:
                print(f"Ошибка чтения JSON: {e}")
                return {}
    print("!!! settings.json не найден в папке проекта !!!")
    return {}

def normalize_name(text):
    if not isinstance(text, str): return ""
    text = re.sub(r'(?i)тестирование по модулю|практическое задание по модулю|практикум|кейс|лекция №\d+|вступление|получено баллов', '', text)
    text = re.sub(r'^\d+[\s\.)]+', '', text)
    text = text.replace('"', '').replace('«', '').replace('»', '').replace('№', '')
    return text.strip().lower()

def extract_score(val):
    if pd.isna(val): return None
    res = re.findall(r'(\d+)', str(val))
    return float(res[0]) if res else None

def get_grade_label(row):
    score = row['Средний процент']
    form = str(row['Форма аттестации'])
    
    if score is None or pd.isna(score): return "-"

    if "Зачет" in form and "оценкой" not in form:
        return "Зачет" if score >= 60 else "Незачет"
    
    if score >= 85: return "5 (отл.)"
    if score >= 70: return "4 (хор.)"
    if score >= 50: return "3 (уд.)"
    return "2 (неуд.)"

def process_student_data(df_utp, df_stud, utp_name, student_name):
    config = load_config()
    
    utp_rules = {}
    target_utp = utp_name.strip().lower()
    for key, rules in config.items():
        if key.strip().lower() == target_utp:
            utp_rules = rules
            break
            
    norm_utp_rules = {normalize_name(k): v for k, v in utp_rules.items()}

    # Подготовка данных конкретного студента из Excel по его ФИО (а не по индексу колонки!)
    df_stud['CleanName'] = df_stud.iloc[:, 0].apply(normalize_name)
    df_stud['ScoreValue'] = df_stud[student_name].apply(extract_score)
    
    if 'Имя_Листа' not in df_stud.columns:
        df_stud['Имя_Листа'] = 'Основной лист'

    # Находим максимальный балл для каждого уникального CleanName, сохраняя имя листа
    df_valid_scores = df_stud.dropna(subset=['ScoreValue']).copy()
    df_sorted = df_valid_scores.sort_values(by='ScoreValue', ascending=False)
    df_grouped = df_sorted.drop_duplicates(subset=['CleanName'], keep='first')
    
    student_results = {
        row['CleanName']: (row['ScoreValue'], row['Имя_Листа']) 
        for _, row in df_grouped.iterrows()
    }

    final_scores = []
    
    print("\n" + "="*70)
    print(f"ЛОГ АНАЛИЗА МОДУЛЕЙ ДЛЯ СТУДЕНТА: {student_name}")
    print(f"УТП: {utp_name}")
    print("="*70)
    
    for _, row in df_utp.iterrows():
        db_module_original = row['Модули']
        norm_db_name = normalize_name(db_module_original)
        best_score = None
        used_sheet = None
        
        rule_max = None
        in_config = False
        
        if norm_db_name in norm_utp_rules:
            rule_max = norm_utp_rules[norm_db_name]
            in_config = True
        else:
            for cfg_norm_name, val in norm_utp_rules.items():
                if fuzz.token_sort_ratio(norm_db_name, cfg_norm_name) >= 90:
                    rule_max = val
                    in_config = True
                    break

        score_found_via_config = False
        if rule_max is not None:
            if rule_max == 0:
                best_score = 100.0
                print(f"[УСПЕХ] Модуль: \"{db_module_original}\" -> Применено правило 0% (автозачет)")
            else:
                found_raw = None
                found_sheet = None
                highest_ratio = 0
                for stud_mod_name, (score, sheet) in student_results.items():
                    ratio = fuzz.token_sort_ratio(norm_db_name, stud_mod_name)
                    if ratio > highest_ratio and ratio >= 80:
                        highest_ratio = ratio
                        found_raw = score
                        found_sheet = sheet
                
                if found_raw is not None:
                    best_score = (found_raw / rule_max) * 100
                    used_sheet = found_sheet
                    score_found_via_config = True
                    print(f"[УСПЕХ] Модуль: \"{db_module_original}\" -> Успешно рассчитан балл ({best_score:.1f}%) [Лист: {used_sheet}]")

        if best_score is None:
            found_raw = None
            found_sheet = None
            highest_ratio = 0
            for stud_mod_name, (score, sheet) in student_results.items():
                ratio = fuzz.token_sort_ratio(norm_db_name, stud_mod_name)
                if ratio > highest_ratio and ratio >= 80:
                    highest_ratio = ratio
                    found_raw = score
                    found_sheet = sheet
            
            if found_raw is not None:
                best_score = found_raw * 10 if found_raw <= 10 else found_raw
                used_sheet = found_sheet

        if best_score is None:
            if in_config:
                print(f"[ПРОБЛЕМА] Модуль: \"{db_module_original}\" (модуль есть в settings.json, но оценка студента в Excel-ведомости не найдена)")
            else:
                print(f"[ПРОБЛЕМА] Модуль: \"{db_module_original}\" (модуль полностью отсутствует в settings.json и в Excel-ведомости оценок по нему нет)")
        else:
            if not score_found_via_config:
                print(f"[ВНИМАНИЕ] Модуль: \"{db_module_original}\" (модуль отсутствует в settings.json, применен автоматический поиск баллов. Результат: {best_score:.1f}%) [Лист: {used_sheet}]")

        if best_score is not None:
            best_score = min(best_score, 100.0)
            
        final_scores.append(best_score)
    
    print("="*70 + "\n")
    
    result_df = df_utp.copy()
    result_df['Средний процент'] = final_scores
    result_df['Итоговая оценка'] = result_df.apply(get_grade_label, axis=1)
    
    return result_df