import pandas as pd
import re
import json
import os
from fuzzywuzzy import fuzz

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
    # 1. Убираем "Тестирование по модулю", "Практикум", "Кейс" и т.д.
    text = re.sub(r'(?i)тестирование по модулю|практическое задание по модулю|практикум|кейс|лекция №\d+|вступление|получено баллов', '', text)
    # 2. Убираем номера в начале строки (например "1. ", "10. ")
    text = re.sub(r'^\d+[\s\.)]+', '', text)
    # 3. Убираем спецсимволы и кавычки
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

def process_student_data(df_utp, df_stud, utp_name):
    config = load_config()
    
    # 1. Поиск правил для конкретного УТП
    utp_rules = {}
    target_utp = utp_name.strip().lower()
    for key, rules in config.items():
        if key.strip().lower() == target_utp:
            utp_rules = rules
            break
            
    # Предварительно нормализуем названия модулей в конфиге для быстрого поиска
    norm_utp_rules = {normalize_name(k): v for k, v in utp_rules.items()}

    # 2. Подготовка данных студента из Excel
    # Используем normalize_name, чтобы "Тестирование по модулю Х" стало просто "х"
    df_stud['CleanName'] = df_stud.iloc[:, 0].apply(normalize_name)
    df_stud['ScoreValue'] = df_stud.iloc[:, 1].apply(extract_score)
    
    # Группируем (если по модулю есть и тест, и лекция, берем максимальный балл)
    student_results = df_stud.dropna(subset=['ScoreValue']).groupby('CleanName')['ScoreValue'].max().to_dict()

    final_scores = []
    
    for _, row in df_utp.iterrows():
        db_module_original = row['Модули']
        norm_db_name = normalize_name(db_module_original)
        best_score = None
        
        # Приоритет
        rule_max = None
        
        # Ищем в нормализованном конфиге
        if norm_db_name in norm_utp_rules:
            rule_max = norm_utp_rules[norm_db_name]
        else:
            # Если прямого совпадения нет, пробуем нечеткий поиск по ключам конфига
            for cfg_norm_name, val in norm_utp_rules.items():
                if fuzz.token_sort_ratio(norm_db_name, cfg_norm_name) >= 90:
                    rule_max = val
                    break

        if rule_max is not None:
            if rule_max == 0:
                best_score = 100.0
                print(f"Применено правило 0% для: {db_module_original}")
            else:
                # Если макс. балл указан (например 20), ищем балл в ведомости
                found_raw = None
                highest_ratio = 0
                for stud_mod_name, score in student_results.items():
                    ratio = fuzz.token_sort_ratio(norm_db_name, stud_mod_name)
                    if ratio > highest_ratio and ratio >= 80:
                        highest_ratio = ratio
                        found_raw = score
                
                if found_raw is not None:
                    best_score = (found_raw / rule_max) * 100
                    print(f"Расчет по конфигу ({rule_max}): {db_module_original} -> {best_score}%")

        # Fallback
        if best_score is None:
            found_raw = None
            highest_ratio = 0
            for stud_mod_name, score in student_results.items():
                ratio = fuzz.token_sort_ratio(norm_db_name, stud_mod_name)
                if ratio > highest_ratio and ratio >= 80:
                    highest_ratio = ratio
                    found_raw = score
            
            if found_raw is not None:
                # Стандартная коррекция для 10-бальной шкалы (если нет в конфиге)
                best_score = found_raw * 10 if found_raw <= 10 else found_raw

        # Ограничиваем результат 100%
        if best_score is not None:
            best_score = min(best_score, 100.0)
            
        final_scores.append(best_score)
    
    result_df = df_utp.copy()
    result_df['Средний процент'] = final_scores
    result_df['Итоговая оценка'] = result_df.apply(get_grade_label, axis=1)
    
    return result_df