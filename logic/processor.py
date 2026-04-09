import pandas as pd
import re
from fuzzywuzzy import fuzz

def normalize_simple(text):
    if not isinstance(text, str): return ""
    text = re.sub(r'по модулю', '', text, flags=re.IGNORECASE)
    return text.replace('"', '').replace('«', '').replace('»', '').strip().lower()

def extract_score(val):
    res = re.findall(r'(\d+)', str(val))
    return float(res[0]) if res else None

def get_grade_label(row):
    score = row['Средний процент']
    module_name = str(row['Модули']).lower()
    form = row['Форма аттестации']
    
    if pd.isna(score): return "-"

    if "защита итогового" in module_name:
        if score >= 45: return "5 (отл.)"
        if score >= 35: return "4 (хор.)"
        if score >= 25: return "3 (уд.)"
        return "2 (неуд.)"

    if "Зачет" in form and "оценкой" not in form:
        return "Зачет" if score >= 60 else "Незачет"
    else:
        if score >= 85: return "5 (отл.)"
        elif score >= 70: return "4 (хор.)"
        elif score >= 50: return "3 (уд.)"
        else: return "2 (неуд.)"

def process_student_data(df_utp, df_stud):
    # Фильтрация и очистка Excel
    mask = df_stud.iloc[:, 0].str.contains('модулю|практикум|дз|итоговое|защита итогового', case=False, na=False)
    df_filtered = df_stud[mask].copy()
    
    df_filtered['CleanName'] = df_filtered.iloc[:, 0].apply(normalize_simple)
    df_filtered['Score'] = df_filtered.iloc[:, 1].apply(extract_score)
    df_filtered = df_filtered.dropna(subset=['Score'])
    
    df_grouped = df_filtered.groupby('CleanName')['Score'].mean().reset_index()

    # Fuzzy matching
    scores_list = []
    for _, row_utp in df_utp.iterrows():
        name_to_find = row_utp['Модули'].lower().strip()
        best_score = None
        highest_ratio = 0
        
        for _, row_excel in df_grouped.iterrows():
            current_ratio = fuzz.token_sort_ratio(name_to_find, row_excel['CleanName'])
            if current_ratio > highest_ratio and current_ratio >= 70:
                highest_ratio = current_ratio
                best_score = row_excel['Score']
        
        # Коррекция шкалы 1-10 в 1-100
        if best_score is not None and best_score <= 10:
            best_score *= 10
            
        scores_list.append(best_score)
    
    result_df = df_utp.copy()
    result_df['Средний процент'] = scores_list
    result_df['Итоговая оценка'] = result_df.apply(get_grade_label, axis=1)
    
    return result_df