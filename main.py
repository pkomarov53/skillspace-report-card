import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
import sqlite3
import os
import re
from fuzzywuzzy import fuzz

# Логические модули
from data.connection import init_db, DB_NAME
from logic.parser import extract_utp_from_pdf, extract_grades_from_pdf
from logic.report import save_report_to_pdf
from logic.processor import process_student_data
from logic.courseID import course_list
from logic.xlsx_parse import filter_skillspace_data

class App:
    def __init__(self, root):
        self.root = root
        self.root.title("Обработчик ведомостей Skillspace")
        self.root.geometry("700x350")
        
        init_db()
        self.excel_path = None
        self.recredits_path = None
        self.setup_ui()

    def setup_ui(self):
        
        # --- Секция БД ---
        frame_db = tk.LabelFrame(self.root, text=" 1. Настройка Учебных Планов (УТП) ", padx=10, pady=10)
        frame_db.pack(fill="x", padx=10, pady=10)
        
        tk.Button(frame_db, text="Загрузить новое УТП из PDF", command=self.import_pdf).pack(side="left")
        
        tk.Label(self.root, text="Выберите Учебный План:").pack(pady=(10,0))
        self.utp_combo = ttk.Combobox(self.root, values=self.get_utp_list(), state="readonly", width=80)
        self.utp_combo.pack(pady=5)
        if self.utp_combo['values']:
            self.utp_combo.current(0)

        # --- Секция данных студента ---
        frame_stud = tk.LabelFrame(self.root, text=" 2. Данные студента и перезачеты ", padx=10, pady=10)
        frame_stud.pack(fill="x", padx=10, pady=10)

        btn_excel = tk.Button(frame_stud, text="Выбрать ведомость Skillspace (.xlsx)", command=self.load_excel, width=35)
        btn_excel.grid(row=0, column=0, padx=5, pady=5)
        self.lbl_file = tk.Label(frame_stud, text="Файл не выбран", fg="gray")
        self.lbl_file.grid(row=0, column=1, sticky="w")

        btn_re = tk.Button(frame_stud, text="Добавить перезачеты (PDF ведомость)", command=self.load_recredits, width=35, bg="#fcf8e3")
        btn_re.grid(row=1, column=0, padx=5, pady=5)
        self.lbl_re = tk.Label(frame_stud, text="Необязательно", fg="gray")
        self.lbl_re.grid(row=1, column=1, sticky="w")

        # --- Кнопка запуска ---
        tk.Button(self.root, text="СФОРМИРОВАТЬ ИТОГОВЫЙ ТАБЕЛЬ", 
                  command=self.process_all, bg="#4CAF50", fg="white", 
                  font=("Arial", 11, "bold"), pady=15).pack(pady=20, fill="x", padx=10)

    def get_utp_list(self):
        try:
            conn = sqlite3.connect(DB_NAME)
            cursor = conn.cursor()
            cursor.execute('SELECT DISTINCT utp_name FROM utp_modules')
            names = [r[0] for r in cursor.fetchall()]
            conn.close()
            return names
        except: return []

    def import_pdf(self):
        path = filedialog.askopenfilename(filetypes=[("PDF files", "*.pdf")])
        if path:
            try:
                name, df = extract_utp_from_pdf(path)
                conn = sqlite3.connect(DB_NAME)
                df.to_sql('utp_modules', conn, if_exists='append', index=False)
                conn.commit()
                conn.close()
                self.utp_combo['values'] = self.get_utp_list()
                messagebox.showinfo("Успех", f"УТП '{name}' успешно добавлено!")
            except Exception as e:
                messagebox.showerror("Ошибка", f"Не удалось прочитать PDF: {e}")

    def load_excel(self):
        path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
        if path:
            raw_utp_name = self.utp_combo.get()
            if not raw_utp_name:
                messagebox.showwarning("Внимание", "Сначала выберите УТП!")
                return

            clean_utp_name = " ".join(raw_utp_name.split())
            course_id = None
            for key, val in course_list.items():
                if " ".join(key.split()) == clean_utp_name:
                    course_id = val
                    break
            
            if not course_id:
                messagebox.showwarning("Внимание", f"ID не найден для: {raw_utp_name}")
                return

            try:
                xl = pd.ExcelFile(path)
                target_sheet = next((s for s in xl.sheet_names if str(course_id) in s), None)
                
                if not target_sheet:
                    messagebox.showwarning("Внимание", f"Лист с ID {course_id} не найден!")
                    return

                # --- ТУТ ПРОИСХОДИТ ОБРАБОТКА ---
                # 1. Читаем сырые данные
                df_raw = pd.read_excel(path, sheet_name=target_sheet)
                
                # 2. Прогоняем через фильтр из xlsx_parse.py
                df_processed = filter_skillspace_data(df_raw)
                
                # 3. Сохраняем результат в промежуточный файл для вашей проверки
                temp_filename = "data/temp_processed_check.xlsx"
                df_processed.to_excel(temp_filename, index=False)
                
                # Сохраняем данные в памяти программы для дальнейшей работы
                self.excel_path = path
                self.target_sheet = target_sheet
                self.current_df = df_processed # Сохраняем уже обработанный DF
                
                self.lbl_file.config(text=f"Выбрано: {target_sheet}", fg="green")
                messagebox.showinfo("Готово", f"Лист обработан и очищен.\nРезультат в {temp_filename}")

            except Exception as e:
                messagebox.showerror("Ошибка", f"Ошибка обработки: {e}")

    def load_recredits(self):
        path = filedialog.askopenfilename(filetypes=[("PDF files", "*.pdf")])
        if path:
            self.recredits_path = path
            self.lbl_re.config(text=f"Прикреплен: {os.path.basename(path)}", fg="green")

    def clean_text(self, text):
        if not text: return ""
        text = str(text).lower()
        text = re.sub(r'^(модуль|тема|раздел|лекция)\s*\d+[\s\.]*', '', text)
        text = re.sub(r'[^а-яёa-z\s]', ' ', text)
        return " ".join(text.split())

    def get_final_grade_text(self, percent, control_form):
        form = str(control_form).lower()
        
        # Если форма контроля подразумевает оценку (Экзамен или Зачет с оценкой)
        if "экзамен" in form or "оценкой" in form or "дифф" in form:
            if percent >= 85: return "5 (отл.)"
            if percent >= 70: return "4 (хор.)"
            if percent >= 50: return "3 (уд.)"
            return "2 (неуд.)"
        
        # Если это обычный зачет
        else:
            return "Зачет" if percent >= 50 else "Незачет"

    def process_all(self):
        utp_name = self.utp_combo.get()
        
        # Проверяем, что УТП выбрано и данные уже загружены и обработаны в load_excel
        if not utp_name or not hasattr(self, 'current_df') or self.current_df is None:
            messagebox.showwarning("Внимание", "Сначала выберите УТП и загрузите файл Excel (кнопка 'Выбрать ведомость')!")
            return

        try:
            # 1. Получаем список студентов из нашей обработанной таблицы
            # Напомню: первая колонка там 'Параметр', остальные — ФИО студентов
            all_students = [col for col in self.current_df.columns if col != 'Параметр']
            
            if not all_students:
                messagebox.showwarning("Ошибка", "Студенты в обработанном листе не найдены!")
                return

            # Если ваша программа рассчитана на создание одного табеля за раз,
            # берем первого студента из списка. 
            student_info = all_students[0] 

            # 2. УТП из базы (остается без изменений)
            conn = sqlite3.connect(DB_NAME)
            df_utp = pd.read_sql(
                f"SELECT module_name as 'Модули', hours as 'Количество часов', control_form as 'Форма аттестации' "
                f"FROM utp_modules WHERE utp_name='{utp_name}'", conn)
            conn.close()

            # 3. Основной расчет (Skillspace)
            # Теперь мы передаем уже отфильтрованный и транспонированный self.current_df
            # ВАЖНО: Функция process_student_data внутри processor.py должна уметь 
            # работать с этим форматом (где параметры в строках).
            report = process_student_data(df_utp, self.current_df, utp_name)

            # 4. Перезачеты (Ваша логика остается нетронутой)
            if self.recredits_path:
                df_re = extract_grades_from_pdf(self.recredits_path)
                
                for idx, row in report.iterrows():
                    cur_val = row['Средний процент']
                    current_score = float(cur_val) if pd.notna(cur_val) and str(cur_val).replace('.','').isdigit() else 0.0
                    
                    utp_mod_clean = self.clean_text(row['Модули'])
                    best_match_ratio = 0
                    best_re_row = None
                    
                    for _, re_row in df_re.iterrows():
                        re_mod_clean = self.clean_text(re_row['module_name'])
                        ratio = fuzz.token_set_ratio(utp_mod_clean, re_mod_clean)
                        if ratio > best_match_ratio:
                            best_match_ratio = ratio
                            best_re_row = re_row
                    
                    if best_match_ratio > 80 and best_re_row is not None:
                        new_score = float(best_re_row['re_score'])
                        if new_score > current_score:
                            report.at[idx, 'Средний процент'] = new_score
                            grade_text = self.get_final_grade_text(new_score, row['Форма аттестации'])
                            report.at[idx, 'Итоговая оценка'] = f"{grade_text} (перезачет)"

            # 5. Сохранение результата (Ваша логика остается нетронутой)
            clean_name = re.sub(r'[\\/*?:"<>|]', "", student_info)
            save_p = filedialog.asksaveasfilename(
                initialfile=f"Табель_{clean_name}.pdf",
                defaultextension=".pdf", filetypes=[("PDF", "*.pdf")]
            )

            if save_p:
                save_report_to_pdf(report, save_p, student_info)
                os.startfile(save_p)
                messagebox.showinfo("Успех", f"Табель для {student_info} готов.")

        except Exception as e:
            messagebox.showerror("Ошибка", f"Произошел сбой при генерации: {e}")

if __name__ == "__main__":
    root = tk.Tk()
    root.resizable(width=False, height=False)
    app = App(root)
    root.mainloop()