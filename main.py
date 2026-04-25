import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
import sqlite3
import os
import re
from pathlib import Path
from fuzzywuzzy import fuzz

# Логические модули
from data.connection import init_db, DB_NAME
from logic.parser import extract_utp_from_pdf, extract_grades_from_pdf
from logic.report import save_report_to_pdf
from logic.processor import process_student_data
from logic.courseID import course_list
from logic.xlsx_parse import filter_skillspace_data

class App:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("Обработчик ведомостей Skillspace")
        self.root.geometry("700x350")
        self.root.resizable(width=False, height=False)
        
        init_db()
        
        self.excel_path: str | None = None
        self.recredits_path: str | None = None
        self.target_sheet: str | None = None
        self.current_df: pd.DataFrame | None = None
        
        self.setup_ui()

    def setup_ui(self):
        frame_db = tk.LabelFrame(self.root, text=" 1. Настройка Учебных Планов (УТП) ", padx=10, pady=10)
        frame_db.pack(fill="x", padx=10, pady=10)
        
        tk.Button(frame_db, text="Загрузить новое УТП из PDF", command=self.import_pdf).pack(side="left")
        
        tk.Label(self.root, text="Выберите Учебный План:").pack(pady=(10,0))
        self.utp_combo = ttk.Combobox(self.root, values=self.get_utp_list(), state="readonly", width=80)
        self.utp_combo.pack(pady=5)
        if self.utp_combo['values']:
            self.utp_combo.current(0)

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

        tk.Button(self.root, text="СФОРМИРОВАТЬ ИТОГОВЫЙ ТАБЕЛЬ", 
                  command=self.process_all, bg="#4CAF50", fg="white", 
                  font=("Arial", 11, "bold"), pady=15).pack(pady=20, fill="x", padx=10)

    def get_utp_list(self) -> list[str]:
        try:
            with sqlite3.connect(DB_NAME) as conn:
                cursor = conn.cursor()
                cursor.execute('SELECT DISTINCT utp_name FROM utp_modules')
                return [r[0] for r in cursor.fetchall()]
        except sqlite3.Error as e:
            print(f"Database error: {e}")
            return []

    def import_pdf(self):
        path = filedialog.askopenfilename(filetypes=[("PDF files", "*.pdf")])
        if not path:
            return

        try:
            name, df = extract_utp_from_pdf(path)
            with sqlite3.connect(DB_NAME) as conn:
                df.to_sql('utp_modules', conn, if_exists='append', index=False)
            
            self.utp_combo['values'] = self.get_utp_list()
            self.utp_combo.set(name) # Автоматически выбираем загруженный план
            messagebox.showinfo("Успех", f"УТП '{name}' успешно добавлено!")
        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось прочитать PDF: {e}")

    def load_excel(self):
        path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
        if not path:
            return

        raw_utp_name = self.utp_combo.get()
        if not raw_utp_name:
            messagebox.showwarning("Внимание", "Сначала выберите УТП!")
            return

        clean_utp_name = " ".join(raw_utp_name.split())
        
        normalized_course_list = {" ".join(k.split()): v for k, v in course_list.items()}
        course_id = normalized_course_list.get(clean_utp_name)
        
        if not course_id:
            messagebox.showwarning("Внимание", f"ID не найден для: {raw_utp_name}")
            return

        try:
            xl = pd.ExcelFile(path)
            target_sheet = next((s for s in xl.sheet_names if str(course_id) in s), None)
            
            if not target_sheet:
                messagebox.showwarning("Внимание", f"Лист с ID {course_id} не найден!")
                return

            # Чтение и обработка
            df_raw = pd.read_excel(path, sheet_name=target_sheet)
            df_processed = filter_skillspace_data(df_raw)
            
            # Сохранение промежуточного файла с использованием pathlib
            temp_filename = Path("data/temp_processed_check.xlsx")
            temp_filename.parent.mkdir(parents=True, exist_ok=True) # Защита, если папки data нет
            df_processed.to_excel(temp_filename, index=False)
            
            # Обновление состояния приложения
            self.excel_path = path
            self.target_sheet = target_sheet
            self.current_df = df_processed
            
            self.lbl_file.config(text=f"Выбрано: {target_sheet}", fg="green")
            messagebox.showinfo("Готово", f"Лист обработан.\nРезультат сохранен в {temp_filename}")

        except Exception as e:
            messagebox.showerror("Ошибка", f"Ошибка обработки Excel: {e}")
            self.current_df = None

    def load_recredits(self):
        path = filedialog.askopenfilename(filetypes=[("PDF files", "*.pdf")])
        if path:
            self.recredits_path = path
            self.lbl_re.config(text=f"Прикреплен: {Path(path).name}", fg="green")

    @staticmethod
    def clean_text(text: str) -> str:
        if not text or pd.isna(text):
            return ""
        text = str(text).lower()
        text = re.sub(r'^(модуль|тема|раздел|лекция)\s*\d+[\s\.]*', '', text)
        text = re.sub(r'[^а-яёa-z\s]', ' ', text)
        return " ".join(text.split())

    @staticmethod
    def get_final_grade_text(percent: float, control_form: str) -> str:
        form = str(control_form).lower()
        
        if any(keyword in form for keyword in ["экзамен", "оценкой", "дифф"]):
            if percent >= 85: return "5 (отл.)"
            if percent >= 70: return "4 (хор.)"
            if percent >= 50: return "3 (уд.)"
            return "2 (неуд.)"
            
        return "Зачет" if percent >= 50 else "Незачет"

    def process_all(self):
        utp_name = self.utp_combo.get()
        
        if not utp_name or self.current_df is None:
            messagebox.showwarning("Внимание", "Выберите УТП и загрузите файл Excel!")
            return

        try:
            all_students = [col for col in self.current_df.columns if col != 'Параметр']
            if not all_students:
                messagebox.showwarning("Ошибка", "Студенты в обработанном листе не найдены!")
                return

            student_info = all_students[0] 

            # Запрашиваем УТП
            query = """
                SELECT module_name as 'Модули', 
                       hours as 'Количество часов', 
                       control_form as 'Форма аттестации' 
                FROM utp_modules 
                WHERE utp_name=?
            """
            with sqlite3.connect(DB_NAME) as conn:
                df_utp = pd.read_sql(query, conn, params=(utp_name,))

            report = process_student_data(df_utp, self.current_df, utp_name)

            # Перезачеты
            if self.recredits_path:
                df_re = extract_grades_from_pdf(self.recredits_path)
                
                df_re['clean_module'] = df_re['module_name'].apply(self.clean_text)
                
                for idx, row in report.iterrows():
                    cur_val = row['Средний процент']
                    
                    try:
                        current_score = float(cur_val) if pd.notna(cur_val) else 0.0
                    except ValueError:
                        current_score = 0.0
                    
                    utp_mod_clean = self.clean_text(row['Модули'])
                    best_match_ratio = 0
                    best_score = current_score
                    
                    for _, re_row in df_re.iterrows():
                        ratio = fuzz.token_set_ratio(utp_mod_clean, re_row['clean_module'])
                        
                        if ratio > best_match_ratio and ratio > 80:
                            best_match_ratio = ratio
                            try:
                                candidate_score = float(re_row['re_score'])
                                if candidate_score > best_score:
                                    best_score = candidate_score
                            except (ValueError, TypeError):
                                pass
                    
                    if best_score > current_score:
                        report.at[idx, 'Средний процент'] = best_score
                        grade_text = self.get_final_grade_text(best_score, row['Форма аттестации'])
                        report.at[idx, 'Итоговая оценка'] = f"{grade_text} (перезачет)"

            clean_name = re.sub(r'[\\/*?:"<>|]', "", student_info)
            save_p = filedialog.asksaveasfilename(
                initialfile=f"Табель_{clean_name}.pdf",
                defaultextension=".pdf", 
                filetypes=[("PDF", "*.pdf")]
            )

            if save_p:
                save_report_to_pdf(report, save_p, student_info)
                os.startfile(save_p)
                messagebox.showinfo("Успех", f"Табель для {student_info} готов.")

        except Exception as e:
            messagebox.showerror("Ошибка", f"Произошел сбой при генерации:\n{e}")

if __name__ == "__main__":
    root = tk.Tk()
    app = App(root)
    root.mainloop()