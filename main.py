import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
import sqlite3
import os
import re

# Импортируем наши модули
from data.connection import init_db, DB_NAME
from logic.parser import extract_utp_from_pdf
from logic.report import save_report_to_pdf
from logic.processor import process_student_data

class App:
    def __init__(self, root):
        self.root = root
        self.root.title("Система учета успеваемости Skillspace")
        self.root.geometry("600x350")
        
        init_db()
        self.excel_path = None
        self.setup_ui()

    def setup_ui(self):
        # --- Секция БД ---
        frame_db = tk.LabelFrame(self.root, text=" База данных УТП ", padx=10, pady=10)
        frame_db.pack(fill="x", padx=10, pady=10)
        tk.Button(frame_db, text="Загрузить новое УТП из PDF", command=self.import_pdf).pack(side="left")
        
        tk.Label(self.root, text="Выберите УТП:").pack(pady=(10,0))
        self.utp_combo = ttk.Combobox(self.root, values=self.get_utp_list(), state="readonly", width=70)
        self.utp_combo.pack(pady=5)
        if self.utp_combo['values']:
            self.utp_combo.current(0)

        # --- Секция студента ---
        frame_stud = tk.LabelFrame(self.root, text=" Данные студента ", padx=10, pady=10)
        frame_stud.pack(fill="x", padx=10, pady=10)
        tk.Button(frame_stud, text="Выбрать файл студента (.xlsx)", command=self.load_excel).pack(side="left")
        self.lbl_file = tk.Label(frame_stud, text="Файл не выбран", fg="gray", padx=10)
        self.lbl_file.pack(side="left")

        # --- Кнопка запуска ---
        tk.Button(self.root, text="СФОРМИРОВАТЬ ТАБЕЛЬ (PDF)", 
                  command=self.process_all, bg="#4CAF50", fg="white", 
                  font=("Arial", 10, "bold"), pady=10).pack(pady=30, fill="x", padx=10)

    def get_utp_list(self):
        conn = sqlite3.connect(DB_NAME)
        cursor = conn.cursor()
        cursor.execute('SELECT DISTINCT utp_name FROM utp_modules')
        names = [r[0] for r in cursor.fetchall()]
        conn.close()
        return names

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
            self.excel_path = path
            self.lbl_file.config(text=os.path.basename(path), fg="black")

    def process_all(self):
            utp_name = self.utp_combo.get()
            if not utp_name or not self.excel_path:
                messagebox.showwarning("Внимание", "Выберите УТП и файл Excel!")
                return

            try:
                # 1. Извлекаем данные студента из ячейки B1
                df_header = pd.read_excel(self.excel_path, header=None, nrows=1)
                student_info = str(df_header.iloc[0, 1]) if not df_header.empty else "Студент"

                # 2. Загружаем структуру УТП из БД
                conn = sqlite3.connect(DB_NAME)
                df_utp = pd.read_sql(
                    f"SELECT module_name as 'Модули', hours as 'Количество часов', control_form as 'Форма аттестации' "
                    f"FROM utp_modules WHERE utp_name='{utp_name}'", conn)
                conn.close()

                # 3. Загружаем ведомость студента
                df_stud = pd.read_excel(self.excel_path)
                
                # ВЫЗОВ ЛОГИКИ ОБРАБОТКИ
                report = process_student_data(df_utp, df_stud, utp_name)

                # 4. Подготовка имени файла по умолчанию
                # Очищаем имя от спецсимволов, чтобы Windows не ругался
                clean_name = re.sub(r'[\\/*?:"<>|]', "", student_info)
                default_filename = f"Табель_{clean_name}.pdf"

                # 5. Сохранение в PDF с предложенным именем
                save_p = filedialog.asksaveasfilename(
                    initialfile=default_filename, # Вот тут магия
                    defaultextension=".pdf", 
                    filetypes=[("PDF", "*.pdf")]
                )

                if save_p:
                    save_report_to_pdf(report, save_p, student_info)
                    
                    found_count = report['Средний процент'].notna().sum()
                    messagebox.showinfo("Готово", f"Табель создан для: {student_info}\nНайдено совпадений: {found_count}")
                    os.startfile(save_p)

            except Exception as e:
                messagebox.showerror("Ошибка", f"Произошел сбой: {e}")

if __name__ == "__main__":
    root = tk.Tk()
    # root.iconbitmap("myIcon.ico") # Убедись, что иконка есть, или закомментируй
    root.resizable(width=False, height=False)
    
    app = App(root)
    root.mainloop()