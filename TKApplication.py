class App:
    def __init__(self, root):
        self.root = root
        self.root.title("Система учета успеваемости Skilspace")
        self.root.geometry("600x350")
        
        init_db()
        self.excel_path = None
        self.setup_ui()

    def setup_ui(self):
        # --- Секция БД ---
        frame_db = tk.LabelFrame(self.root, text=" База данных УТП ", padx=10, pady=10)
        frame_db.pack(fill="x", padx=10, pady=10)

        tk.Button(frame_db, text="➕ Загрузить новое УТП из PDF", command=self.import_pdf).pack(side="left")
        
        tk.Label(self.root, text="Выберите учебный план (УТП):").pack(pady=(10,0))
        self.utp_combo = ttk.Combobox(self.root, values=self.get_utp_list(), state="readonly", width=70)
        self.utp_combo.pack(pady=5)

        # --- Секция Студента ---
        frame_stud = tk.LabelFrame(self.root, text=" Данные студента ", padx=10, pady=10)
        frame_stud.pack(fill="x", padx=10, pady=10)

        tk.Button(frame_stud, text="📂 Выбрать файл студента (.xlsx)", command=self.load_excel).pack(side="left")
        self.lbl_file = tk.Label(frame_stud, text="Файл не выбран", fg="gray", padx=10)
        self.lbl_file.pack(side="left")

        self.btn_run = tk.Button(self.root, text="СФОРМИРОВАТЬ ТАБЕЛЬ (PDF)", 
                                 command=self.process_all, bg="#4CAF50", fg="white", 
                                 font=("Arial", 10, "bold"), pady=10)
        self.btn_run.pack(pady=30, fill="x", padx=10)

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
                messagebox.showinfo("Успех", f"УТП '{name}' успешно добавлено в базу!")
            except Exception as e:
                messagebox.showerror("Ошибка", f"Не удалось прочитать PDF: {e}")

    def load_excel(self):
        path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
        if path:
            self.excel_path = path
            self.lbl_file.config(text=os.path.basename(path), fg="black")

    def grade_logic(self, row):
        score = row['Средний процент']
        form = row['Форма аттестации']
        
        if "Зачет" in form and "оценкой" not in form:
            return "Зачет" if score >= 60 else "Незачет"
        else:
            if score >= 85: return "5 (отл.)"
            elif score >= 70: return "4 (хор.)"
            elif score >= 50: return "3 (уд.)"
            else: return "2 (неуд.)"

    def process_all(self):
        utp_name = self.utp_combo.get()
        if not utp_name or not self.excel_path:
            messagebox.showwarning("Внимание", "Выберите УТП и файл Excel!")
            return

        try:
            # Загрузка данных из БД (УТП)
            conn = sqlite3.connect(DB_NAME)
            df_utp = pd.read_sql(
                f"SELECT module_name as 'Модули', hours as 'Количество часов', control_form as 'Форма аттестации' "
                f"FROM utp_modules WHERE utp_name='{utp_name}'", conn)
            conn.close()

            # Загрузка и очистка данных из Excel (Студент)
            df_stud = pd.read_excel(self.excel_path)
            
            mask = df_stud.iloc[:, 0].str.contains('модулю|практикум|дз|тестирование|защита итогового', case=False, na=False)
            df_filtered = df_stud[mask].copy()
            
            # Чистим названия в Excel для более точного сравнения
            def normalize_simple(text):
                if not isinstance(text, str): return ""
                text = re.sub(r'тестирование по модулю', '', text, flags=re.IGNORECASE)
                text = text.replace('"', '').replace('«', '').replace('»', '').strip().lower()
                return text

            df_filtered['CleanName'] = df_filtered.iloc[:, 0].apply(normalize_simple)
            
            # Извлекаем баллы
            def extract_score(val):
                res = re.findall(r'(\d+)', str(val))
                return float(res[0]) if res else None
            df_filtered['Score'] = df_filtered.iloc[:, 1].apply(extract_score)
            df_filtered = df_filtered.dropna(subset=['Score'])
            df_grouped = df_filtered.groupby('CleanName')['Score'].mean().reset_index()

            # FUZZY MATCHING
            scores_list = []
            
            for index, row_utp in df_utp.iterrows():
                name_to_find = row_utp['Модули'].lower().strip()
                best_score = None
                highest_ratio = 0
                
                for _, row_excel in df_grouped.iterrows():
                    current_ratio = fuzz.token_sort_ratio(name_to_find, row_excel['CleanName'])
                    
                    if current_ratio > highest_ratio and current_ratio >= 70:
                        highest_ratio = current_ratio
                        best_score = row_excel['Score']
                
                scores_list.append(best_score)
            
            report = df_utp.copy()
            report['Средний процент'] = scores_list
            
            # Завершение обработки
            report['Итоговая оценка'] = report.apply(self.grade_logic, axis=1)

            save_p = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF", "*.pdf")])
            if save_p:
                save_report_to_pdf(report, save_p)
                found_count = report['Средний процент'].notna().sum()
                messagebox.showinfo("Готово", f"Табель создан!\nНайдено совпадений: {found_count} из {len(report)}")
                os.startfile(save_p)

        except Exception as e:
            import traceback
            print(traceback.format_exc())
            messagebox.showerror("Ошибка", f"Ошибка обработки: {str(e)}")
