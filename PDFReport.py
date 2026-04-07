class PDFReport(FPDF):
    def header(self):
        font_path = r"C:\Windows\Fonts\arial.ttf"
        if os.path.exists(font_path):
            self.add_font('Arial', '', font_path, uni=True)
            self.set_font('Arial', '', 14)
        else:
            self.set_font('Helvetica', '', 14)
        
        self.cell(0, 10, 'Табель успеваемости студента', ln=True, align='C')
        self.ln(5)

def save_report_to_pdf(df, save_path):
    pdf = PDFReport()
    pdf.add_page()
    
    pdf.set_font('Arial', '', 9)
    col_widths = [80, 35, 20, 25, 30]
    headers = ['Предмет', 'Форма контроля', 'Часы', 'Результат %', 'Оценка']
    
    for i in range(len(headers)):
        pdf.cell(col_widths[i], 10, headers[i], border=1, align='C')
    pdf.ln()

    for _, row in df.iterrows():
        pdf.cell(col_widths[0], 10, str(row['Модули'])[:45], border=1)
        pdf.cell(col_widths[1], 10, str(row['Форма аттестации']), border=1, align='C')
        pdf.cell(col_widths[2], 10, str(row['Количество часов']), border=1, align='C')
        pdf.cell(col_widths[3], 10, f"{row['Средний процент']:.1f}%" if pd.notna(row['Средний процент']) else "-", border=1, align='C')
        pdf.cell(col_widths[4], 10, str(row['Итоговая оценка']), border=1, align='C')
        pdf.ln()
    
    pdf.output(save_path)
