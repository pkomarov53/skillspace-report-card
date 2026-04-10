import os
import pandas as pd
from fpdf import FPDF

class PDFReport(FPDF):
    def __init__(self, student_info=""):
        super().__init__()
        # Сохраняем информацию о студенте для заголовка
        self.student_info = student_info 
        self.col_widths = [80, 30, 30, 25, 30]
        self.headers = ['Предмет', 'Форма контроля', 'Трудоемкость', 'Результат', 'Оценка']

    def header(self):
        # Настройка шрифтов (поддержка кириллицы)
        font_path = r"C:\Windows\Fonts\arial.ttf"
        if os.path.exists(font_path):
            self.add_font('Arial', '', font_path, uni=True)
            self.set_font('Arial', '', 12)
        else:
            self.set_font('Helvetica', 'B', 12)
        
        # Вывод основной надписи и данных студента
        title = f'Табель успеваемости: {self.student_info}'
        self.cell(0, 10, title, ln=True, align='C') 
        self.ln(5)
        
        # Отрисовка шапки таблицы на каждой новой странице
        self.draw_table_header()

    def draw_table_header(self):
        self.set_font('Arial', '', 9)
        self.set_fill_color(230, 230, 230)
        for i in range(len(self.headers)):
            self.cell(self.col_widths[i], 10, self.headers[i], border=1, align='C', fill=True)
        self.ln()

    def get_row_height(self, text, width, line_height):
        """Вспомогательная функция для расчета высоты, которую займет текст"""
        lines = self.multi_cell(width, line_height, text, split_only=True)
        return len(lines) * line_height

def save_report_to_pdf(df, save_path, student_info=""):
    # Передаем student_info в конструктор класса
    pdf = PDFReport(student_info=student_info)
    pdf.set_auto_page_break(auto=True, margin=15)
    pdf.add_page()
    
    line_height = 7
    
    for _, row in df.iterrows():
        text = str(row['Модули'])
        
        # 1. Расчет высоты строки для контроля переноса страницы
        calculated_height = pdf.get_row_height(text, pdf.col_widths[0], line_height)
        
        # 2. Перенос на новую страницу, если строка не влезает
        if pdf.get_y() + calculated_height > 270:
            pdf.add_page() 
            
        x_start = pdf.get_x()
        y_start = pdf.get_y()

        # 3. Отрисовка первой колонки (название модуля)
        pdf.multi_cell(pdf.col_widths[0], line_height, text, border=1, align='L')
        
        y_end = pdf.get_y()
        total_row_height = y_end - y_start
        
        # 4. Отрисовка остальных колонок с выравниванием по высоте первой
        current_x = x_start + pdf.col_widths[0]
        other_cols = [
            str(row['Форма аттестации']),
            str(row['Количество часов']),
            f"{row['Средний процент']:.1f}%" if pd.notna(row['Средний процент']) else "-",
            str(row['Итоговая оценка'])
        ]
        
        for i, content in enumerate(other_cols):
            pdf.set_xy(current_x, y_start)
            pdf.cell(pdf.col_widths[i+1], total_row_height, content, border=1, align='C')
            current_x += pdf.col_widths[i+1]
        
        # Возврат курсора в начало следующей строки
        pdf.set_xy(x_start, y_end)

    pdf.output(save_path)