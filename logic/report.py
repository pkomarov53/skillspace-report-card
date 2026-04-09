import os
import pandas as pd
from fpdf import FPDF

class PDFReport(FPDF):
    def __init__(self):
        super().__init__()
        self.col_widths = [80, 30, 30, 25, 30]
        self.headers = ['Предмет', 'Форма контроля', 'Трудоемкость', 'Результат', 'Оценка']

    def header(self):
        font_path = r"C:\Windows\Fonts\arial.ttf"
        if os.path.exists(font_path):
            self.add_font('Arial', '', font_path, uni=True)
            self.set_font('Arial', '', 12)
        else:
            self.set_font('Helvetica', 'B', 12)
        
        self.cell(0, 10, 'Табель успеваемости студента', ln=True, align='C')
        self.ln(5)
        # Рисуем шапку таблицы при создании новой страницы
        self.draw_table_header()

    def draw_table_header(self):
        self.set_font('Arial', '', 9)
        self.set_fill_color(230, 230, 230)
        for i in range(len(self.headers)):
            self.cell(self.col_widths[i], 10, self.headers[i], border=1, align='C', fill=True)
        self.ln()

    def get_row_height(self, text, width, line_height):
        """Вспомогательная функция для расчета высоты, которую займет текст"""
        # Считаем, сколько строк займет текст в ячейке заданной ширины
        lines = self.multi_cell(width, line_height, text, split_only=True)
        return len(lines) * line_height

def save_report_to_pdf(df, save_path):
    pdf = PDFReport()
    pdf.set_auto_page_break(auto=True, margin=15)
    pdf.add_page()
    
    line_height = 7
    
    for _, row in df.iterrows():
        text = str(row['Модули'])
        
        # 1. Считаем высоту будущей строки заранее
        calculated_height = pdf.get_row_height(text, pdf.col_widths[0], line_height)
        
        # 2. Проверяем: если текущая позиция Y + высота строки > лимита страницы
        # Лимит страницы обычно ~270-280мм для A4
        if pdf.get_y() + calculated_height > 270:
            pdf.add_page() # Это автоматически вызовет header() и нарисует шапку
            
        # Запоминаем координаты после того, как убедились, что страница верная
        x_start = pdf.get_x()
        y_start = pdf.get_y()

        # 3. Рисуем первую колонку (название)
        pdf.multi_cell(pdf.col_widths[0], line_height, text, border=1, align='L')
        
        # Получаем реальную высоту, которую заняла ячейка
        y_end = pdf.get_y()
        total_row_height = y_end - y_start
        
        # 4. Дорисовываем остальные колонки
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
        
        # Переходим строго в начало следующей строки
        pdf.set_xy(x_start, y_end)

    pdf.output(save_path)