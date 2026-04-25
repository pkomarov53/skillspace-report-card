import os
import pandas as pd

from fpdf import FPDF
from fpdf.fonts import FontFace

class PDFReport(FPDF):
    def __init__(self, student_info: str = "", font_path: str = r"C:\Windows\Fonts\arial.ttf"):
        super().__init__()
        self.student_info = student_info 
        
        if os.path.exists(font_path):
            self.add_font('Arial', '', font_path, uni=True)
            self.default_font = 'Arial'
        else:
            self.default_font = 'Helvetica'

    def header(self):
        self.set_font(self.default_font, '', 12)
        title = f'Табель успеваемости: {self.student_info}'
        self.cell(0, 10, title, ln=True, align='C') 
        self.ln(5)


def save_report_to_pdf(df: pd.DataFrame, save_path: str, student_info: str = "", font_path: str = r"C:\Windows\Fonts\arial.ttf"):
    pdf = PDFReport(student_info=student_info, font_path=font_path)
    pdf.set_auto_page_break(auto=True, margin=15)
    pdf.add_page()

    pdf.set_font(pdf.default_font, '', 9)
    
    col_widths = (80, 30, 25, 20, 40)
    headers = ('Предмет', 'Форма контроля', 'Трудоемкость', 'Результат', 'Оценка')
    
    with pdf.table(
        col_widths=col_widths,
        text_align=("LEFT", "CENTER", "CENTER", "CENTER", "CENTER"),
        borders_layout="ALL",
        first_row_as_headings=True,
        line_height=7,
        headings_style=FontFace(emphasis="") 
    ) as table:
        
        row = table.row()
        for header in headers:
            row.cell(header)

        for data in df.to_dict('records'):
            row = table.row()
            
            avg_percent = data.get('Средний процент')
            avg_str = f"{avg_percent:.1f}%" if pd.notna(avg_percent) else "-"
            
            row.cell(str(data.get('Модули', '')))
            row.cell(str(data.get('Форма аттестации', '')))
            row.cell(str(data.get('Количество часов', '')))
            row.cell(avg_str)
            row.cell(str(data.get('Итоговая оценка', '')))

    pdf.output(save_path)