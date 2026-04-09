import pdfplumber
import os
import re
import pandas as pd

def extract_utp_from_pdf(file_path):
    data = []
    utp_name = os.path.basename(file_path).replace('.pdf', '')
    
    with pdfplumber.open(file_path) as pdf:
        for page in pdf.pages:
            table = page.extract_table()
            if not table:
                continue
            
            for row in table:
                clean_row = [str(cell).replace('\n', ' ').strip() if cell else "" for cell in row]
                
                # Фильтр: берем строки, где есть форма контроля
                if any(word in clean_row[-1] for word in ['Зачет', 'Экзамен']):
                    module_name = clean_row[1] if len(clean_row) > 1 else clean_row[0]
                    module_name = re.sub(r'^\d+\.?\s*', '', module_name)
                    
                    data.append({
                        'utp_name': utp_name,
                        'module_name': module_name,
                        'hours': clean_row[-2] if len(clean_row) > 2 else "0",
                        'control_form': clean_row[-1]
                    })
    return utp_name, pd.DataFrame(data)