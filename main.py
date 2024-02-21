import docx
import pandas as pd
import re

import warnings
warnings.filterwarnings('ignore')

Input_excel_file = 'Input.xlsx'
ExcelData_df = pd.read_excel(Input_excel_file)

def detect_hyperlink_column(df):
    for col in ExcelData_df.columns:
        for value in df[col]:
            if isinstance(value, str) and re.match(r'https?://\S+', value):
                return col
    return None

def add_hyperlink(cell, url, text, color, underline):
    paragraph = cell.paragraphs[0]
    paragraph.clear()
    part = paragraph.part
    r_id = part.relate_to(url, docx.opc.constants.RELATIONSHIP_TYPE.HYPERLINK, is_external=True)
    hyperlink = docx.oxml.shared.OxmlElement('w:hyperlink')
    hyperlink.set(docx.oxml.shared.qn('r:id'), r_id)
    new_run = docx.oxml.shared.OxmlElement('w:r')
    rPr = docx.oxml.shared.OxmlElement('w:rPr')
    if not color is None:
        c = docx.oxml.shared.OxmlElement('w:color')
        c.set(docx.oxml.shared.qn('w:val'), color)
        rPr.append(c)
    if not underline:
        u = docx.oxml.shared.OxmlElement('w:u')
        u.set(docx.oxml.shared.qn('w:val'), 'none')
        rPr.append(u)
    new_run.append(rPr)
    new_run.text = text
    hyperlink.append(new_run)
    paragraph._p.append(hyperlink)
    return hyperlink

hyperlink_column = detect_hyperlink_column(ExcelData_df)

if hyperlink_column is None:
    print("No URL-like column detected in the DataFrame.")
else:
    Word = docx.Document()
    table = Word.add_table(rows=1, cols=len(ExcelData_df.columns))
    table.style = 'Table Grid'
    hdr_cells = table.rows[0].cells
    for i, col_name in enumerate(ExcelData_df.columns):
        hdr_cells[i].text = col_name
        for paragraph in hdr_cells[i].paragraphs:
            for run in paragraph.runs:
                run.bold = True


    for _, row in ExcelData_df.iterrows():
        row_cells = table.add_row().cells
        for i, value in enumerate(row):
            row_cells[i].text = str(value)

            for cell in row_cells:
                paragraphs = cell.paragraphs
                for paragraph in paragraphs:
                    paragraph.alignment = docx.enum.text.WD_ALIGN_PARAGRAPH.LEFT

        if ExcelData_df.columns[-1] == hyperlink_column:
            cell = row_cells[-1]
            hyperlink = add_hyperlink(cell, row[-1], row[-1], '0000FF', False)


    Word.save('Output.docx')
    print("Done , file saved as Output.docx at cwd ")
