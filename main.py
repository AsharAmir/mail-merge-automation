import pandas as pd
from docx import Document
import os


source = 'StateCommissions_PHA_pulled 6.13.2024_1214pm.xlsx' #change source file here
df = pd.read_excel(source, sheet_name='DATA') #change sheet name here


df = df.fillna('')

'''
format: 
word_placeholder : excel_column
'''
placeholders = { #change placeeholders heree
    '< AUTHORIZED_REP >': 'auth name',
    '< ORGANIZATION >': 'Org',
    '< Org Street Addr 1 >': 'Adress1',
    '< Org Street Addr 2 >': 'Address 2',
    '< ORG CITY>': 'City',
    '< ORG STATE >': 'State',
    '<ORG ZIP >': 'zip',
    '<Authorized Rep Name>': 'auth name',
}

def replace_placeholders(doc, replacements):
    for paragraph in doc.paragraphs:
        for key, value in replacements.items():
            if key in paragraph.text:                
                paragraph.text = paragraph.text.replace(key, str(value) if value is not None else "")
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for key, value in replacements.items():
                    if key in cell.text:
                        cell.text = cell.text.replace(key, str(value) if value is not None else "")


output_dir = 'outputs' #change output directory here
os.makedirs(output_dir, exist_ok=True)


for index, row in df.iterrows():
    doc = Document('FY_24_PHAComm_Notification__Ltr_TEMPLATE (3).docx')  #change template file here
    
    replacements = {key: row[value] for key, value in placeholders.items()}
    replace_placeholders(doc, replacements)
    file_name = row['File_Name'] + '.docx'
    doc.save(os.path.join(output_dir, file_name))

print("Mail merge completed!")
