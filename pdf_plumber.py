import pdfplumber
import pandas as pd 

path='C:/Users/RaphaelThiney/OneDrive- Unison Energy/Cyxterra BOS 1A/'

# Function to extract tables from a PDF using pdfplumber
def extract_tables_with_pdfplumber(pdf_path):
    tables = []
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            tables.extend(page.extract_tables())
    return tables

# Extract tables from both PDFs
pdf_path_1 = f'{path}1_Interfacelist Master BOS1-A-High lighted.pdf'
pdf_path_2 = f'{path}2_Interfacelist CHP Control High lighted.pdf'

tables_1 = extract_tables_with_pdfplumber(pdf_path_1)
tables_2 = extract_tables_with_pdfplumber(pdf_path_2)

# Convert tables to DataFrames
dfs_1 = [pd.DataFrame(table) for table in tables_1]
dfs_2 = [pd.DataFrame(table) for table in tables_2]

# Combine all DataFrames into one Excel file
combined_output_path_all = '/mnt/data/Complete_Extracted_Data_All.xlsx'
with pd.ExcelWriter(combined_output_path_all) as writer:
    for i, df in enumerate(dfs_1, start=1):
        df.to_excel(writer, sheet_name=f'Table1_{i}', index=False)
    for i, df in enumerate(dfs_2, start=1):
        df.to_excel(writer, sheet_name=f'Table2_{i}', index=False)

combined_output_path_all
