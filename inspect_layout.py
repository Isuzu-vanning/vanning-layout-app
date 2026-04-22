import pandas as pd
file_path = 'vanning_layout_2026.xlsx'
xl = pd.ExcelFile(file_path)
print(f"Sheets: {xl.sheet_names}")
for sheet in xl.sheet_names[:3]:
    df = pd.read_excel(file_path, sheet_name=sheet)
    print(f"\nSheet: {sheet}")
    print(df.head())
    print(df.columns)
