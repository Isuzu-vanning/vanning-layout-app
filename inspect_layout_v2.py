import pandas as pd
file_path = 'vanning_layout_2026.xlsx'
xl = pd.ExcelFile(file_path)
for sheet in xl.sheet_names[:1]:
    df = pd.read_excel(file_path, sheet_name=sheet, header=None)
    print(f"\nSheet: {sheet}")
    print(df.to_csv(index=False)) # CSV output for cleaner view
