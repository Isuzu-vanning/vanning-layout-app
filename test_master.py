import pandas as pd
df = pd.read_excel('parts_master.xlsx')
print("Columns:", df.columns.tolist())
for index, row in df.head(3).iterrows():
    print(f"Row {index}: {row.iloc[0]}, {row.iloc[1]}, {row.iloc[2]}")
