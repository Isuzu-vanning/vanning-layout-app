import pandas as pd
import re

file_path = 'vanning_layout_2026.xlsx'
xl = pd.ExcelFile(file_path)

for sheet_name in xl.sheet_names[:2]: # Check first two months
    df = pd.read_excel(file_path, sheet_name=sheet_name, header=None)
    print(f"--- Sheet: {sheet_name} ---")
    
    # Look for container headers like "2026/01/15 Container-1"
    container_pattern = re.compile(r'(\d{4}/\d{2}/\d{2})\s+Container-(\d+)')
    
    current_date = None
    current_container = None
    
    for i, row in df.iterrows():
        cell_0 = str(row[0])
        match = container_pattern.search(cell_0)
        if match:
            date_str = match.group(1)
            container_id = match.group(2)
            print(f"Line {i}: Date={date_str}, Container={container_id}")
            current_date = date_str
            current_container = container_id
        elif current_date:
            # Check if this is a data row
            # Usually data rows have an ID in column 1
            try:
                part_id = int(row[1])
                # print(f"  Item {part_id} in {current_date} C-{current_container}")
            except:
                pass
