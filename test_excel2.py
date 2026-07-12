# -*- coding: utf-8 -*-
import pandas as pd
file_path = 'data/atm004b 04.05 1.xlsx'

for sheet in ['12', '13', 'GD lỗi']:
    print(f"\n--- {sheet} ---")
    df = pd.read_excel(file_path, sheet_name=sheet, engine='calamine', header=None)
    
    # Find the row that contains 'TRN_DT'
    header_idx = -1
    for i in range(min(10, len(df))):
        if 'TRN_DT' in df.iloc[i].values:
            header_idx = i
            break
            
    if header_idx != -1:
        print(f"Found header at row {header_idx}")
        df_real = pd.read_excel(file_path, sheet_name=sheet, engine='calamine', header=header_idx)
        print("Columns:", df_real.columns.tolist())
        print("Dates:", df_real['TRN_DT'].dropna().unique())
    else:
        print("Could not find 'TRN_DT' in first 10 rows")
        print(df.head())
