# -*- coding: utf-8 -*-
import pandas as pd
file_path = 'data/atm004b 04.05 1.xlsx'
print("Sheet 12:")
try:
    df12 = pd.read_excel(file_path, sheet_name='12', engine='calamine')
    print(df12.head())
    print("Unique TRN_DT in 12:", df12['TRN_DT'].unique())
except Exception as e:
    print(e)
print("\nSheet 13:")
try:
    df13 = pd.read_excel(file_path, sheet_name='13', engine='calamine')
    print(df13.head())
    print("Unique TRN_DT in 13:", df13['TRN_DT'].unique())
except Exception as e:
    print(e)
print("\nGD Lỗi:")
try:
    df_err = pd.read_excel(file_path, sheet_name='GD lỗi', engine='calamine')
    print(df_err.head())
    print("AC_NO head:", df_err['AC_NO'].head().tolist())
except Exception as e:
    print(e)
