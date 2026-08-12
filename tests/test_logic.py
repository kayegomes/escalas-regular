import pandas as pd
import sys

file_path = "Check pre envio-Macro Colunas - 15052026.xlsm"
try:
    xl = pd.ExcelFile(file_path)
    # Read base MP, skip row 0 if it's messy, let's use header=1
    df_base = xl.parse("base MP", header=1)
    
    # print columns
    print("BASE MP COLUMNS:")
    print(df_base.columns.tolist())
    
    print("\nSAMPLE ROWS from base MP:")
    print(df_base[['Nome', 'Data', 'Evento/Programa', 'Canal']].head(10).to_string())
    
    print("\nSAMPLE from premiere:")
    df_prem = xl.parse("premiere")
    print(df_prem.columns.tolist())
    print(df_prem[['EVENTO', 'MANDANTE', 'VISITANTE', 'concat', 'DATA']].head(5).to_string())
    
except Exception as e:
    print(e)
