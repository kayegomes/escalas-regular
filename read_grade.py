import pandas as pd
import json

def print_grade_info(file_path):
    print(f"\n--- {file_path} ---")
    try:
        xl = pd.ExcelFile(file_path)
        sheet_names = xl.sheet_names
        print(f"Sheets: {sheet_names}")
        df = xl.parse(sheet_names[0], nrows=10)
        print(df.to_string())
    except Exception as e:
        print(f"Error: {e}")

print_grade_info("GRADE DE EVENTOS COMBATE 2026 - JULHO  (V4).xlsx")
print_grade_info("PPV 2026 (Jul 10ª versão).xlsx")
