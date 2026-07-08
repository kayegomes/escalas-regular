import pandas as pd
import json

file_path = "Check pre envio-Macro Colunas - 15052026.xlsm"
try:
    xl = pd.ExcelFile(file_path)
    sheet_names = xl.sheet_names
    
    info = {"sheets": sheet_names, "columns": {}}
    for sheet in sheet_names:
        try:
            df = xl.parse(sheet, nrows=10)
            df.columns = [str(c) for c in df.columns]
            info["columns"][sheet] = list(df.columns)
            info[sheet + "_sample"] = df.astype(str).fillna("").head(3).to_dict(orient="records")
        except Exception as e:
            info["columns"][sheet] = f"Error reading: {e}"
            
    with open("excel_info.json", "w", encoding="utf-8") as f:
        json.dump(info, f, indent=4, ensure_ascii=False)
    print("Success. Wrote to excel_info.json")
except Exception as e:
    print(f"Error: {e}")
