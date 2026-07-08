import pandas as pd

df = pd.read_excel(r'c:\Users\ligomes\Downloads\escalas_regular\Enviador de Escalas V7.7 (Coluna Colaboradores) - 26062026_Escala 06 a 12 jul.xlsm', sheet_name='Lista e-mails', header=None)

emails_found = []
for row_idx, row in df.iterrows():
    for col_idx, val in enumerate(row):
        val_str = str(val)
        if '@' in val_str:
            name = str(row.iloc[col_idx-1]) if col_idx > 0 else ""
            emails_found.append(f"Name: {name} -> Email: {val_str}")
            if len(emails_found) >= 30:
                break
    if len(emails_found) >= 30:
        break

with open(r'c:\Users\ligomes\Downloads\escalas_regular\emails_out.txt', 'w', encoding='utf-8') as f:
    for e in emails_found:
        f.write(e + "\n")
