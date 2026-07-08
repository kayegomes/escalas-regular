import pandas as pd
import numpy as np
import os
from engine_2468 import process_2468_base
from engine_grades import process_all_grades
from openpyxl import load_workbook
from openpyxl.styles import PatternFill

def run_etapa1(path_2468, path_sportv, path_premiere, path_combate):
    """
    Executes the entire Etapa 1 logic:
    1. Read and consolidate 2468 base.
    2. Read and flatten all grids.
    3. Cross reference and apply rules (replays, multimodalidades, viagens).
    4. Export to Excel with Visual Alerts.
    """
    print("Iniciando Motor 2468...")
    df_2468 = process_2468_base(path_2468)
    
    print("Iniciando Motor de Grades...")
    df_grades = process_all_grades(path_sportv, path_premiere, path_combate)
    
    print("Iniciando Cruzamento (Match)...")
    
    # Ignore travels
    if 'Atividade/Descrição' in df_2468.columns:
        df_2468 = df_2468[~df_2468['Atividade/Descrição'].astype(str).str.contains('Viagem', case=False, na=False)]
        
    # We will build a unified output list
    output_rows = []
    
    for _, row_2468 in df_2468.iterrows():
        out_row = row_2468.to_dict()
        out_row['Status Revisão'] = '🟢 OK' # Default
        
        # Extracted basic info
        evento_col = 'Atividade/Descrição' if 'Atividade/Descrição' in row_2468 else 'Evento'
        evento_2468 = str(row_2468.get(evento_col, '')).strip()
        
        plat_2468 = str(row_2468.get('Plataforma', row_2468.get('Canal (Master Room)', ''))).strip().upper()
        if plat_2468 == 'NAN': plat_2468 = ''
        
        if df_grades.empty:
            out_row['Status Revisão'] = '🔴 Sem Grades Fornecidas'
            output_rows.append(out_row)
            continue
            
        # Match logic only if we have platform
        match = pd.DataFrame()
        if plat_2468:
            match = df_grades[
                (df_grades['Plataforma'].str.upper().str.contains(plat_2468[:4], na=False, case=False) |
                 (plat_2468 == 'PREMIERE' and df_grades['Plataforma'] == 'PREMIERE'))
            ]
            
        # Try to find a match in grades
        # Match criteria: Same date (approx), Platform matches, Evento string contains some keyword
        # Since we don't have perfect IDs, we do a fuzzy check.
        
        if not match.empty:
            # Try to match the event name (very simplistic for now)
            # In a real scenario we'd use fuzzy string matching or WO extraction if present in both
            event_words = evento_2468.split()
            if len(event_words) > 0:
                keyword = event_words[0].upper()
                match_evento = match[match['Evento'].str.upper().str.contains(keyword, na=False)]
                if not match_evento.empty:
                    match = match_evento
                    
            best_match = match.iloc[0]
            
            # Apply rules
            # 1. Fallback for Multimodality / Long events
            if 'SURF' in evento_2468.upper() or 'TENIS' in evento_2468.upper() or 'TÊNIS' in evento_2468.upper():
                # Keep 2468 times
                out_row['Status Revisão'] = '🟡 Fallback (Multimodalidade)'
            else:
                # Update times from grade
                if pd.notnull(best_match.get('Início')): out_row['Início'] = best_match['Início']
                if pd.notnull(best_match.get('Pré')): out_row['Pré'] = best_match['Pré']
                if pd.notnull(best_match.get('Fim')): out_row['Fim'] = best_match['Fim']
                
            # A Confirmar status
            if 'CONFIRMAR' in str(best_match.get('Evento', '')).upper():
                out_row['Status Revisão'] = '🟡 A Confirmar'
        else:
            out_row['Status Revisão'] = '🔴 Não encontrado na Grade'
            
        output_rows.append(out_row)
        
    df_out = pd.DataFrame(output_rows)
    
    # Save to Excel
    out_path = os.path.join(os.path.dirname(path_2468), "Check_Pre_Envio_Gerado.xlsx")
    df_out.to_excel(out_path, index=False)
    
    # Apply openpyxl conditional formatting
    wb = load_workbook(out_path)
    ws = wb.active
    
    red_fill = PatternFill(start_color='FFFF0000', end_color='FFFF0000', fill_type='solid')
    yellow_fill = PatternFill(start_color='FFFFFF00', end_color='FFFFFF00', fill_type='solid')
    
    # Assuming 'Status Revisão' is the first column (A)
    for row in range(2, ws.max_row + 1):
        status_cell = ws.cell(row=row, column=1)
        if '🔴' in str(status_cell.value):
            for col in range(1, ws.max_column + 1):
                ws.cell(row=row, column=col).fill = red_fill
        elif '🟡' in str(status_cell.value):
            for col in range(1, ws.max_column + 1):
                ws.cell(row=row, column=col).fill = yellow_fill
                
    wb.save(out_path)
    return out_path
