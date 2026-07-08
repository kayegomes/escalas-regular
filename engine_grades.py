import pandas as pd
import numpy as np

def flatten_sportv_grade(file_path):
    """
    Reads the horizontal SporTV grid and flattens it.
    """
    try:
        xl = pd.ExcelFile(file_path)
        # Assume first sheet is the grade
        df_raw = xl.parse(xl.sheet_names[0], header=None)
        
        # In the SporTV grid, row 3 (index 2) usually contains the dates? No, let's search for "DATA GRADE"
        start_row = 0
        for i, row in df_raw.iterrows():
            if 'DATA' in str(row.values).upper() or 'EVENTO' in str(row.values).upper():
                start_row = i
                break
                
        df = xl.parse(xl.sheet_names[0], header=start_row)
        
        flat_events = []
        
        # In the horizontal layout, we have repeating blocks for each channel.
        # We look for columns that have 'Evento' or 'Programa' in their name
        evento_cols = [c for c in df.columns if 'EVENTO' in str(c).upper() or 'PROGRAMA' in str(c).upper()]
        
        for c in evento_cols:
            # We need to find the related Data, Hora, V/I for this specific Evento column.
            # Usually, they are the columns immediately to the left.
            idx = df.columns.get_loc(c)
            # Ensure idx is an integer, if there are duplicate column names, get_loc returns a slice or boolean array
            if isinstance(idx, (slice, np.ndarray)):
                # If there are duplicates, we need to iterate over indices. 
                # Pandas handles duplicate columns by adding .1, .2, etc. if read properly, 
                # but just in case, we grab the first match if it's an array.
                if isinstance(idx, slice):
                    idx_list = range(idx.start, idx.stop, idx.step or 1)
                else:
                    idx_list = np.where(idx)[0]
                
                for i in idx_list:
                    extract_sportv_channel_block(df, i, flat_events)
            else:
                extract_sportv_channel_block(df, idx, flat_events)
                
        df_flat = pd.DataFrame(flat_events)
        if not df_flat.empty:
            # Clean up
            df_flat = df_flat.dropna(subset=['Evento'])
            df_flat = df_flat[df_flat['Evento'] != '']
        return df_flat
    except Exception as e:
        print(f"Erro processando Grade SporTV: {e}")
        return pd.DataFrame()
        
def extract_sportv_channel_block(df, evento_col_idx, flat_events):
    """ Helper to extract one channel block """
    col_evento = df.columns[evento_col_idx]
    
    # Try to find Data, Hora, V/I to the left
    col_data = df.columns[evento_col_idx - 3] if evento_col_idx >= 3 else None
    col_hora = df.columns[evento_col_idx - 2] if evento_col_idx >= 2 else None
    col_vi = df.columns[evento_col_idx - 1] if evento_col_idx >= 1 else None
    
    # Canal is tricky, it might be in the header above the data, or we can infer from the index 
    # (e.g. first block is SporTV, second is SporTV 2...)
    # Let's infer based on position for now, or just default to SPORTV (the exact channel can be matched via WO/Platform later)
    # A better way is checking the row above the headers if we kept it, but let's use a heuristic
    if evento_col_idx < 15:
        canal = "SPORTV"
    elif evento_col_idx < 20:
        canal = "SPORTV2"
    elif evento_col_idx < 25:
        canal = "SPORTV3"
    else:
        canal = "SPORTV"
        
    for _, row in df.iterrows():
        evento = str(row[col_evento]).strip()
        if evento.lower() in ['nan', 'none', '']:
            continue
            
        data = row[col_data] if col_data else None
        hora = row[col_hora] if col_hora else None
        vi = row[col_vi] if col_vi else 'V'
        
        flat_events.append({
            'Plataforma': canal,
            'Data': data,
            'Início': hora,
            'Pré': hora, # usually SporTV grid doesn't have Pre, just Hora
            'Fim': None, # Duraçao is on the right usually
            'Evento': evento,
            'V/I': vi
        })

def process_premiere_grade(file_path):
    try:
        xl = pd.ExcelFile(file_path)
        df = xl.parse(xl.sheet_names[0])
        # Find header row
        start_row = 0
        for i, row in df.iterrows():
            if 'EVENTO' in str(row.values).upper() or 'DATA' in str(row.values).upper():
                start_row = i
                break
        df = xl.parse(xl.sheet_names[0], header=start_row)
        
        events = []
        for _, row in df.iterrows():
            evento = str(row.get('EVENTO', '')).strip()
            if evento.lower() in ['nan', 'none', '']:
                continue
                
            # Construct Event name from Mandante X Visitante if Evento is just "BRASILEIRO"
            mandante = str(row.get('MANDANTE', '')).strip()
            visitante = str(row.get('VISITANTE', '')).strip()
            if mandante and visitante and mandante.lower() != 'nan' and visitante.lower() != 'nan':
                evento_full = f"{mandante} X {visitante} - {evento}"
            else:
                evento_full = evento
                
            canal = str(row.get('CANAL', 'PREMIERE'))
            
            events.append({
                'Plataforma': 'PREMIERE', # we'll extract TV Globo later
                'Data': row.get('DATA'),
                'Início': row.get('HORA') if 'HORA' in df.columns else row.get(':ORA'), # typo in excel
                'Pré': row.get('PRÉ'),
                'Fim': row.get('PÓS'),
                'Evento': evento_full,
                'V/I': 'V',
                'Raw_Canal': canal
            })
            
        return pd.DataFrame(events)
    except Exception as e:
        print(f"Erro processando Grade Premiere: {e}")
        return pd.DataFrame()

def process_combate_grade(file_path):
    try:
        xl = pd.ExcelFile(file_path)
        df = xl.parse(xl.sheet_names[0])
        start_row = 0
        for i, row in df.iterrows():
            if 'EVENTO' in str(row.values).upper() or 'DATA' in str(row.values).upper():
                start_row = i
                break
        df = xl.parse(xl.sheet_names[0], header=start_row)
        
        events = []
        for _, row in df.iterrows():
            evento = str(row.get('EVENTO', '')).strip()
            if evento.lower() in ['nan', 'none', '']:
                continue
                
            canal_sportv = str(row.get('SPORTV', ''))
            
            events.append({
                'Plataforma': 'COMBATE',
                'Data': row.get('DATA'),
                'Início': row.get('INÍCIO'),
                'Pré': row.get('PRÉ'),
                'Fim': row.get('FIM'),
                'Evento': evento,
                'V/I': 'V',
                'Raw_Sportv': canal_sportv
            })
            
        return pd.DataFrame(events)
    except Exception as e:
        print(f"Erro processando Grade Combate: {e}")
        return pd.DataFrame()

def process_all_grades(path_sportv, path_premiere, path_combate):
    """ Processes all grids, extracts TV Globo shared events, and returns a single DataFrame """
    df_sp = flatten_sportv_grade(path_sportv)
    df_pr = process_premiere_grade(path_premiere)
    df_co = process_combate_grade(path_combate)
    
    all_events = []
    
    if not df_sp.empty:
        all_events.append(df_sp)
        
    if not df_pr.empty:
        # Extract TV Globo from Premiere
        # e.g., 'CANAL' == 'PRE/TVG/GE TV'
        df_pr_globo = df_pr[df_pr['Raw_Canal'].astype(str).str.contains('TVG|GLOBO|GE TV', case=False, na=False)].copy()
        if not df_pr_globo.empty:
            df_pr_globo['Plataforma'] = 'TV GLOBO'
            all_events.append(df_pr_globo)
        all_events.append(df_pr)
        
    if not df_co.empty:
        # Extract TV Globo / SporTV from Combate
        df_co_sp = df_co[df_co['Raw_Sportv'].astype(str).str.contains('SPORTV', case=False, na=False)].copy()
        if not df_co_sp.empty:
            # We can't know exactly if it's SP1, SP2 or SP3 without parsing the text carefully,
            # but usually it says 'COMPARTILHADO COM O SPORTV 3'
            df_co_sp['Plataforma'] = 'SPORTV' # generic
            all_events.append(df_co_sp)
        all_events.append(df_co)
        
    if not all_events:
        return pd.DataFrame()
        
    df_all = pd.concat(all_events, ignore_index=True)
    
    # Filter out Replays
    if 'V/I' in df_all.columns:
        df_all = df_all[~df_all['V/I'].astype(str).str.upper().isin(['R', 'REPRISE'])]
        
    return df_all
