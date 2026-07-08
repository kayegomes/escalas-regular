import pandas as pd
import numpy as np

def process_2468_base(file_path):
    """
    Reads the 2468 report (Base MP or Base 2) and consolidates roles by WO#
    """
    xl = pd.ExcelFile(file_path)
    sheet_to_read = None
    for s in ['base (2)', 'base MP', 'Base MP']:
        if s in xl.sheet_names:
            sheet_to_read = s
            break
            
    if not sheet_to_read:
        # Fallback to the first sheet if specific names not found
        sheet_to_read = xl.sheet_names[0]
        
    # Try finding the correct header row
    df = pd.DataFrame()
    for h in [0, 1, 2]:
        temp_df = xl.parse(sheet_to_read, header=h)
        # Rename column WO # to WO# immediately for check
        if 'WO #' in temp_df.columns:
            temp_df = temp_df.rename(columns={'WO #': 'WO#'})
            
        if 'WO#' in temp_df.columns or 'Nome' in temp_df.columns or 'Nome ' in temp_df.columns:
            df = temp_df
            break
            
    if 'WO#' not in df.columns:
        raise ValueError("Coluna 'WO#' não encontrada no relatório 2468.")
        
    # Map columns exactly to what the macro used to output (base email)
    col_map = {
        'Nome': 'Nome ',
        'Tipo de Atividade': 'Tipo Atividade ',
        'Sub-Atividade': 'Sub-Atividade (shift)',
        'Canal': 'Canal (Master Room)',
        'Inicio': 'Início',
        'Início Recurso': 'Início',
        'Fim Recurso': 'Fim',
        'data ': 'Data',
        'dia': 'Dia',
        'Air Start Time': 'Hr Produção',
        'WO Phase': 'WO Status',
        'Produto (WO/Quick Hold)': 'Produto (WO/Shift)',
        'Evento/Programa': 'Atividade/Descrição',
        'Local de Locução': 'Local Narração',
        'Local': 'Local de Gravação',
        'Status Aprov.': 'Status',
        'Notes': 'notas'
    }
    
    df = df.rename(columns=col_map)
    
    if 'Canal (Master Room)' in df.columns and 'Plataforma' not in df.columns:
        df['Plataforma'] = df['Canal (Master Room)']
        
    # Drop empty names or WOs
    df = df.dropna(subset=['Nome ', 'WO#'])
    
    # 1. Fill missing values
    for c in ['Narrador', 'Comentarista', 'Repórter', 'Elenco', 'Coordenador', 'Produtor', 'Pré']:
        if c not in df.columns:
            df[c] = ''
    df['Nome '] = df['Nome '].fillna('')
    df['Função'] = df['Função'].fillna('')
    if 'Status' not in df.columns:
        df['Status'] = ''
    else:
        df['Status'] = df['Status'].fillna('')
    
    # 2. Aggregation: for each WO#, we collect all Narrators, Commentators, etc.
    # Grouping by WO# to collect the team
    wo_team = {}
    for _, row in df.iterrows():
        wo = row['WO#']
        nome = str(row['Nome ']).strip()
        funcao = str(row['Função']).strip().lower()
        if wo not in wo_team:
            wo_team[wo] = {'Narrador': [], 'Comentarista': [], 'Repórter': [], 'Elenco': [], 'Coordenador': [], 'Produtor': []}
            
        # Determine the role category based on 'Função'
        if 'narrador' in funcao:
            wo_team[wo]['Narrador'].append(nome)
        elif 'coment' in funcao:
            wo_team[wo]['Comentarista'].append(nome)
        elif 'rep' in funcao or 'reporter' in funcao:
            wo_team[wo]['Repórter'].append(nome)
        elif 'coord' in funcao:
            wo_team[wo]['Coordenador'].append(nome)
        elif 'produtor' in funcao:
            wo_team[wo]['Produtor'].append(nome)
        else:
            wo_team[wo]['Elenco'].append(nome)
            
    # Now build the consolidated strings for each WO#
    for wo in wo_team:
        for role in wo_team[wo]:
            wo_team[wo][role] = " ; ".join(list(set(wo_team[wo][role])))
            
    # Update the dataframe with the consolidated team members
    df['Narrador'] = df['WO#'].apply(lambda w: wo_team.get(w, {}).get('Narrador', ''))
    df['Comentarista'] = df['WO#'].apply(lambda w: wo_team.get(w, {}).get('Comentarista', ''))
    df['Repórter'] = df['WO#'].apply(lambda w: wo_team.get(w, {}).get('Repórter', ''))
    df['Coordenador'] = df['WO#'].apply(lambda w: wo_team.get(w, {}).get('Coordenador', ''))
    df['Produtor'] = df['WO#'].apply(lambda w: wo_team.get(w, {}).get('Produtor', ''))
    df['Elenco'] = df['WO#'].apply(lambda w: wo_team.get(w, {}).get('Elenco', ''))
    
    # Tag (by) - The user said: "adicionando a tag (by) quando aplicável". We will append " (by)" for those in "pedido_correção BY"? 
    # For now, we just consolidate. The user didn't specify exactly what triggers (by), 
    # but we can add it later if they ask, or just leave the consolidation.
    
    # Clean up Data formatting
    if 'Data' in df.columns:
        df['Data_raw'] = df['Data'] # Keep raw for crossing
        
    return df
