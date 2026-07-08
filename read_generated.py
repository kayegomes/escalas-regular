import sys
import pandas as pd
sys.path.append(r'c:\Users\ligomes\Downloads\escalas_regular')
from engine_cross import run_etapa1

try:
    path_2468 = r'c:\Users\ligomes\Downloads\escalas_regular\(2468) Esporte - Atividades de Equipe – Sub-Atividades_v2_ (9).xlsx'
    path_sportv = r'c:\Users\ligomes\Downloads\escalas_regular\GRADE DE JULHO - 11ª VERSÃO.xlsm'
    path_premiere = r'c:\Users\ligomes\Downloads\escalas_regular\PPV 2026 (Jul 10ª versão).xlsx'
    path_combate = r'c:\Users\ligomes\Downloads\escalas_regular\GRADE DE EVENTOS COMBATE 2026 - JULHO  (V4).xlsx'
    
    out_path = run_etapa1(path_2468, path_sportv, path_premiere, path_combate)
    df = pd.read_excel(out_path)
    
    with open(r'c:\Users\ligomes\Downloads\escalas_regular\output_generated.txt', 'w', encoding='utf-8') as f:
        f.write("=== FINAL COLUMNS ===\n")
        f.write(str(df.columns.tolist()) + "\n\n")
        
        f.write("=== STATUS REVISÃO COUNTS ===\n")
        f.write(str(df['Status Revisão'].value_counts()) + "\n\n")
        
        f.write("=== SAMPLE MATCHES (OK) ===\n")
        df_ok = df[df['Status Revisão'] == '🟢 OK'].head(3)
        for _, row in df_ok.iterrows():
            f.write(f"WO: {row.get('WO#', '')} | Evento: {row.get('Atividade/Descrição', '')} | Plat: {row.get('Plataforma', '')} | Inicio: {row.get('Início', '')} | Fim: {row.get('Fim', '')}\n")
            f.write(f"   Equipe -> Narrador: {row.get('Narrador', '')}, Comentarista: {row.get('Comentarista', '')}, Produtor: {row.get('Produtor', '')}\n")
            
except Exception as e:
    print(f"Erro ao ler arquivo: {e}")
