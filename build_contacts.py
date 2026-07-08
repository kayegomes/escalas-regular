import pandas as pd
import openpyxl

def main():
    print("Lendo Lista e-mails original...")
    df_old = pd.read_excel('Enviador de Escalas V7.7 (Coluna Colaboradores) - 26062026_Escala 06 a 12 jul.xlsm', sheet_name='Lista e-mails', header=None)
    
    # Headers are in row 0
    headers = {
        'Narrador': 11,
        'Comentarista Futebol': 16,
        'Comentaristas (outros)': 21,
        'Comentaristas Arbitragem': 26,
        'Colaboradores': 31
    }
    
    # Mapear nome para email e nome para grupo
    name_to_email = {}
    name_to_group = {}
    
    for group, col_idx in headers.items():
        # Iterate from row 3 downwards
        for r in range(3, len(df_old)):
            name = str(df_old.iloc[r, col_idx + 1]).strip()
            email = str(df_old.iloc[r, col_idx + 2]).strip()
            
            if name and name not in ['nan', 'NaT', '-']:
                name_to_group[name] = group
                if email and email not in ['nan', 'NaT', '-']:
                    name_to_email[name] = email
                else:
                    name_to_email[name] = ''
    
    print(f"Extraídos {len(name_to_group)} nomes da planilha original.")
    
    # Agora lendo os nomes da escala gerada
    try:
        df_scale = pd.read_excel('Check_Pre_Envio_Gerado.xlsx')
        scale_names = df_scale['Nome '].dropna().unique() if 'Nome ' in df_scale.columns else df_scale['Nome'].dropna().unique()
    except Exception as e:
        print("Aviso: Nao consegui ler Check_Pre_Envio_Gerado.xlsx", e)
        scale_names = []
        
    for name in scale_names:
        name = str(name).strip()
        if name and name not in ['nan', '-'] and name not in name_to_group:
            print(f"Novo nome detectado na escala: {name}")
            name_to_group[name] = 'Outros / Desconhecidos'
            name_to_email[name] = 'a definir'
            
    # Criar DataFrame para a nova aba Lista e-mails
    # Formato desejado: 
    # Coluna A: Grupo
    # Coluna B: Nome
    # Coluna C: Email
    
    data = []
    for name, group in name_to_group.items():
        data.append({
            'Grupo': group,
            'Nome': name,
            'Email': name_to_email.get(name, 'a definir')
        })
        
    df_new = pd.DataFrame(data)
    df_new = df_new.sort_values(by=['Grupo', 'Nome'])
    
    output_path = 'contatos_nova_versao.xlsx'
    df_new.to_excel(output_path, index=False, sheet_name='Lista e-mails')
    print(f"Planilha de contatos inicializada com sucesso em: {output_path}")

if __name__ == '__main__':
    main()
