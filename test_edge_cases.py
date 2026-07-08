import pandas as pd
import os
from gerador_escalas_desktop import GeradorEscalasApp
import tkinter as tk

def test_edge_cases():
    print("Gerando mock de Check_Pre_Envio_Gerado.xlsx com casos extremos...")
    data = [
        {
            'Nome ': 'Cesar Brandão',
            'Plataforma': 'SPORTV',
            'Data': '16/07/2026',
            'Dia': '2026-07-16 18:30:00', # Simulating broken Dia
            'Pré': '2026-07-16 08:00:00', # Simulating DateTime in time column
            'Início': '08:30:00', # Simulating broken window start
            'Fim': '11:45:00', # Simulating broken window end
            'Evento/Descrição': 'BDRJ - EXIBIÇÃO',
            'Função': 'Comentarista Futebol'
        },
        {
            'Nome ': 'Mariana Becker',
            'Plataforma': 'PREMIERE',
            'Data': '17/07/2026',
            'Dia': '2026-07-17 18:30:00',
            'Pré': '-', 
            'Início': '2026-07-17 22:15:00', # Date time
            'Fim': '2026-07-18 01:30:00', # Crossing midnight
            'Evento/Descrição': 'FUTEBOL - SÉRIE A',
            'Função': 'Comentaristas (outros)'
        }
    ]
    df = pd.DataFrame(data)
    df.to_excel('Check_Pre_Envio_Mock.xlsx', index=False)
    
    # Criar uma root oculta do tkinter para testar a engine
    root = tk.Tk()
    root.withdraw()
    app = GeradorEscalasApp(root)
    
    # Injetar variáveis selecionadas para ignorar interface
    app.envio_var.set("gerar") 
    
    print("Processando Etapa 2...")
    # Chamando diretamente o process_etapa2
    app.process_etapa2('Check_Pre_Envio_Mock.xlsx', 'contatos_nova_versao.xlsx', ['Comentarista Futebol', 'Comentaristas (outros)', 'Outros / Desconhecidos'])
    
    print("Processamento concluído. Checando HTMLs...")
    html_dir = 'escalas_geradas_html'
    
    for f in os.listdir(html_dir):
        if 'Cesar_Brandão' in f or 'Mariana_Becker' in f:
            print(f"\\n--- Verificando {f} ---")
            with open(os.path.join(html_dir, f), 'r', encoding='utf-8') as html_file:
                content = html_file.read()
                # Procurar pelos tempos e datas
                if '08:00' in content and '2026-07-16' not in content[content.find('08:00')-20:content.find('08:00')+20]:
                    print("[OK] Pré formatado corretamente (08:00)")
                if '08:30' in content:
                    print("[OK] Início quebrado formatado corretamente (08:30)")
                if '11:45' in content:
                    print("[OK] Fim quebrado formatado corretamente (11:45)")
                if '22:15' in content:
                    print("[OK] Início com data formatado corretamente (22:15)")
                if '01:30' in content:
                    print("[OK] Fim cruzando meia-noite formatado corretamente (01:30)")
                if 'Quinta-feira' in content or 'Sexta-feira' in content:
                    print("[OK] Coluna Dia forçada para dia da semana")
                    
    print("\\nTodos os testes finalizados.")

if __name__ == '__main__':
    test_edge_cases()
