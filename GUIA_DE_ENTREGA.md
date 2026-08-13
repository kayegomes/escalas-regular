# Sistema de Escalas — Pacote Atualizado

Este pacote contém a versão atualizada do sistema desktop de conferência e envio de escalas. A aplicação mantém o fluxo de criação de rascunhos no Outlook; ela não faz envio automático de e-mails.

## Instalação

Em um computador Windows com Python 3.11 ou superior, abra o Prompt de Comando na pasta do sistema e execute:

```bat
py -m pip install -r requirements.txt
py gerador_escalas_desktop.py
```

O `pywin32` é usado somente para integração com o Outlook. Caso ele não esteja disponível, os rascunhos de e-mail não poderão ser abertos, mas o processamento das planilhas continua funcionando.

## Fluxo de uso

1. Abra o aplicativo por `gerador_escalas_desktop.py`.
2. Selecione a Base 2468 e as grades Sportv, PPV/Premiere e Combate que serão conferidas.
3. Execute a etapa de cruzamento para gerar `Check_Pre_Envio_Gerado.xlsx`.
4. Revise as linhas sinalizadas por `Conferir Pré`, `Horário não encontrado na Grade` ou `Fallback (Multimodalidade)`.
5. Acompanhe o processamento na aba **Log de Execução**; ela permite copiar, limpar ou salvar o log sem reduzir a área das etapas.
6. Use a etapa de geração de rascunhos somente após a conferência da planilha.

## Regras incorporadas

| Regra | Comportamento |
|---|---|
| GE TV | Excluído deste fluxo, pois possui envio separado. |
| GE.com | Consulta a grade principal do Sportv para programas associados, como PANELA. |
| Conteúdo `V/I = I` | Evento é encontrado; sem Pré válido, recebe `Conferir Pré`. |
| Programas sem Pré | Permanecem `OK` quando encontrados na grade. |
| PPV com Pré-Hora | A linha de Pré separada é associada ao próximo evento da mesma data/canal. |
| Eventos repetidos | São consolidados até o início do próximo evento distinto. |
| Datas brasileiras | Datas `DD/MM/YYYY` são interpretadas corretamente no matching. |

## Testes

Execute a validação automatizada com:

```bat
py -m unittest discover -s . -p "test_*.py" -q
```

Os testes de regressão cobrem os cenários tratados nesta atualização, incluindo Sportv 2/3, GE.com, PPV com Pré separado, programas repetidos e eventos inéditos.

## Conteúdo da pasta

| Pasta/arquivo | Finalidade |
|---|---|
| `gerador_escalas_desktop.py` | Aplicativo principal, incluindo a aba exclusiva de Log de Execução. |
| `engine_*.py` | Motores de leitura, cruzamento e tratamento das grades. |
| `app_config.json` | Nomes e caminhos configuráveis da aplicação. |
| `test_*.py` | Suíte de testes automatizados. |
| `amostras_agosto_2026/` | Arquivos de agosto usados na validação e a saída gerada. |
| `RELATORIO_DE_CONCLUSAO.md` | Registro geral das melhorias anteriores. |
