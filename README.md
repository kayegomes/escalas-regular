# Gerador de Escalas

Aplicação desktop em Python/Tkinter para validar, cruzar e preparar escalas de transmissões esportivas para conferência e envio manual no Outlook.

## Fluxo operacional

O sistema está organizado em duas abas principais. A aba **Gerador de Escalas** conduz as três etapas do processo, enquanto a aba **Gestão de Contatos** mantém a planilha de destinatários.

### Etapa 1 — Checagem 2405 versus grades

A aplicação lê o relatório de WOs independentes, lê as grades de programação dos canais informados, cruza os eventos válidos e gera `Check_2405_Gerado.xlsx`, com as abas `Checagem` e `Resumo`. A saída é colorida conforme o resultado da correspondência.

### Etapa 2 — Motor de cruzamento 2468 versus grades

O motor lê a Base 2468, achata as grades, realiza a correspondência automática e corrige os horários de Pré, Início e Fim. Também trata multimodalidades, eventos extensos, programas sem grade, madrugadas e alertas visuais. O resultado é `Check_Pre_Envio_Gerado.xlsx`.

Se a saída estiver aberta no Excel, o programa tenta gravar uma alternativa com sufixo `_novo`. Se o arquivo original e a alternativa estiverem bloqueados, a interface apresenta uma mensagem explicando que é necessário fechar o Excel.

Na revisão de horários, programas podem permanecer com status `OK` sem Pré, pois não necessariamente possuem bloco de Pré. Eventos esportivos com Início e Fim, mas sem horário de Pré válido, recebem status `Conferir Pré`, mesmo quando a grade marca o campo Pré como `X`. Produções classificadas como `VT no Controle` são tratadas como conteúdo transmissivo quando a linha correspondente da grade possui `V/I = I` (inédito) ou `V` (ao vivo); nesse caso, a ausência de Pré gera `Conferir Pré`. Linhas sem esse indicador permanecem sujeitas à classificação de programa até confirmação. Quando o mesmo evento aparece repetido em blocos consecutivos, o parser consolida a janela desde a primeira ocorrência até o próximo evento distinto. Se nenhum item correspondente for encontrado na grade, o status recebe `Horário não encontrado na Grade`, mesmo quando a atividade for um programa; nesse caso, os horários do relatório permanecem apenas como referência para conferência.

### Etapa 3 — Geração de HTML e preparação de rascunhos

A aplicação lê o check revisado, permite filtrar o período, seleciona grupos e permite configurar exceções de profissionais. Em seguida, gera um HTML por profissional na pasta `escalas_geradas_html`.

O modo **Somente Gerar HTML** não interage com o Outlook. O modo **Enviar Teste** exige um endereço de e-mail válido e prepara um rascunho de teste para esse endereço. O modo **Disparo Oficial** prepara rascunhos para os e-mails encontrados na planilha de contatos. **Nenhum modo envia automaticamente**: todas as mensagens são abertas para conferência humana no Outlook.

Contatos sem endereço válido são registrados como ignorados e não geram rascunhos incompletos.

## Histórico e logs

Cada execução concluída registra uma linha em `historico_execucoes.jsonl`, na pasta dos arquivos de entrada. O registro contém data e hora, etapa, status, resumo e detalhes relevantes. A aplicação também mantém `gerador_escalas.log` na pasta do programa, com rotação automática dos arquivos antigos. O botão **Salvar Log** continua disponível para exportar o conteúdo visível da interface.

A interface possui uma barra de progresso e atualiza o estado por etapas sem bloquear a janela principal. As operações demoradas continuam sendo executadas em segundo plano, enquanto as alterações visuais são encaminhadas com segurança para a interface Tkinter.

## Gestão de contatos

A aba de contatos mantém `contatos_nova_versao.xlsx`, na aba `Lista e-mails`, com as colunas `Grupo`, `Nome` e `Email`. É possível adicionar, alterar e excluir registros. Feche a planilha no Excel antes de salvar alterações.

## Requisitos

O uso completo requer Windows, Microsoft Outlook instalado e configurado, Python 3.10 ou superior e as dependências abaixo:

```cmd
pip install pandas openpyxl pywin32 tkcalendar
```

A aplicação e os testes utilitários podem ser importados em ambientes sem Outlook. Nesses ambientes, a abertura de rascunhos permanece indisponível, mas os motores e os testes das regras centrais continuam verificáveis.

## Como iniciar

Na pasta do projeto, execute:

```cmd
python gerador_escalas_desktop.py
```

Selecione os arquivos de entrada, execute as etapas na ordem recomendada e revise a planilha da Etapa 2 antes de preparar os HTMLs. No modo oficial, revise todos os rascunhos do Outlook antes de enviá-los manualmente.

## Como testar

Execute a suíte de regressão sem abrir a interface gráfica:

```cmd
python -m unittest discover -s . -p "test_*.py" -v
```

Os testes cobrem nomes de arquivo seguros, validação de e-mail, fallback para arquivos Excel bloqueados, histórico JSON Lines, normalização de canais, matching de eventos, regras de folga e quick hold e validação de horários.

## Estrutura dos arquivos principais

| Arquivo | Responsabilidade |
|---|---|
| `gerador_escalas_desktop.py` | Interface Tkinter, filtros, contatos, geração de HTML e rascunhos Outlook. |
| `engine_2405.py` | Checagem 2405 versus grades e resumo da primeira etapa. |
| `engine_cross.py` | Motor de cruzamento 2468 versus grades e geração do check de pré-envio. |
| `engine_grades.py` | Leitura e normalização das grades dos canais. |
| `engine_2468.py` | Leitura e preparação da Base 2468. |
| `app_support.py` | Histórico, validações, nomes seguros e gravação protegida. |
| `test_regressions.py` | Testes automatizados das regressões e utilitários principais. |

As planilhas de operação, HTMLs e logs podem conter dados sensíveis e devem ser tratados de acordo com as políticas internas da organização.
