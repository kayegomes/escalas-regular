# GSD — Escalas Regular

**Projeto:** Teste escalas regular
**Repositório:** [kayegomes/escalas-regular](https://github.com/kayegomes/escalas-regular)
**Branch principal:** `main`
**Última atualização deste GSD:** 08/09/2026
**Idioma de trabalho:** Português do Brasil
**Objetivo:** manter uma memória persistente, verificável e operacional do sistema desktop Python/Tkinter usado para validar e enviar internamente escalas.

> Este documento deve ser lido antes de qualquer nova alteração relevante. Ele registra não apenas o estado desejado, mas também as regressões já identificadas e as diferenças entre o diretório de trabalho local e o código efetivamente publicado.

## 1. Visão geral do sistema

O sistema é uma aplicação desktop em Python/Tkinter para validação e envio interno de escalas. O fluxo possui três etapas principais:

| Etapa | Entrada principal | Resultado |
|---|---|---|
| Etapa 1 | Relatório 2405 ou `WorkOrdersExport` e grades Sportv, PPV/Premiere e Combate | Checagem de existência dos eventos da grade no relatório, com período limitado ao intervalo do relatório. |
| Etapa 2 | Relatório 2468 e grades Sportv, PPV/Premiere e Combate | `Check_Pre_Envio_Gerado.xlsx`, com horários encontrados, alertas e `Elenco`. |
| Etapa 3 | Check/relatório processado e contatos | HTMLs individuais e rascunhos Outlook. |

A interface inclui abas independentes para as etapas, gestão de contatos e **Log de Execução**, que não deve voltar a ficar escondido na tela principal quando a janela for pequena.

## 2. Estado funcional consolidado

As seguintes regras foram definidas e devem ser preservadas:

| Regra | Comportamento esperado |
|---|---|
| `OK` | Evento/programa encontrado na grade sem alerta pendente; programas podem ficar OK sem Pré. |
| `Conferir Pré` | Evento ou conteúdo transmissivo/inédito (`V/I = V` ou `I`) encontrado sem Pré válido. Conteúdo `VT no Controle` inédito continua sinalizado para conferência. |
| `Pré igual ao Início` | O Pré e o Início do evento ficaram iguais; deve indicar conferência manual, exceto quando a nova linha separada de `PRÉ-HORA` tiver sido corretamente associada. |
| `Horário não encontrado na Grade` | Não houve correspondência confiável na grade. Não usar o horário do relatório para declarar que o evento foi encontrado. |
| `A Confirmar` | Refere-se ao próprio evento da grade marcado como `A CONFIRMAR`, não a um horário que esteja descrito como a confirmar. |
| `Fallback (Multimodalidade)` | Match encontrado em modalidade com maior risco de variação de nomenclatura, principalmente surfe e tênis; exige revisão. |
| `Mudança de Canal` | Mesmo evento encontrado no mesmo dia em canal interno diferente, principalmente entre Sportv 2/3/4; o horário é considerado encontrado, mas deve ser alertado. |
| `GE TV` | Excluir da Etapa 2 porque o envio desse canal é separado. |
| `GE.com` | Tratar como equivalente à grade Sportv quando o evento corresponder, com alerta apenas quando aplicável. |
| Quick Holds | Incluir marketing/institucional, Podcast e Cabine do Jogo; excluir compromisso pessoal/particular. |
| Folgas | Manter no Check e nos HTMLs, normalizar para `FOLGA` e mostrar `-` em Pré, Início e Fim. |
| Elenco | Agrupar por WO, data, plataforma, evento e janela de Início/Fim; excluir da coluna o próprio profissional que recebe aquela escala. |
| Local no HTML | Usar prioritariamente `Local de Locução`; manter campos antigos somente como fallback. |
| Virada de dia | Remover linhas vazias de transição; preservar atividades reais que cruzem meia-noite. |
| Período da Etapa 1 | Filtrar grades mensais ao intervalo mínimo/máximo do relatório 2405/`WorkOrdersExport`. |

## 3. Arquivos centrais

| Arquivo | Responsabilidade |
|---|---|
| `engine_2405.py` | Etapa 1, aliases de relatório, datas BR/ISO, equivalências de canal, confrontos, filtro de período e alerta de mudança de canal. |
| `engine_2468.py` | Normalização do 2468, exclusão de GE TV, Quick Holds e agrupamento de Elenco por janela. Contém `_team_key`. |
| `engine_grades.py` | Leitura e normalização das grades Sportv, PPV/Premiere e Combate; consolidação de janelas, Pré/Aquecimento e eventos repetidos. |
| `engine_cross.py` | Cruzamento do 2468 com as grades, definição de `Status Revisão`, gravação do Check e criação da aba `Legenda`. |
| `gerador_escalas_desktop.py` | GUI Tkinter, Etapa 3, HTML, contatos, aba Log de Execução e rascunhos Outlook. |
| `app_support.py` | Configuração externa, histórico de execução, nomes seguros, validação de e-mail e fallback de gravação. |
| `app_config.json` | Parâmetros externos e nomes de saídas. Não codificar caminhos variáveis diretamente no código. |
| `tests/test_regressions.py` | Regressões de matching, canais, datas, PPV, folgas, HTML e outros casos reais. |
| `CHANGELOG.md` | Histórico resumido das alterações publicadas. |
| `GUIA_DE_ENTREGA.md` | Orientações de instalação e operação. |
| `GSD_ESCALAS_REGULAR.md` | Este documento de memória persistente. |

## 4. Correções já realizadas

O histórico consolidado inclui as seguintes alterações implementadas ao longo do projeto:

1. Histórico/logs/configuração externa e robustez de gravação.
2. Match por data e plataforma, consolidação de eventos repetidos e alertas de Pré.
3. Exclusão de GE TV da Etapa 2.
4. Tratamento de GE.com como grade Sportv.
5. Equivalência entre canais internos Sportv, com alerta de `Mudança de Canal`.
6. Inclusão de Quick Holds de marketing/institucional, Podcast e Cabine do Jogo; exclusão de compromissos pessoais.
7. Etapa 3 com HTML contendo Elenco, Produto, Local de Locução, folgas, contatos, período da escala e limpeza de transições vazias.
8. Aba separada `Log de Execução` restaurada na GUI.
9. Etapa 1 compatível com 2405 e `WorkOrdersExport`, com grades limitadas ao período do relatório.
10. Correções de Etapa 1 para PANELA/GE.com, Grand Slam de Judô com canal diferente e confrontos como Remo x Flamengo.
11. A aba `Legenda` foi adicionada ao Check e deve ser recriada em toda execução.
12. A linha separada `PRÉ-HORA` do PPV/Premiere passou a ser associada por **data + mandante + visitante**, sem exigir igualdade entre os textos de canal `PRE/GE TV` e `PREMIERE`.

## 5. Caso real do Premiere validado

Na amostra de setembro, o jogo **Palmeiras x São Paulo**, em 12/09/2026, possui:

| Linha | Horário | Canal | Observação |
|---|---:|---|---|
| `PRÉ-HORA` | 17:30 | `PRE/GE TV` | Linha separada do Pré. |
| Evento principal | 18:30–20:40 | `PREMIERE` | Evento com mandante Palmeiras e visitante São Paulo. |

O resultado correto é **Pré 17:30, Início 18:30, Fim 20:40**, com `Status Revisão = OK` quando a linha principal for encontrada.

## 6. Regressão identificada em 08/09/2026

### Sintoma

Após a publicação do commit `4e4a7eb`, uma execução usando o clone/entrega publicada apresentou muitos registros adicionais como `Horário não encontrado na Grade`. A planilha original do usuário tinha poucos casos desse status.

### Evidência quantitativa

O arquivo original enviado pelo usuário, disponível localmente como `check_user.xlsx` e também como `/home/ubuntu/upload/Check_Pre_Envio_Gerado.xlsx`, apresentou:

| Status | Quantidade original |
|---|---:|
| `OK` | 499 |
| `Conferir Pré` | 44 |
| `Fallback (Multimodalidade)` | 26 |
| `Horário não encontrado na Grade` | 23 |
| `Pré igual ao Início` | 5 |

A execução com o código publicado incompleto apresentou aproximadamente:

| Status | Quantidade regressiva |
|---|---:|
| `OK` | 345 |
| `Conferir Pré` | 43 |
| `Fallback (Multimodalidade)` | 25 |
| `Horário não encontrado na Grade` | 199 |

A execução controlada com o diretório de trabalho local, que continha as correções Sportv ainda não sincronizadas, apresentou:

| Status | Quantidade após correções locais |
|---|---:|
| `OK` | 504 |
| `Conferir Pré` | 44 |
| `Fallback (Multimodalidade)` | 26 |
| `Horário não encontrado na Grade` | 23 |

A diferença entre 499 e 504 corresponde aos cinco registros PPV que deixaram de ficar como `Pré igual ao Início` depois da associação correta do `PRÉ-HORA`.

### Causa técnica

A causa não foi a chave de confronto do PPV isoladamente. O problema foi que o clone/ZIP publicado não continha correções anteriores que estavam no diretório de trabalho local, mas ainda não tinham sido sincronizadas com o GitHub.

No `engine_grades.py` local correto, a leitura da grade Sportv usa o número do bloco horizontal para determinar `SPORTV`, `SPORTV2`, `SPORTV3` e `SPORTV4`, sem interpretar indevidamente a coluna adjacente como canal. A lógica de `AQUECIMENTO`/Pré também encerra a sequência repetida correta do evento e evita que o cálculo posterior de fim substitua esse limite.

A versão publicada incompleta ainda possuía uma lógica antiga que:

- inferia o canal por faixas de índices menos adequadas;
- lia uma coluna adjacente como possível canal e descartava linhas quando o conteúdo não era um canal esperado;
- aplicava o limite do Pré somente à última ocorrência anterior, em vez de toda a sequência repetida do mesmo evento;
- permitia que o cálculo posterior de janelas substituísse o fim definido pelo AQUECIMENTO/Pré.

Isso alterava as janelas Sportv e fazia o cruzamento procurar eventos em horários diferentes dos presentes no Check original. Por isso a regressão atingiu muitos registros Sportv, não apenas os eventos Premiere.

### Regra de publicação aprendida

Antes de gerar ZIP ou publicar, comparar o clone Git com o diretório de trabalho completo. Nenhuma correção funcional deve permanecer apenas em `/home/ubuntu/github_escalas_regular` ou em `/home/ubuntu/escalas_regular/escalas_regular`. O teste mínimo de release deve reproduzir as contagens da amostra real antes e depois da alteração.

## 7. Evidências e arquivos da investigação

| Arquivo | Conteúdo |
|---|---|
| `/home/ubuntu/escalas_regular/amostra_setembro_12/check_user.xlsx` | Check original do usuário, com 612 linhas e 23 casos não encontrados. |
| `/home/ubuntu/upload/Check_Pre_Envio_Gerado.xlsx` | Cópia do Check original enviado. |
| `/home/ubuntu/escalas_regular/amostra_setembro_12/report_2468.xlsx` | Relatório 2468 da amostra. |
| `/home/ubuntu/escalas_regular/amostra_setembro_12/grade_sportv.xlsm` | Grade Sportv da amostra. |
| `/home/ubuntu/escalas_regular/amostra_setembro_12/grade_ppv.xlsx` | Grade PPV/Premiere da amostra. |
| `/home/ubuntu/escalas_regular/amostra_setembro_12/grade_combate.xlsx` | Grade Combate da amostra. |
| `/home/ubuntu/escalas_regular/amostra_setembro_12/compare_check_versions.py` | Comparação de status entre Check original e saída regenerada. |
| `/home/ubuntu/escalas_regular/amostra_setembro_12/verify_stage2_module_paths.py` | Prova de que as versões local e anterior carregam módulos de caminhos diferentes. |
| `/home/ubuntu/escalas_regular/amostra_setembro_12/compare_sportv_flatten_direct.log` | Teste direto mostrando que os parsers Sportv local e anterior geram as mesmas 2.265 linhas quando usados isoladamente. |
| `/home/ubuntu/escalas_regular/workorders_export_2026/Sistema_Escalas_WorkOrdersExport_periodo/codigo` | Entrega de 01/09 que reproduz as contagens originais da amostra. |

## 8. Testes e validações

Com o código local correto, foram executados:

```bash
python3 -m py_compile engine_cross.py engine_grades.py tests/test_regressions.py
PYTHONPATH=. python3 -m unittest discover -s tests -p 'test_*.py' -q
```

Resultado registrado: **42 testes executados, OK, 1 ignorado** no diretório local que contém a suíte ampliada.

No clone publicado antes da sincronização completa, foram executados 39 testes OK com 3 ignorados, mas essa execução não é suficiente para validar a amostra porque o clone estava sem correções funcionais locais de Sportv.

O teste sintético de PPV cobre a seguinte regra: uma linha `PRÉ-HORA` em `PRE/GE TV` deve ser associada à linha de evento em `PREMIERE` quando data, mandante e visitante forem iguais.

## 9. Próxima correção obrigatória

A próxima alteração deve ser feita no clone que será publicado e deve:

1. sincronizar do diretório local correto a lógica Sportv de `engine_grades.py`;
2. manter a chave PPV por data + mandante + visitante;
3. preservar a aba `Legenda`, incluindo `Mudança de Canal`;
4. manter as correções de `engine_2468.py`, especialmente `_team_key`;
5. executar a amostra real e exigir `Horário não encontrado na Grade = 23`, `Conferir Pré = 44` e `Fallback (Multimodalidade) = 26`;
6. exigir `OK = 504` após o PPV corrigido, ou explicar qualquer diferença real de dados;
7. executar a suíte completa;
8. somente depois atualizar o ZIP e publicar no GitHub.

## 10. Protocolo para futuras sessões

Ao iniciar uma nova sessão, ler este GSD e verificar:

```bash
cd /home/ubuntu/github_escalas_regular_git
git status --short
git log -5 --oneline
```

Depois comparar o clone publicado com o diretório local de trabalho, sem assumir que são iguais:

```bash
for file in engine_2405.py engine_2468.py engine_cross.py engine_grades.py gerador_escalas_desktop.py tests/test_regressions.py; do
  cmp -s /home/ubuntu/github_escalas_regular_git/$file /home/ubuntu/github_escalas_regular/$file || echo "DIFERENTE: $file"
done
```

Para validação de Etapa 2, usar as entradas reais da pasta `amostra_setembro_12` e comparar as contagens com o Check original. Não substituir a planilha original do usuário sem preservar uma cópia de referência.

Para qualquer modificação de matching, adicionar ou atualizar uma regressão antes de publicar. Não incluir no Git planilhas de entrada, Check operacional, logs, `__pycache__`, scripts de diagnóstico temporários ou dados pessoais.

## 11. Estado atual ao final deste registro

A correção do `PRÉ-HORA` por confronto/data está conceitualmente correta e resolve Palmeiras x São Paulo. A regressão observada decorre da publicação de um clone que não recebeu correções Sportv anteriores presentes no diretório local. O próximo passo técnico é sincronizar essas correções no clone Git, executar a validação quantitativa da amostra e somente então publicar a nova versão.

## 12. Atualização após a sincronização do clone

Esta seção complementa e, em caso de conflito, substitui a pendência descrita nas seções 9 e 11.

Em 08/09/2026, as correções funcionais que estavam apenas no diretório local foram sincronizadas no clone Git que será publicado:

| Arquivo | Correção sincronizada |
|---|---|
| `engine_grades.py` | Lógica correta dos blocos Sportv, determinação de canal por bloco horizontal, limites de AQUECIMENTO/Pré e associação PPV por confronto/data. |
| `engine_2468.py` | Chave de Elenco por WO, data, plataforma, evento e janela de horário. |
| `engine_cross.py` | Equivalência entre canais Sportv, alerta `Mudança de Canal` e aba `Legenda` formatada. |
| `GSD_ESCALAS_REGULAR.md` | Memória persistente do projeto e da regressão. |

A validação do clone sincronizado com a amostra real produziu:

| Status principal | Resultado validado |
|---|---:|
| `OK` | 504 |
| `Conferir Pré` | 44 |
| `Fallback (Multimodalidade)` | 26 |
| `Horário não encontrado na Grade` | 23 |

Os cinco registros do evento exato `PALMEIRAS X SÃO PAULO` foram validados com Pré `17:30`, Início `18:30`, Fim `20:40` e status `OK`. O filtro diagnóstico foi ajustado para não confundir outros eventos que contêm o texto São Paulo em seus nomes.

A causa da regressão publicada anteriormente foi confirmada: o clone tinha recebido a mudança PPV, mas não havia recebido simultaneamente correções anteriores de Sportv, equivalência de canais e matching do 2468. A mudança isolada de PPV não era responsável pelos 176 casos adicionais; a publicação parcial do código era.

No momento deste registro, a sincronização foi validada, mas ainda deve passar pela suíte completa, revisão final do diff, commit e push antes de ser considerada a versão oficial do GitHub.

## 13. Correção de 12/09/2026 — sequência consecutiva de AQUECIMENTO

Foi identificado e corrigido o caso de **Santos x Cruzeiro**, WO `2561032-1`, para Alline Calandrini. O relatório 2468 informa janela operacional de Início `20:00` até Fim `23:00`, enquanto a grade Sportv apresenta o evento principal às `21:00` e duas linhas consecutivas de `AQUECIMENTO SPORTV`, às `20:00` e `20:30`.

A regra correta é preservar o início da sequência de aquecimento como Pré do evento principal. Portanto, o resultado esperado e validado é:

| Profissional | Pré | Início | Fim | Status |
|---|---:|---:|---:|---|
| Alline Calandrini | 20:00 | 21:00 | 23:00 | `OK` |

A causa era a substituição do primeiro aquecimento pelo segundo enquanto o parser percorria linhas consecutivas. O parser agora mantém o primeiro horário da sequência por canal/data e só o substitui depois que o evento seguinte é processado ou quando a data muda.

Foi adicionada a regressão `test_consecutive_aquecimento_keeps_first_pre_time`, que reproduz duas linhas de AQUECIMENTO às 20:00 e 20:30 antes de Santos x Cruzeiro às 21:00. A suíte passou com 40 testes OK e 3 ignorados por arquivos/dependências opcionais.

## 14. Correção de 07/09/2026 — BDRJ/BOM DIA RIO sem grade TV Globo

Foi investigada a linha de **André Loffredo**, WO `2571311-1`, referente a `BDRJ - EXIBIÇÃO` / `BOM DIA RIO` em 07/09/2026. O relatório 2468 informa Início `06:00` e Fim `07:30`.

A grade normalizada da amostra não contém `BOM DIA RIO`, `BDRJ` ou uma linha TV Globo correspondente nesse dia. O horário `06:30–07:30` do Check anterior veio de um falso match: sem plataforma preenchida no relatório, o cruzamento escolheu `SPORTV3 — DIAMOND LEAGUE - 15ª ETAPA - DIA 1`, às 06:30, por coincidência da palavra genérica `DIA` e proximidade de horário. Não era uma confirmação real do evento.

A correção possui duas partes. Quando `Canal/Plataforma` estiver vazio, o motor passa a derivar `TV GLOBO` do campo `Cliente` quando ele contiver Globo. Além disso, `DIA` foi incluído entre as palavras genéricas, impedindo que um único termo comum crie um match entre programas diferentes.

O resultado validado para André Loffredo passou a ser:

| Pré | Início | Fim | Status |
|---:|---:|---:|---|
| `-` | `06:00` | `07:30` | `Horário não encontrado na Grade` |

A amostra passou a apresentar 503 registros `OK` e 24 `Horário não encontrado na Grade`, com os demais alertas preservados. Foram adicionadas as regressões `test_tv_globo_event_without_globo_grade_is_not_matched_to_sportv` e `test_generic_dia_does_not_create_event_match`.
