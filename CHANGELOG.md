# Histórico de Alterações

## Atualização de agosto de 2026

### Interface de Log de Execução

A interface desktop agora possui a aba exclusiva **Log de Execução**. O log não ocupa mais a tela principal das etapas, permitindo trabalhar em telas menores sem perder a leitura das mensagens. A nova aba oferece rolagem vertical e horizontal, além de ações para copiar, limpar e salvar o conteúdo.

Esta atualização consolida as correções aplicadas durante a validação com as grades de agosto de 2026.

| Área | Atualização |
|---|---|
| Datas e horários | Datas brasileiras no formato `DD/MM/YYYY` passam a ser interpretadas corretamente; o matching usa `Início`/`Air Start Time` quando a coluna de data não contém hora. |
| Grades Sportv | O parser identifica `SPORTV`, `SPORTV2` e `SPORTV3` por linha e calcula janelas somente dentro do mesmo canal. |
| Eventos repetidos | Blocos repetidos são consolidados até o próximo evento distinto, incluindo os limites de TROCA DE PASSES e SPORTV NEWS. |
| Conteúdos inéditos | Ocorrências `V/I = I` são reconhecidas; se não houver Pré válido, recebem `Conferir Pré`. |
| PPV/Premiere | Linhas `PRÉ-HORA` são associadas ao próximo evento da mesma data/canal, evitando `Pré igual ao Início`. |
| Plataformas digitais | `GE.com` consulta a grade principal do Sportv para eventos associados; `GE TV` continua excluído deste fluxo por possuir envio separado. |
| Segurança do matching | Rótulos genéricos, como `VT DE EVENTO`, e ocorrências de outros dias não podem confirmar horários indevidamente. |
| Testes | A suíte de regressão cobre os cenários de canais Sportv, PPV, GE.com, eventos repetidos, Pré separado e datas brasileiras. |

## Execução de testes

```bash
python -m unittest discover -s tests -p "test_*.py" -q
```

Consulte o [guia de entrega](GUIA_DE_ENTREGA.md) para instalação e operação do aplicativo.

### Legenda dos Status Revisão

A planilha `Check_Pre_Envio_Gerado.xlsx` agora inclui uma aba `Legenda` com a descrição de `OK`, `Conferir Pré`, `Pré igual ao Início`, `Horário não encontrado na Grade`, `Fallback (Multimodalidade)`, `A Confirmar`, `Local Ausente` e `Sem Grades Fornecidas`, além da explicação para combinações de alertas.

### Cabeçalho de contato nos HTMLs

Cada escala HTML agora começa com uma saudação personalizada, a mensagem de envio da escala e o bloco de dúvidas ou problemas com os contatos de **Leticia Alvares — (21) 97951-2324** e **Carlla Amara — (21) 99242-1837**.

### Período da escala no cabeçalho

O cabeçalho de cada HTML agora informa automaticamente o intervalo da escala no formato `Escala consolidada: DD/MM/AAAA a DD/MM/AAAA`, calculado a partir das datas presentes na escala.

### Filtro de linhas vazias de virada de dia

A Etapa 3 agora remove dos HTMLs registros com data em formato de intervalo, como `10/08/2026 para 11/08/2026`, quando não possuem plataforma, evento, produto, local ou horário de atividade. Atividades reais com virada de madrugada continuam sendo preservadas.
