# Gerador de Escalas - Motor de Cruzamento V2

Uma aplicação desktop em Python (Tkinter) desenvolvida para automatizar, cruzar e disparar e-mails de escalas para narradores, comentaristas e colaboradores de transmissões esportivas (SporTV, Premiere, Combate).

## Funcionalidades Principais

O sistema é dividido em duas abas principais, que separam o processo de formatação dos dados da comunicação real.

### Aba 1: Gerador de Escalas
Esta é a aba central (pipeline de cruzamento e envio).
* **Etapa 1: Motor de Cruzamento**
  * Absorve o relatório bruto de escalas (Base 2468).
  * Lê e achata as tabelas das grades de programação de cada canal (SporTV, Premiere, Combate).
  * Realiza o "match" automático corrigindo e inserindo os horários de **Pré**, **Início** e **Fim** baseando-se nos eventos da grade.
  * Trata multimodalidades, eventos extensos, alertas visuais para revisão (cores) e gera a planilha `Check_Pre_Envio_Gerado.xlsx`.
* **Etapa 2: Gerador de HTML e Disparo**
  * Recebe o Check revisado e formata as escalas de forma isolada por profissional.
  * Possui filtro nativo por período de datas.
  * **Grupos para Enviar**: Filtragem inteligente (ex: gerar apenas para Narradores).
  * **Gerenciar Exceções**: Interface que lista dinamicamente todos os nomes presentes na escala da semana, permitindo que você barre o envio individual de certos profissionais com apenas alguns cliques.
  * Conecta-se nativamente à API do Microsoft Outlook (COM object) para gerar os rascunhos (*Drafts*) com os HTMLs customizados contendo todas as formatações perfeitas.

### Aba 2: Gestão de Contatos
Um painel administrativo integrado (CRUD) para gerenciar a base de nomes e e-mails (`contatos_nova_versao.xlsx`).
* Não é necessário abrir o Excel manualmente. O painel lê a planilha na hora.
* Tabela interativa onde é possível visualizar Grupo, Nome e E-mail.
* Permite **Adicionar**, **Atualizar**, e **Excluir** contatos.
* Salva nativamente no formato suportado pela Etapa 2.

## Tratamento de Dados (Edge Cases Suportados)
A aplicação possui lógicas robustas de extração de horário ("Time formatting"):
* Se o Excel fornecer horários embutidos com datas completas (ex: `2026-07-16 08:30:00`), o sistema renderizará nos e-mails apenas o trecho útil `08:30`.
* Suporte nativo para janelas de eventos contínuos, escalas atravessando a madrugada (ex: `23:30` - `02:00`), dias da semana, e eliminação absoluta de formatações científicas do Pandas no momento de enviar ao HTML.

## Requisitos
* Windows OS (Obrigatório para o disparo no Outlook Desktop nativo via `win32com.client`)
* Microsoft Outlook configurado na máquina (Obrigatório para os disparos).
* Python 3.10+
* Módulos requeridos (disponíveis via pip):
  * `pandas`
  * `openpyxl`
  * `pywin32`
  * `tkcalendar`

## Como Iniciar

1. Clone o repositório.
2. Certifique-se de instalar as dependências:
   ```cmd
   pip install pandas openpyxl pywin32 tkcalendar
   ```
3. Inicie o sistema executando:
   ```cmd
   python gerador_escalas_desktop.py
   ```

## Notas de Uso
Todas as planilhas brutas ou com dados confidenciais estão ignoradas nativamente pelo `.gitignore`, protegendo bases privadas.
