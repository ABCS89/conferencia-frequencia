# Guia de Implementação do Dashboard no Power BI — Monitoramento de Frequência

Este documento orienta passo a passo como montar o painel de **Monitoramento de Atrasos (Minutos Perdidos) e Faltas Acumuladas** no Power BI, a partir dos dados gerados pelo sistema de conferência de frequência.

---

## 1. Origem dos Dados (Duas Formas de Conectar)

O sistema oferece **duas alternativas** para conectar os dados no Power BI:

### Opção A (Recomendada — Conexão Direta e Rápida)
O script `src/export_powerbi.py` (executado automaticamente pelo `main.py`) já consolida todos os arquivos de todas as secretarias e meses em um arquivo único:
- **Arquivo Excel**: `output/powerbi/monitoramento_consolidado.xlsx`
- **Ou arquivos CSV**: `output/powerbi/fato_atrasos.csv`, `output/powerbi/fato_faltas.csv` e `output/powerbi/resumo_secretarias.csv`

**Como importar no Power BI:**
1. No Power BI Desktop, clique em **Obter Dados** > **Pasta de Trabalho do Excel**.
2. Selecione o arquivo: `.../conferencia-frequencia/output/powerbi/monitoramento_consolidado.xlsx`.
3. Marque as 3 abas:
   - `Resumo_Secretarias`
   - `Fato_Atrasos`
   - `Fato_Faltas`
4. Clique em **Carregar**.

---

### Opção B (Conexão Dinâmica na Pasta via Power Query M)
Se preferir que o Power BI leia diretamente a pasta `output` e combine automaticamente os arquivos `monitoramento_*.xlsx` de qualquer mês:

1. No Power BI, clique em **Obter Dados** > **Pasta**.
2. Aponte para o caminho da pasta: `.../conferencia-frequencia/output`.
3. Clique em **Transformar Dados**.
4. No Editor do Power Query, clique em **Página Inicial** > **Editor Avançado** e cole a consulta abaixo para cada tabela:

#### Consulta M: `Fato_Atrasos`
```powerquery
let
    Fonte = Folder.Files("C:\Users\abusilva\Desktop\Github\conferencia-frequencia\output"),
    ApenasExcelMonitoramento = Table.SelectRows(Fonte, each Text.Contains([Name], "monitoramento_") and Text.EndsWith([Name], ".xlsx") and not Text.Contains([Folder Path], "powerbi")),
    AdicionarAbaAtrasos = Table.AddColumn(ApenasExcelMonitoramento, "DadosAtrasos", each let
        excel = Excel.Workbook([Content], true),
        aba = Table.SelectRows(excel, each Text.Contains([Item], "Atraso") or Text.Contains([Item], "Minuto"))
    in
        if Table.IsEmpty(aba) then null else aba{0}[Data]),
    FiltrarNulos = Table.SelectRows(AdicionarAbaAtrasos, each [DadosAtrasos] <> null),
    ExpandirColunas = Table.ExpandTableColumn(FiltrarNulos, "DadosAtrasos", 
        {"Matrícula", "Nome", "Minutos Acumulados", "Equivalência", "Datas / Ocorrências", "Situação no Sistema"},
        {"Matricula", "NomeServidor", "MinutosAcumulados", "EquivalenciaTexto", "OcorrenciasDatas", "SituacaoSistema"}
    ),
    ExtrairAnoMes = Table.AddColumn(ExpandirColunas, "AnoMes", each Text.BetweenDelimiters([Folder Path], "\output\", "\monitoramento\"), type text),
    ExtrairSecretaria = Table.AddColumn(ExtrairAnoMes, "Secretaria", each Text.BetweenDelimiters([Name], "monitoramento_", ".xlsx"), type text),
    RemoverOutrasColunas = Table.SelectColumns(ExtrairSecretaria, {"AnoMes", "Secretaria", "Matricula", "NomeServidor", "MinutosAcumulados", "EquivalenciaTexto", "OcorrenciasDatas", "SituacaoSistema"}),
    Tipagem = Table.TransformColumnTypes(RemoverOutrasColunas, {
        {"MinutosAcumulados", Int64.Type},
        {"AnoMes", type text},
        {"Secretaria", type text},
        {"Matricula", type text},
        {"NomeServidor", type text}
    })
in
    Tipagem
```

---

## 2. Estrutura do Modelo de Dados

### Tabelas Principais:
1. **`Resumo_Secretarias`**:
   - `AnoMes`: identificador do mês (ex.: `2026-08`)
   - `CodigoSecretaria`: código numérico (ex.: `103`, `107`)
   - `Secretaria`: nome da secretaria
   - `OrgaoCompleto`: código e nome formatados
   - `TotalMinutosPerdidos`: volume de minutos perdidos
   - `TotalHorasPerdidas`: total em horas decimais (`Minutos / 60`)
   - `TotalDiasFalta`: volume de dias de falta acumulados
   - `QtdServidoresAtraso`: servidores com minutos perdidos
   - `QtdServidoresFalta`: servidores com faltas

2. **`Fato_Atrasos`**:
   - `AnoMes`, `CodigoSecretaria`, `Secretaria`, `OrgaoCompleto`
   - `Matricula`, `NomeServidor`
   - `MinutosAcumulados`, `HorasEquivalentes`, `EquivalenciaTexto`
   - `OcorrenciasDatas`: datas específicas com os minutos de cada dia
   - `SituacaoSistema`: `Lançado no sistema` ou `Sem registro no sistema`

3. **`Fato_Faltas`**:
   - `AnoMes`, `CodigoSecretaria`, `Secretaria`, `OrgaoCompleto`
   - `Matricula`, `NomeServidor`
   - `DiasFalta`, `TiposFalta`
   - `OcorrenciasDatas`: datas e períodos das faltas
   - `SituacaoSistema`: `Lançado no sistema` ou `Sem registro no sistema`

---

## 3. Medidas DAX Essenciais

Crie uma nova tabela de medidas no Power BI chamada `_Medidas` e adicione as fórmulas DAX abaixo:

### Totais e Horas
```dax
Total Minutos Perdidos = 
SUM(Fato_Atrasos[MinutosAcumulados])
```

```dax
Total Horas Perdidas = 
DIVIDE([Total Minutos Perdidos], 60, 0)
```

```dax
Total Horas Perdidas Formatado = 
VAR _Horas = INT([Total Horas Perdidas])
VAR _Minutos = INT(MOD([Total Minutos Perdidos], 60))
RETURN
IF(_Horas > 0, _Horas & "h " & _Minutos & "min", _Minutos & "min")
```

```dax
Total Dias de Falta = 
SUM(Fato_Faltas[DiasFalta])
```

### Contagens de Servidores
```dax
Qtd Servidores com Atraso = 
DISTINCTCOUNT(Fato_Atrasos[Matricula])
```

```dax
Qtd Servidores com Falta = 
DISTINCTCOUNT(Fato_Faltas[Matricula])
```

### Médias por Servidor
```dax
Média Minutos por Servidor = 
DIVIDE([Total Minutos Perdidos], [Qtd Servidores com Atraso], 0)
```

```dax
Média Faltas por Servidor = 
DIVIDE([Total Dias de Falta], [Qtd Servidores com Falta], 0)
```

### Indicadores de Conformidade (Auditória de Lançamento)
```dax
Minutos Sem Registro no Sistema = 
CALCULATE(
    [Total Minutos Perdidos],
    Fato_Atrasos[SituacaoSistema] = "Sem registro no sistema"
)
```

```dax
Faltas Sem Registro no Sistema = 
CALCULATE(
    [Total Dias de Falta],
    Fato_Faltas[SituacaoSistema] = "Sem registro no sistema"
)
```

---

## 4. Proposta de Layout e Telas do Dashboard

### 📌 Tela 1: Visão Executiva & Ranking de Secretarias (Tela Principal)
*Objetivo: Permitir aos gestores identificar instantaneamente os maiores gargalos de pontualidade e assiduidade em todas as secretarias do município.*

1. **Barra Superior de Filtros / Slicers**:
   - Segmentador de **Mês de Referência** (`AnoMes` ou `MesReferencia`).
   - Segmentador de **Secretaria / Órgão** (`OrgaoCompleto`).
2. **Cards / KPIs no Topo (Cartões Grandes)**:
   - **Card 1**: Total de Horas Perdidas (`Total Horas Perdidas Formatado`).
   - **Card 2**: Total de Dias de Falta (`Total Dias de Falta`).
   - **Card 3**: Servidores com Atraso (`Qtd Servidores com Atraso`).
   - **Card 4**: Servidores com Falta (`Qtd Servidores com Falta`).
3. **Visuais Centrais (Rankings Principais)**:
   - **Gráfico de Barras Horizontais 1 (Esquerda)**:
     - *Título*: **Ranking de Secretarias — Maior Volume de Horas Perdidas**
     - *Eixo Y*: `Secretaria`
     - *Eixo X*: `Total Horas Perdidas`
     - *Dica de Ferramenta*: `Total Minutos Perdidos`, `Qtd Servidores com Atraso`
   - **Gráfico de Barras Horizontais 2 (Direita)**:
     - *Título*: **Ranking de Secretarias — Maior Número de Faltas (Dias)**
     - *Eixo Y*: `Secretaria`
     - *Eixo X*: `Total Dias de Falta`
     - *Dica de Ferramenta*: `Qtd Servidores com Falta`
4. **Visual Inferior de Conformidade**:
   - **Gráfico de Rosca / Donut**:
     - *Legenda*: `SituacaoSistema` (`Lançado no sistema` vs. `Sem registro no sistema`)
     - *Valores*: `Total Minutos Perdidos` (ou `Total Dias de Falta`)
     - *Objetivo*: Mostrar a proporção de ocorrências já lançadas no sistema vs. ocorrências pendentes de inserção pelo RH.

---

### 📌 Tela 2: Detalhamento por Servidor & Auditoria
*Objetivo: Consulta nominal detalhada para RH e chefias acompanharem os servidores e cobrarem justificativas.*

1. **Filtros e Caixas de Busca**:
   - Segmentadores de `Mês`, `Secretaria` e `Situação no Sistema`.
   - Campo de pesquisa de texto por `Matrícula` ou `Nome do Servidor`.
2. **Top 10 Servidores com Maiores Atrasos (Gráfico de Barras)**:
   - *Eixo Y*: `NomeServidor`
   - *Eixo X*: `MinutosAcumulados`
3. **Top 10 Servidores com Mais Faltas (Gráfico de Barras)**:
   - *Eixo Y*: `NomeServidor`
   - *Eixo X*: `DiasFalta`
4. **Tabela Analítica Completa (Grade com Detalhes)**:
   - Colunas: `Matrícula`, `Nome do Servidor`, `Secretaria`, `Total Horas/Minutos`, `Total Faltas`, `Datas / Ocorrências Detalhadas`, `Situação no Sistema`.
   - Formatação condicional: Destacar em amarelo/vermelho servidores com mais de 300 minutos ou mais de 5 dias de falta.

---

## 5. Como Atualizar os Dados Mensalmente
Sempre que rodar a conferência mensal (`python src/main.py`):
1. O script gera os relatórios individuais na pasta `output/<mes>/monitoramento/`.
2. O script executa automaticamente a consolidação, atualizando `output/powerbi/monitoramento_consolidado.xlsx`.
3. No Power BI Desktop, basta clicar no botão **Atualizar** (Refresh) na barra superior e todos os gráficos e rankings serão recalculados instantaneamente!
