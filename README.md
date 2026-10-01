# Conferência de Frequência — Secretarias Municipais

Confere, mês a mês, o relatório de frequência gerado pela **secretaria**
("Frequência dos Funcionários") contra o relatório gerado pelo **sistema**
("Relatório Ocorrência Geral"), e gera uma lista pronta de retificações:
"nesta data, o tipo lançado deveria ser outro".

O código e nome da secretaria são lidos diretamente do cabeçalho oficial do PDF
(ex: `116 GUARDA CIVIL DO MUNICIPIO DE PIRACICABA`, `103 PROCURADORIA GERAL`),
e os relatórios finais são salvos com essa identificação clara.

## Estrutura

```
conferencia-frequencia/
  .venv/                    # ambiente virtual
  input/
    secretaria/COD - secretaria.pdf  # PDF que a secretaria manda (ex: 116 - secretaria.pdf)
    sistema/COD - sistema.pdf        # PDF que o sistema gera (ex: 116 - sistema.pdf)
  output/
    conferencia_<COD - NOME DA SECRETARIA>.xlsx      # aba "Retificações" + aba "Resumo"
    retificacoes_<COD - NOME DA SECRETARIA>.md       # só se houver retificação
    retificacoes_<COD - NOME DA SECRETARIA>.txt      # idem, texto puro
  src/
    extract_secretaria.py   # lê o PDF de texto corrido da secretaria e extrai cabeçalho
    extract_sistema.py      # lê o PDF em tabela do sistema
    compare.py              # compara dia a dia dentro do mês de referência
    main.py                 # processa em lote ou individualmente e gera Excel + md + txt
  requirements.txt
```

## Uso Mensal

### Modo 1: Processamento Automático em Lote (Recomendado)

Basta salvar os PDFs na pasta `input/`:
- `input/secretaria/COD - secretaria.pdf`
- `input/sistema/COD - sistema.pdf`

E rodar no terminal (com a `.venv` ativada):
```bash
python src/main.py
```
O script encontra todos os pares disponíveis automaticamente, extrai o nome de cada secretaria e o mês de referência, e processa tudo de uma única vez.

### Modo 2: Processamento Individual

Se quiser rodar apenas um par específico:
```bash
python src/main.py "input/secretaria/116 - secretaria.pdf" "input/sistema/116 - sistema.pdf"
```
Se quiser forçar um mês específico, utilize o parâmetro `--mes AAAA-MM`:
```bash
python src/main.py "input/secretaria/116 - secretaria.pdf" "input/sistema/116 - sistema.pdf" --mes 2026-07
```

### Saídas Geradas em `output/`:
- `conferencia_<COD - NOME DA SECRETARIA>.xlsx` — Planilha Excel com aba **Retificações** (colorida por tipo de erro) + aba **Resumo**.
- `retificacoes_<COD - NOME DA SECRETARIA>.md` e `.txt` — **Só gerados se houver retificação**, já prontos para copiar e colar em e-mails/chamados.

## Como a comparação funciona (dia a dia)

Em vez de somar a quantidade total de dias por tipo de ocorrência, o
script:

1. **Expande** cada ocorrência de cada relatório em dias individuais
   (ex.: "Férias regulamentares, 02/04 a 16/04" vira 15 dias marcados
   como "Férias regulamentares"), recortando para dentro do mês de
   referência.
2. Para cada matrícula, **percorre dia a dia** o mês inteiro comparando
   o tipo lançado pela secretaria com o tipo lançado pelo sistema
   naquele dia específico.
3. Quando os dois batem (inclusive quando os dois não têm nada = dia
   normal), não gera nada. Quando divergem, cria uma retificação.
4. **Agrupa dias consecutivos** com o mesmo "de → para" num único
   intervalo (é por isso que "28 a 30/04, 3 dias" sai como uma linha só,
   não três).

Isso resolve sozinho o problema de eventos que também ocupam dias de
outro mês (ex.: férias que começaram em março): nos dias que caem dentro
de abril, o tipo bate normalmente entre os dois relatórios, mesmo que a
duração total do evento (contando os dias de março) não bata — então
não é mais preciso calcular ou sinalizar essa diferença separadamente.

Há 3 "sabores" de retificação (cores diferentes no Excel):

- **Amarelo — SEM REGISTRO → tipo do sistema**: a secretaria não tinha
  nada lançado nesse dia (ou só "Frequência normal"), mas o sistema tem
  uma ocorrência. Precisa lançar na secretaria.
- **Azul — tipo da secretaria → sem registro em sistema**: a secretaria
  lançou algo, mas o sistema não tem nada para aquele dia. Vale conferir
  se é pra lançar no sistema ou se a secretaria lançou errado.
- **Laranja — tipo A → tipo B**: os dois têm algo lançado, mas tipos
  diferentes (o caso mais comum é "Aguardando perícia sempem" na
  secretaria virando "Tratamento De Saúde" no sistema).

## Ajustando o "de-para" de tipos de ocorrência

Se aparecer um tipo de ocorrência escrito de formas diferentes nos dois
relatórios mas que na prática é o mesmo tipo (tipo "Falta" na secretaria
virando "Faltas Efetivos" no sistema, que eu já mapeei — sem isso o
script ia comparar dia a dia e ver como se fossem tipos diferentes),
adicione o par em `ALIASES`, no topo de `src/compare.py`:

```python
ALIASES = {
    "falta": "faltas efetivos",
    # "outro tipo escrito diferente": "tipo equivalente no sistema",
}
```
(As chaves/valores devem estar em minúsculas e sem acento — é assim que
o código normaliza antes de comparar.)

## O que o script NÃO consegue pegar sozinho

Se a classificação correta de um dia não existe em **nenhum dos dois**
relatórios (por exemplo, você sabe por um atestado à parte que um dia
deveria ser "Doença em Pessoa da Família", mas nem a secretaria nem o
sistema têm isso registrado), o script só consegue apontar "secretaria
tem X, sistema não tem nada" — ele não tem como adivinhar qual é o tipo
correto quando essa informação vem de fora dos dois PDFs. Nesses casos a
retificação sai como "... --> sem registro em sistema" e cabe a você
completar manualmente com a informação que só você tem.

## Limitações conhecidas

- O PDF da secretaria vem em texto corrido (não é uma tabela de verdade),
  e o rodapé/marca d'água de algumas páginas vaza como "linhas não
  reconhecidas" nos logs — isso é normal e não afeta os dados extraídos
  (já são descartadas). Se quiser conferir, rode
  `./venv/bin/python src/extract_secretaria.py input/secretaria/AAAA-MM.pdf`
  direto e olhe a lista impressa.
- A comparação depende de a matrícula ser idêntica nos dois relatórios
  (ela é, nos dois modelos vistos até agora — `NN.NNN-N`).
- Para ocorrências de vários dias, a secretaria só informa a data de
  início + quantidade — o script assume que os dias são corridos a
  partir dali (`data + quantidade - 1`). Isso bateu em todos os casos
  conferidos até agora, mas se algum tipo de ocorrência não for
  contínuo dessa forma, pode gerar um recorte errado.
