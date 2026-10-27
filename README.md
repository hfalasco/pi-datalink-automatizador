# Automatizador PI DataLink

Preenche, em lote, a coluna de data/hora e as fórmulas `PIAdvCalcVal` do
PI DataLink em várias planilhas Excel — sem abrir o Excel e sem alterar os
arquivos originais.

Tem também uma segunda página, **Tirar fórmulas**, que faz o caminho de volta:
recebe a planilha já calculada e devolve uma cópia só com os números (seção 5-E).

---

## 1. O que o programa faz

Para cada arquivo da pasta de entrada:

1. Descobre as abas e, em cada uma, as colunas que têm **TAG válida** na linha 2
   — mais as colunas que combinam **duas ou mais TAGs** na mesma conta, como a
   vazão medida por dois canos (ver seção 5-C).
2. Gera os carimbos de tempo do período (por exemplo 00:00, 00:05 … 23:55).
3. Escreve na coluna A as datas como **valores reais de data/hora** do Excel.
4. Escreve em cada coluna com TAG a fórmula:

   ```
   =PIAdvCalcVal(B$2;$A3;$A4;"average (time-weighted)";"time-weighted";0;1;0;"")
   ```

   A TAG **nunca** entra como texto: a fórmula referencia a célula `B$2`, `C$2`,
   `D$2`… Trocar a TAG na linha 2 muda o resultado sem mexer nas fórmulas.
5. Salva uma **cópia** na pasta de saída, com o mesmo nome.

O Python **não calcula nada do PI**. Ele só escreve as fórmulas. Quem executa é
o Excel com o PI DataLink, quando você abre o arquivo.

---

## 2. Estrutura reconhecida

Confirmada nos três templates analisados (Brix do clarificado fábrica,
Caldo Misto Destilaria e Vazão Clarificado Fábrica):

| Linha / coluna | Conteúdo |
|---|---|
| Linha 1 | Nome legível do indicador (`Vazão Clarificado Fábrica (m³/h)`) |
| Linha 2 | TAG do PI (`\\azr-pida130.ad.dom\USM-FT_421016`); `A2` = rótulo `data` |
| Linha 3 em diante | `A` = data/hora, demais colunas = fórmula PI |
| Coluna A | Carimbos de tempo |

Tudo isso é **configurável** em `config.py` (`LINHA_TAGS_PADRAO`,
`PRIMEIRA_LINHA_DADOS_PADRAO`, `COLUNA_DATA_PADRAO`). A página usa esses
valores como estão — planilhas com layout diferente pedem um ajuste ali.

### Quantas linhas são geradas

Para 1 dia com intervalo de 5 minutos: **288 fórmulas** (linhas 3 a 290) e
**289 datas** (linhas 3 a 291). A última data existe apenas para servir de
*fim* do último intervalo — exatamente como nos templates originais.

### Como uma TAG é reconhecida

Uma coluna só é processada se a célula da linha 2:

- for **texto** (não número, não vazio);
- não for um rótulo como `data`, `hora`, `timestamp`.

Isso reproduz o comportamento correto no caso real da aba *Indicadores por
hora* do Caldo Misto, onde `B2 = 0` (número) e a coluna B, de propósito, não
recebe fórmula.

Uma coluna **sem** TAG na linha 2 ainda assim é preenchida quando a fórmula que
já está nela combina TAGs de outras colunas — é o caso da vazão somada de dois
canos, explicado na seção 5-C.

Para aceitar **somente** caminhos PI, use em `config.py`:

```python
padrao_tag = r"^\\\\"
```

---

## 3. Por que não openpyxl

Um teste de ida e volta com openpyxl nos seus arquivos reais perdeu estas
partes do pacote:

```
xl/metadata.xml                     (propriedades de matriz dinâmica)
xl/webextensions/taskpanes.xml      (painel do PI DataLink)
xl/webextensions/webextension1.xml
xl/printerSettings/printerSettings1.bin
xl/worksheets/_rels/sheet1.xml.rels
```

e também os atributos `aca="1"`, `ca="1"` e `cm="1"` das fórmulas de matriz —
sendo que `ca="1"` é justamente o sinalizador de *recalcular sempre* que o
PI DataLink usa.

Este programa usa **edição direta do XML** do `.xlsx`: copia todas as partes
byte a byte e reescreve apenas o `<sheetData>` das abas-alvo.

| | openpyxl | Excel via COM | **XML direto (usado aqui)** |
|---|---|---|---|
| Preservação | perde 6 partes + flags | total | **total** |
| Velocidade | rápida | 3–10 s por arquivo | **~50 ms por arquivo** |
| Precisa do Excel instalado | não | **sim** | não |
| Diálogos / travas | — | sim | não |

Resultado medido nos seus três arquivos: a única parte removida é
`xl/calcChain.xml`, e isso é **proposital** — é um cache da ordem de cálculo;
mantê-lo desatualizado depois de inserir células é o que provoca o aviso
"conteúdo ilegível" do Excel. Removido, o Excel o reconstrói em silêncio.

### O que mais muda no arquivo

- `xl/workbook.xml` ganha `fullCalcOnLoad="1"`, para o Excel recalcular ao abrir.
- Os valores em cache `#NAME?` (gravados numa máquina sem o PI DataLink) são
  removidos junto com as fórmulas antigas. Assim o Excel calcula de verdade em
  vez de exibir o erro antigo.
- `xl/styles.xml` só é tocado quando uma aba **não tinha** formato de data —
  aí um formato `dd/mm/yyyy hh:mm` é acrescentado ao final, sem alterar nenhum
  estilo existente.
- `xl/metadata.xml` é criada quando o arquivo não tem essa parte. É ela que
  permite a **matriz dinâmica**; sem ela o Excel trava o bloco de fórmulas.
  Ver *Matriz dinâmica × matriz travada*, na seção 5-B.

---

## 4. Instalação

```bash
python --version
```

Precisa ser 3.9 ou superior. Depois, na pasta do projeto:

```bash
pip install -r requirements.txt
```

Isso instala o **Streamlit**, que é a interface do programa. O motor continua
usando só a biblioteca padrão do Python — a única dependência obrigatória é a
tela. (`openpyxl` é opcional, para conferir os arquivos gerados por fora.)

A versão do Streamlit está **fixa** no `requirements.txt`, de propósito: com
`>=` o servidor da nuvem instalava a versão mais nova enquanto o computador de
quem desenvolve ficava numa antiga, e a página só quebrava depois de publicada.
Para atualizar, troque o número, rode o `pip install` de novo, teste localmente
e só então publique.

---

## 5. Como executar

Na pasta do projeto:

```bash
streamlit run app.py
```

O navegador abre sozinho em `http://localhost:8501`. Se não abrir, copie o
endereço que aparece no terminal. Para encerrar, feche a aba e pressione
`Ctrl+C` no terminal.

### Passo a passo na página

A página tem duas partes: a **barra lateral**, à esquerda, com os módulos e o
que muda de rodada para rodada (período e opções), e o **miolo**, com os três
passos na ordem em que são feitos.

Na barra lateral:

- **Módulos**
  - **Processar planilhas** e **Tirar fórmulas** — os links das duas páginas
    do programa. O da página aberta fica em destaque (branco, com o texto
    verde). Período e opções continuam como estavam quando você vai para a
    outra página e volta; os arquivos enviados, não — ao voltar, envie de
    novo.
  - **Planilha base** — baixa o modelo em branco das nossas planilhas
    (`planilha_base_pi_datalink.xlsx`), com cada célula da estrutura
    destacada numa cor:

    | Célula | Conteúdo | Cor |
    |---|---|---|
    | `A2` | `Data` — rótulo da coluna de datas | azul |
    | `B1` | `Nome` — nome do indicador | verde |
    | `B2` | `Tag` — TAG do PI | dourado |

    Troque `Nome` e `Tag` pelos valores reais; para mais indicadores, repita o
    par nas colunas C, D, E… A coluna A, da linha 3 para baixo, fica vazia — é
    o programa que preenche as datas e as fórmulas. A barra lateral mostra uma
    miniatura da planilha com as mesmas cores.
- **Período** — data inicial, data final e intervalo.
- **Opções** — *Buscar tudo de uma vez* (modo rápido), *Abrir sem calcular*,
  *Matriz dinâmica* e *Manter contas da planilha*. Todas já vêm marcadas nos
  valores recomendados.

No miolo:

1. **Arquivos** — arraste as planilhas `.xlsx`/`.xlsm` para a caixa (ou clique
   em *Browse files*). Pode enviar várias de uma vez.
2. **Abas** — deixe *Processar todas as abas* marcado, ou desmarque e escolha
   as abas na lista. Para dar período próprio a alguma aba, marque *Período
   próprio para alguma aba* (seção 5-A).
3. **Executar** — *Testar 1 arquivo* mostra no *Registro* o que seria feito, sem
   gravar nada; **PROCESSAR** roda o lote, com barra de progresso e registro
   acompanhando.

Entre os passos 2 e 3 ficam as quatro cartelas de números — arquivos, abas com
TAG, linhas por arquivo e consultas ao PI do arquivo mais pesado. Elas mudam na
hora a cada data, intervalo ou opção trocada.

No fim, **Baixar arquivos processados (.zip)** — o `.zip` traz as cópias
processadas e a pasta `logs/` com o registro completo da execução.

### Onde ficam os arquivos

Os arquivos enviados vão para uma pasta temporária própria da sessão. Os
originais no seu computador **nunca são tocados** — o que a página recebe já é
uma cópia. O resultado só existe dentro do `.zip` que você baixa.

Cada clique em **PROCESSAR** refaz a saída do zero, então o `.zip` traz sempre
só o resultado daquela rodada.

> **Interromper no meio:** a página processa o lote de uma vez, sem botão de
> cancelar. Para interromper, pare o `streamlit run` no terminal (`Ctrl+C`) —
> nada terá sido gravado no seu computador.

---

## 5-A. Período diferente para cada aba

Um mesmo arquivo pode ter a aba **por hora** de 5 em 5 minutos e a aba
**diária** de um em um dia. Mexer em uma **não altera** a outra.

### Na página

Na caixa **Abas**, depois de enviar os arquivos, marque *Período próprio para
alguma aba*. Aparece uma tabela com uma linha por aba encontrada e as colunas
*Data inicial*, *Data final* e *Intervalo (min)*.

A tabela fica atrás dessa caixa de marcar de propósito: ela usa o componente
mais pesado da página, e quase toda rodada usa o mesmo período em todas as abas.

1. Clique na célula da aba que quer tratar diferente (ex.: `Indicadores Diários`).
2. Preencha **só o que muda**. Campo em branco herda o valor geral da caixa
   **Período**. Para uma aba diária, basta pôr `1440` no intervalo.
3. Pronto — o valor vale no próximo *Testar 1 arquivo* ou *PROCESSAR*.

Para desfazer, apague o conteúdo da célula: a aba volta ao período geral.

Datas em `dd/mm/aaaa`. Atalhos de intervalo: `60` = 1 hora, `720` = 12 horas,
`1440` = 1 dia.

O nome da aba é comparado **sem acento e sem diferenciar maiúsculas**, então
`Indicadores Diarios` encontra tanto `Indicadores Diários` quanto
`Indicadores diários`.

### Resultado

Com o período geral de 01/10/2026 a 07/10/2026 a cada 5 minutos e `1440` no
intervalo da aba `Indicadores Diários`, no mesmo arquivo:

| Aba | Linhas de data | Fórmulas | Passo |
|---|---|---|---|
| Indicadores por hora | 2017 | 2016 | 5 minutos |
| Indicadores Diários | 8 | 7 | 1 dia |

---

## 5-B. Desempenho: por que a planilha abria devagar

O programa grava os arquivos em ~50 ms. A demora acontecia **dentro do Excel**,
e a causa era o número de consultas ao servidor PI.

Uma fórmula `PIAdvCalcVal` devolve **um** valor e faz **uma** ida e volta à
rede. Na aba "Indicadores por hora" da Vazão isso dava:

```
12 TAGs x 288 linhas = 3.456 consultas separadas ao PI
```

O PI DataLink tem a função irmã `PIAdvCalcDat`, que devolve **todos os
intervalos de uma vez**. As assinaturas, lidas do próprio add-in instalado
(`C:\Program Files\PIPC\Excel\pipc64.xll`), diferem por um único argumento:

```
PIAdvCalcVal(data item, stime, etime,           modestring, weightmode, minpctgood, cfactor, outcode, root path)
PIAdvCalcDat(data item, stime, etime, interval, modestring, weightmode, minpctgood, cfactor, outcode, root path)
```

Com isso, uma fórmula por **coluna** substitui 288 fórmulas por coluna:

| | Modo linha (antigo) | Modo bloco (padrão) |
|---|---|---|
| Função | `PIAdvCalcVal` | `PIAdvCalcDat` |
| Fórmulas na coluna B | 288 | 1 (matriz `B3:B290`) |
| Consultas ao PI (aba por hora) | 3.456 | **12** |

A fórmula gravada fica assim:

```
=PIAdvCalcDat(B$2;$A$3;$A$291;"5m";"average (time-weighted)";"time-weighted";0;1;0;"")
```

O intervalo acompanha o da aba: `5m`, `1h`, `1d`. Se a aba diária estiver
configurada para 1440 min, a fórmula dela sai com `"1d"`.

### Matriz dinâmica × matriz travada

Como a fórmula agora cobre a coluna inteira, ela vira uma **matriz**. O Excel
não deixa alterar pedaços de uma matriz — daí a mensagem:

> Não é possível alterar parte de uma matriz.

Por isso o programa grava **matriz dinâmica** por padrão: a fórmula mora só em
`B3` e o resultado se expande sozinho para baixo. O bloco nasce e some junto
com `B3`, o que evita o aviso no uso normal.

As suas **datas (coluna A)** e **TAGs (linha 2)** ficam fora da matriz, então
continuam totalmente editáveis nos dois casos.

| | Matriz dinâmica (padrão) | Matriz CSE |
|---|---|---|
| Onde fica a fórmula | só em `B3` | `B3:B290` |
| Para limpar a coluna | apague `B3` | selecione o bloco inteiro |
| Compatibilidade | Excel 365 / 2021+ | qualquer versão |

#### O que fazia a mensagem continuar aparecendo

Para o Excel entender a fórmula como matriz **dinâmica**, a célula precisa do
atributo `cm="N"`, que aponta para um registro `XLDAPR` dentro de
`xl/metadata.xml`. Antes o programa só reaproveitava o `cm` que já existisse no
arquivo — e cerca de **metade dos templates não tem essa parte**:

```
com xl/metadata.xml : Moenda A, Moenda B, Moenda C, Tratamento de Caldo, Pol da torta
sem xl/metadata.xml : Concentração de Caldo, Destilação, Fábrica de Açúcar,
                      Secagem de Levedura, SAFRA_COMPLETA
```

Nesses arquivos não havia `cm` para herdar, o programa caía para a matriz CSE e
a planilha saía com o bloco `B3:B290` travado. Como as abas de 5 em 5 minutos
são as de bloco grande, era nelas que a mensagem aparecia.

Agora `xl/metadata.xml` é **criada quando não existe** (e completada quando
existe sem o tipo `XLDAPR`), junto com as declarações em `[Content_Types].xml`
e em `xl/_rels/workbook.xml.rels`. O índice de `cm` é calculado a partir do
arquivo, e não chutado: se `XLDAPR` for o segundo tipo da lista, sai `cm="2"`.
Ver `core/metadata.py`.

Resultado: a matriz CSE só aparece se você pedir por ela.

Para forçar a forma clássica, desmarque **Matriz dinamica** na caixa
*Opções*.

Se precisar mesmo editar célula a célula, o caminho é desmarcar **Buscar tudo
de uma vez**, que volta a uma fórmula independente por célula.

### Abrir sem calcular

Além disso, os arquivos passam a ser gravados em **modo de cálculo manual**
(`calcMode="manual"`). O arquivo abre na hora e nada é consultado até você:

- apertar **F9**, ou
- usar **PI DataLink > Refresh**.

Assim dá para abrir vários arquivos e só então disparar a busca.

### Como desligar

Na página, na caixa **Opções**:

- **Buscar tudo de uma vez (muito mais rapido)** — desmarque para voltar a
  `PIAdvCalcVal`, exatamente como era antes.
- **Abrir sem calcular (buscar com F9 no Excel)** — desmarque para o arquivo
  calcular sozinho ao abrir.

A página e o log mostram quantas consultas cada modo geraria, para você comparar.

---

## 5-C. Colunas que somam duas TAGs (vazão medida por mais de um cano)

Às vezes a vazão de um equipamento é medida por **dois** transmissores e a
coluna mostra a soma dos dois. É o caso de *Template Pol da torta composta.xlsx*
(pasta **Filtração do Lodo**), na aba *Indicadores de Horas*:

```
J3 = PIAdvCalcVal($AA$2;$A3;$A4;...) + PIAdvCalcVal($AB$2;$A3;$A4;...)
K3 = PIAdvCalcVal($Y$2 ;$A3;$A4;...) + PIAdvCalcVal($Z$2 ;$A3;$A4;...)
L3 = IFERROR(N(PIAdvCalcVal($AC$2;...));0) + IFERROR(N(PIAdvCalcVal($AD$2;...));0)
```

Repare no arranjo:

| | |
|---|---|
| `J2`, `K2`, `L2` | **vazias** — a coluna não tem TAG própria |
| `Y2` … `AD2` | as TAGs de verdade, em colunas de apoio fora da área visível |
| linha 1 de `Y`…`AD` | vazia — não são colunas de indicador |

### O que o programa faz

Na primeira linha de dados ele lê a fórmula que já está na planilha e a
classifica:

- **coluna comum** — uma chamada só, lendo a TAG da própria coluna
  (`PIAdvCalcVal(B$2;...)`). Recebe a fórmula padrão, como sempre.
- **coluna com conta própria** — qualquer outra coisa: duas ou mais chamadas,
  `IFERROR`, `N`, `AVERAGE`, fator de conversão, filtro (`PIAdvCalcFilVal`), TAG
  lida de outra coluna. A conta é **reaproveitada como está** e o programa troca
  **apenas o início e o fim** de cada consulta para as novas linhas.
- **coluna de apoio** — tem TAG na linha 2, **não** tem fórmula própria e é
  citada pela conta de outra coluna (`Y`…`AD`). Fica **intacta**. Vale também
  para a célula que guarda a **expressão de filtro** do `PIAdvCalcFilVal`, que
  também não é coluna de dados.

Antes disso, `J`, `K` e `L` saíam em branco (sem TAG na linha 2 o programa nem
olhava para elas) e `Y`…`AD` viravam colunas de dados, gastando 6 consultas por
linha à toa.

As TAGs continuam morando nas células da linha 2: trocar `AA2` continua mudando
o resultado de `J`, exatamente como antes.

### Modo bloco nessas colunas

Essas colunas **também vão em matriz**, como o resto do arquivo. Cada
`PIAdvCalc...Val` vira a irmã `PIAdvCalc...Dat`, que devolve a coluna inteira de
uma vez:

```
=PIAdvCalcDat($AA$2;$A$3;$A$291;"5m";...)+PIAdvCalcDat($AB$2;$A$3;$A$291;"5m";...)

=IFERROR(N(PIAdvCalcDat($AC$2;$A$3;$A$291;"5m";...));0)
+IFERROR(N(PIAdvCalcDat($AD$2;$A$3;$A$291;"5m";...));0)
```

Somar duas matrizes do mesmo tamanho devolve outra matriz, e `IFERROR`, `N` e
companhia trabalham **valor a valor** — a conta continua valendo linha por linha
e a coluna passa a fazer **uma consulta por TAG** no lugar de uma por linha.

As quatro famílias têm par, e as assinaturas foram lidas do próprio add-in
instalado (`C:\Program Files\PIPC\Excel\pipc64.xll`):

| Uma linha | Coluna inteira | Diferença |
|---|---|---|
| `PIAdvCalcVal` | `PIAdvCalcDat` | `interval` no 4º argumento |
| `PIAdvCalcFilVal` | `PIAdvCalcFilDat` | idem — o filtro continua no 5º |
| `PIAdvCalcExpVal` | `PIAdvCalcExpDat` | idem |
| `PIAdvCalcExpFilVal` | `PIAdvCalcExpFilDat` | idem |

#### A exceção: funções que resumem a coluna

`AVERAGE`, `SUM`, `MAX`, `MIN`… **recebem** um conjunto e devolvem **um número**.
Uma coluna montada assim — é o caso das planilhas da Moenda —

```
O3 = AVERAGE(PIAdvCalcFilVal($O$2;...);PIAdvCalcFilVal($P$2;...))
```

em matriz viraria a média do **período inteiro**, num valor só, em vez de uma
média por linha. Por isso essas colunas continuam com **uma fórmula por linha**,
e o log avisa:

```
Por hora: coluna(s) O, Q, S, U, W, Y ficam com uma formula por linha porque a
conta usa AVERAGE, que resume a coluna inteira em um numero so
```

Se a conta for mesmo a média de duas TAGs, dá para trocar `AVERAGE(a;b)` por
`(a+b)/2` na planilha: aí ela é aritmética pura e passa a caber na matriz. A
diferença aparece quando uma das TAGs volta sem dado — `AVERAGE` ignora o texto
`[-11059] No Good Data` e usa só a outra, enquanto `(a+b)/2` devolve erro.

Dá para forçar o comportamento dessas colunas em `config.py`, no campo
`modo_composta` do `Config`: `auto` (padrão, acompanha o modo geral), `linha`
(essas colunas linha a linha mesmo no modo bloco) ou `bloco` (matriz mesmo com
o resto do arquivo no modo linha). Nem `bloco` põe em matriz uma coluna com
`AVERAGE`: o programa nunca grava uma fórmula que devolveria um número no lugar
da coluna.

### Como desligar

Na caixa **Abas**, desmarque **Manter contas da planilha (colunas que somam
duas ou mais TAGs)**. Aí o programa volta ao comportamento antigo: fórmula padrão em toda coluna com TAG na linha 2 —
inclusive nas de apoio — e nada é preservado.

### O que aparece no teste e no log

```
[X] Indicadores de Horas           23 TAG(s)  (3 coluna(s) somando mais de uma TAG)

COLUNAS COM CONTA PROPRIA (a formula da planilha e reaproveitada;
so o inicio e o fim de cada consulta sao trocados):
  J3 (matriz, uma consulta por TAG)
    AA2      \\azr-pida130.ad.dom\USM-FT_415301
    AB2      \\azr-pida130.ad.dom\USM-FT_415302
    =PIAdvCalcDat($AA$2;$A$3;$A$27;"1h";...)+PIAdvCalcDat($AB$2;$A$3;$A$27;"1h";...)

COLUNAS DE APOIO (so guardam a TAG das contas acima - ficam intactas):
  Y2 = \\azr-pida130.ad.dom\USM-FT_415402
```

### O filtro das planilhas da Moenda

As planilhas de Moenda usam `PIAdvCalcFilVal`, que só considera o tempo em que
uma condição é verdadeira — a expressão fica numa célula da linha 2
(`AD2 = "'…USM-STS_MA_STOP'=""Moendo"""`), fora da área visível.

Antes o programa trocava essas fórmulas pela padrão `PIAdvCalcVal`/`Dat` e
**o filtro se perdia** — a média passava a incluir o tempo de moenda parada — e
ainda escrevia fórmula na coluna do filtro (`AD`). Agora a função filtrada é
preservada (virando `PIAdvCalcFilDat` no modo bloco) e `AD` fica intacta.

### Limites

O programa só reescreve as oito funções da tabela acima. Se a coluna usar outra
função do PI (`PIArcVal`, `PICompDat`…) ou se o início/fim da consulta for um
texto fixo em vez de uma célula (`"*-1h"`), ele **não mexe na coluna** — melhor
deixar como está do que reescrever errado.

Colunas com conta que **não consulta o PI** (por exemplo a aba *CALCULO AUX*,
com `AVERAGEIFS` sobre a aba por hora) também ficam como estão: o programa não
estende nem encurta essas fórmulas quando o período muda. O log avisa:

```
Por dia: coluna(s) F tem conta que o programa nao ajusta (conta que nao
consulta o PI) - confira se o numero de linhas ainda bate
```

---

## 5-D. O que a página faz para não cair

O Streamlit roda a página **inteira** (`app.py` e a página aberta) a cada clique, tecla digitada ou
data trocada. Com a versão anterior, cada um desses gestos:

- reabria todas as planilhas enviadas — duas vezes, uma para listar as abas e
  outra para estimar as consultas;
- montava a lista completa de carimbos de tempo só para escrever quantos eram
  (um ano de 1 em 1 minuto são 525.600 datas, ~26 MB, jogados fora em seguida).

E, durante o processamento, **cada linha do log** redesenhava a caixa do
registro, ou seja, uma mensagem pela conexão por linha escrita, carregando todo
o texto acumulado até ali. Era o que travava a aba do navegador e derrubava a
página nos lotes grandes.

O que mudou:

| Antes | Agora |
|---|---|
| Planilhas reabertas a cada interação | Lidas **uma vez** por conjunto enviado; trocar data ou opção não abre arquivo nenhum |
| Lista de datas montada a cada interação | Só a **conta**, sem montar lista |
| Uma mensagem de tela por linha de log | Tela atualizada no máximo a cada 0,4 s, com as últimas 400 linhas |
| `.zip` do resultado na memória **e** na sessão | `.zip` escrito em disco; a sessão guarda só o caminho |
| Pasta de saída mantida junto com o `.zip` | Apagada assim que o `.zip` fica pronto |
| Pastas temporárias acumulando para sempre | Apagadas depois de 6 h paradas |
| Envio de até 500 MB por arquivo | 50 MB — o suficiente com folga para um template (a página *Tirar fórmulas* aceita 500 MB, ver 5-E) |

Medido com 20 planilhas e período de um ano de 1 em 1 minuto: o custo de trocar
uma data caiu de cerca de **1,4 s** (1 s reabrindo arquivos + 0,4 s montando
datas) para menos de **1 ms**.

### Limites que a página impõe

- **Acima de 100.000 linhas por arquivo** aparece um aviso: a planilha fica
  pesada e o processamento pode faltar memória.
- **Acima de 1.048.576 linhas** a página recusa a rodada — é o limite de linhas
  do próprio Excel, o arquivo gerado não abriria.

Em ambos os casos a saída é a mesma: aumentar o intervalo ou dividir o período
em mais de uma rodada.

---

## 5-E. Tirar fórmulas (deixar só os números)

A página **Tirar fórmulas** (link em *Módulos*, na barra lateral) é a versão em
página do antigo `tirar_formulas.py`: você envia a planilha, ela troca cada
fórmula pelo valor que o Excel já calculou e devolve uma cópia só com os
números, com o sufixo `_SEM_FORMULAS` no nome (`Safra.xlsx` →
`Safra_SEM_FORMULAS.xlsx`).

### Antes de enviar

O programa usa o valor que o Excel **gravou no arquivo** — ele não calcula
nada. Então a planilha precisa ter sido:

1. aberta no Excel, com o PI DataLink atualizado;
2. **salva** (`Ctrl+S`).

Fórmula sem valor gravado vira célula vazia. O registro avisa quantas ficaram
assim em cada arquivo (`ATENCAO: ... ficaram VAZIAS`), e a cartela *Fórmulas
trocadas* também. Uma planilha que acabou de sair da página *Processar
planilhas* ainda não foi calculada — abra, atualize e salve antes.

### Na página

1. **Arquivos** — arraste as planilhas (até **500 MB** cada).
2. **Executar** — **TIRAR FORMULAS** converte os arquivos **um de cada vez**,
   com barra de progresso e registro.

Com **um** arquivo, o botão de download entrega a própria planilha. Com
**vários**, entrega um `.zip` com as planilhas e a pasta `logs/`.

### O que muda no arquivo

Só as fórmulas. Diferente do script antigo, que montava uma planilha nova com
openpyxl e levava só valor, formato de número, mesclagem, largura e painel
congelado, aqui o arquivo é copiado parte por parte (o mesmo princípio da
seção 3) e todo o resto fica como estava: cores, fontes, bordas, gráficos,
formatação condicional, abas ocultas.

| Célula antes | Depois |
|---|---|
| Fórmula com número gravado | o número |
| Fórmula com texto gravado | o texto |
| Fórmula com erro gravado (`#N/A`) | o erro, como valor |
| Fórmula sem valor gravado | célula vazia (o formato fica) |
| Matriz do PI (`PIAdvCalcDat` em B3, derramando até B290) | os números de B3 a B290 |

Também sai o `xl/calcChain.xml`, o cache da ordem de cálculo — sem fórmulas ele
apontaria para contas que não existem mais.

### Arquivos grandes

As planilhas calculadas chegam a centenas de MB, e o script antigo abria cada
uma **duas vezes** com openpyxl, uma delas inteira na memória. Aqui cada aba é
lida em pedaços de 8 MB, então a memória usada não depende do tamanho do
arquivo. Medido numa aba de 337 MB de XML com 2,4 milhões de fórmulas:
**13,5 s e 75 MB de memória**. Numa planilha de 20.000 linhas, o script levou
8 s e a página 1 s — com os mesmos valores, célula a célula.

O limite de envio desta página é 500 MB por arquivo (`LIMITE_ENVIO_MB` em
`paginas/tirar_formulas.py` e `maxUploadSize` em `.streamlit/config.toml`). O
que pesa agora é o próprio Streamlit, que guarda na memória o arquivo enquanto
ele está na caixa de envio: num servidor pequeno, como a nuvem gratuita, poucos
envios desse tamanho ao mesmo tempo já apertam.

---

## 6. Testar primeiro com 1 arquivo

Envie **um** arquivo, escolha-o em *Arquivo usado no teste* e clique em
**Testar 1 arquivo**. Nada é gravado; o *Registro* mostra:

```
================================================================
MODO DE TESTE - nenhum arquivo sera alterado ou criado
================================================================
Arquivo analisado: Template Brix do clarificado fabrica.xlsx
Linhas de data: 289   |   Linhas de formula: 288

[X] Indicadores Diarios - 2 TAG(s)
    Periodo: 11/09/2026 a 11/09/2026, a cada 5 min
      B2 = \\azr-pida130.ad.dom\NN_USM-EM-039006
      C2 = \\azr-pida130.ad.dom\NN_USM-EM-039001
    Formulas que serao inseridas:
      B3 =PIAdvCalcDat(B$2;$A$3;$A$291;"5m";"average (time-weighted)";...)  (uma so, em matriz)

[ ] Resumo - ignorada (nenhuma TAG valida na linha 2)
Nada foi gravado em disco.
```

Depois de aprovar, clique em **PROCESSAR** e abra o arquivo do `.zip` no Excel
para conferir que o PI DataLink preenche os valores.

---

## 7. Processar vários arquivos de uma vez

A caixa de envio aceita **várias planilhas juntas** — arraste todas, ou
selecione-as de uma vez no *Browse files*. Elas são processadas em sequência e
saem no mesmo `.zip`, cada uma com o **mesmo nome** que tinha.

São ignorados automaticamente os temporários `~$arquivo.xlsx` e qualquer coisa
que não seja `.xlsx`/`.xlsm` — a própria caixa de envio já recusa os outros
formatos. Um erro em um arquivo **não interrompe** os demais: ele aparece como
`Status: ERRO` no registro e os outros seguem.

Desempenho medido: ~50 ms por arquivo. Centenas de arquivos levam segundos — o
que demora é o envio pelo navegador, não o processamento.

---

## 8. Separador da fórmula: vírgula ou ponto e vírgula

Dentro do `.xlsx` o separador é **sempre vírgula** — é o que o padrão OOXML
exige, e é o que o próprio PI DataLink grava. O Excel em português **exibe**
ponto e vírgula. As duas formas abaixo são o mesmo conteúdo:

```
no arquivo:  PIAdvCalcVal(B$2,$A3,$A4,"average (time-weighted)",...)
na tela:    =PIAdvCalcVal(B$2;$A3;$A4;"average (time-weighted)";...)
```

Não altere isso manualmente.

---

## 9. Log

Tudo que a execução escreve vai para `logs\log_AAAAMMDD_HHMMSS.txt`, dentro do
`.zip` baixado. O **Registro da execução** na tela mostra as últimas 400 linhas
desse mesmo texto — o navegador não aguenta desenhar mais do que isso de uma vez,
e o arquivo tem sempre o registro inteiro:

```
PROCESSAMENTO INICIADO
Data: 14/09/2026 10:43:09
Periodo: 11/09/2026 a 11/09/2026  |  intervalo: 5 min
Arquivos encontrados: 3

Arquivo: Template Caldo Misto Destilaria.xlsx
Abas processadas: Indicadores por hora, Indicadores Diarios
TAGs encontradas: 7
Linhas geradas: 288
Status: OK

PROCESSAMENTO FINALIZADO
Sucesso: 3
Erros: 0
```

---

## 10. Compartilhar a página com outras pessoas

Por padrão o `streamlit run app.py` atende só o seu computador. Para deixar
outras máquinas da rede abrirem a mesma página:

```bash
streamlit run app.py --server.address 0.0.0.0
```

O terminal passa a mostrar também um *Network URL* (`http://10.x.x.x:8501`) —
é esse endereço que os colegas usam. O computador que roda o comando precisa
ficar ligado, e o firewall do Windows costuma pedir liberação na primeira vez.

Como cada pessoa envia e baixa os próprios arquivos, uma sessão não enxerga a
outra.

> Não existe mais o `.exe`: quem usa a página não precisa instalar nada além do
> navegador — só o computador que a hospeda precisa de Python.

### Publicar na nuvem (Streamlit Community Cloud)

O campo **Main file path** é `app.py`. Não há mais nada a configurar: o
`requirements.txt` e o `.streamlit/config.toml` já estão no repositório.

Dois pontos antes de publicar:

- **O endereço é público.** Qualquer pessoa com o link abre a página e pode
  enviar planilhas, mesmo que o repositório seja privado. Restrinja os
  espectadores por e-mail nas configurações do app, em *Settings › Sharing*.
- **As planilhas passam por um servidor de terceiros.** Cada arquivo enviado é
  processado na infraestrutura do Streamlit, não na sua rede. Se os dados forem
  sensíveis, prefira rodar na rede interna, como descrito acima.

---

## 11. Estrutura do projeto

```
pi_datalink_automatizador/
    app.py                  ponto de entrada: tema, marca, módulos e navegação
    comum.py                o que as duas páginas usam igual (envio, registro, .zip)
    paginas/
        processar.py        página Processar planilhas
        tirar_formulas.py   página Tirar fórmulas (seção 5-E)
    tema.py                 paleta de cores e CSS (mesma identidade visual do AutoFill)
    assets/
        logo_sao_martinho.png   logotipo do topo da barra lateral e ícone da aba
    config.py               toda a configuração e validação
    logger.py               log em arquivo e na tela
    requirements.txt
    README.md
    .streamlit/
        config.toml         tema do Streamlit e limite de envio
    core/
        xml_utils.py        leitura/escrita do pacote OOXML, conversão de colunas
        timeline.py         geração dos carimbos e conversão para serial do Excel
        formula.py          montagem da fórmula PIAdvCalcVal
        composta.py         colunas que combinam mais de uma TAG (ver 5-C)
        inspector.py        leitura da estrutura (abas, TAGs, estilos) - só leitura
        styles.py           criação segura de formato de data quando não existe
        writer.py           motor de escrita (cirurgia no XML)
        processor.py        orquestração pasta -> arquivos -> abas
        modelo.py           planilha base em branco (módulo da barra lateral)
        sem_formulas.py     troca as fórmulas pelos valores gravados (seção 5-E)
```

Cada módulo tem uma responsabilidade só. Trocar a fórmula do PI mexe apenas em
`core/formula.py` e em `config.py`; mudar a estrutura da planilha mexe apenas
em `config.py`; mudar a aparência mexe apenas em `tema.py`,
`.streamlit/config.toml` e `assets/`.

A pasta do projeto é a **raiz da aplicação**, não um pacote: os módulos são
importados direto dela (`import tema`, `from core import processor`). É de
propósito — ao publicar, o servidor clona o repositório numa pasta com o nome
do repositório, e um nome com hífen (`pi-datalink-automatizador`) não seria um
nome de pacote válido. Assim nada depende de como a pasta se chama.

O `core/` não conhece a interface — foi por isso que trocar a janela Tkinter
pela página não exigiu tocar em nenhum módulo dele.

---

## 12. Problemas comuns

| Sintoma | Causa | Solução |
|---|---|---|
| `Erro: arquivo aberto no Excel` | existe `~$nome.xlsx` na pasta | feche o arquivo no Excel |
| `nenhuma aba com TAG válida na linha 2` | layout diferente | ajuste `LINHA_TAGS_PADRAO` / `PRIMEIRA_LINHA_DADOS_PADRAO` em `config.py` |
| A caixa de envio recusa o arquivo | formato `.xls`/`.xlsb`, não suportado | abra no Excel e salve como `.xlsx` |
| Células mostram `#NOME?` no Excel | PI DataLink não está carregado | habilite o suplemento em *Arquivo › Opções › Suplementos* |
| Valores em branco ao abrir | modo de cálculo manual (proposital) | pressione **F9** ou use *PI DataLink › Refresh* |
| `#SPILL!` numa coluna | algo ocupa as células da matriz | limpe a coluna abaixo da linha 3 e reprocesse |
| *Não é possível alterar parte de uma matriz* | a fórmula cobre a coluna inteira | normal no modo bloco; apague `B3` para limpar tudo, ou desmarque *Buscar tudo de uma vez* |
| Coluna traz data em vez de valor | `PIAdvCalcDat` devolveu timestamps | desmarque *Buscar tudo de uma vez* e avise, para ajustar a matriz |
| A página não abre no navegador | o `streamlit run` não está de pé | rode `streamlit run app.py` e use o endereço mostrado no terminal |
| A página cai ou reinicia sozinha | período grande demais para a memória do servidor | aumente o intervalo, divida em rodadas ou veja a seção 5-D |
| `File must be 50.0MB or smaller` | planilha acima do limite de envio | aumente `maxUploadSize` em `.streamlit/config.toml` |

---

## 13. Licença

Uso interno, todos os direitos reservados. Ver [LICENSE](LICENSE).

PI System, PI DataLink e PI Server são marcas de seus respectivos titulares.
Este projeto não é afiliado a eles nem por eles patrocinado.
