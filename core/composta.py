# -*- coding: utf-8 -*-
"""Colunas cuja formula NAO e a formula padrao de uma TAG so.

Caso real encontrado em "Template Pol da torta composta.xlsx"
(pasta Filtracao do Lodo): a vazao de um filtro e medida por DOIS canos, e a
coluna mostra a soma dos dois:

    J3 = PIAdvCalcVal($AA$2,$A3,$A4,...) + PIAdvCalcVal($AB$2,$A3,$A4,...)
    L3 = IFERROR(N(PIAdvCalcVal($AC$2,...)),0) + IFERROR(N(PIAdvCalcVal($AD$2,...)),0)

As duas TAGs moram em colunas auxiliares fora da area visivel (Y2:AD2) e a
celula da linha de TAGs da propria coluna (J2, K2, L2) fica VAZIA.

Sem este modulo o programa:
  * ignorava J, K e L (nao tem TAG na linha 2) - a coluna saia em branco;
  * preenchia Y..AD com formulas (tem TAG na linha 2) - colunas de apoio
    viravam colunas de dados e cada uma consultava o PI a toa.

A estrategia aqui e conservadora: em vez de tentar adivinhar a intencao, o
programa REAPROVEITA a formula que ja esta na planilha e troca apenas as
referencias de tempo (inicio e fim) pelas novas linhas. Assim a conta - soma,
IFERROR, fator de conversao, o que for - continua exatamente a que o usuario
montou, e as TAGs continuam vindo das celulas da linha 2.
"""

from __future__ import annotations

import re
from dataclasses import dataclass, field

# Funcoes do PI DataLink que este modulo sabe reescrever. Cada uma existe em
# dois sabores: "Val" devolve UM valor (uma formula por linha) e "Dat" devolve
# TODOS os intervalos de uma vez (uma formula por coluna, em matriz).
#
# Assinaturas lidas do proprio add-in instalado
# (C:\Program Files\PIPC\Excel\pipc64.xll):
#
#   PIAdvCalcVal      data item, stime, etime,           modestring, weightmode, minpctgood, cfactor, outcode, root path
#   PIAdvCalcDat      data item, stime, etime, interval, modestring, weightmode, minpctgood, cfactor, outcode, root path
#   PIAdvCalcFilVal   data item, stime, etime,           filtexp, modestring, weightmode, sampmode, sampfreq, ...
#   PIAdvCalcFilDat   data item, stime, etime, interval, filtexp, modestring, weightmode, sampmode, sampfreq, ...
#   PIAdvCalcExpVal   expression, stime, etime,           modestring, weightmode, sampmode, sampfreq, ...
#   PIAdvCalcExpDat   expression, stime, etime, interval, modestring, weightmode, sampmode, sampfreq, ...
#   PIAdvCalcExpFilVal expression, stime, etime,           filterexp, modestring, ...
#   PIAdvCalcExpFilDat expression, stime, etime, interval, filterexp, modestring, ...
#   PITimeFilterVal   exp, stime, etime,           timeunites, outcode, root path
#   PITimeFilter      exp, stime, etime, interval, timeunites, outcode, root path
#
# Em TODAS elas o inicio e o 2o argumento, o fim e o 3o, e a diferenca entre as
# duas versoes e exatamente um "interval" na 4a posicao. E por isso que o
# programa consegue trocar uma pela outra sem tocar no resto da conta - o
# filtro, o modo, o sampfreq: tudo continua como o usuario montou.
#
# A familia PITimeFilter foge do padrao em duas coisas, e as duas estao
# resolvidas na classe Assinatura:
#
#   * o nome e invertido - quem devolve UM valor leva o sufixo "Val", e quem
#     devolve a coluna inteira e a PITimeFilter, sem sufixo;
#   * o 1o argumento nao e uma TAG, e sim uma EXPRESSAO DE FILTRO. Ela pode vir
#     escrita na propria formula ('USM-XV_411010'="Aberta") ou, melhor, como
#     referencia a celula que a guarda (F$2) - e assim que o programa grava.
FUNCAO_VALOR = "PIAdvCalcVal"
FUNCAO_BLOCO = "PIAdvCalcDat"
FUNCAO_TEMPO = "PITimeFilterVal"
FUNCAO_TEMPO_BLOCO = "PITimeFilter"


@dataclass(frozen=True)
class Assinatura:
    """Como cada funcao do PI aparece nos dois modos de gravacao."""

    valor: str           # nome da versao "um valor por linha"
    bloco: str           # nome da versao "coluna inteira" (matriz)
    tem_intervalo: bool  # a funcao escrita na planilha ja recebe "interval"
    # True quando o 1o argumento e uma expressao de filtro, e nao uma TAG
    # (familia PITimeFilter). Muda so a leitura: a reescrita e a mesma.
    expressao: bool = False


def _par(valor: str, bloco: str, expressao: bool = False) -> dict:
    return {
        valor.lower(): Assinatura(valor, bloco, False, expressao),
        bloco.lower(): Assinatura(valor, bloco, True, expressao),
    }


# Funcoes do Excel que trabalham VALOR A VALOR. Numa formula de matriz elas se
# aplicam a cada linha do resultado, entao a conta continua valendo quando as
# chamadas do PI passam a devolver a coluna inteira:
#
#   IFERROR(N(PIAdvCalcDat(...)),0)  -> uma coluna de resultados
#
# Ja funcoes que RESUMEM um conjunto - AVERAGE, SUM, MAX, MIN... - fazem o
# contrario: recebem a coluna inteira e devolvem UM numero so. Uma coluna
# montada como AVERAGE(PIAdvCalcVal(A);PIAdvCalcVal(B)) so funciona linha a
# linha; em matriz ela viraria a media do periodo todo. Por isso qualquer
# funcao fora desta lista joga a coluna para o modo linha.
FUNCOES_VALOR_A_VALOR = frozenset({
    "IF", "IFERROR", "IFNA", "N", "ABS", "INT", "TRUNC", "SIGN", "VALUE",
    "ROUND", "ROUNDUP", "ROUNDDOWN", "SQRT", "EXP", "LN", "LOG", "LOG10",
    "POWER", "ISERROR", "ISERR", "ISNUMBER", "ISBLANK", "ISTEXT", "NOT",
})

# Nome de funcao dentro da formula: MAIUSCULA, minuscula ou com "_xll.".
RE_NOME_FUNCAO = re.compile(r"([A-Za-z_][A-Za-z0-9_.]*)\s*\(")

ASSINATURAS = {}
for _nomes in (
    ("PIAdvCalcVal", "PIAdvCalcDat"),
    ("PIAdvCalcFilVal", "PIAdvCalcFilDat"),
    ("PIAdvCalcExpVal", "PIAdvCalcExpDat"),
    ("PIAdvCalcExpFilVal", "PIAdvCalcExpFilDat"),
):
    ASSINATURAS.update(_par(*_nomes))
ASSINATURAS.update(_par(FUNCAO_TEMPO, FUNCAO_TEMPO_BLOCO, expressao=True))

# Formulas que o proprio programa sabe gerar do zero. Reconhece-las evita
# preserva-las como se fossem conta do usuario: a coluna volta a ser comum e
# acompanha as opcoes da rodada (unidade de tempo, modo bloco...).
FUNCOES_PADRAO = frozenset(
    nome.lower()
    for nome in (FUNCAO_VALOR, FUNCAO_BLOCO, FUNCAO_TEMPO, FUNCAO_TEMPO_BLOCO)
)

# Posicao do argumento "interval": logo depois do fim, em todas as funcoes.
POSICAO_INTERVALO = 3

ASPAS = chr(34)
APOSTROFO = chr(39)

# Qualquer funcao comecada em "PI" que nao esteja na lista acima faz o programa
# desistir da coluna - melhor deixar como esta do que reescrever errado.
RE_FUNCAO_PI = re.compile(r"(?:_xll\.)?(PI[A-Za-z0-9_]*)\s*\(", re.I)

# Referencia de celula, com ou sem prefixo de aba: 'Indicadores'!$A$3, B$2, $AA$2
RE_REFERENCIA = re.compile(
    r"^\s*(?P<prefixo>(?:'[^']*'|[A-Za-z0-9_.]+)!)?"
    r"\$?(?P<coluna>[A-Za-z]{1,3})\$?(?P<linha>\d+)\s*$"
)

# Comparacao dentro de uma expressao de filtro do PI: =, <>, >=, <, ...
RE_COMPARACAO = re.compile(r"(?:<>|>=|<=|[=<>])")


def e_expressao_de_filtro(texto: str) -> bool:
    """True quando o texto parece uma expressao de filtro, e nao uma TAG.

    E assim que uma coluna e marcada como "tempo por filtro": em vez da TAG, a
    celula da linha de TAGs guarda a condicao, no formato do proprio PI -

        'USM-XV_411010'="Aberta" AND 'USM-XV_411012'="Aberta"

    A regra e de proposito restritiva: exige uma comparacao E uma aspa (simples
    ou dupla). Um caminho de TAG - \\\\AZR-PIDA130.AD.DOM\\USM-FT_411012 - nao tem
    nenhuma das duas, entao nunca e confundido com uma expressao.
    """
    limpo = (texto or "").strip()
    if not limpo:
        return False
    if APOSTROFO not in limpo and ASPAS not in limpo:
        return False
    return bool(RE_COMPARACAO.search(limpo))


@dataclass
class ChamadaPI:
    """Uma chamada a PIAdvCalcVal/PIAdvCalcDat dentro de uma formula."""

    prefixo: str        # "_xll." ou ""
    nome: str           # "PIAdvCalcVal", "PIAdvCalcDat", "PIAdvCalcFilVal"
    inicio: int         # posicao do primeiro caractere na formula original
    fim: int            # posicao logo apos o ")" de fechamento
    argumentos: list = field(default_factory=list)

    @property
    def assinatura(self) -> Assinatura:
        return ASSINATURAS[self.nome.lower()]

    @property
    def tem_intervalo(self) -> bool:
        """A funcao escrita na planilha ja recebe 'interval' no 4o argumento."""
        return self.assinatura.tem_intervalo

    @property
    def cabe_em_bloco(self) -> bool:
        return bool(self.assinatura.bloco)

    @property
    def primeiro(self) -> str:
        """Primeiro argumento, seja ele a TAG ou a expressao de filtro."""
        return self.argumentos[0].strip() if self.argumentos else ""

    @property
    def tag(self) -> str:
        """Referencia da celula que guarda a TAG. Vazio no PITimeFilter."""
        return "" if self.assinatura.expressao else self.primeiro

    @property
    def expressao(self) -> str:
        """Expressao de filtro do PITimeFilter: um texto ou a celula que o guarda."""
        return self.primeiro if self.assinatura.expressao else ""


@dataclass
class FormulaPersonalizada:
    """A formula original de uma coluna, pronta para ser reemitida."""

    original: str                  # texto sem "=", ja desescapado do XML
    chamadas: list = field(default_factory=list)
    # Funcoes do Excel que envolvem as chamadas do PI (IFERROR, AVERAGE...).
    funcoes_em_volta: frozenset = frozenset()

    @property
    def total_chamadas(self) -> int:
        return len(self.chamadas)

    @property
    def composta(self) -> bool:
        """Mais de uma TAG combinada na mesma celula."""
        return len(self.chamadas) > 1

    @property
    def celulas_tag(self) -> list:
        """Referencias das celulas de TAG usadas, na ordem em que aparecem.

        Chamadas da familia PITimeFilter nao entram: o 1o argumento delas e a
        expressao de filtro (ver a propriedade expressoes).
        """
        return [chamada.tag for chamada in self.chamadas if chamada.tag]

    @property
    def expressoes(self) -> list:
        """Expressoes de filtro usadas pelas chamadas PITimeFilter."""
        return [chamada.expressao for chamada in self.chamadas if chamada.expressao]

    @property
    def celulas_usadas(self) -> list:
        """Toda celula citada pelas chamadas: TAG, filtro, inicio, fim...

        Serve para reconhecer as colunas de apoio - inclusive a que guarda a
        expressao de filtro do PIAdvCalcFilVal, que tambem nao e uma coluna de
        dados e nao pode receber formula.
        """
        celulas = []
        for chamada in self.chamadas:
            for argumento in chamada.argumentos:
                if RE_REFERENCIA.match(argumento):
                    celulas.append(argumento.strip())
        return celulas

    @property
    def cabe_em_bloco(self) -> bool:
        """True quando a conta continua valendo com as chamadas em matriz.

        Exige duas coisas: toda chamada do PI ter versao PIAdvCalc...Dat e toda
        funcao em volta trabalhar valor a valor. Uma funcao que resume o
        conjunto (AVERAGE, SUM, MAX...) devolveria UM numero para o periodo
        inteiro no lugar da coluna - essa coluna fica no modo linha.
        """
        return (
            all(chamada.cabe_em_bloco for chamada in self.chamadas)
            and self.funcoes_em_volta <= FUNCOES_VALOR_A_VALOR
        )

    @property
    def funcoes_que_impedem_bloco(self) -> list:
        return sorted(self.funcoes_em_volta - FUNCOES_VALOR_A_VALOR)

    @property
    def funcoes(self) -> list:
        return sorted({chamada.nome for chamada in self.chamadas})


# ---------------------------------------------------------------------------
# Leitura da formula
# ---------------------------------------------------------------------------

def _dividir_argumentos(texto: str, abre: int) -> tuple:
    """Separa os argumentos de nivel 1 a partir do "(" na posicao 'abre'.

    Respeita textos entre aspas ("average (time-weighted)") e nomes de aba
    entre apostrofos ('Indicadores de Horas'!), que podem conter virgulas e
    parenteses. Devolve (lista_de_argumentos, posicao_apos_o_fecha).
    """
    argumentos = []
    atual = []
    profundidade = 0
    em_aspas = False
    em_apostrofo = False

    for posicao in range(abre, len(texto)):
        caractere = texto[posicao]

        if em_aspas:
            atual.append(caractere)
            em_aspas = caractere != ASPAS
            continue
        if em_apostrofo:
            atual.append(caractere)
            em_apostrofo = caractere != APOSTROFO
            continue

        if caractere == ASPAS:
            em_aspas = True
            atual.append(caractere)
        elif caractere == APOSTROFO:
            em_apostrofo = True
            atual.append(caractere)
        elif caractere == "(":
            profundidade += 1
            if profundidade > 1:
                atual.append(caractere)
        elif caractere == ")":
            profundidade -= 1
            if profundidade == 0:
                argumentos.append("".join(atual))
                return argumentos, posicao + 1
            atual.append(caractere)
        elif caractere == "," and profundidade == 1:
            argumentos.append("".join(atual))
            atual = []
        else:
            atual.append(caractere)

    raise ValueError("parentese sem fechamento na formula")


def analisar(formula: str):
    """Le a formula de uma celula. Devolve None quando nao ha chamada do PI.

    Tambem devolve None quando aparece alguma funcao do PI que este modulo nao
    sabe reescrever (PIArcVal, PICompDat...) ou quando o inicio/fim nao sao
    referencias de celula - nesses casos a coluna e deixada como esta.
    """
    if not formula:
        return None
    texto = formula.strip()
    if texto.startswith("="):
        texto = texto[1:]

    chamadas = []
    posicao = 0
    while True:
        achado = RE_FUNCAO_PI.search(texto, posicao)
        if not achado:
            break
        assinatura = ASSINATURAS.get(achado.group(1).lower())
        if not assinatura:
            return None  # funcao do PI fora do escopo deste modulo
        nome_real = achado.group(1)
        abre = texto.index("(", achado.end() - 1)
        try:
            argumentos, fim = _dividir_argumentos(texto, abre)
        except ValueError:
            return None
        if len(argumentos) < 3:
            return None
        chamada = ChamadaPI(
            prefixo="_xll." if achado.group(0).lower().startswith("_xll.") else "",
            nome=nome_real,
            inicio=achado.start(),
            fim=fim,
            argumentos=argumentos,
        )
        if chamada.tem_intervalo and len(argumentos) < 4:
            return None
        # O inicio e o fim precisam ser celulas: e so isso que o programa
        # reescreve. Se alguem usou texto fixo ("*-1h"), a coluna fica intacta.
        if not RE_REFERENCIA.match(argumentos[1]) or not RE_REFERENCIA.match(argumentos[2]):
            return None
        chamadas.append(chamada)
        posicao = fim

    if not chamadas:
        return None

    return FormulaPersonalizada(
        original=texto,
        chamadas=chamadas,
        funcoes_em_volta=_funcoes_em_volta(texto, chamadas),
    )


def _funcoes_em_volta(texto: str, chamadas: list) -> frozenset:
    """Nomes das funcoes do Excel que envolvem as chamadas do PI.

    Le apenas o que sobra da formula depois de tirar as chamadas do PI - o
    miolo delas ja foi conferido argumento por argumento.
    """
    sobra = []
    anterior = 0
    for chamada in chamadas:
        sobra.append(texto[anterior:chamada.inicio])
        anterior = chamada.fim
    sobra.append(texto[anterior:])
    return frozenset(
        nome.upper() for nome in RE_NOME_FUNCAO.findall("".join(sobra))
    )


# ---------------------------------------------------------------------------
# Referencias
# ---------------------------------------------------------------------------

def referencia_de_celula(argumento: str) -> tuple:
    """'$AA$2' -> ('', 'AA', 2). Devolve (prefixo_de_aba, coluna, linha)."""
    achado = RE_REFERENCIA.match(argumento or "")
    if not achado:
        return ("", "", 0)
    return (
        achado.group("prefixo") or "",
        achado.group("coluna").upper(),
        int(achado.group("linha")),
    )


def _prefixo_de_aba(argumento: str) -> str:
    return referencia_de_celula(argumento)[0]


def e_formula_padrao(analise, coluna: str, linha_tags: int) -> bool:
    """True quando a formula e exatamente a que o programa geraria sozinho.

    Ou seja: uma unica chamada, nada em volta dela, e a TAG (ou a expressao de
    filtro, no PITimeFilter) lida da propria coluna na linha de TAGs. Nesses
    casos nada precisa ser preservado - a formula normal do programa reproduz a
    coluna (e ainda corrige referencias de tempo travadas que porventura
    existam).
    """
    if not analise or len(analise.chamadas) != 1:
        return False
    chamada = analise.chamadas[0]
    if chamada.nome.lower() not in FUNCOES_PADRAO:
        # PIAdvCalcFilVal e outras variantes tem argumentos que o programa nao
        # sabe montar (o filtro, por exemplo). A formula da planilha e que vale.
        return False
    if chamada.inicio != 0 or chamada.fim != len(analise.original):
        return False
    # chamada.primeiro, e nao chamada.tag: no PITimeFilter o que aponta para a
    # celula da coluna e a expressao de filtro.
    _, coluna_tag, linha = referencia_de_celula(chamada.primeiro)
    return coluna_tag == coluna.upper() and linha == linha_tags


# ---------------------------------------------------------------------------
# Reescrita
# ---------------------------------------------------------------------------

def _reconstruir(analise, novas: list) -> str:
    """Troca cada chamada pelo texto novo, preservando tudo o que esta em volta."""
    partes = []
    anterior = 0
    for chamada, nova in zip(analise.chamadas, novas):
        partes.append(analise.original[anterior:chamada.inicio])
        partes.append(nova)
        anterior = chamada.fim
    partes.append(analise.original[anterior:])
    return "".join(partes)


def _chamada_em_texto(chamada: ChamadaPI, nome: str, argumentos: list) -> str:
    return "%s%s(%s)" % (chamada.prefixo, nome, ",".join(argumentos))


def montar_linha(analise, linha: int, coluna_data: str) -> str:
    """Mesma formula da planilha, apontando para o inicio/fim desta linha.

    Cada chamada sai na sua forma "um valor por linha": PIAdvCalcDat volta a
    ser PIAdvCalcVal, e PIAdvCalcFilVal continua sendo PIAdvCalcFilVal.
    """
    novas = []
    for chamada in analise.chamadas:
        argumentos = list(chamada.argumentos)
        if chamada.tem_intervalo:
            # PIAdvCalcDat -> PIAdvCalcVal: sai o "interval"
            del argumentos[POSICAO_INTERVALO]
        argumentos[1] = "%s$%s%d" % (
            _prefixo_de_aba(chamada.argumentos[1]), coluna_data, linha
        )
        argumentos[2] = "%s$%s%d" % (
            _prefixo_de_aba(chamada.argumentos[2]), coluna_data, linha + 1
        )
        novas.append(
            _chamada_em_texto(chamada, chamada.assinatura.valor, argumentos)
        )
    return _reconstruir(analise, novas)


def montar_bloco(analise, primeira_linha: int, ultima_linha: int,
                 coluna_data: str, intervalo: str) -> str:
    """Mesma formula da planilha, com uma unica chamada por TAG (PIAdvCalcDat).

    Cada chamada devolve a coluna inteira de uma vez; somar duas matrizes do
    mesmo tamanho devolve outra matriz, entao a conta original continua valida.

    So vale para as chamadas que tem versao em matriz - ver cabe_em_bloco.
    """
    novas = []
    for chamada in analise.chamadas:
        argumentos = list(chamada.argumentos)
        if chamada.tem_intervalo:
            argumentos[POSICAO_INTERVALO] = '"%s"' % intervalo
        else:
            argumentos.insert(POSICAO_INTERVALO, '"%s"' % intervalo)
        argumentos[1] = "%s$%s$%d" % (
            _prefixo_de_aba(chamada.argumentos[1]), coluna_data, primeira_linha
        )
        argumentos[2] = "%s$%s$%d" % (
            _prefixo_de_aba(chamada.argumentos[2]), coluna_data, ultima_linha
        )
        novas.append(
            _chamada_em_texto(chamada, chamada.assinatura.bloco, argumentos)
        )
    return _reconstruir(analise, novas)
