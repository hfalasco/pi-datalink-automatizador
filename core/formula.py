# -*- coding: utf-8 -*-
"""Montagem da formula do PI DataLink.

A formula gravada no arquivo tem esta forma (exatamente como o proprio
PI DataLink grava):

    PIAdvCalcVal(B$2,$A3,$A4,"average (time-weighted)","time-weighted",0,1,0,"")

Observacoes importantes:

* O separador de argumentos DENTRO DO ARQUIVO e sempre a virgula. Em maquinas
  com Excel em portugues a formula sera EXIBIDA com ponto e virgula - isso e
  apenas apresentacao, o arquivo continua com virgula.
* A TAG nunca e escrita como texto: usamos a referencia a celula da linha de
  TAGs, com a linha travada (B$2). Assim a mesma formula serve para qualquer
  TAG e continua valida se a TAG for trocada na linha 2.
* O inicio e o fim usam a coluna travada ($A3 e $A4), permitindo arrastar a
  formula para os lados sem quebrar a referencia de tempo.
"""

from __future__ import annotations

from .xml_utils import escapar_xml


def montar_formula(
    coluna_tag: str,
    linha_tags: int,
    linha_inicio: int,
    coluna_data: str,
    funcao: str,
    argumentos_fixos: tuple,
) -> str:
    """Devolve o texto da formula SEM o sinal de igual (formato do XML)."""
    referencia_tag = "%s$%d" % (coluna_tag, linha_tags)
    referencia_inicio = "$%s%d" % (coluna_data, linha_inicio)
    referencia_fim = "$%s%d" % (coluna_data, linha_inicio + 1)
    argumentos = [referencia_tag, referencia_inicio, referencia_fim]
    argumentos.extend(argumentos_fixos)
    return "%s(%s)" % (funcao, ",".join(argumentos))


def montar_formula_xml(
    coluna_tag: str,
    linha_tags: int,
    linha_inicio: int,
    coluna_data: str,
    funcao: str,
    argumentos_fixos: tuple,
) -> str:
    """Mesma formula, ja escapada para ser inserida no XML."""
    return escapar_xml(
        montar_formula(
            coluna_tag, linha_tags, linha_inicio, coluna_data, funcao, argumentos_fixos
        )
    )


def para_exibicao(formula_xml: str) -> str:
    """Converte para a forma que o usuario ve no Excel em pt-BR (com ';')."""
    resultado = []
    dentro_de_texto = False
    for caractere in formula_xml:
        if caractere == chr(34):
            dentro_de_texto = not dentro_de_texto
        if caractere == "," and not dentro_de_texto:
            resultado.append(";")
        else:
            resultado.append(caractere)
    return "=" + "".join(resultado)


# ---------------------------------------------------------------------------
# Formula em bloco (PIAdvCalcDat)
# ---------------------------------------------------------------------------
# PIAdvCalcVal devolve UM valor e exige uma formula por linha. PIAdvCalcDat
# devolve TODOS os intervalos de uma vez, entao basta uma formula por coluna.
#
# Assinaturas reais, lidas do add-in instalado (pipc64.xll):
#
#   PIAdvCalcVal(data item, stime, etime,           modestring, weightmode,
#                minpctgood, cfactor, outcode, root path)
#   PIAdvCalcDat(data item, stime, etime, interval, modestring, weightmode,
#                minpctgood, cfactor, outcode, root path)
#
# Ou seja: mesma lista de argumentos, mais "interval" na 4a posicao.

MINUTOS_POR_HORA = 60
MINUTOS_POR_DIA = 1440


def intervalo_pi(minutos: int) -> str:
    """Converte minutos para a sintaxe de tempo do PI: 5 -> 5m, 60 -> 1h, 1440 -> 1d."""
    if minutos <= 0:
        raise ValueError("Intervalo invalido: %r" % minutos)
    if minutos % MINUTOS_POR_DIA == 0:
        return "%dd" % (minutos // MINUTOS_POR_DIA)
    if minutos % MINUTOS_POR_HORA == 0:
        return "%dh" % (minutos // MINUTOS_POR_HORA)
    return "%dm" % minutos


def montar_formula_bloco(
    coluna_tag: str,
    linha_tags: int,
    linha_inicial: int,
    linha_final: int,
    coluna_data: str,
    funcao: str,
    intervalo_minutos: int,
    argumentos_fixos: tuple,
) -> str:
    """Formula unica que preenche a coluna inteira, SEM o sinal de igual.

    As referencias de tempo sao absolutas ($A$3 e $A$291) porque existe apenas
    uma formula na coluna - ela nao e arrastada linha a linha.
    """
    argumentos = [
        "%s$%d" % (coluna_tag, linha_tags),
        "$%s$%d" % (coluna_data, linha_inicial),
        "$%s$%d" % (coluna_data, linha_final),
        '"%s"' % intervalo_pi(intervalo_minutos),
    ]
    argumentos.extend(argumentos_fixos)
    return "%s(%s)" % (funcao, ",".join(argumentos))


def montar_formula_bloco_xml(
    coluna_tag: str,
    linha_tags: int,
    linha_inicial: int,
    linha_final: int,
    coluna_data: str,
    funcao: str,
    intervalo_minutos: int,
    argumentos_fixos: tuple,
) -> str:
    """Mesma formula em bloco, ja escapada para ser inserida no XML."""
    return escapar_xml(
        montar_formula_bloco(
            coluna_tag, linha_tags, linha_inicial, linha_final,
            coluna_data, funcao, intervalo_minutos, argumentos_fixos,
        )
    )
