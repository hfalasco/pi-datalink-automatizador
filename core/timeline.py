# -*- coding: utf-8 -*-
"""Geracao da sequencia de carimbos de tempo e conversao para serial do Excel."""

from __future__ import annotations

from datetime import date, datetime, timedelta

# O Excel conta os dias a partir de 1899-12-30 (inclui o bug do ano bissexto 1900).
EPOCA_EXCEL = datetime(1899, 12, 30)

MINUTOS_POR_DIA = 1440


def gerar_carimbos(
    data_inicial: date,
    data_final: date,
    intervalo_minutos: int,
    carimbo_final_extra: bool = True,
) -> list:
    """Gera os carimbos de tempo do periodo.

    Para 11/09/2026 a 11/09/2026 com intervalo de 5 min devolve 289 itens:
    00:00, 00:05, ... 23:55 (288 inicios) + 12/09/2026 00:00 (fim do ultimo
    intervalo). O numero de FORMULAS e sempre len(carimbos) - 1.
    """
    if intervalo_minutos <= 0:
        raise ValueError("O intervalo precisa ser maior que zero.")
    if data_final < data_inicial:
        raise ValueError("A data final e anterior a data inicial.")

    inicio = datetime.combine(data_inicial, datetime.min.time())
    # O periodo termina as 00:00 do dia seguinte ao ultimo dia solicitado.
    fim = datetime.combine(data_final, datetime.min.time()) + timedelta(days=1)
    passo = timedelta(minutes=intervalo_minutos)

    carimbos = []
    atual = inicio
    while atual < fim:
        carimbos.append(atual)
        atual += passo

    if carimbo_final_extra:
        carimbos.append(fim)
    return carimbos


def quantidade_carimbos(
    data_inicial: date,
    data_final: date,
    intervalo_minutos: int,
    carimbo_final_extra: bool = True,
) -> int:
    """Quantos carimbos o periodo teria - por conta, sem montar a lista.

    Serve para a pagina: ela so precisa do NUMERO para mostrar a previsao, e
    montar a lista inteira a cada clique era desperdicio (um ano de 1 em 1
    minuto sao mais de 500 mil datas criadas so para contar quantas eram).
    """
    if intervalo_minutos <= 0:
        raise ValueError("O intervalo precisa ser maior que zero.")
    if data_final < data_inicial:
        raise ValueError("A data final e anterior a data inicial.")

    minutos = ((data_final - data_inicial).days + 1) * MINUTOS_POR_DIA
    total = -(-minutos // intervalo_minutos)   # divisao arredondada para cima
    return total + (1 if carimbo_final_extra else 0)


def ultimo_carimbo(
    data_inicial: date,
    data_final: date,
    intervalo_minutos: int,
    carimbo_final_extra: bool = True,
) -> datetime:
    """Ultimo carimbo do periodo, tambem sem montar a lista."""
    fim = datetime.combine(data_final, datetime.min.time()) + timedelta(days=1)
    if carimbo_final_extra:
        return fim
    total = quantidade_carimbos(data_inicial, data_final, intervalo_minutos, False)
    return (datetime.combine(data_inicial, datetime.min.time())
            + timedelta(minutes=intervalo_minutos * (total - 1)))


def quantidade_formulas(carimbos, carimbo_final_extra: bool = True) -> int:
    """Quantas linhas de formula correspondem a estes carimbos.

    Aceita tanto a lista de carimbos quanto so a quantidade deles.
    """
    total = carimbos if isinstance(carimbos, int) else len(carimbos)
    if not total:
        return 0
    return total - 1 if carimbo_final_extra else total


def para_serial_excel(momento: datetime) -> float:
    """Converte datetime em numero serial do Excel (valor real, nao texto)."""
    delta = momento - EPOCA_EXCEL
    return delta.days + delta.seconds / 86400.0 + delta.microseconds / 86400_000_000.0


def formatar_serial(valor: float) -> str:
    """Serializa o numero para o XML sem notacao cientifica e sem zeros inuteis."""
    if valor == int(valor):
        return str(int(valor))
    texto = repr(float(valor))
    if "e" in texto or "E" in texto:
        texto = "%.10f" % valor
    return texto.rstrip("0").rstrip(".") if "." in texto else texto
