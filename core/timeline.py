# -*- coding: utf-8 -*-
"""Geracao da sequencia de carimbos de tempo e conversao para serial do Excel."""

from __future__ import annotations

from datetime import date, datetime, timedelta

# O Excel conta os dias a partir de 1899-12-30 (inclui o bug do ano bissexto 1900).
EPOCA_EXCEL = datetime(1899, 12, 30)


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


def quantidade_formulas(carimbos: list, carimbo_final_extra: bool = True) -> int:
    """Quantas linhas de formula correspondem a esta lista de carimbos."""
    if not carimbos:
        return 0
    return len(carimbos) - 1 if carimbo_final_extra else len(carimbos)


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
