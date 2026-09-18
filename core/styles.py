# -*- coding: utf-8 -*-
"""Ajustes minimos e seguros em xl/styles.xml.

So e usado quando a coluna de data ainda NAO possui um estilo de data (aba
nova, sem linhas de dados) ou quando o usuario pede explicitamente para forcar
o formato dd/mm/yyyy hh:mm.

Em todos os outros casos o estilo existente e reaproveitado e o styles.xml nao
e tocado - essa e a forma mais segura de preservar a formatacao original.
"""

from __future__ import annotations

import re

from .xml_utils import PacoteXlsx, escapar_xml

PARTE_ESTILOS = "xl/styles.xml"

RE_NUMFMT = re.compile(r"<numFmt\b[^>]*/>")
RE_NUMFMT_ID = re.compile(r'\bnumFmtId="(\d+)"')
RE_NUMFMT_CODIGO = re.compile(r'\bformatCode="([^"]*)"')
RE_BLOCO_NUMFMTS = re.compile(r"<numFmts\b[^>]*>.*?</numFmts>|<numFmts\b[^>]*/>", re.S)
RE_BLOCO_CELLXFS = re.compile(r"<cellXfs\b[^>]*>.*?</cellXfs>", re.S)
RE_XF = re.compile(r"<xf\b[^>]*/>|<xf\b[^>]*>.*?</xf>", re.S)

# Formatos internos do Excel que ja representam data + hora.
NUMFMT_INTERNOS_DATA_HORA = {22}
PRIMEIRO_NUMFMT_CUSTOMIZADO = 164


def _numfmts_existentes(xml: str) -> dict:
    """Mapeia formatCode -> numFmtId dos formatos personalizados."""
    bloco = RE_BLOCO_NUMFMTS.search(xml)
    if not bloco:
        return {}
    mapa = {}
    for item in RE_NUMFMT.finditer(bloco.group(0)):
        ident = RE_NUMFMT_ID.search(item.group(0))
        codigo = RE_NUMFMT_CODIGO.search(item.group(0))
        if ident and codigo:
            mapa[codigo.group(1)] = int(ident.group(1))
    return mapa


def _proximo_numfmt_id(xml: str) -> int:
    usados = {int(m) for m in RE_NUMFMT_ID.findall(xml)}
    candidato = PRIMEIRO_NUMFMT_CUSTOMIZADO
    while candidato in usados:
        candidato += 1
    return candidato


def _garantir_numfmt(xml: str, formato: str) -> tuple:
    """Garante que o formatCode exista. Devolve (novo_xml, numFmtId)."""
    formato_xml = escapar_xml(formato).replace(chr(34), "&quot;")
    existentes = _numfmts_existentes(xml)
    if formato_xml in existentes:
        return xml, existentes[formato_xml]

    novo_id = _proximo_numfmt_id(xml)
    novo_item = '<numFmt numFmtId="%d" formatCode="%s"/>' % (novo_id, formato_xml)

    bloco = RE_BLOCO_NUMFMTS.search(xml)
    if bloco is None:
        # Nao existe <numFmts>. Ele precisa ser o primeiro filho de <styleSheet>.
        abertura = re.search(r"<styleSheet\b[^>]*>", xml)
        if not abertura:
            raise ValueError("styles.xml sem elemento <styleSheet>.")
        insercao = '<numFmts count="1">%s</numFmts>' % novo_item
        return xml[: abertura.end()] + insercao + xml[abertura.end():], novo_id

    texto = bloco.group(0)
    if texto.endswith("/>"):
        novo_bloco = '<numFmts count="1">%s</numFmts>' % novo_item
    else:
        quantidade = len(RE_NUMFMT.findall(texto)) + 1
        novo_bloco = texto.replace("</numFmts>", novo_item + "</numFmts>")
        novo_bloco = re.sub(r'\bcount="\d+"', 'count="%d"' % quantidade, novo_bloco, count=1)
    return xml[: bloco.start()] + novo_bloco + xml[bloco.end():], novo_id


def garantir_estilo_data(pacote: PacoteXlsx, formato: str) -> str:
    """Devolve o indice de um <xf> que aplica o formato de data pedido.

    Cria o numFmt e o xf apenas se ainda nao existirem.
    """
    xml = pacote.texto(PARTE_ESTILOS)
    xml, numfmt_id = _garantir_numfmt(xml, formato)

    bloco = RE_BLOCO_CELLXFS.search(xml)
    if not bloco:
        raise ValueError("styles.xml sem elemento <cellXfs>.")
    texto = bloco.group(0)
    xfs = RE_XF.findall(texto)

    for indice, xf in enumerate(xfs):
        ident = RE_NUMFMT_ID.search(xf)
        if ident and int(ident.group(1)) == numfmt_id:
            pacote.definir_texto(PARTE_ESTILOS, xml)
            return str(indice)

    # Clona o primeiro xf (herda fonte/borda/preenchimento padrao) e troca o numFmt.
    modelo = xfs[0] if xfs else '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>'
    if not modelo.endswith("/>"):
        modelo = modelo[: modelo.find(">") + 1].replace(">", "/>")
    novo_xf = re.sub(r'\bnumFmtId="\d+"', 'numFmtId="%d"' % numfmt_id, modelo)
    if "numFmtId=" not in novo_xf:
        novo_xf = novo_xf.replace("<xf", '<xf numFmtId="%d"' % numfmt_id, 1)
    if "applyNumberFormat=" in novo_xf:
        novo_xf = re.sub(r'\bapplyNumberFormat="\d+"', 'applyNumberFormat="1"', novo_xf)
    else:
        novo_xf = novo_xf.replace("/>", ' applyNumberFormat="1"/>', 1)

    novo_indice = len(xfs)
    novo_texto = texto.replace("</cellXfs>", novo_xf + "</cellXfs>")
    novo_texto = re.sub(
        r'\bcount="\d+"', 'count="%d"' % (novo_indice + 1), novo_texto, count=1
    )
    xml = xml[: bloco.start()] + novo_texto + xml[bloco.end():]
    pacote.definir_texto(PARTE_ESTILOS, xml)
    return str(novo_indice)


def formato_ja_e_data(pacote: PacoteXlsx, indice_xf: str | None) -> bool:
    """Diz se o estilo informado ja aplica algum formato de data/hora."""
    if indice_xf is None:
        return False
    try:
        alvo = int(indice_xf)
    except (TypeError, ValueError):
        return False

    xml = pacote.texto(PARTE_ESTILOS)
    bloco = RE_BLOCO_CELLXFS.search(xml)
    if not bloco:
        return False
    xfs = RE_XF.findall(bloco.group(0))
    if alvo >= len(xfs):
        return False
    ident = RE_NUMFMT_ID.search(xfs[alvo])
    if not ident:
        return False
    numfmt_id = int(ident.group(1))
    if numfmt_id in NUMFMT_INTERNOS_DATA_HORA or 14 <= numfmt_id <= 22:
        return True

    for codigo, valor in _numfmts_existentes(xml).items():
        if valor == numfmt_id:
            minusculo = codigo.lower()
            return ("y" in minusculo or "d" in minusculo) and (
                "h" in minusculo or "m" in minusculo
            )
    return False
