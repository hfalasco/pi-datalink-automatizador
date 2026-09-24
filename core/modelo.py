# -*- coding: utf-8 -*-
"""Planilha base: o modelo em branco que o programa sabe processar.

Serve para quem vai montar uma planilha nova do zero. O arquivo sai com a
estrutura reconhecida ja marcada, cada celula com uma cor:

    A2 = "Data"  -> rotulo da coluna de datas (o programa preenche A3 para baixo)
    B1 = "Nome"  -> nome legivel do indicador
    B2 = "Tag"   -> TAG do PI que a formula consulta

Para mais indicadores, repete-se o par Nome/Tag nas colunas C, D, E...

Como o resto do motor, o arquivo e montado so com a biblioteca padrao: sao
meia duzia de partes XML dentro de um ZIP, sem openpyxl. As linhas e colunas
vem de config.py, entao o modelo acompanha qualquer ajuste feito la.
"""

from __future__ import annotations

import io
import zipfile

from config import COLUNA_DATA_PADRAO, LINHA_TAGS_PADRAO
from .xml_utils import (escapar_xml, indice_para_letra, letra_para_indice,
                        separar_referencia)

NOME_ARQUIVO = "planilha_base_pi_datalink.xlsx"
NOME_ABA = "Indicadores"

ROTULO_DATA = "Data"
ROTULO_NOME = "Nome"
ROTULO_TAG = "Tag"

# Cores das celulas marcadas - as mesmas da barra lateral (tema.py), em ARGB.
_COR_DATA = "FF1F4E79"        # azul da marca, texto branco
_COR_NOME = "FFE8F3EC"        # verde claro, texto verde escuro
_COR_NOME_TEXTO = "FF166139"
_COR_TAG = "FFFBF1CC"         # dourado claro, texto escuro
_COR_TAG_TEXTO = "FF5C4A00"
_COR_BORDA = "FFB9C6D4"

# Larguras pensadas para o conteudo real: "dd/mm/aaaa hh:mm" na coluna de
# datas e um caminho de TAG como \\azr-pida130.ad.dom\USM-FT_421016 na de dados.
_LARGURA_DATA = 20.7
_LARGURA_TAG = 48.7

_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
_NS_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
_NS_PKG = "http://schemas.openxmlformats.org/package/2006/relationships"
_CABECALHO_XML = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'

_TIPOS = (
    '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
    '<Default Extension="rels" '
    'ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
    '<Default Extension="xml" ContentType="application/xml"/>'
    '<Override PartName="/xl/workbook.xml" ContentType="application/'
    'vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
    '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/'
    'vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
    '<Override PartName="/xl/styles.xml" ContentType="application/'
    'vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>'
    '<Override PartName="/xl/sharedStrings.xml" ContentType="application/'
    'vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>'
    "</Types>"
)

_RELACOES_PACOTE = (
    '<Relationships xmlns="%s">'
    '<Relationship Id="rId1" Type="%s/officeDocument" Target="xl/workbook.xml"/>'
    "</Relationships>" % (_NS_PKG, _NS_R)
)

_RELACOES_PASTA = (
    '<Relationships xmlns="%s">'
    '<Relationship Id="rId1" Type="%s/worksheet" Target="worksheets/sheet1.xml"/>'
    '<Relationship Id="rId2" Type="%s/styles" Target="styles.xml"/>'
    '<Relationship Id="rId3" Type="%s/sharedStrings" Target="sharedStrings.xml"/>'
    "</Relationships>" % (_NS_PKG, _NS_R, _NS_R, _NS_R)
)

# Estilos: 0 = padrao (o writer clona este para criar o formato de data),
# 1 = Data, 2 = Nome, 3 = Tag.
_ESTILO_DATA, _ESTILO_NOME, _ESTILO_TAG = 1, 2, 3

_ESTILOS = (
    '<styleSheet xmlns="%(ns)s">'
    '<fonts count="4">'
    '<font><sz val="11"/><name val="Calibri"/><family val="2"/></font>'
    '<font><b/><sz val="11"/><color rgb="FFFFFFFF"/><name val="Calibri"/>'
    '<family val="2"/></font>'
    '<font><b/><sz val="11"/><color rgb="%(nome_texto)s"/><name val="Calibri"/>'
    '<family val="2"/></font>'
    '<font><b/><sz val="11"/><color rgb="%(tag_texto)s"/><name val="Calibri"/>'
    '<family val="2"/></font>'
    "</fonts>"
    '<fills count="5">'
    '<fill><patternFill patternType="none"/></fill>'
    '<fill><patternFill patternType="gray125"/></fill>'
    '<fill><patternFill patternType="solid"><fgColor rgb="%(data)s"/>'
    '<bgColor indexed="64"/></patternFill></fill>'
    '<fill><patternFill patternType="solid"><fgColor rgb="%(nome)s"/>'
    '<bgColor indexed="64"/></patternFill></fill>'
    '<fill><patternFill patternType="solid"><fgColor rgb="%(tag)s"/>'
    '<bgColor indexed="64"/></patternFill></fill>'
    "</fills>"
    '<borders count="2">'
    "<border><left/><right/><top/><bottom/><diagonal/></border>"
    '<border><left style="thin"><color rgb="%(borda)s"/></left>'
    '<right style="thin"><color rgb="%(borda)s"/></right>'
    '<top style="thin"><color rgb="%(borda)s"/></top>'
    '<bottom style="thin"><color rgb="%(borda)s"/></bottom><diagonal/></border>'
    "</borders>"
    '<cellStyleXfs count="1">'
    '<xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>'
    "</cellStyleXfs>"
    '<cellXfs count="4">'
    '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>'
    '<xf numFmtId="0" fontId="1" fillId="2" borderId="1" xfId="0" applyFont="1"'
    ' applyFill="1" applyBorder="1" applyAlignment="1">'
    '<alignment horizontal="center" vertical="center"/></xf>'
    '<xf numFmtId="0" fontId="2" fillId="3" borderId="1" xfId="0" applyFont="1"'
    ' applyFill="1" applyBorder="1" applyAlignment="1">'
    '<alignment vertical="center"/></xf>'
    '<xf numFmtId="0" fontId="3" fillId="4" borderId="1" xfId="0" applyFont="1"'
    ' applyFill="1" applyBorder="1" applyAlignment="1">'
    '<alignment vertical="center"/></xf>'
    "</cellXfs>"
    '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/>'
    "</cellStyles>"
    '<dxfs count="0"/>'
    '<tableStyles count="0" defaultTableStyle="TableStyleMedium2"'
    ' defaultPivotStyle="PivotStyleLight16"/>'
    "</styleSheet>"
) % {
    "ns": _NS,
    "data": _COR_DATA,
    "nome": _COR_NOME,
    "nome_texto": _COR_NOME_TEXTO,
    "tag": _COR_TAG,
    "tag_texto": _COR_TAG_TEXTO,
    "borda": _COR_BORDA,
}


def celulas_do_modelo() -> dict:
    """Onde fica cada rotulo: {"data": "A2", "nome": "B1", "tag": "B2"}.

    Derivado de config.py - a linha de TAGs e a coluna de datas sao as mesmas
    que o programa procura na hora de processar.
    """
    coluna_dados = indice_para_letra(letra_para_indice(COLUNA_DATA_PADRAO) + 1)
    return {
        "data": "%s%d" % (COLUNA_DATA_PADRAO, LINHA_TAGS_PADRAO),
        "nome": "%s%d" % (coluna_dados, LINHA_TAGS_PADRAO - 1),
        "tag": "%s%d" % (coluna_dados, LINHA_TAGS_PADRAO),
    }


def _celula_texto(referencia: str, indice_string: int, estilo: int) -> str:
    return '<c r="%s" s="%d" t="s"><v>%d</v></c>' % (referencia, estilo, indice_string)


def _aba() -> str:
    celulas = celulas_do_modelo()
    # (referencia, indice em sharedStrings, estilo)
    marcadas = [
        (celulas["data"], 0, _ESTILO_DATA),
        (celulas["nome"], 1, _ESTILO_NOME),
        (celulas["tag"], 2, _ESTILO_TAG),
    ]

    por_linha = {}
    for referencia, indice, estilo in marcadas:
        coluna, linha = separar_referencia(referencia)
        por_linha.setdefault(linha, []).append(
            (coluna, _celula_texto(referencia, indice, estilo)))

    linhas = "".join(
        '<row r="%d" ht="20" customHeight="1">%s</row>'
        % (numero, "".join(xml for _, xml in sorted(por_linha[numero])))
        for numero in sorted(por_linha)
    )

    # Paineis congelados logo abaixo e a direita do cabecalho: as datas e o
    # par Nome/Tag continuam visiveis ao rolar milhares de linhas.
    coluna_data = letra_para_indice(COLUNA_DATA_PADRAO)
    canto = "%s%d" % (indice_para_letra(coluna_data + 1), LINHA_TAGS_PADRAO + 1)
    return (
        '<worksheet xmlns="%s" xmlns:r="%s">'
        '<dimension ref="A1:%s"/>'
        '<sheetViews><sheetView tabSelected="1" workbookViewId="0">'
        '<pane xSplit="%d" ySplit="%d" topLeftCell="%s" activePane="bottomRight"'
        ' state="frozen"/>'
        '<selection pane="topRight"/><selection pane="bottomLeft"/>'
        '<selection pane="bottomRight" activeCell="%s" sqref="%s"/>'
        "</sheetView></sheetViews>"
        '<sheetFormatPr defaultRowHeight="15"/>'
        '<cols>'
        '<col min="%d" max="%d" width="%s" customWidth="1"/>'
        '<col min="%d" max="%d" width="%s" customWidth="1"/>'
        "</cols>"
        "<sheetData>%s</sheetData>"
        '<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75"'
        ' header="0.3" footer="0.3"/>'
        "</worksheet>"
        % (_NS, _NS_R, celulas["tag"],
           coluna_data, LINHA_TAGS_PADRAO, canto, canto, canto,
           coluna_data, coluna_data, _LARGURA_DATA,
           coluna_data + 1, coluna_data + 1, _LARGURA_TAG,
           linhas)
    )


def _pasta() -> str:
    return (
        '<workbook xmlns="%s" xmlns:r="%s">'
        '<bookViews><workbookView/></bookViews>'
        '<sheets><sheet name="%s" sheetId="1" r:id="rId1"/></sheets>'
        "</workbook>" % (_NS, _NS_R, escapar_xml(NOME_ABA))
    )


def _textos() -> str:
    textos = (ROTULO_DATA, ROTULO_NOME, ROTULO_TAG)
    itens = "".join("<si><t>%s</t></si>" % escapar_xml(texto) for texto in textos)
    return ('<sst xmlns="%s" count="%d" uniqueCount="%d">%s</sst>'
            % (_NS, len(textos), len(textos), itens))


def montar_planilha_base() -> bytes:
    """Devolve o .xlsx do modelo, pronto para download."""
    partes = (
        ("[Content_Types].xml", _TIPOS),
        ("_rels/.rels", _RELACOES_PACOTE),
        ("xl/workbook.xml", _pasta()),
        ("xl/_rels/workbook.xml.rels", _RELACOES_PASTA),
        ("xl/worksheets/sheet1.xml", _aba()),
        ("xl/styles.xml", _ESTILOS),
        ("xl/sharedStrings.xml", _textos()),
    )
    buffer = io.BytesIO()
    with zipfile.ZipFile(buffer, "w", zipfile.ZIP_DEFLATED) as pacote:
        for nome, conteudo in partes:
            pacote.writestr(nome, (_CABECALHO_XML + conteudo).encode("utf-8"))
    return buffer.getvalue()
