# -*- coding: utf-8 -*-
"""Garantia de suporte a MATRIZ DINAMICA (parte xl/metadata.xml).

O problema que este modulo resolve
----------------------------------
No modo bloco existe UMA formula PIAdvCalcDat por coluna, e ela devolve todos
os intervalos de uma vez. Isso e uma matriz, e o Excel tem duas formas de
representar matriz:

* matriz DINAMICA - a formula mora so em B3 e o resultado "derrama" sozinho
  para baixo. O bloco nasce e some junto com B3, e a planilha continua
  editavel normalmente.

* matriz CSE (classica) - a formula ocupa explicitamente B3:B290 e o Excel
  TRAVA o bloco inteiro. Qualquer tentativa de mexer em uma celula de dentro
  devolve a mensagem:

      Nao e possivel alterar parte de uma matriz.

Para o Excel entender a formula como DINAMICA, a celula precisa do atributo
cm="N", e esse N e um indice para um registro dentro de xl/metadata.xml que
descreve a propriedade "dynamicArrayProperties" (tipo XLDAPR).

Nem todos os templates trazem essa parte. Levantamento feito nas planilhas
reais: cerca de metade dos arquivos NAO tem xl/metadata.xml (por exemplo os
templates de Concentracao de Caldo, Destilacao e Fabrica de Acucar). Nesses
arquivos nao havia cm= para reaproveitar, o programa caia para a matriz CSE e
a planilha saia travada - exatamente o erro relatado nas abas de 5 em 5
minutos, que sao as que tem bloco grande.

A solucao e nao depender do que o arquivo ja tinha: quando a parte nao existe
(ou existe sem o tipo XLDAPR), ela e criada/completada aqui, e o indice certo
de cm= e devolvido para o writer.
"""

from __future__ import annotations

import re

PARTE = "xl/metadata.xml"
PARTE_TIPOS = "[Content_Types].xml"
PARTE_RELS = "xl/_rels/workbook.xml.rels"

TIPO_CONTEUDO = (
    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheetMetadata+xml"
)
TIPO_RELACAO = (
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sheetMetadata"
)

NOME_TIPO = "XLDAPR"

# Registro de propriedade de matriz dinamica, identico ao que o proprio Excel
# grava. O uri do <ext> e fixo, definido pela Microsoft.
XML_TIPO = (
    '<metadataType name="XLDAPR" minSupportedVersion="120000" copy="1"'
    ' pasteAll="1" pasteValues="1" merge="1" splitFirst="1" rowColShift="1"'
    ' clearFormats="1" clearComments="1" assign="1" coerce="1" cellMeta="1"/>'
)
XML_FUTURO_BK = (
    '<bk><extLst><ext uri="{bdbb8cdc-fa1e-496e-a857-3c3f30c029c3}">'
    '<xda:dynamicArrayProperties fDynamic="1" fCollapsed="0"/>'
    "</ext></extLst></bk>"
)
XML_COMPLETO = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
    '<metadata xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"'
    ' xmlns:xda="http://schemas.microsoft.com/office/spreadsheetml/2017/dynamicarray">'
    '<metadataTypes count="1">' + XML_TIPO + "</metadataTypes>"
    '<futureMetadata name="XLDAPR" count="1">' + XML_FUTURO_BK + "</futureMetadata>"
    '<cellMetadata count="1"><bk><rc t="1" v="0"/></bk></cellMetadata>'
    "</metadata>"
)

ESPACO_XDA = "http://schemas.microsoft.com/office/spreadsheetml/2017/dynamicarray"

RE_BK = re.compile(r"<bk\b[^>]*>.*?</bk>|<bk\b[^>]*/>", re.S)
RE_TIPO = re.compile(
    r"<metadataType\b[^>]*/>|<metadataType\b[^>]*>.*?</metadataType>", re.S
)
RE_NOME = re.compile(r'\bname="([^"]*)"')


# ---------------------------------------------------------------------------
# Leitura dos indices dentro do metadata.xml
# ---------------------------------------------------------------------------

def _interior(xml: str, tag: str, filtro: str = "") -> tuple:
    """Devolve (inicio_do_interior, fim_do_interior, texto) do elemento pedido.

    "filtro" permite exigir um atributo na tag de abertura - usado para achar o
    <futureMetadata name="XLDAPR"> certo. Devolve (None, None, "") se o
    elemento nao existir ou estiver vazio (auto-fechado).
    """
    padrao = r"<%s\b([^>]*%s[^>]*)>" % (re.escape(tag), re.escape(filtro))
    abertura = re.search(padrao, xml)
    if not abertura or abertura.group(0).endswith("/>"):
        return None, None, ""
    fechamento = xml.find("</%s>" % tag, abertura.end())
    if fechamento == -1:
        return None, None, ""
    return abertura.end(), fechamento, xml[abertura.end(): fechamento]


def _indice_do_tipo(xml: str) -> "int | None":
    """Posicao (base 1) do <metadataType name="XLDAPR">."""
    _, _, corpo = _interior(xml, "metadataTypes")
    if not corpo:
        return None
    for posicao, tipo in enumerate(RE_TIPO.finditer(corpo), start=1):
        nome = RE_NOME.search(tipo.group(0))
        if nome and nome.group(1) == NOME_TIPO:
            return posicao
    return None


def _indice_do_futuro(xml: str) -> "int | None":
    """Posicao (base 0) do <bk> de matriz dinamica dentro de futureMetadata."""
    _, _, corpo = _interior(xml, "futureMetadata", 'name="%s"' % NOME_TIPO)
    if not corpo:
        return None
    for posicao, bloco in enumerate(RE_BK.finditer(corpo)):
        if 'fDynamic="1"' in bloco.group(0):
            return posicao
    return None


def _indice_da_celula(xml: str, tipo: int) -> "int | None":
    """Posicao (base 1) do <bk> de cellMetadata que aponta para o tipo dado."""
    _, _, corpo = _interior(xml, "cellMetadata")
    if not corpo:
        return None
    for posicao, bloco in enumerate(RE_BK.finditer(corpo), start=1):
        if re.search(r'<rc\b[^>]*\bt="%d"' % tipo, bloco.group(0)):
            return posicao
    return None


# ---------------------------------------------------------------------------
# Escrita: acrescenta o que estiver faltando
# ---------------------------------------------------------------------------

def _acrescentar(xml: str, tag: str, conteudo: str, filtro: str = "",
                 abertura_nova: str = "") -> str:
    """Insere conteudo no fim do elemento e atualiza o atributo count."""
    inicio, fim, corpo = _interior(xml, tag, filtro)
    if inicio is None:
        # Elemento inexistente: cria ja com o conteudo, antes de </metadata>.
        novo = '<%s%s count="1">%s</%s>' % (tag, abertura_nova, conteudo, tag)
        posicao = xml.rfind("</metadata>")
        if posicao == -1:
            return xml
        return xml[:posicao] + novo + xml[posicao:]

    novo_corpo = corpo + conteudo
    xml = xml[:inicio] + novo_corpo + xml[fim:]

    # O count precisa acompanhar o numero real de itens, senao o Excel reclama.
    regex = RE_TIPO if tag == "metadataTypes" else RE_BK
    total = sum(1 for _ in regex.finditer(novo_corpo))
    padrao_com_count = r"(<%s\b[^>]*?)\scount=\"\d+\"" % re.escape(tag)
    if re.search(padrao_com_count, xml):
        return re.sub(padrao_com_count, r'\1 count="%d"' % total, xml, count=1)
    return re.sub(
        r"(<%s\b[^>]*?)>" % re.escape(tag), r'\1 count="%d">' % total, xml, count=1
    )


def _garantir_namespace(xml: str) -> str:
    if "xmlns:xda=" in xml:
        return xml
    return re.sub(
        r"(<metadata\b[^>]*?)>", r'\1 xmlns:xda="%s">' % ESPACO_XDA, xml, count=1
    )


def _proximo_rid(rels: str) -> str:
    usados = [int(numero) for numero in re.findall(r'\bId="rId(\d+)"', rels)]
    return "rId%d" % ((max(usados) + 1) if usados else 1)


def _registrar_parte(pacote) -> None:
    """Declara xl/metadata.xml no [Content_Types] e nas relacoes da pasta."""
    if PARTE_TIPOS in pacote.partes:
        tipos = pacote.texto(PARTE_TIPOS)
        if '/%s"' % PARTE not in tipos:
            marca = '<Override PartName="/%s" ContentType="%s"/>' % (
                PARTE, TIPO_CONTEUDO
            )
            pacote.definir_texto(
                PARTE_TIPOS, tipos.replace("</Types>", marca + "</Types>", 1)
            )

    if PARTE_RELS in pacote.partes:
        rels = pacote.texto(PARTE_RELS)
        if 'Target="metadata.xml"' not in rels:
            marca = '<Relationship Id="%s" Type="%s" Target="metadata.xml"/>' % (
                _proximo_rid(rels), TIPO_RELACAO
            )
            pacote.definir_texto(
                PARTE_RELS,
                rels.replace("</Relationships>", marca + "</Relationships>", 1),
            )


# ---------------------------------------------------------------------------
# Ponto de entrada
# ---------------------------------------------------------------------------

def garantir_cm_dinamico(pacote, aplicar: bool = True) -> "str | None":
    """Devolve o valor de cm= que marca uma formula como matriz dinamica.

    Cria ou completa xl/metadata.xml quando necessario. Com aplicar=False nada
    e gravado no pacote (modo de teste) - so o indice e calculado.

    Devolve None apenas quando o arquivo tem um metadata.xml que nao foi
    possivel interpretar; nesse caso o writer volta para a matriz CSE.
    """
    try:
        if PARTE not in pacote.partes:
            if aplicar:
                pacote.definir_texto(PARTE, XML_COMPLETO)
                _registrar_parte(pacote)
            return "1"

        xml = pacote.texto(PARTE)
        original = xml

        tipo = _indice_do_tipo(xml)
        if tipo is None:
            xml = _acrescentar(xml, "metadataTypes", XML_TIPO)
            tipo = _indice_do_tipo(xml)
            if tipo is None:
                return None

        futuro = _indice_do_futuro(xml)
        if futuro is None:
            xml = _garantir_namespace(xml)
            xml = _acrescentar(
                xml, "futureMetadata", XML_FUTURO_BK,
                filtro='name="%s"' % NOME_TIPO,
                abertura_nova=' name="%s"' % NOME_TIPO,
            )
            futuro = _indice_do_futuro(xml)
            if futuro is None:
                return None

        celula = _indice_da_celula(xml, tipo)
        if celula is None:
            xml = _acrescentar(
                xml, "cellMetadata", '<bk><rc t="%d" v="%d"/></bk>' % (tipo, futuro)
            )
            celula = _indice_da_celula(xml, tipo)
            if celula is None:
                return None

        if aplicar:
            if xml != original:
                pacote.definir_texto(PARTE, xml)
            # Barato e idempotente: cobre o caso raro de um arquivo que tem a
            # parte mas perdeu a declaracao no [Content_Types] ou nas relacoes.
            _registrar_parte(pacote)
        return str(celula)
    except Exception:
        # Nunca derrubar o processamento por causa da metadata: sem ela o
        # writer apenas volta para a matriz CSE.
        return None
