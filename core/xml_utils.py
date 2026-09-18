# -*- coding: utf-8 -*-
"""Utilitarios de baixo nivel para manipular o pacote OOXML de um .xlsx.

Um arquivo .xlsx e um ZIP com varias "partes" XML. A estrategia deste projeto e
copiar TODAS as partes byte a byte e reescrever apenas o <sheetData> das abas
que precisam ser alteradas. Isso preserva formatacao, larguras, estilos,
congelamento de paineis, configuracoes de impressao, o painel do PI DataLink
(xl/webextensions) e a metadata de matriz dinamica (xl/metadata.xml).
"""

from __future__ import annotations

import re
import shutil
import zipfile
from dataclasses import dataclass
from pathlib import Path


# ---------------------------------------------------------------------------
# Conversao de colunas
# ---------------------------------------------------------------------------

def letra_para_indice(letra: str) -> int:
    """'A' -> 1, 'B' -> 2, 'AA' -> 27."""
    letra = letra.strip().upper()
    if not letra or not letra.isalpha():
        raise ValueError("Coluna invalida: %r" % letra)
    indice = 0
    for caractere in letra:
        indice = indice * 26 + (ord(caractere) - 64)
    return indice


def indice_para_letra(indice: int) -> str:
    """1 -> 'A', 27 -> 'AA'."""
    if indice < 1:
        raise ValueError("Indice de coluna invalido: %r" % indice)
    letra = ""
    while indice:
        indice, resto = divmod(indice - 1, 26)
        letra = chr(65 + resto) + letra
    return letra


_RE_REF = re.compile(r"^([A-Z]{1,3})(\d+)$")


def separar_referencia(ref: str) -> tuple:
    """'B12' -> (coluna=2, linha=12)."""
    correspondencia = _RE_REF.match(ref.strip().upper().replace("$", ""))
    if not correspondencia:
        raise ValueError("Referencia invalida: %r" % ref)
    return letra_para_indice(correspondencia.group(1)), int(correspondencia.group(2))


# ---------------------------------------------------------------------------
# Escape XML
# ---------------------------------------------------------------------------

def escapar_xml(texto: str) -> str:
    return texto.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")


def desescapar_xml(texto: str) -> str:
    return (
        texto.replace("&lt;", "<")
        .replace("&gt;", ">")
        .replace("&quot;", chr(34))
        .replace("&apos;", chr(39))
        .replace("&amp;", "&")
    )


# ---------------------------------------------------------------------------
# Pacote OOXML
# ---------------------------------------------------------------------------

@dataclass
class PacoteXlsx:
    """Todas as partes de um .xlsx carregadas em memoria."""

    caminho: Path
    partes: dict
    ordem: list

    @classmethod
    def abrir(cls, caminho) -> "PacoteXlsx":
        caminho = Path(caminho)
        partes = {}
        ordem = []
        with zipfile.ZipFile(caminho, "r") as zf:
            for nome in zf.namelist():
                if nome.endswith("/"):
                    continue
                ordem.append(nome)
                partes[nome] = zf.read(nome)
        if "xl/workbook.xml" not in partes:
            raise ValueError("Nao parece um arquivo Excel valido (xl/workbook.xml ausente).")
        return cls(caminho=caminho, partes=partes, ordem=ordem)

    def texto(self, nome: str) -> str:
        return self.partes[nome].decode("utf-8")

    def definir_texto(self, nome: str, conteudo: str) -> None:
        self.partes[nome] = conteudo.encode("utf-8")

    def remover_parte(self, nome: str) -> bool:
        """Remove uma parte e as referencias a ela ([Content_Types] e .rels)."""
        if nome not in self.partes:
            return False
        del self.partes[nome]
        self.ordem = [item for item in self.ordem if item != nome]

        tipos = "[Content_Types].xml"
        if tipos in self.partes:
            conteudo = self.texto(tipos)
            conteudo = re.sub(
                r'<Override\s+PartName="/%s"[^>]*/>' % re.escape(nome), "", conteudo
            )
            self.definir_texto(tipos, conteudo)

        alvo_curto = nome.split("/")[-1]
        for parte in list(self.partes):
            if not parte.endswith(".rels"):
                continue
            conteudo = self.texto(parte)
            novo = re.sub(
                r'<Relationship\b[^>]*Target="[^"]*%s"[^>]*/>' % re.escape(alvo_curto),
                "",
                conteudo,
            )
            if novo != conteudo:
                self.definir_texto(parte, novo)
        return True

    def salvar(self, destino) -> None:
        """Grava o pacote de forma atomica (arquivo temporario + troca)."""
        destino = Path(destino)
        destino.parent.mkdir(parents=True, exist_ok=True)
        temporario = destino.with_name(destino.name + ".tmp")
        try:
            with zipfile.ZipFile(temporario, "w", zipfile.ZIP_DEFLATED) as zf:
                nomes = [n for n in self.ordem if n in self.partes]
                nomes += [n for n in self.partes if n not in nomes]
                nomes.sort(key=lambda n: 0 if n == "[Content_Types].xml" else 1)
                for nome in nomes:
                    zf.writestr(nome, self.partes[nome])
            if destino.exists():
                destino.unlink()
            shutil.move(str(temporario), str(destino))
        finally:
            if temporario.exists():
                temporario.unlink()


# ---------------------------------------------------------------------------
# Parsing do sheetData
# ---------------------------------------------------------------------------
# O XML do Excel escapa obrigatoriamente "<" dentro de textos, portanto e seguro
# delimitar <row>/<c> com expressoes regulares nao-gulosas.

RE_LINHA = re.compile(r"<row\b[^>]*/>|<row\b[^>]*>.*?</row>", re.S)
RE_CELULA = re.compile(r"<c\b[^>]*/>|<c\b[^>]*>.*?</c>", re.S)
RE_ATRIBUTO_R = re.compile(r'\br="([^"]+)"')
RE_ATRIBUTO_S = re.compile(r'\bs="(\d+)"')
RE_ATRIBUTO_T = re.compile(r'\bt="([^"]+)"')
RE_ATRIBUTO_CM = re.compile(r'\bcm="(\d+)"')
RE_VALOR = re.compile(r"<v>(.*?)</v>", re.S)
RE_TEXTO_INLINE = re.compile(r"<is>.*?<t[^>]*>(.*?)</t>.*?</is>", re.S)


def dividir_sheet(xml: str) -> tuple:
    """Separa o XML da aba em (cabecalho, miolo do sheetData, rodape)."""
    vazio = re.search(r"<sheetData\s*/>", xml)
    if vazio:
        return xml[: vazio.start()] + "<sheetData>", "", "</sheetData>" + xml[vazio.end():]

    abertura = re.search(r"<sheetData[^>]*>", xml)
    fechamento = xml.rfind("</sheetData>")
    if not abertura or fechamento == -1:
        raise ValueError("Aba sem elemento <sheetData>.")
    return xml[: abertura.end()], xml[abertura.end(): fechamento], xml[fechamento:]


def abertura_da_linha(linha_xml: str) -> str:
    """Devolve a tag de abertura <row ...> normalizada (sempre nao auto-fechada)."""
    fim = linha_xml.find(">")
    abertura = linha_xml[: fim + 1]
    if abertura.endswith("/>"):
        abertura = abertura[:-2] + ">"
    return abertura


def atributos_sem(abertura: str, remover: tuple) -> str:
    """Remove atributos indesejados de uma tag de abertura <row ...>."""
    interior = abertura[len("<row"): -1]
    for nome in remover:
        interior = re.sub(r'\s%s="[^"]*"' % re.escape(nome), "", interior)
    return interior.strip()
