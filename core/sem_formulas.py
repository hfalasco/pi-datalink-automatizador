# -*- coding: utf-8 -*-
"""Tirar formulas: troca cada formula pelo valor que o Excel ja calculou.

E a versao em pagina do antigo tirar_formulas.py. O resultado e o mesmo - uma
copia da planilha so com os numeros, com o sufixo _SEM_FORMULAS no nome - mas
o caminho ate ele e outro.

Por que nao openpyxl, como no script
------------------------------------
O script abria cada arquivo DUAS vezes com openpyxl, e uma delas inteira na
memoria (o modo normal, necessario para ler mesclagens, larguras e paineis).
Com planilhas de centenas de MB isso passa de varios GB - num servidor, derruba
a pagina para todo mundo. E a copia levava so valor, formato de numero,
mesclagem, largura e painel congelado: cores, fontes, bordas e graficos ficavam
para tras.

Aqui o arquivo e copiado parte por parte, como no resto do motor, e so as
celulas com formula mudam:

    <c r="B3" s="4" cm="1"><f t="array" ref="B3:B290">PIAdvCalcDat(...)</f><v>812.5</v></c>
    <c r="B3" s="4"><v>812.5</v></c>

A <f> sai e o <v> - o valor que o Excel gravou da ultima vez que calculou -
fica. Todo o resto do arquivo passa intacto. As abas sao lidas em pedacos, de
linha em linha, entao a memoria usada nao depende do tamanho da planilha.

IMPORTANTE (igual ao script): o valor precisa estar gravado no arquivo. A
planilha tem de ter sido aberta, calculada (PI DataLink atualizado) e SALVA no
Excel antes de ser enviada. Uma formula sem valor gravado vira celula vazia, e
o resultado conta quantas ficaram assim.
"""

from __future__ import annotations

import codecs
import re
import shutil
import zipfile
from dataclasses import dataclass, field
from pathlib import Path

from . import inspector
from . import xml_utils as xu

SUFIXO_SAIDA = "_SEM_FORMULAS"

PARTE_TIPOS = "[Content_Types].xml"
PARTE_PASTA = "xl/workbook.xml"
PARTE_RELS = "xl/_rels/workbook.xml.rels"
# Cache da ordem de calculo. Sem formulas ele aponta para celulas que nao tem
# mais conta nenhuma, e o Excel acusaria "conteudo ilegivel" - sai do arquivo.
PARTE_CALCCHAIN = "xl/calcChain.xml"

TIPO_ABA = "spreadsheetml.worksheet+xml"

# Tamanho de cada pedaco lido da aba. As celulas sao tratadas de linha em
# linha; o pedaco so define quanto texto passa pela memoria de cada vez.
TAMANHO_PEDACO = 8 * 1024 * 1024

# Acima disto a parte descompactada precisa do formato ZIP64 (limite de 4 GB
# do ZIP comum). A margem cobre o pouco que a troca de texto pode crescer.
LIMITE_ZIP64 = 2 * 1024 ** 3

RE_ABRE_F = re.compile(r"<f[\s/>]")
RE_ELEMENTO_F = re.compile(r"<f\b[^>]*/>|<f\b[^>]*>.*?</f>", re.S)
RE_ATRIBUTO_CM = re.compile(r'\scm="[^"]*"')
RE_ATRIBUTO_T = re.compile(r'\st="([^"]*)"')
RE_NUMERO_LINHA = re.compile(r'<row\b[^>]*?\sr="(\d+)"')
RE_DIMENSAO = re.compile(r'<dimension\b[^>]*?\sref="([^"]+)"')
RE_OVERRIDE = re.compile(r"<Override\b[^>]*/>")
RE_PARTNAME = re.compile(r'\bPartName="/?([^"]+)"')
RE_CONTENTTYPE = re.compile(r'\bContentType="([^"]+)"')


@dataclass
class AbaSemFormulas:
    nome: str
    parte: str
    linhas: int = 0
    colunas: int = 0
    formulas: int = 0     # celulas que tinham formula
    sem_valor: int = 0    # das quais nao havia valor gravado (ficaram vazias)


@dataclass
class ResultadoSemFormulas:
    arquivo: Path
    saida: Path | None = None
    erro: str = ""
    abas: list = field(default_factory=list)

    @property
    def sucesso(self) -> bool:
        return not self.erro

    @property
    def formulas(self) -> int:
        return sum(aba.formulas for aba in self.abas)

    @property
    def sem_valor(self) -> int:
        return sum(aba.sem_valor for aba in self.abas)


def nome_de_saida(arquivo) -> str:
    """'Safra.xlsx' -> 'Safra_SEM_FORMULAS.xlsx' (sem repetir o sufixo)."""
    arquivo = Path(arquivo)
    if arquivo.stem.endswith(SUFIXO_SAIDA):
        return arquivo.name
    return arquivo.stem + SUFIXO_SAIDA + arquivo.suffix


# ---------------------------------------------------------------------------
# Celulas
# ---------------------------------------------------------------------------

def _celula_vazia(abertura: str) -> str:
    """Mantem a celula (e o estilo dela), sem valor nenhum."""
    return RE_ATRIBUTO_T.sub("", abertura)[:-1] + "/>"


def _sem_formula(celula: str, aba: AbaSemFormulas) -> str:
    """Tira a <f> de UMA celula e deixa so o valor gravado."""
    fim = celula.find(">")
    abertura = celula[: fim + 1]
    if abertura.endswith("/>"):
        return celula
    corpo = RE_ELEMENTO_F.sub("", celula[fim + 1:])
    aba.formulas += 1

    # cm= liga a celula ao registro de matriz dinamica de xl/metadata.xml -
    # e propriedade da formula, e vai embora com ela.
    abertura = RE_ATRIBUTO_CM.sub("", abertura)
    valor = xu.RE_VALOR.search(corpo)
    if valor is None:
        aba.sem_valor += 1
        return _celula_vazia(abertura)

    tipo = RE_ATRIBUTO_T.search(abertura)
    if tipo and tipo.group(1) == "str":
        # t="str" e o texto devolvido por uma formula; sem a formula, o jeito
        # certo de guardar texto na propria celula e o inlineStr. O texto do
        # <v> ja vem escapado, entao passa como esta.
        if not valor.group(1):
            return _celula_vazia(abertura)
        abertura = RE_ATRIBUTO_T.sub(' t="inlineStr"', abertura)
        return (abertura + '<is><t xml:space="preserve">%s</t></is></c>'
                % valor.group(1))
    return abertura + corpo


def _limpar_trecho(texto: str, aba: AbaSemFormulas) -> str:
    """Troca as formulas de um trecho de linhas COMPLETAS do <sheetData>.

    Nao passa celula por celula: procura direto cada <f> e so remonta a celula
    em volta dela. A <f> e sempre o primeiro filho da celula, entao o "<c"
    mais proximo antes dela e a abertura da propria celula.
    """
    ultima = texto.rfind("<row")
    if ultima != -1:
        numero = RE_NUMERO_LINHA.match(texto, ultima)
        if numero:
            aba.linhas = max(aba.linhas, int(numero.group(1)))

    partes = []
    cursor = 0
    for achado in RE_ABRE_F.finditer(texto):
        if achado.start() < cursor:
            continue  # segunda <f> da mesma celula (nao acontece no Excel)
        inicio = texto.rfind("<c", cursor, achado.start())
        fim = texto.find("</c>", achado.start())
        if inicio == -1 or fim == -1:
            continue
        fim += len("</c>")
        partes.append(texto[cursor:inicio])
        partes.append(_sem_formula(texto[inicio:fim], aba))
        cursor = fim
    if not partes:
        return texto
    partes.append(texto[cursor:])
    return "".join(partes)


# ---------------------------------------------------------------------------
# Abas - lidas em pedacos
# ---------------------------------------------------------------------------

def _limpar_aba(origem, destino, aba: AbaSemFormulas, avancar) -> None:
    """Copia uma aba de origem para destino, sem as formulas.

    Tres trechos: o cabecalho ate <sheetData>, as linhas, e o rodape depois de
    </sheetData>. So as linhas sao mexidas - fora delas ha outras "formulas"
    (formatacao condicional, validacao de dados) que devem ficar como estao.
    """
    leitor = codecs.getincrementaldecoder("utf-8")()
    estado = "cabecalho"
    pendente = ""
    # O cabecalho (dimensao, paineis, larguras) tem poucos KB; fica guardado
    # inteiro so para ler a <dimension> mesmo que ela venha cortada ao meio.
    cabecalho = []

    def gravar(texto):
        if texto:
            destino.write(texto.encode("utf-8"))

    while True:
        bruto = origem.read(TAMANHO_PEDACO)
        acabou = not bruto
        pendente += leitor.decode(bruto, final=acabou)
        avancar(len(bruto))

        while True:
            if estado == "cabecalho":
                achado = pendente.find("<sheetData")
                fecha = pendente.find(">", achado) if achado != -1 else -1
                if fecha == -1:
                    # Guarda o finalzinho: a tag pode ter vindo cortada.
                    corte = max(0, len(pendente) - 16) if not acabou else len(pendente)
                    if achado != -1:
                        corte = min(corte, achado)
                    gravar(pendente[:corte])
                    cabecalho.append(pendente[:corte])
                    pendente = pendente[corte:]
                    break
                cabecalho.append(pendente[:achado])
                _ler_dimensao("".join(cabecalho), aba)
                gravar(pendente[:fecha + 1])
                vazia = pendente[fecha - 1] == "/"
                pendente = pendente[fecha + 1:]
                estado = "rodape" if vazia else "linhas"
            elif estado == "linhas":
                fim = pendente.find("</sheetData>")
                if fim != -1:
                    gravar(_limpar_trecho(pendente[:fim], aba))
                    pendente = pendente[fim:]
                    estado = "rodape"
                    continue
                corte = pendente.rfind("</row>")
                if corte != -1:
                    corte += len("</row>")
                    gravar(_limpar_trecho(pendente[:corte], aba))
                    pendente = pendente[corte:]
                break
            else:
                gravar(pendente)
                pendente = ""
                break

        if acabou:
            gravar(pendente)  # arquivo truncado: devolve o que sobrou como veio
            return


def _ler_dimensao(cabecalho: str, aba: AbaSemFormulas) -> None:
    """Numero de colunas a partir de <dimension ref="A1:AB290"/>."""
    achado = RE_DIMENSAO.search(cabecalho)
    if not achado:
        return
    try:
        coluna, _ = xu.separar_referencia(achado.group(1).split(":")[-1])
    except ValueError:
        return
    aba.colunas = coluna


# ---------------------------------------------------------------------------
# Pacote
# ---------------------------------------------------------------------------

def _texto(pacote: zipfile.ZipFile, nome: str) -> str:
    return pacote.read(nome).decode("utf-8") if nome in pacote.NameToInfo else ""


def _partes_de_aba(tipos: str, nomes) -> list:
    """Partes declaradas como aba em [Content_Types].xml."""
    partes = []
    for item in RE_OVERRIDE.finditer(tipos):
        parte = RE_PARTNAME.search(item.group(0))
        tipo = RE_CONTENTTYPE.search(item.group(0))
        if parte and tipo and tipo.group(1).endswith(TIPO_ABA) and parte.group(1) in nomes:
            partes.append(parte.group(1))
    return partes


def _sem_calcchain(nome: str, texto: str) -> str:
    """Tira as referencias a xl/calcChain.xml dos tipos e das relacoes."""
    if nome == PARTE_TIPOS:
        return re.sub(r'<Override\b[^>]*PartName="/%s"[^>]*/>'
                      % re.escape(PARTE_CALCCHAIN), "", texto)
    return re.sub(r'<Relationship\b[^>]*Target="[^"]*calcChain\.xml"[^>]*/>',
                  "", texto)


def nomes_das_abas(caminho) -> list:
    """Nomes das abas, lendo so xl/workbook.xml e as relacoes (poucos KB)."""
    with zipfile.ZipFile(caminho, "r") as pacote:
        return [nome for nome, _, _ in inspector.mapear_abas_xml(
            _texto(pacote, PARTE_PASTA), _texto(pacote, PARTE_RELS),
            set(pacote.NameToInfo))]


def _info_de_saida(info: zipfile.ZipInfo) -> zipfile.ZipInfo:
    novo = zipfile.ZipInfo(info.filename, date_time=info.date_time)
    novo.compress_type = zipfile.ZIP_DEFLATED
    return novo


def tirar_formulas(caminho, destino, progresso=None) -> ResultadoSemFormulas:
    """Grava em destino uma copia de caminho so com os valores.

    progresso(fracao) e chamado enquanto as abas sao lidas, de 0.0 a 1.0.
    Nunca altera o original; um erro vira ResultadoSemFormulas.erro.
    """
    caminho = Path(caminho)
    destino = Path(destino)
    resultado = ResultadoSemFormulas(arquivo=caminho)
    temporario = destino.with_name(destino.name + ".tmp")

    try:
        if destino.resolve() == caminho.resolve():
            raise ValueError("a saida apontaria para o proprio arquivo original")
        try:
            entrada = zipfile.ZipFile(caminho, "r")
        except zipfile.BadZipFile:
            raise ValueError("nao e um arquivo .xlsx/.xlsm valido "
                             "(.xls e .xlsb precisam ser convertidos antes)")

        with entrada:
            nomes = set(entrada.NameToInfo)
            if PARTE_PASTA not in nomes:
                raise ValueError("Nao parece um arquivo Excel valido "
                                 "(xl/workbook.xml ausente).")
            tipos = _texto(entrada, PARTE_TIPOS)
            partes_aba = _partes_de_aba(tipos, nomes)

            nome_da_parte = {
                parte: nome for nome, parte, _ in inspector.mapear_abas_xml(
                    _texto(entrada, PARTE_PASTA), _texto(entrada, PARTE_RELS), nomes)
            }
            # Na ordem das abas na pasta, e depois alguma que so esteja nos tipos.
            ordem = [parte for parte in nome_da_parte if parte in partes_aba]
            ordem += [parte for parte in partes_aba if parte not in ordem]
            resultado.abas = [AbaSemFormulas(nome=nome_da_parte.get(parte, parte),
                                             parte=parte) for parte in ordem]
            aba_da_parte = {aba.parte: aba for aba in resultado.abas}

            total = sum(entrada.getinfo(parte).file_size for parte in ordem) or 1
            lidos = [0]

            def avancar(quantidade):
                lidos[0] += quantidade
                if progresso:
                    progresso(min(1.0, lidos[0] / float(total)))

            # [Content_Types].xml vai primeiro no pacote, como o Excel grava.
            infos = sorted(entrada.infolist(),
                           key=lambda info: info.filename != PARTE_TIPOS)
            destino.parent.mkdir(parents=True, exist_ok=True)
            with zipfile.ZipFile(temporario, "w", zipfile.ZIP_DEFLATED) as saida:
                for info in infos:
                    nome = info.filename
                    if nome.endswith("/") or nome == PARTE_CALCCHAIN:
                        continue
                    novo = _info_de_saida(info)
                    if nome in (PARTE_TIPOS, PARTE_RELS):
                        saida.writestr(novo, _sem_calcchain(
                            nome, _texto(entrada, nome)).encode("utf-8"))
                        continue
                    grande = info.file_size >= LIMITE_ZIP64
                    with entrada.open(info) as origem, \
                            saida.open(novo, "w", force_zip64=grande) as copia:
                        if nome in aba_da_parte:
                            _limpar_aba(origem, copia, aba_da_parte[nome], avancar)
                        else:
                            shutil.copyfileobj(origem, copia, TAMANHO_PEDACO)

        if destino.exists():
            destino.unlink()
        shutil.move(str(temporario), str(destino))
        resultado.saida = destino
    except Exception as exc:  # arquivo corrompido, disco cheio, etc.
        resultado.erro = str(exc) or exc.__class__.__name__
    finally:
        if temporario.exists():
            temporario.unlink()
    return resultado
