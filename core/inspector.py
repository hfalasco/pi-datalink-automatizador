# -*- coding: utf-8 -*-
"""Leitura (somente leitura) da estrutura de uma planilha.

Descobre, sem alterar nada:
  * quais abas existem e qual parte XML corresponde a cada uma;
  * quais colunas possuem TAG valida na linha configurada;
  * quais colunas trazem uma EXPRESSAO DE FILTRO no lugar da TAG - essas saem
    com PITimeFilter, contando o tempo em que a condicao foi verdadeira;
  * quais colunas combinam MAIS DE UMA TAG na mesma celula (ver core/composta.py)
    e quais colunas existem apenas para guardar a TAG dessas contas;
  * o estilo das celulas de data e de formula (para reaproveitar formatacao);
  * o valor do atributo "cm" usado pelo PI DataLink nas formulas de matriz.
"""

from __future__ import annotations

import re
import unicodedata
from dataclasses import dataclass, field

from . import composta
from . import xml_utils as xu
from .xml_utils import PacoteXlsx

RE_ABA = re.compile(r"<sheet\b[^>]*/>")
RE_NOME_ABA = re.compile(r'\bname="([^"]*)"')
RE_RID_ABA = re.compile(r'\br:id="([^"]+)"')
RE_ESTADO = re.compile(r'\bstate="([^"]+)"')
RE_RELACAO = re.compile(r"<Relationship\b[^>]*/>")
RE_REL_ID = re.compile(r'\bId="([^"]+)"')
RE_REL_ALVO = re.compile(r'\bTarget="([^"]+)"')
RE_STRING_COMPARTILHADA = re.compile(r"<si\b[^>]*>(.*?)</si>|<si\b[^>]*/>", re.S)
RE_TEXTO = re.compile(r"<t\b[^>]*>(.*?)</t>", re.S)
RE_FORMULA = re.compile(r"<f\b[^>]*>(.*?)</f>", re.S)

TIPOS_TEXTO = ("s", "str", "inlineStr")


def _normalizar(texto: str) -> str:
    """Minusculas, sem acento e sem espacos - usado para comparar rotulos."""
    base = unicodedata.normalize("NFKD", texto or "")
    base = "".join(c for c in base if not unicodedata.combining(c))
    return base.lower().strip().replace(" ", "")


@dataclass
class TagInfo:
    """Uma coluna que sera preenchida: TAG propria ou conta de varias TAGs."""

    coluna: str                       # letra da coluna, exemplo: B
    indice: int                       # indice numerico, exemplo: 2
    texto: str                        # conteudo lido da linha de TAGs
    titulo: str = ""                  # conteudo da linha 1 (nome do indicador)
    estilo_dados: str | None = None   # atributo s= das celulas de dados

    # Formula propria da planilha, preservada como esta (ver core/composta.py).
    # None = coluna comum, que recebe a formula padrao do programa.
    personalizada: object = None
    # Texto das TAGs usadas por essa formula, na ordem em que aparecem.
    fontes: list = field(default_factory=list)
    # True quando a celula da linha de TAGs guarda uma EXPRESSAO DE FILTRO no
    # lugar da TAG. A coluna sai com PITimeFilter: quanto tempo a condicao foi
    # verdadeira em cada intervalo, e nao a media de uma TAG.
    filtro: bool = False

    @property
    def composta(self) -> bool:
        """True quando a coluna combina mais de uma TAG na mesma celula."""
        return bool(self.personalizada) and self.personalizada.composta

    @property
    def consultas(self) -> int:
        """Quantas TAGs esta coluna consulta por formula."""
        return self.personalizada.total_chamadas if self.personalizada else 1

    def descricao(self) -> str:
        """Como a coluna aparece nos relatorios."""
        if self.personalizada and self.fontes:
            return " + ".join(self.fontes) if self.composta else self.fontes[0]
        return self.texto


@dataclass
class AbaInfo:
    nome: str
    parte: str                        # exemplo: xl/worksheets/sheet1.xml
    oculta: bool = False
    tags: list = field(default_factory=list)
    # Colunas que so guardam a TAG de uma coluna composta: tem TAG na linha 2,
    # nao tem formula propria e sao citadas pela formula de outra coluna.
    # Ficam intactas - preenche-las criaria colunas de dados indesejadas.
    auxiliares: list = field(default_factory=list)
    # [(coluna, motivo)] - colunas cuja formula o programa nao sabe reescrever
    # (conta que nao consulta o PI, funcao do PI fora do escopo...). Servem de
    # aviso: o programa nao as estende nem as ajusta ao novo periodo.
    nao_tratadas: list = field(default_factory=list)
    estilo_data: str | None = None    # atributo s= da coluna de data
    cm: str | None = None             # atributo cm= das formulas de matriz
    ultima_linha: int = 0
    motivo_ignorada: str = ""

    @property
    def processavel(self) -> bool:
        return bool(self.tags)

    @property
    def compostas(self) -> list:
        return [tag for tag in self.tags if tag.composta]

    @property
    def personalizadas(self) -> list:
        return [tag for tag in self.tags if tag.personalizada]

    @property
    def filtradas(self) -> list:
        """Colunas de tempo por filtro geradas pelo programa (PITimeFilter)."""
        return [tag for tag in self.tags if tag.filtro and not tag.personalizada]


@dataclass
class ArquivoInfo:
    caminho: object
    abas: list = field(default_factory=list)
    erro: str = ""

    @property
    def abas_processaveis(self) -> list:
        return [aba for aba in self.abas if aba.processavel]

    @property
    def total_tags(self) -> int:
        return sum(len(aba.tags) for aba in self.abas_processaveis)


# ---------------------------------------------------------------------------
# Leitura das partes
# ---------------------------------------------------------------------------

def ler_strings_compartilhadas(pacote: PacoteXlsx) -> list:
    """Le xl/sharedStrings.xml (onde o Excel guarda os textos das celulas)."""
    nome = "xl/sharedStrings.xml"
    if nome not in pacote.partes:
        return []
    xml = pacote.texto(nome)
    strings = []
    for bloco in RE_STRING_COMPARTILHADA.finditer(xml):
        conteudo = bloco.group(1) or ""
        # Uma <si> pode estar quebrada em varios <r><t>..</t></r> (rich text).
        strings.append(xu.desescapar_xml("".join(RE_TEXTO.findall(conteudo))))
    return strings


def mapear_abas(pacote: PacoteXlsx) -> list:
    """Devolve [(nome, parte, oculta)] na ordem em que aparecem na pasta."""
    return mapear_abas_xml(pacote.texto("xl/workbook.xml"),
                           pacote.texto("xl/_rels/workbook.xml.rels"),
                           pacote.partes)


def mapear_abas_xml(workbook: str, rels: str, partes) -> list:
    """O mesmo que mapear_abas, a partir do XML das duas partes.

    Serve a quem nao quer carregar o pacote inteiro so para saber o nome das
    abas (ver core/sem_formulas.py). partes = nomes das partes do arquivo.
    """
    destino_por_id = {}
    for relacao in RE_RELACAO.finditer(rels):
        trecho = relacao.group(0)
        ident = RE_REL_ID.search(trecho)
        alvo = RE_REL_ALVO.search(trecho)
        if not ident or not alvo:
            continue
        caminho = alvo.group(1)
        if caminho.startswith("/"):
            caminho = caminho.lstrip("/")
        elif not caminho.startswith("xl/"):
            caminho = "xl/" + caminho
        destino_por_id[ident.group(1)] = caminho.replace("xl/../", "")

    abas = []
    for bruto in RE_ABA.finditer(workbook):
        trecho = bruto.group(0)
        nome = RE_NOME_ABA.search(trecho)
        rid = RE_RID_ABA.search(trecho)
        if not nome or not rid:
            continue
        parte = destino_por_id.get(rid.group(1))
        if not parte or parte not in partes:
            continue
        estado = RE_ESTADO.search(trecho)
        abas.append(
            (
                xu.desescapar_xml(nome.group(1)),
                parte,
                bool(estado and estado.group(1) != "visible"),
            )
        )
    return abas


def _cabecalho(celula_xml: str) -> str:
    """Somente a tag de abertura da celula, onde ficam os atributos."""
    return celula_xml[: celula_xml.find(">") + 1]


def _tipo_da_celula(celula_xml: str) -> str:
    tipo = xu.RE_ATRIBUTO_T.search(_cabecalho(celula_xml))
    return tipo.group(1) if tipo else "n"


def valor_da_celula(celula_xml: str, strings: list) -> str:
    """Extrai o valor textual de uma celula, resolvendo sharedStrings."""
    tipo = _tipo_da_celula(celula_xml)
    if tipo == "inlineStr":
        achado = xu.RE_TEXTO_INLINE.search(celula_xml)
        return xu.desescapar_xml(achado.group(1)) if achado else ""
    valor = xu.RE_VALOR.search(celula_xml)
    if not valor:
        return ""
    bruto = valor.group(1)
    if tipo == "s":
        try:
            return strings[int(bruto)]
        except (ValueError, IndexError):
            return ""
    if tipo in ("str", "e"):
        return xu.desescapar_xml(bruto)
    return bruto  # numerico ou booleano, devolvido como esta


def formula_da_celula(celula_xml: str) -> str:
    """Texto da formula de uma celula, ja desescapado. Vazio se nao houver."""
    achado = RE_FORMULA.search(celula_xml or "")
    return xu.desescapar_xml(achado.group(1)) if achado else ""


def e_tag_valida(texto: str, tipo: str, cfg) -> bool:
    """Regra de deteccao derivada da analise real dos templates.

    Rejeita celula vazia, celula numerica (o caso real B2 = 0 na aba
    "Indicadores por hora" do Caldo Misto) e rotulos como "data".
    Aceita caminhos PI e qualquer outro texto nao vazio.
    """
    if tipo not in TIPOS_TEXTO:
        return False
    if not texto or not texto.strip():
        return False
    limpo = texto.strip()
    ignorados = {_normalizar(rotulo) for rotulo in cfg.rotulos_ignorados}
    if _normalizar(limpo) in ignorados:
        return False
    if cfg.padrao_tag and not re.search(cfg.padrao_tag, limpo):
        return False
    return True


# ---------------------------------------------------------------------------
# Formulas que a planilha ja traz prontas
# ---------------------------------------------------------------------------

def _indice_de_coluna(letra: str):
    try:
        return xu.letra_para_indice(letra)
    except ValueError:
        return None


def _mesma_aba(prefixo: str, nome_da_aba: str) -> bool:
    """True quando 'Indicadores de Horas'! aponta para a propria aba."""
    if not prefixo:
        return True
    limpo = prefixo.rstrip("!").strip().strip(chr(39))
    return _normalizar(limpo) == _normalizar(nome_da_aba)


def _analisar_formulas_existentes(nome_da_aba: str, celulas_dados: dict,
                                  indice_data: int, cfg) -> tuple:
    """Le a primeira linha de dados e separa dois grupos de colunas.

    personalizadas : {indice: FormulaPersonalizada} - colunas cuja formula NAO
        e a formula padrao de uma TAG so (soma de duas vazoes, IFERROR, fator
        de conversao...). Elas sao reemitidas com a mesma conta.
    auxiliares : {indice} - colunas que so existem para guardar a TAG usada por
        uma dessas contas. Ficam intactas.
    nao_tratadas : [(coluna, motivo)] - colunas com formula que o programa nao
        sabe reescrever. So viram aviso.
    """
    if not cfg.preservar_formulas_personalizadas:
        return {}, set(), []

    analises = {}
    nao_tratadas = []
    for indice, celula_xml in celulas_dados.items():
        if indice == indice_data or indice > cfg.max_colunas_varredura:
            continue
        formula = formula_da_celula(celula_xml)
        if not formula:
            continue
        analise = composta.analisar(formula)
        if analise is not None:
            analises[indice] = analise
        else:
            nao_tratadas.append(
                (
                    xu.indice_para_letra(indice),
                    "usa uma funcao do PI fora do escopo do programa"
                    if composta.RE_FUNCAO_PI.search(formula)
                    else "conta que nao consulta o PI",
                )
            )

    personalizadas = {
        indice: analise
        for indice, analise in analises.items()
        if not composta.e_formula_padrao(
            analise, xu.indice_para_letra(indice), cfg.linha_tags
        )
    }

    auxiliares = set()
    if cfg.ignorar_colunas_auxiliares:
        for indice, analise in personalizadas.items():
            # Vale para qualquer celula citada pela conta - a que guarda a TAG
            # e tambem a que guarda a expressao de filtro do PIAdvCalcFilVal.
            for referencia in analise.celulas_usadas:
                prefixo, coluna, linha = composta.referencia_de_celula(referencia)
                if linha != cfg.linha_tags or not _mesma_aba(prefixo, nome_da_aba):
                    continue
                alvo = _indice_de_coluna(coluna)
                # Uma coluna que tem formula propria e coluna de dados de
                # verdade, mesmo sendo citada por outra. So entra aqui a que
                # existe unicamente para alimentar a conta de outra coluna.
                if alvo is None or alvo == indice or alvo == indice_data:
                    continue
                if alvo in analises:
                    continue
                auxiliares.add(alvo)
    return personalizadas, auxiliares, nao_tratadas


def _tags_de_origem(analise, celulas_tag: dict, strings: list) -> list:
    """Texto do que cada chamada consulta, para mostrar nos relatorios.

    Normalmente e a TAG. Nas chamadas da familia PITimeFilter e a expressao de
    filtro - que pode estar escrita na propria formula ou guardada numa celula,
    caso em que o texto da celula e que aparece.
    """
    textos = []
    for chamada in analise.chamadas:
        referencia = chamada.tag or chamada.expressao
        if not referencia:
            continue
        _, coluna, _linha = composta.referencia_de_celula(referencia)
        indice = _indice_de_coluna(coluna) if coluna else None
        celula = celulas_tag.get(indice) if indice else None
        valor = valor_da_celula(celula, strings).strip() if celula else ""
        textos.append(valor or referencia.replace("$", ""))
    return textos


# ---------------------------------------------------------------------------
# Inspecao
# ---------------------------------------------------------------------------

def inspecionar_aba(pacote: PacoteXlsx, nome: str, parte: str, oculta: bool,
                    cfg, strings: list) -> AbaInfo:
    aba = AbaInfo(nome=nome, parte=parte, oculta=oculta)
    xml = pacote.texto(parte)
    _, miolo, _ = xu.dividir_sheet(xml)

    indice_data = xu.letra_para_indice(cfg.coluna_data)
    linhas_de_interesse = {1, cfg.linha_tags, cfg.primeira_linha_dados}
    coletadas = {numero: {} for numero in linhas_de_interesse}

    for linha_m in xu.RE_LINHA.finditer(miolo):
        linha_xml = linha_m.group(0)
        ref_linha = xu.RE_ATRIBUTO_R.search(xu.abertura_da_linha(linha_xml))
        if not ref_linha:
            continue
        numero = int(ref_linha.group(1))
        if numero > aba.ultima_linha:
            aba.ultima_linha = numero
        if numero not in linhas_de_interesse:
            continue
        for celula_m in xu.RE_CELULA.finditer(linha_xml):
            celula_xml = celula_m.group(0)
            ref = xu.RE_ATRIBUTO_R.search(_cabecalho(celula_xml))
            if not ref:
                continue
            coluna, _ = xu.separar_referencia(ref.group(1))
            coletadas[numero][coluna] = celula_xml

    celulas_titulo = coletadas.get(1, {})
    celulas_tag = coletadas.get(cfg.linha_tags, {})
    celulas_dados = coletadas.get(cfg.primeira_linha_dados, {})

    personalizadas, auxiliares, aba.nao_tratadas = _analisar_formulas_existentes(
        nome, celulas_dados, indice_data, cfg
    )

    # --- colunas que serao preenchidas -----------------------------------
    for indice in sorted(set(celulas_tag) | set(personalizadas)):
        if indice == indice_data:
            continue
        if indice > cfg.max_colunas_varredura:
            break

        celula_xml = celulas_tag.get(indice, "")
        texto = valor_da_celula(celula_xml, strings) if celula_xml else ""
        tipo = _tipo_da_celula(celula_xml) if celula_xml else ""
        # Celula com expressao de filtro no lugar da TAG: a coluna vira tempo
        # por filtro. A checagem vem antes de e_tag_valida porque um padrao_tag
        # restrito a caminhos PI (^\\\\) recusaria a expressao.
        filtro = bool(
            cfg.tempo_filtrado
            and tipo in TIPOS_TEXTO
            and composta.e_expressao_de_filtro(texto)
        )
        tag_propria = bool(celula_xml) and (
            filtro or e_tag_valida(texto, tipo, cfg)
        )
        analise = personalizadas.get(indice)

        if indice in auxiliares:
            # Coluna de apoio: guarda a TAG de uma coluna composta. Fica como
            # esta, para nao virar uma coluna de dados e nao gastar consulta.
            aba.auxiliares.append(
                TagInfo(
                    coluna=xu.indice_para_letra(indice),
                    indice=indice,
                    texto=texto.strip(),
                )
            )
            continue

        if analise is None and not tag_propria:
            continue

        celula_dados = celulas_dados.get(indice, "")
        estilo = xu.RE_ATRIBUTO_S.search(_cabecalho(celula_dados)) if celula_dados else None
        titulo_xml = celulas_titulo.get(indice, "")
        aba.tags.append(
            TagInfo(
                coluna=xu.indice_para_letra(indice),
                indice=indice,
                texto=texto.strip(),
                titulo=valor_da_celula(titulo_xml, strings).strip() if titulo_xml else "",
                estilo_dados=estilo.group(1) if estilo else None,
                personalizada=analise,
                fontes=_tags_de_origem(analise, celulas_tag, strings) if analise else [],
                filtro=filtro,
            )
        )

    if not aba.tags:
        aba.motivo_ignorada = "nenhuma TAG valida na linha %d" % cfg.linha_tags

    # --- estilo ja usado na coluna de data -------------------------------
    celula_data = celulas_dados.get(indice_data, "")
    if celula_data:
        estilo = xu.RE_ATRIBUTO_S.search(_cabecalho(celula_data))
        aba.estilo_data = estilo.group(1) if estilo else None

    # --- atributo cm das formulas de matriz ------------------------------
    for indice, celula_xml in celulas_dados.items():
        if indice == indice_data:
            continue
        cm = xu.RE_ATRIBUTO_CM.search(_cabecalho(celula_xml))
        if cm:
            aba.cm = cm.group(1)
            break

    return aba


def inspecionar_pacote(pacote: PacoteXlsx, cfg) -> ArquivoInfo:
    info = ArquivoInfo(caminho=pacote.caminho)
    strings = ler_strings_compartilhadas(pacote)

    cm_global = None
    for nome, parte, oculta in mapear_abas(pacote):
        aba = inspecionar_aba(pacote, nome, parte, oculta, cfg, strings)
        if cm_global is None:
            cm_global = aba.cm
        info.abas.append(aba)

    # Abas ainda sem formulas herdam o "cm" de outra aba do mesmo arquivo.
    if cm_global is None and "xl/metadata.xml" in pacote.partes:
        if "XLDAPR" in pacote.texto("xl/metadata.xml"):
            cm_global = "1"
    for aba in info.abas:
        if aba.cm is None:
            aba.cm = cm_global
    return info


def inspecionar_arquivo(caminho, cfg) -> ArquivoInfo:
    """Le a estrutura completa de um arquivo. Nunca escreve nada."""
    try:
        pacote = PacoteXlsx.abrir(caminho)
    except Exception as exc:  # zip corrompido, arquivo bloqueado, etc.
        return ArquivoInfo(caminho=caminho, erro=str(exc))
    return inspecionar_pacote(pacote, cfg)
