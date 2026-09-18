# -*- coding: utf-8 -*-
"""Motor de escrita: cirurgia direta no XML da planilha.

Por que nao openpyxl
--------------------
Um teste de ida e volta com openpyxl nos templates reais mostrou a perda das
seguintes partes do pacote:

    xl/metadata.xml                    (propriedades de matriz dinamica)
    xl/webextensions/taskpanes.xml     (painel do PI DataLink)
    xl/webextensions/webextension1.xml
    xl/printerSettings/printerSettings1.bin
    xl/worksheets/_rels/sheet1.xml.rels

e tambem dos atributos aca="1", ca="1" e cm="1" das formulas de matriz - sendo
que ca="1" e justamente o sinalizador de "recalcular sempre" que o PI DataLink
usa.

Este modulo copia TODAS as partes byte a byte e reescreve apenas o <sheetData>
das abas alvo. Tudo o que nao foi pedido para mudar permanece identico:
formatacao, larguras, alturas, mesclagens, filtros, tabelas, congelamento de
paineis, estilos, nomes das abas, macros (.xlsm) e impressao.

A unica parte deliberadamente removida e xl/calcChain.xml, um cache da ordem de
calculo. Mante-lo desatualizado depois de inserir celulas e o que provoca o
aviso "conteudo ilegivel" do Excel; removido, o Excel o reconstroi em silencio.
"""

from __future__ import annotations

import re
from dataclasses import dataclass, field

from . import composta
from . import metadata
from . import xml_utils as xu
from . import styles
from ..config import MODO_BLOCO, MODO_LINHA
from .formula import intervalo_pi, montar_formula_bloco_xml, montar_formula_xml
from .inspector import AbaInfo
from .timeline import formatar_serial, para_serial_excel
from .xml_utils import PacoteXlsx

PARTE_CALCCHAIN = "xl/calcChain.xml"


@dataclass
class ResultadoAba:
    nome: str
    tags: int = 0
    linhas_formula: int = 0
    linhas_data: int = 0
    formulas_escritas: int = 0
    consultas_pi: int = 0        # quantas idas ao servidor PI estas formulas geram
    modo: str = ""
    matriz: str = ""             # "dinamica" ou "cse" (so no modo bloco)
    compostas: int = 0           # colunas que combinam mais de uma TAG
    personalizadas: int = 0      # colunas com conta propria preservada
    auxiliares: int = 0          # colunas que so guardam TAG (mantidas intactas)
    amostra: list = field(default_factory=list)


# ---------------------------------------------------------------------------
# Construcao de celulas
# ---------------------------------------------------------------------------

def _celula_data(referencia: str, estilo, serial: str) -> str:
    atributo = ' s="%s"' % estilo if estilo else ""
    return "<c r=%s%s><v>%s</v></c>" % (_aspas(referencia), atributo, serial)


def _celula_formula(referencia: str, estilo, cm, formula_xml: str,
                    matriz: bool, cache: str = "", ref_matriz: str = "") -> str:
    """Celula com formula. ref_matriz permite uma matriz multicelula (B3:B290)."""
    atributos = ""
    if estilo:
        atributos += ' s="%s"' % estilo
    if cache:
        atributos += cache  # devolve algo como ' t="e"'
    if cm:
        atributos += ' cm="%s"' % cm
    if matriz:
        interno = '<f t="array" aca="1" ref=%s ca="1">%s</f>' % (
            _aspas(ref_matriz or referencia), formula_xml
        )
    else:
        interno = '<f ca="1">%s</f>' % formula_xml
    return "<c r=%s%s>%s</c>" % (_aspas(referencia), atributos, interno)


def _celula_vazia(referencia: str, estilo) -> str:
    """Celula sem conteudo, preservando o estilo.

    E assim que o Excel representa o corpo de uma matriz multicelula: so a
    celula do canto superior esquerdo carrega o <f>.
    """
    atributo = ' s="%s"' % estilo if estilo else ""
    return "<c r=%s%s/>" % (_aspas(referencia), atributo)


def _aspas(valor: str) -> str:
    return chr(34) + valor + chr(34)


# ---------------------------------------------------------------------------
# Leitura e reconstrucao do sheetData
# ---------------------------------------------------------------------------

def _indexar_linhas(miolo: str) -> dict:
    """{numero_da_linha: {"abertura": str, "celulas": {indice_coluna: xml}}}."""
    linhas = {}
    for linha_m in xu.RE_LINHA.finditer(miolo):
        linha_xml = linha_m.group(0)
        abertura = xu.abertura_da_linha(linha_xml)
        ref = xu.RE_ATRIBUTO_R.search(abertura)
        if not ref:
            continue
        numero = int(ref.group(1))
        celulas = {}
        for celula_m in xu.RE_CELULA.finditer(linha_xml):
            celula_xml = celula_m.group(0)
            cabecalho = celula_xml[: celula_xml.find(">") + 1]
            ref_celula = xu.RE_ATRIBUTO_R.search(cabecalho)
            if not ref_celula:
                continue
            coluna, _ = xu.separar_referencia(ref_celula.group(1))
            celulas[coluna] = celula_xml
        linhas[numero] = {"abertura": abertura, "celulas": celulas}
    return linhas


def _serializar(linhas: dict) -> str:
    partes = []
    for numero in sorted(linhas):
        dados = linhas[numero]
        celulas = dados["celulas"]
        if not celulas:
            # Linha vazia so e mantida se carregava formatacao propria.
            if "ht=" not in dados["abertura"] and "s=" not in dados["abertura"]:
                continue
        colunas = sorted(celulas)
        atributos = xu.atributos_sem(dados["abertura"], ("r", "spans"))
        spans = ""
        if colunas:
            spans = ' spans="%d:%d"' % (colunas[0], colunas[-1])
        sufixo = " " + atributos if atributos else ""
        partes.append(
            "<row r=%s%s%s>%s</row>"
            % (
                _aspas(str(numero)),
                spans,
                sufixo,
                "".join(celulas[coluna] for coluna in colunas),
            )
        )
    return "".join(partes)


def _atualizar_dimensao(cabecalho: str, linhas: dict) -> str:
    if not linhas:
        return cabecalho
    numeros = sorted(linhas)
    colunas = [c for dados in linhas.values() for c in dados["celulas"]]
    if not colunas:
        return cabecalho
    referencia = "%s%d:%s%d" % (
        xu.indice_para_letra(min(colunas)),
        numeros[0],
        xu.indice_para_letra(max(colunas)),
        numeros[-1],
    )
    novo = "<dimension ref=%s/>" % _aspas(referencia)
    if re.search(r"<dimension\b[^>]*/>", cabecalho):
        return re.sub(r"<dimension\b[^>]*/>", novo, cabecalho, count=1)
    abertura = re.search(r"<worksheet\b[^>]*>", cabecalho)
    if abertura:
        return cabecalho[: abertura.end()] + novo + cabecalho[abertura.end():]
    return cabecalho


# ---------------------------------------------------------------------------
# Reescrita de uma aba
# ---------------------------------------------------------------------------

def _intervalo_dos_carimbos(carimbos: list, cfg) -> int:
    """Minutos entre dois carimbos consecutivos.

    Lido da propria sequencia gerada, e nao de cfg.intervalo_minutos, porque a
    aba pode ter intervalo proprio (ver Config.configuracoes_por_aba).
    """
    if len(carimbos) >= 2:
        return int((carimbos[1] - carimbos[0]).total_seconds() // 60)
    return cfg.intervalo_minutos


def reescrever_aba(pacote: PacoteXlsx, aba: AbaInfo, carimbos: list, cfg,
                   estilo_data_extra=None, simular: bool = False,
                   cm_dinamico=None) -> ResultadoAba:
    """Reescreve a coluna de data e as colunas com TAG de uma aba.

    Quando simular=True nada e gravado no pacote; apenas o relatorio e a
    amostra de formulas sao devolvidos (usado pelo modo de teste).

    cm_dinamico e o indice de metadata que marca a formula como matriz
    dinamica (ver core/metadata.py). Ele vem pronto de processar_pacote, que ja
    criou a parte xl/metadata.xml caso o arquivo nao a tivesse.
    """
    resultado = ResultadoAba(nome=aba.nome, tags=len(aba.tags))
    if not aba.tags or not carimbos:
        return resultado

    xml = pacote.texto(aba.parte)
    cabecalho, miolo, rodape = xu.dividir_sheet(xml)
    linhas = _indexar_linhas(miolo)

    indice_data = xu.letra_para_indice(cfg.coluna_data)
    primeira = cfg.primeira_linha_dados
    total_carimbos = len(carimbos)
    total_formulas = total_carimbos - 1 if cfg.carimbo_final_extra else total_carimbos
    ultima_linha_gerada = primeira + total_carimbos - 1

    resultado.linhas_data = total_carimbos
    resultado.linhas_formula = total_formulas

    estilo_data = estilo_data_extra or aba.estilo_data
    colunas_gerenciadas = {indice_data} | {tag.indice for tag in aba.tags}
    modelo_abertura = linhas.get(primeira, {}).get("abertura", "<row>")

    modo_bloco = cfg.modo_formula == MODO_BLOCO
    resultado.modo = cfg.modo_formula
    # O indice garantido pelo pacote tem prioridade sobre o que a aba ja usava:
    # arquivos sem xl/metadata.xml nao tinham nenhum, e era justamente isso que
    # empurrava o programa para a matriz CSE (bloco travado).
    cm_matriz = cm_dinamico or aba.cm
    ultima_linha_formula = primeira + total_formulas - 1
    intervalo_minutos = _intervalo_dos_carimbos(carimbos, cfg)

    # Colunas com conta propria (soma de duas vazoes, IFERROR, fator...) sao
    # regravadas com a MESMA conta, so com as datas trocadas. Ver core/composta.py.
    comuns = [tag for tag in aba.tags if not tag.personalizada]
    proprias = [tag for tag in aba.tags if tag.personalizada]
    resultado.compostas = sum(1 for tag in proprias if tag.composta)
    resultado.personalizadas = len(proprias)
    resultado.auxiliares = len(aba.auxiliares)

    # --- passo 1: coluna de data (identica nos dois modos) ---------------
    for deslocamento, carimbo in enumerate(carimbos):
        numero = primeira + deslocamento
        linha = linhas.get(numero)
        if linha is None:
            linha = {"abertura": modelo_abertura, "celulas": {}}
            linhas[numero] = linha

        linha["celulas"][indice_data] = _celula_data(
            "%s%d" % (cfg.coluna_data, numero),
            estilo_data,
            formatar_serial(para_serial_excel(carimbo)),
        )

        if deslocamento >= total_formulas:
            # Ultimo carimbo: so existe como fim do intervalo anterior.
            for tag in aba.tags:
                linha["celulas"].pop(tag.indice, None)

    # --- passo 2: colunas com TAG ----------------------------------------
    if modo_bloco:
        # UMA formula por coluna. Duas formas de representar o resultado:
        #
        # dinamica (preferida): a formula mora so em B3 e o resultado "derrama"
        #   para baixo sozinho. O bloco inteiro nasce e some junto com B3, o que
        #   evita o aviso "Nao e possivel alterar parte de uma matriz" no uso
        #   normal. Exige o marcador cm="1", que depende de xl/metadata.xml.
        #
        # CSE (reserva): matriz multicelula explicita em B3:B290. Funciona em
        #   qualquer versao do Excel, mas trava o bloco inteiro para edicao.
        dinamica = bool(cm_matriz) and cfg.matriz_dinamica
        resultado.matriz = "dinamica" if dinamica else "cse"

        for tag in comuns:
            referencia = "%s%d" % (tag.coluna, primeira)
            if dinamica or ultima_linha_formula <= primeira:
                # Ancora unica: o proprio Excel expande o resultado.
                ref_matriz = referencia
            else:
                ref_matriz = "%s%d:%s%d" % (
                    tag.coluna, primeira, tag.coluna, ultima_linha_formula
                )
            formula_xml = montar_formula_bloco_xml(
                tag.coluna,
                cfg.linha_tags,
                primeira,
                primeira + total_carimbos - 1,
                cfg.coluna_data,
                cfg.funcao_pi_bloco,
                intervalo_minutos,
                cfg.argumentos_fixos,
            )
            linhas[primeira]["celulas"][tag.indice] = _celula_formula(
                referencia,
                tag.estilo_dados,
                cm_matriz if dinamica else None,
                formula_xml,
                True,
                ref_matriz=ref_matriz,
            )
            # As linhas de baixo ficam vazias com o estilo preservado. Celula
            # apenas formatada conta como vazia, entao nao provoca #SPILL!.
            for deslocamento in range(1, total_formulas):
                numero = primeira + deslocamento
                linhas[numero]["celulas"][tag.indice] = _celula_vazia(
                    "%s%d" % (tag.coluna, numero), tag.estilo_dados
                )
            resultado.formulas_escritas += 1
            resultado.consultas_pi += 1
            if len(resultado.amostra) < 6:
                resultado.amostra.append((ref_matriz, tag.texto, formula_xml))
    else:
        # Uma formula PIAdvCalcVal por linha (comportamento original). Cada
        # formula cobre uma unica celula, entao nao existe bloco travado; o
        # cm= so serve para o Excel trata-la como matriz dinamica de 1 celula.
        dinamica = bool(cm_matriz) and cfg.matriz_dinamica
        resultado.matriz = "dinamica" if cm_matriz else "cse"
        for deslocamento in range(total_formulas):
            numero = primeira + deslocamento
            linha = linhas[numero]
            for tag in comuns:
                referencia = "%s%d" % (tag.coluna, numero)
                formula_xml = montar_formula_xml(
                    tag.coluna,
                    cfg.linha_tags,
                    numero,
                    cfg.coluna_data,
                    cfg.funcao_pi,
                    cfg.argumentos_fixos,
                )
                cache = _cache_preservado(
                    linha["celulas"].get(tag.indice), formula_xml, cfg
                )
                linha["celulas"][tag.indice] = _celula_formula(
                    referencia,
                    tag.estilo_dados,
                    cm_matriz,
                    formula_xml,
                    cfg.formula_matriz,
                    cache,
                )
                resultado.formulas_escritas += 1
                resultado.consultas_pi += 1
                if len(resultado.amostra) < 6:
                    resultado.amostra.append((referencia, tag.texto, formula_xml))

    # --- passo 3: colunas com conta propria ------------------------------
    for tag in proprias:
        _escrever_coluna_propria(
            linhas, tag, primeira, total_formulas, total_carimbos,
            cfg, cm_matriz, dinamica, intervalo_minutos, resultado,
        )

    # Linhas abaixo do novo intervalo: limpa apenas as colunas gerenciadas.
    for numero in list(linhas):
        if numero <= ultima_linha_gerada or numero < primeira:
            continue
        for coluna in colunas_gerenciadas:
            linhas[numero]["celulas"].pop(coluna, None)
        if not linhas[numero]["celulas"]:
            del linhas[numero]

    if simular:
        return resultado

    novo_miolo = _serializar(linhas)
    novo_cabecalho = _atualizar_dimensao(cabecalho, linhas)
    pacote.definir_texto(aba.parte, novo_cabecalho + novo_miolo + rodape)
    return resultado


def usar_bloco_na_coluna_propria(analise, cfg) -> bool:
    """Decide se a conta propria de uma coluna vai em matriz ou linha a linha.

    No modo "auto" (padrao) ela acompanha o resto do arquivo: com o modo bloco
    ligado, a conta tambem vai em matriz, trocando cada PIAdvCalc...Val pela
    irma PIAdvCalc...Dat. Somar (ou passar por IFERROR, N, AVERAGE...) duas
    matrizes do mesmo tamanho devolve outra matriz, entao a conta continua
    valendo - e a coluna faz uma consulta por TAG no lugar de uma por linha.

    So cai para linha a linha quando alguma chamada nao tem versao em matriz
    conhecida: e melhor consultar mais vezes do que inventar a assinatura.
    """
    if not analise.cabe_em_bloco:
        return False
    if cfg.modo_composta == MODO_BLOCO:
        return True
    if cfg.modo_composta == MODO_LINHA:
        return False
    return cfg.modo_formula == MODO_BLOCO


def _escrever_coluna_propria(linhas: dict, tag, primeira: int, total_formulas: int,
                             total_carimbos: int, cfg, cm_matriz, dinamica: bool,
                             intervalo_minutos: int, resultado: ResultadoAba) -> None:
    """Regrava uma coluna reaproveitando a conta que ja estava na planilha.

    O unico trecho alterado sao as referencias de inicio e fim de cada chamada
    do PI. As TAGs continuam vindo das mesmas celulas da linha de TAGs, entao
    trocar a TAG la continua mudando o resultado.
    """
    analise = tag.personalizada
    if not analise or not total_formulas:
        return

    if usar_bloco_na_coluna_propria(analise, cfg):
        referencia = "%s%d" % (tag.coluna, primeira)
        ultima_linha_formula = primeira + total_formulas - 1
        if dinamica or ultima_linha_formula <= primeira:
            ref_matriz = referencia
        else:
            ref_matriz = "%s%d:%s%d" % (
                tag.coluna, primeira, tag.coluna, ultima_linha_formula
            )
        formula_xml = xu.escapar_xml(
            composta.montar_bloco(
                analise, primeira, primeira + total_carimbos - 1,
                cfg.coluna_data, intervalo_pi(intervalo_minutos),
            )
        )
        linhas[primeira]["celulas"][tag.indice] = _celula_formula(
            referencia,
            tag.estilo_dados,
            cm_matriz if dinamica else None,
            formula_xml,
            True,
            ref_matriz=ref_matriz,
        )
        for deslocamento in range(1, total_formulas):
            numero = primeira + deslocamento
            linhas[numero]["celulas"][tag.indice] = _celula_vazia(
                "%s%d" % (tag.coluna, numero), tag.estilo_dados
            )
        resultado.formulas_escritas += 1
        resultado.consultas_pi += analise.total_chamadas
        if len(resultado.amostra) < 6:
            resultado.amostra.append((ref_matriz, tag.descricao(), formula_xml))
        return

    for deslocamento in range(total_formulas):
        numero = primeira + deslocamento
        referencia = "%s%d" % (tag.coluna, numero)
        formula_xml = xu.escapar_xml(
            composta.montar_linha(analise, numero, cfg.coluna_data)
        )
        linhas[numero]["celulas"][tag.indice] = _celula_formula(
            referencia,
            tag.estilo_dados,
            cm_matriz,
            formula_xml,
            cfg.formula_matriz,
        )
        resultado.formulas_escritas += 1
        resultado.consultas_pi += analise.total_chamadas
        if len(resultado.amostra) < 6 and deslocamento == 0:
            resultado.amostra.append((referencia, tag.descricao(), formula_xml))


def _cache_preservado(celula_anterior, formula_xml: str, cfg) -> str:
    """Reaproveita o valor em cache quando o usuario pede para nao limpar."""
    if cfg.limpar_valores_cache or not celula_anterior:
        return ""
    if formula_xml not in celula_anterior:
        return ""
    cabecalho = celula_anterior[: celula_anterior.find(">") + 1]
    tipo = xu.RE_ATRIBUTO_T.search(cabecalho)
    return ' t="%s"' % tipo.group(1) if tipo else ""


# ---------------------------------------------------------------------------
# Ajustes no nivel da pasta de trabalho
# ---------------------------------------------------------------------------

def configurar_calculo(pacote: PacoteXlsx, cfg) -> None:
    """Define como o Excel calcula ao abrir o arquivo.

    calculo_manual=True (padrao): o arquivo abre instantaneamente e nada e
    consultado no PI ate o usuario apertar F9 ou usar PI DataLink > Refresh.
    Sem isso, abrir a planilha dispara todas as consultas de uma vez, que e
    justamente a causa da demora.

    calculo_manual=False: marca fullCalcOnLoad para o Excel calcular sozinho.
    """
    nome = "xl/workbook.xml"
    xml = pacote.texto(nome)
    achado = re.search(r"<calcPr\b[^>]*/>", xml)

    if cfg.calculo_manual:
        atributos = ' calcMode="manual" calcOnSave="0"'
        if achado:
            trecho = achado.group(0)
            # Remove marcas conflitantes de recalculo automatico.
            for atributo in ("calcMode", "fullCalcOnLoad", "calcOnSave", "forceFullCalc"):
                trecho = re.sub(r'\s%s="[^"]*"' % atributo, "", trecho)
            novo_trecho = trecho[:-2] + atributos + "/>"
        else:
            novo_trecho = '<calcPr calcId="191029"%s/>' % atributos
    else:
        if achado:
            trecho = achado.group(0)
            trecho = re.sub(r'\scalcMode="[^"]*"', "", trecho)
            if "fullCalcOnLoad=" in trecho:
                novo_trecho = re.sub(
                    r'\bfullCalcOnLoad="[^"]*"', 'fullCalcOnLoad="1"', trecho
                )
            else:
                novo_trecho = trecho[:-2] + ' fullCalcOnLoad="1"/>'
        else:
            novo_trecho = '<calcPr calcId="191029" fullCalcOnLoad="1"/>'

    if achado:
        pacote.definir_texto(
            nome, xml[: achado.start()] + novo_trecho + xml[achado.end():]
        )
        return

    fechamento = xml.rfind("</workbook>")
    if fechamento != -1:
        pacote.definir_texto(nome, xml[:fechamento] + novo_trecho + xml[fechamento:])


def preparar_estilo_data(pacote: PacoteXlsx, abas: list, cfg):
    """Garante um estilo de data valido quando ele ainda nao existe.

    Devolve o indice do <xf> a ser usado, ou None quando o estilo original de
    cada aba deve ser preservado (caso normal, que nao toca em styles.xml).
    """
    if cfg.forcar_formato_data:
        return styles.garantir_estilo_data(pacote, cfg.formato_data)

    precisa = any(
        not styles.formato_ja_e_data(pacote, aba.estilo_data) for aba in abas
    )
    if precisa:
        return styles.garantir_estilo_data(pacote, cfg.formato_data)
    return None


def processar_pacote(pacote: PacoteXlsx, abas: list, carimbos, cfg,
                     simular: bool = False) -> list:
    """Aplica a reescrita em todas as abas alvo de um pacote ja aberto.

    "carimbos" pode ser:
      * um dicionario {nome_da_aba: lista_de_carimbos} - cada aba com o seu
        proprio periodo/intervalo (caso normal);
      * uma lista unica - a mesma sequencia para todas as abas.
    """
    if not abas:
        return []

    def carimbos_da(aba) -> list:
        if isinstance(carimbos, dict):
            return carimbos.get(aba.nome, [])
        return carimbos

    estilo_novo = preparar_estilo_data(pacote, abas, cfg)

    # Garante o suporte a matriz dinamica ANTES de escrever qualquer formula.
    # Sem isso, arquivos sem xl/metadata.xml saiam com matriz CSE e o Excel
    # recusava qualquer edicao com "Nao e possivel alterar parte de uma matriz".
    cm_dinamico = None
    if cfg.matriz_dinamica:
        cm_dinamico = metadata.garantir_cm_dinamico(pacote, aplicar=not simular)

    resultados = []
    for aba in abas:
        estilo = estilo_novo
        if estilo is not None and not cfg.forcar_formato_data:
            # So usa o estilo criado nas abas que realmente nao tinham um.
            if styles.formato_ja_e_data(pacote, aba.estilo_data):
                estilo = None
        resultados.append(
            reescrever_aba(pacote, aba, carimbos_da(aba), cfg, estilo,
                           simular=simular, cm_dinamico=cm_dinamico)
        )

    if not simular:
        configurar_calculo(pacote, cfg)
        pacote.remover_parte(PARTE_CALCCHAIN)
    return resultados
