# -*- coding: utf-8 -*-
"""Orquestracao: pasta -> arquivos -> abas.

Regras de seguranca implementadas aqui:
  * os arquivos originais NUNCA sao abertos para escrita;
  * arquivos temporarios do Excel (~$nome.xlsx) sao ignorados;
  * extensoes nao suportadas sao ignoradas e registradas;
  * um erro em um arquivo nao interrompe os demais;
  * o processamento pode ser cancelado a qualquer momento;
  * por padrao um arquivo ja existente na saida nao e sobrescrito.
"""

from __future__ import annotations

import os
from dataclasses import dataclass, field
from pathlib import Path

from config import EXTENSOES_SUPORTADAS, MODO_BLOCO
from . import inspector, writer
from .timeline import gerar_carimbos
from .xml_utils import PacoteXlsx


@dataclass
class ResultadoArquivo:
    caminho: Path
    status: str = "OK"              # OK | ERRO | IGNORADO
    mensagem: str = ""
    abas: list = field(default_factory=list)   # lista de ResultadoAba
    avisos: list = field(default_factory=list)

    @property
    def nome(self) -> str:
        return self.caminho.name

    @property
    def total_tags(self) -> int:
        return sum(aba.tags for aba in self.abas)

    @property
    def total_linhas(self) -> int:
        return max((aba.linhas_formula for aba in self.abas), default=0)

    @property
    def total_consultas(self) -> int:
        """Quantas idas ao servidor PI as formulas deste arquivo geram."""
        return sum(aba.consultas_pi for aba in self.abas)

    @property
    def modo(self) -> str:
        return self.abas[0].modo if self.abas else ""

    @property
    def total_compostas(self) -> int:
        """Colunas que combinam mais de uma TAG (ver core/composta.py)."""
        return sum(aba.compostas for aba in self.abas)


@dataclass
class ResumoProcessamento:
    resultados: list = field(default_factory=list)
    cancelado: bool = False

    @property
    def sucesso(self) -> int:
        return sum(1 for r in self.resultados if r.status == "OK")

    @property
    def erros(self) -> int:
        return sum(1 for r in self.resultados if r.status == "ERRO")

    @property
    def ignorados(self) -> int:
        return sum(1 for r in self.resultados if r.status == "IGNORADO")


# ---------------------------------------------------------------------------
# Descoberta de arquivos
# ---------------------------------------------------------------------------

def listar_arquivos(pasta) -> list:
    """Lista os arquivos Excel processaveis de uma pasta (sem recursao)."""
    pasta = Path(pasta)
    if not pasta.is_dir():
        raise NotADirectoryError("Pasta inexistente: %s" % pasta)

    encontrados = []
    for item in sorted(pasta.iterdir(), key=lambda p: p.name.lower()):
        if not item.is_file():
            continue
        if item.name.startswith("~$"):
            continue  # arquivo temporario criado pelo Excel quando aberto
        if item.name.startswith("."):
            continue
        if item.suffix.lower() not in EXTENSOES_SUPORTADAS:
            continue
        encontrados.append(item)
    return encontrados


def listar_ignorados(pasta) -> list:
    """Arquivos que parecem Excel mas nao podem ser processados."""
    pasta = Path(pasta)
    if not pasta.is_dir():
        return []
    ignorados = []
    for item in sorted(pasta.iterdir(), key=lambda p: p.name.lower()):
        if not item.is_file() or item.name.startswith("~$"):
            continue
        sufixo = item.suffix.lower()
        if sufixo in (".xls", ".xlsb"):
            ignorados.append(
                (item, "formato %s nao suportado - converta para .xlsx" % sufixo)
            )
    return ignorados


def abas_disponiveis(arquivos: list, cfg) -> dict:
    """Mapeia nome da aba -> quantidade de arquivos em que ela tem TAGs.

    Usado pela interface para montar a lista de abas selecionaveis.
    """
    contagem = {}
    for caminho in arquivos:
        info = inspector.inspecionar_arquivo(caminho, cfg)
        if info.erro:
            continue
        for aba in info.abas_processaveis:
            contagem[aba.nome] = contagem.get(aba.nome, 0) + 1
    return contagem


def _verificar_bloqueio(caminho: Path) -> None:
    """Levanta erro se o arquivo estiver aberto no Excel ou sem permissao."""
    if (caminho.parent / ("~$" + caminho.name)).exists():
        raise PermissionError("arquivo aberto no Excel")
    try:
        with open(caminho, "rb") as handle:
            handle.read(4)
    except PermissionError:
        raise PermissionError("arquivo bloqueado ou sem permissao de leitura")


def avisos_de_colunas_proprias(abas: list, cfg) -> list:
    """Resume, por aba, as colunas com conta propria e as colunas de apoio.

    Sao as colunas descritas em core/composta.py: a que combina duas TAGs
    (vazao medida por dois canos, por exemplo) e a que so existe para guardar
    a TAG usada nessa conta.
    """
    avisos = []
    for aba in abas:
        proprias = aba.personalizadas

        # Colunas que nao cabem em matriz: a conta usa uma funcao que resume a
        # coluna inteira em um numero so (AVERAGE, SUM, MAX...). Elas continuam
        # linha a linha, e isso pesa nas consultas - por isso vale o aviso.
        travadas = [
            tag for tag in proprias
            if tag.personalizada.funcoes_que_impedem_bloco
            and not writer.usar_bloco_na_coluna_propria(tag.personalizada, cfg)
        ]
        if travadas:
            funcoes = sorted({
                funcao for tag in travadas
                for funcao in tag.personalizada.funcoes_que_impedem_bloco
            })
            avisos.append(
                "%s: coluna(s) %s ficam com uma formula por linha porque a conta "
                "usa %s, que resume a coluna inteira em um numero so"
                % (aba.nome, ", ".join(tag.coluna for tag in travadas),
                   "/".join(funcoes))
            )

        if proprias:
            detalhe = ", ".join(
                "%s (%d TAGs)" % (tag.coluna, tag.personalizada.total_chamadas)
                if tag.composta else tag.coluna
                for tag in proprias
            )
            avisos.append(
                "%s: conta propria da planilha preservada em %s"
                % (aba.nome, detalhe)
            )
        if aba.auxiliares:
            avisos.append(
                "%s: colunas de apoio mantidas intactas (so guardam TAG): %s"
                % (aba.nome, ", ".join(tag.coluna for tag in aba.auxiliares))
            )

        # Colunas marcadas com uma expressao de filtro na linha de TAGs: saem
        # com PITimeFilter, contando o tempo em que a condicao foi verdadeira.
        if aba.filtradas:
            por_linha, em_bloco, _ = cfg.receita_da_coluna(True)
            avisos.append(
                "%s: tempo por filtro (%s, em %s) na(s) coluna(s) %s - a "
                "expressao continua na celula da linha %d e pode ser trocada "
                "direto no Excel"
                % (aba.nome,
                   em_bloco if cfg.modo_formula == MODO_BLOCO else por_linha,
                   cfg.unidade_tempo_filtrado,
                   ", ".join(tag.coluna for tag in aba.filtradas),
                   cfg.linha_tags)
            )

        # Colunas com formula que o programa nao reescreve. Quando nao tem TAG
        # propria elas ficam como estao - inclusive com o numero de linhas
        # antigo, se o periodo mudou. Vale avisar em vez de mexer no escuro.
        geridas = {tag.coluna for tag in aba.tags}
        intactas = [item for item in aba.nao_tratadas if item[0] not in geridas]
        substituidas = [item for item in aba.nao_tratadas if item[0] in geridas]
        if intactas:
            avisos.append(
                "%s: coluna(s) %s tem conta que o programa nao ajusta "
                "(%s) - confira se o numero de linhas ainda bate"
                % (aba.nome, ", ".join(coluna for coluna, _ in intactas),
                   intactas[0][1])
            )
        if substituidas:
            avisos.append(
                "%s: formula anterior trocada pela padrao em %s"
                % (aba.nome, ", ".join(coluna for coluna, _ in substituidas))
            )
    return avisos


def _selecionar_abas(info, cfg) -> list:
    """Filtra as abas processaveis de acordo com a escolha do usuario."""
    processaveis = info.abas_processaveis
    if not cfg.abas_selecionadas:
        return processaveis
    escolhidas = {nome.strip().lower() for nome in cfg.abas_selecionadas}
    return [aba for aba in processaveis if aba.nome.strip().lower() in escolhidas]


# ---------------------------------------------------------------------------
# Processamento
# ---------------------------------------------------------------------------

def carimbos_por_aba(abas: list, cfg) -> dict:
    """Gera a sequencia de datas de CADA aba conforme os ajustes dela.

    Abas sem ajuste proprio usam o periodo geral da execucao. Assim a aba
    "Indicadores Diarios" pode ser gerada de dia em dia enquanto a aba
    "Indicadores por hora" do mesmo arquivo continua de 5 em 5 minutos.
    """
    cache = {}
    resultado = {}
    for aba in abas:
        inicio, fim, intervalo = cfg.periodo_da_aba(aba.nome)
        chave = (inicio, fim, intervalo)
        if chave not in cache:
            cache[chave] = gerar_carimbos(inicio, fim, intervalo, cfg.carimbo_final_extra)
        resultado[aba.nome] = cache[chave]
    return resultado


def processar_arquivo(caminho, destino, cfg, simular: bool = False) -> ResultadoArquivo:
    """Processa um unico arquivo. Nunca altera o original."""
    caminho = Path(caminho)
    resultado = ResultadoArquivo(caminho=caminho)

    try:
        _verificar_bloqueio(caminho)

        destino = Path(destino)
        if not simular:
            if destino.resolve() == caminho.resolve():
                raise ValueError("a saida apontaria para o proprio arquivo original")
            if destino.exists() and not cfg.sobrescrever_saida:
                resultado.status = "IGNORADO"
                resultado.mensagem = "ja existe na pasta de saida (marque 'sobrescrever' para refazer)"
                return resultado

        pacote = PacoteXlsx.abrir(caminho)
        info = inspector.inspecionar_pacote(pacote, cfg)
        alvo = _selecionar_abas(info, cfg)

        if not alvo:
            resultado.status = "IGNORADO"
            if info.abas_processaveis:
                resultado.mensagem = "nenhuma das abas selecionadas existe neste arquivo"
            else:
                resultado.mensagem = (
                    "nenhuma aba com TAG valida na linha %d" % cfg.linha_tags
                )
            return resultado

        ignoradas = [
            aba.nome for aba in info.abas if not aba.processavel and not aba.oculta
        ]
        if ignoradas:
            resultado.avisos.append(
                "abas sem TAG mantidas intactas: %s" % ", ".join(ignoradas)
            )

        resultado.avisos.extend(avisos_de_colunas_proprias(alvo, cfg))

        carimbos = carimbos_por_aba(alvo, cfg)
        resultado.abas = writer.processar_pacote(pacote, alvo, carimbos, cfg, simular)

        if not simular:
            pacote.salvar(destino)

    except PermissionError as exc:
        resultado.status = "ERRO"
        resultado.mensagem = str(exc)
    except Exception as exc:
        resultado.status = "ERRO"
        resultado.mensagem = "%s: %s" % (type(exc).__name__, exc)

    return resultado


def processar_pasta(cfg, logger, progresso=None, cancelado=None) -> ResumoProcessamento:
    """Processa todos os arquivos da pasta de entrada.

    progresso : callable(indice, total, nome_do_arquivo)
    cancelado : callable() -> bool, consultado antes de cada arquivo
    """
    resumo = ResumoProcessamento()

    arquivos = listar_arquivos(cfg.pasta_entrada)
    logger.inicio(cfg, len(arquivos))

    for caminho, motivo in listar_ignorados(cfg.pasta_entrada):
        item = ResultadoArquivo(caminho=caminho, status="IGNORADO", mensagem=motivo)
        resumo.resultados.append(item)
        logger.arquivo_ignorado(item.nome, motivo)

    if not arquivos:
        logger.escrever("Nenhum arquivo .xlsx ou .xlsm encontrado na pasta de entrada.")
        logger.linha_em_branco()
        logger.fim(0, 0, resumo.ignorados)
        return resumo

    # Os carimbos sao gerados por aba, dentro de processar_arquivo, porque cada
    # aba pode ter periodo e intervalo proprios.
    pasta_saida = Path(cfg.pasta_saida)
    pasta_saida.mkdir(parents=True, exist_ok=True)

    total = len(arquivos)
    for indice, caminho in enumerate(arquivos, start=1):
        if cancelado and cancelado():
            resumo.cancelado = True
            break
        if progresso:
            progresso(indice, total, caminho.name)

        resultado = processar_arquivo(caminho, pasta_saida / caminho.name, cfg)
        resumo.resultados.append(resultado)

        if resultado.status == "OK":
            logger.arquivo_ok(
                resultado.nome,
                [aba.nome for aba in resultado.abas],
                resultado.total_tags,
                resultado.total_linhas,
                resultado.avisos,
                resultado.total_consultas,
                resultado.modo,
            )
        elif resultado.status == "IGNORADO":
            logger.arquivo_ignorado(resultado.nome, resultado.mensagem)
        else:
            logger.arquivo_erro(resultado.nome, resultado.mensagem)

    logger.fim(resumo.sucesso, resumo.erros, resumo.ignorados, resumo.cancelado)
    return resumo


def estimar_consultas(arquivos: list, cfg) -> tuple:
    """Estima as consultas ao PI do arquivo mais pesado da pasta.

    Devolve (consultas_modo_bloco, consultas_modo_linha). Respeita o periodo
    proprio de cada aba, entao uma aba diaria pesa muito menos que uma de
    5 em 5 minutos.
    """
    pior_bloco = 0
    pior_linha = 0
    for caminho in arquivos:
        info = inspector.inspecionar_arquivo(caminho, cfg)
        if info.erro:
            continue
        alvo = _selecionar_abas(info, cfg)
        if not alvo:
            continue
        try:
            carimbos = carimbos_por_aba(alvo, cfg)
        except ValueError:
            continue

        bloco = linha = 0
        for aba in alvo:
            sequencia = carimbos.get(aba.nome, [])
            formulas = max(0, len(sequencia) - (1 if cfg.carimbo_final_extra else 0))
            for tag in aba.tags:
                # Uma coluna com conta propria consulta uma TAG por chamada; e,
                # quando a conta nao cabe em matriz, ela continua linha a linha
                # mesmo com o resto do arquivo em modo bloco.
                if tag.personalizada and not writer.usar_bloco_na_coluna_propria(
                    tag.personalizada, cfg
                ):
                    bloco += tag.consultas * formulas
                else:
                    bloco += tag.consultas
                linha += tag.consultas * formulas
        pior_bloco = max(pior_bloco, bloco)
        pior_linha = max(pior_linha, linha)
    return pior_bloco, pior_linha
