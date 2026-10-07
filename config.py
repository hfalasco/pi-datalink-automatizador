# -*- coding: utf-8 -*-
"""Configuracao central do Automatizador PI DataLink.

Todos os parametros estruturais (linha das TAGs, coluna de data, primeira linha
de dados, formato da formula) ficam concentrados aqui para que a estrutura das
planilhas possa ser ajustada sem tocar no restante do codigo.

Os valores padrao foram extraidos da analise real dos templates
"Template Brix do clarificado fabrica.xlsx", "Template Caldo Misto
Destilaria.xlsx" e "Template Vazao Clarificado Fabrica.xlsx".
"""

from __future__ import annotations

import re
import unicodedata
from dataclasses import dataclass, field
from datetime import date
from pathlib import Path


def normalizar_nome(texto: str) -> str:
    """Minusculas, sem acento e sem espacos - para comparar nomes de abas."""
    base = unicodedata.normalize("NFKD", texto or "")
    base = "".join(c for c in base if not unicodedata.combining(c))
    return base.lower().strip().replace(" ", "")

APP_NOME = "AUTOMATIZADOR PI DATALINK"
APP_VERSAO = "1.0.0"

# Extensoes que o motor de cirurgia XML sabe processar.
# .xls (formato binario antigo) NAO e suportado - precisa ser convertido antes.
EXTENSOES_SUPORTADAS = (".xlsx", ".xlsm")

# ---------------------------------------------------------------------------
# Estrutura padrao da planilha (confirmada nos tres templates analisados)
# ---------------------------------------------------------------------------
LINHA_TAGS_PADRAO = 2           # linha onde ficam as TAGs do PI
PRIMEIRA_LINHA_DADOS_PADRAO = 3  # primeira linha de data/formula
COLUNA_DATA_PADRAO = "A"        # coluna que recebe os carimbos de tempo
MAX_COLUNAS_VARREDURA = 200     # limite de seguranca ao varrer a linha de TAGs

# ---------------------------------------------------------------------------
# Formula do PI DataLink
# ---------------------------------------------------------------------------
# IMPORTANTE: dentro do arquivo .xlsx o separador de argumentos e SEMPRE virgula,
# independentemente do idioma do Excel. O Excel exibira ";" em maquinas pt-BR.
FUNCAO_PI_PADRAO = "PIAdvCalcVal"

# Funcao irmã que devolve TODOS os intervalos de uma vez. A assinatura, extraida
# do proprio add-in (C:\Program Files\PIPC\Excel\pipc64.xll), difere da de
# PIAdvCalcVal apenas pelo argumento "interval" na 4a posicao:
#
#   PIAdvCalcVal(data item, stime, etime,           modestring, weightmode, ...)
#   PIAdvCalcDat(data item, stime, etime, interval, modestring, weightmode, ...)
#
# Uma formula por coluna no lugar de uma por linha reduz as consultas ao
# servidor PI na mesma proporcao do numero de linhas.
FUNCAO_PI_BLOCO_PADRAO = "PIAdvCalcDat"

# Modos de gravacao das formulas
MODO_BLOCO = "bloco"   # 1 formula de matriz por coluna  (PIAdvCalcDat) - rapido
MODO_LINHA = "linha"   # 1 formula por linha por coluna  (PIAdvCalcVal) - original
MODOS_VALIDOS = (MODO_BLOCO, MODO_LINHA)

# Colunas com conta propria, inclusive as que combinam mais de uma TAG
# (ver core/composta.py):
#   auto   = acompanha o modo geral; com o modo bloco ligado, a conta tambem vai
#            em matriz (PIAdvCalcVal -> PIAdvCalcDat, PIAdvCalcFilVal ->
#            PIAdvCalcFilDat, e assim por diante)
#   linha  = sempre uma formula por linha
#   bloco  = sempre em matriz, mesmo com o resto do arquivo em modo linha
COMPOSTA_AUTO = "auto"
MODOS_COMPOSTA = (COMPOSTA_AUTO, MODO_LINHA, MODO_BLOCO)
ARGUMENTOS_FIXOS_PADRAO = (
    '"average (time-weighted)"',
    '"time-weighted"',
    "0",
    "1",
    "0",
    '""',
)

# ---------------------------------------------------------------------------
# Tempo por filtro (PITimeFilter)
# ---------------------------------------------------------------------------
# A familia PIAdvCalc responde "qual foi a media da TAG no intervalo". A familia
# PITimeFilter responde outra pergunta: "quanto TEMPO, dentro do intervalo, a
# condicao foi verdadeira" - quantas horas a valvula ficou aberta. Por isso o
# 1o argumento nao e uma TAG, e sim a expressao de filtro, que a planilha guarda
# na celula da linha de TAGs:
#
#   'USM-XV_411010'="Aberta" AND 'USM-XV_411012'="Aberta"
#
# Assinaturas (a PITimeFilter confere com a formula que o proprio Excel grava;
# a PITimeFilterVal segue o mesmo padrao, sem o "interval" na 4a posicao):
#
#   PITimeFilterVal   exp, stime, etime,           timeunites, outcode, root path
#   PITimeFilter      exp, stime, etime, interval, timeunites, outcode, root path
#
# O nome e invertido em relacao ao resto: quem devolve UM valor leva o sufixo
# "Val"; PITimeFilter (sem sufixo) e a que devolve todos os intervalos.
FUNCAO_TEMPO_FILTRADO_PADRAO = "PITimeFilterVal"
FUNCAO_TEMPO_FILTRADO_BLOCO_PADRAO = "PITimeFilter"

# Em que unidade o tempo volta. "hours" num intervalo de 1h devolve algo entre
# 0 e 1; "minutes" no mesmo intervalo, algo entre 0 e 60.
UNIDADES_TEMPO_FILTRADO = ("seconds", "minutes", "hours", "days")
UNIDADE_TEMPO_FILTRADO_PADRAO = "hours"

# Formato de data/hora aplicado quando a coluna de data ainda nao possui estilo.
FORMATO_DATA_PADRAO = "dd/mm/yyyy hh:mm"

# Datas anteriores a esta sao consideradas suspeitas (evita erro de digitacao).
ANO_MINIMO = 1990
ANO_MAXIMO = 2100


MINUTOS_POR_DIA = 1440


def validar_periodo(data_inicial: date, data_final: date, intervalo: int) -> list:
    """Checa um par de datas e um intervalo. Devolve a lista de problemas."""
    erros = []

    if data_final < data_inicial:
        erros.append("a data final e anterior a data inicial.")
    for rotulo, valor in (("inicial", data_inicial), ("final", data_final)):
        if not (ANO_MINIMO <= valor.year <= ANO_MAXIMO):
            erros.append("data %s fora da faixa aceitavel: %s" % (rotulo, valor))

    if intervalo is None or intervalo <= 0:
        erros.append("o intervalo precisa ser maior que zero.")
    elif intervalo < MINUTOS_POR_DIA and MINUTOS_POR_DIA % intervalo != 0:
        erros.append(
            "o intervalo de %d min nao divide o dia de forma exata "
            "(use 1, 5, 10, 15, 30, 60, 720 ou 1440)." % intervalo
        )
    elif intervalo > MINUTOS_POR_DIA and intervalo % MINUTOS_POR_DIA != 0:
        erros.append(
            "intervalos maiores que um dia precisam ser multiplos de 1440 min "
            "(1440 = 1 dia, 2880 = 2 dias...)."
        )
    return erros


@dataclass
class ConfigAba:
    """Periodo e intervalo proprios de UMA aba.

    Serve para tratar abas diferentes do mesmo arquivo de formas diferentes -
    por exemplo, "Indicadores por hora" de 5 em 5 minutos e "Indicadores
    Diarios" uma vez por dia.

    Todo campo deixado como None herda o valor geral da execucao, portanto
    basta preencher o que muda naquela aba.
    """

    data_inicial: date | None = None
    data_final: date | None = None
    intervalo_minutos: int | None = None

    def combinar(self, cfg: "Config") -> tuple:
        """Resolve os valores finais desta aba (proprios ou herdados)."""
        return (
            self.data_inicial or cfg.data_inicial,
            self.data_final or cfg.data_final,
            self.intervalo_minutos or cfg.intervalo_minutos,
        )

    @property
    def personalizada(self) -> bool:
        return any(
            valor is not None
            for valor in (self.data_inicial, self.data_final, self.intervalo_minutos)
        )

    def resumo(self) -> str:
        partes = []
        if self.data_inicial:
            partes.append("inicio " + self.data_inicial.strftime("%d/%m/%Y"))
        if self.data_final:
            partes.append("fim " + self.data_final.strftime("%d/%m/%Y"))
        if self.intervalo_minutos:
            partes.append("%d min" % self.intervalo_minutos)
        return ", ".join(partes) if partes else "usa a configuracao geral"


@dataclass
class Config:
    """Parametros de uma execucao completa."""

    pasta_entrada: Path
    pasta_saida: Path
    data_inicial: date
    data_final: date
    intervalo_minutos: int = 5

    # None ou lista vazia = processar todas as abas que tiverem TAGs validas.
    abas_selecionadas: list[str] | None = None

    # Ajustes por aba: {nome_da_aba: ConfigAba}. A busca ignora caixa e acento,
    # entao "Indicadores Diarios" encontra a aba "Indicadores Diários".
    configuracoes_por_aba: dict = field(default_factory=dict)

    # --- estrutura da planilha -------------------------------------------
    linha_tags: int = LINHA_TAGS_PADRAO
    primeira_linha_dados: int = PRIMEIRA_LINHA_DADOS_PADRAO
    coluna_data: str = COLUNA_DATA_PADRAO
    max_colunas_varredura: int = MAX_COLUNAS_VARREDURA

    # --- formula ----------------------------------------------------------
    funcao_pi: str = FUNCAO_PI_PADRAO
    funcao_pi_bloco: str = FUNCAO_PI_BLOCO_PADRAO
    argumentos_fixos: tuple[str, ...] = ARGUMENTOS_FIXOS_PADRAO
    formula_matriz: bool = True   # grava como formula de matriz (igual ao DataLink)

    # True = chave unica: TODAS as colunas que o programa grava saem com
    # PITimeFilter no lugar de PIAdvCalcVal/Dat. Nenhuma coluna e analisada para
    # decidir - a planilha inteira muda junto, e a linha das TAGs precisa trazer
    # a expressao de filtro no lugar da TAG. Contas proprias preservadas
    # (preservar_formulas_personalizadas) continuam como estao.
    tempo_filtrado: bool = False
    funcao_tempo_filtrado: str = FUNCAO_TEMPO_FILTRADO_PADRAO
    funcao_tempo_filtrado_bloco: str = FUNCAO_TEMPO_FILTRADO_BLOCO_PADRAO
    unidade_tempo_filtrado: str = UNIDADE_TEMPO_FILTRADO_PADRAO

    # MODO_BLOCO = uma formula PIAdvCalcDat por coluna (muito menos consultas);
    # MODO_LINHA = uma formula PIAdvCalcVal por linha (comportamento original).
    modo_formula: str = MODO_BLOCO

    # True = arquivo abre sem calcular nada; os dados vem com F9 / DataLink Refresh.
    calculo_manual: bool = True

    # No modo bloco, prefere matriz DINAMICA (a formula mora so na primeira
    # celula e o resultado se expande sozinho). E o que evita o aviso
    # "Nao e possivel alterar parte de uma matriz" ao mexer nas celulas.
    # Se a planilha nao tiver suporte (sem xl/metadata.xml), cai para a matriz
    # multicelula classica automaticamente.
    matriz_dinamica: bool = True

    # --- colunas com conta propria ----------------------------------------
    # True = quando a coluna ja traz uma formula diferente da padrao (soma de
    # duas vazoes medidas por canos diferentes, IFERROR, fator de conversao...),
    # o programa reaproveita a conta e troca apenas as datas. Ver core/composta.py.
    preservar_formulas_personalizadas: bool = True
    # True = colunas que so guardam a TAG usada por essas contas ficam intactas
    # (nao viram coluna de dados nem gastam consulta ao PI).
    ignorar_colunas_auxiliares: bool = True
    # Como gravar essas colunas: auto, linha ou bloco.
    modo_composta: str = COMPOSTA_AUTO

    # --- deteccao de TAGs -------------------------------------------------
    # Regex opcional. Vazio = aceita qualquer texto nao numerico.
    # Para aceitar somente caminhos PI use:  r"^\\\\"
    padrao_tag: str = ""
    # Textos que nunca sao considerados TAG (comparacao sem acento/caixa).
    rotulos_ignorados: tuple[str, ...] = ("data", "date", "hora", "datahora", "timestamp")

    # --- datas ------------------------------------------------------------
    # Grava um carimbo extra no fim para servir de "fim" do ultimo intervalo.
    carimbo_final_extra: bool = True
    # True = sempre aplica FORMATO_DATA; False = reaproveita o estilo existente.
    forcar_formato_data: bool = False
    formato_data: str = FORMATO_DATA_PADRAO

    # --- arquivos ---------------------------------------------------------
    sobrescrever_saida: bool = False
    limpar_valores_cache: bool = True  # nao grava <v>, forca o Excel a recalcular

    # --- log --------------------------------------------------------------
    pasta_log: Path | None = None

    # ------------------------------------------------------------------
    @property
    def receita_da_formula(self) -> tuple:
        """(funcao por linha, funcao em matriz, argumentos fixos) das colunas.

        E o unico ponto em que o programa escolhe a familia da formula. Com a
        chave "tempo por filtro" desligada sai PIAdvCalcVal/Dat; ligada, sai
        PITimeFilterVal/PITimeFilter - para todas as colunas, sem excecao.
        """
        if self.tempo_filtrado:
            return (
                self.funcao_tempo_filtrado,
                self.funcao_tempo_filtrado_bloco,
                ('"%s"' % self.unidade_tempo_filtrado, "0", '""'),
            )
        return (self.funcao_pi, self.funcao_pi_bloco, self.argumentos_fixos)

    # ------------------------------------------------------------------
    def config_da_aba(self, nome_aba: str) -> ConfigAba:
        """Devolve os ajustes especificos de uma aba (vazio se nao houver)."""
        if not self.configuracoes_por_aba:
            return ConfigAba()
        procurado = normalizar_nome(nome_aba)
        for nome, ajuste in self.configuracoes_por_aba.items():
            if normalizar_nome(nome) == procurado:
                return ajuste
        return ConfigAba()

    def periodo_da_aba(self, nome_aba: str) -> tuple:
        """(data_inicial, data_final, intervalo) finais para uma aba."""
        return self.config_da_aba(nome_aba).combinar(self)

    # ------------------------------------------------------------------
    def validar(self) -> list[str]:
        """Retorna a lista de problemas encontrados (vazia = tudo certo)."""
        erros: list[str] = []

        if not self.pasta_entrada or not Path(self.pasta_entrada).is_dir():
            erros.append(f"Pasta de entrada inexistente: {self.pasta_entrada}")
        if not self.pasta_saida:
            erros.append("Pasta de saida nao informada.")

        try:
            entrada = Path(self.pasta_entrada).resolve()
            saida = Path(self.pasta_saida).resolve()
            if entrada == saida:
                erros.append(
                    "A pasta de saida nao pode ser a mesma da entrada "
                    "(os arquivos originais nunca sao sobrescritos)."
                )
        except OSError:
            pass

        erros.extend(
            validar_periodo(self.data_inicial, self.data_final, self.intervalo_minutos)
        )

        for nome, ajuste in (self.configuracoes_por_aba or {}).items():
            inicio, fim, intervalo = ajuste.combinar(self)
            for problema in validar_periodo(inicio, fim, intervalo):
                erros.append(f"Aba {nome!r}: {problema}")

        if self.modo_formula not in MODOS_VALIDOS:
            erros.append(
                "Modo de formula invalido: %r (use %s)."
                % (self.modo_formula, " ou ".join(MODOS_VALIDOS))
            )

        if self.modo_composta not in MODOS_COMPOSTA:
            erros.append(
                "Modo das colunas compostas invalido: %r (use %s)."
                % (self.modo_composta, " ou ".join(MODOS_COMPOSTA))
            )

        if self.unidade_tempo_filtrado not in UNIDADES_TEMPO_FILTRADO:
            erros.append(
                "Unidade de tempo invalida: %r (use %s)."
                % (self.unidade_tempo_filtrado, ", ".join(UNIDADES_TEMPO_FILTRADO))
            )

        if self.linha_tags < 1:
            erros.append("A linha das TAGs precisa ser >= 1.")
        if self.primeira_linha_dados <= self.linha_tags:
            erros.append("A primeira linha de dados precisa ficar abaixo da linha das TAGs.")
        if not re.fullmatch(r"[A-Za-z]{1,3}", self.coluna_data or ""):
            erros.append(f"Coluna de data invalida: {self.coluna_data!r}")

        if self.padrao_tag:
            try:
                re.compile(self.padrao_tag)
            except re.error as exc:
                erros.append(f"Expressao regular de TAG invalida: {exc}")

        return erros
