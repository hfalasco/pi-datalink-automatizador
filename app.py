# -*- coding: utf-8 -*-
"""Pagina Streamlit do Automatizador PI DataLink.

Para abrir:

    streamlit run app.py

Como a pagina roda no navegador, nao existe o seletor de pastas do Windows:
o usuario envia os arquivos e recebe de volta um .zip com as copias
processadas e o log da execucao. Os arquivos enviados ficam em uma pasta
temporaria propria de cada sessao - os originais no computador do usuario
nunca sao tocados.

Todo o motor (core/, config.py, logger.py) e o mesmo de antes; esta pagina e
so a casca que substitui a antiga janela Tkinter.

O Streamlit roda este arquivo INTEIRO a cada clique, tecla ou data trocada.
Por isso tudo que e caro fica fora do caminho de um clique comum:

  * as planilhas sao lidas UMA vez por conjunto enviado (resumo_dos_arquivos);
    trocar data, intervalo ou opcao nao reabre arquivo nenhum;
  * a previsao de linhas e feita por conta, sem montar a lista de datas;
  * o registro e redesenhado algumas vezes por segundo, nao a cada linha -
    uma mensagem por linha entupia a conexao e derrubava a pagina;
  * o .zip do resultado e montado em disco e fica em disco; a memoria da
    sessao guarda so o caminho dele.
"""

from __future__ import annotations

import gc
import shutil
import sys
import tempfile
import time
import traceback
import zipfile
from datetime import date, datetime
from pathlib import Path

import streamlit as st

# Esta pasta e a raiz da aplicacao, e os modulos sao importados direto dela.
# Nada depende do NOME da pasta: ao publicar, o servidor clona o repositorio
# numa pasta com o nome do repositorio ("pi-datalink-automatizador", com
# hifens), que nao seria um nome de pacote valido.
_RAIZ = Path(__file__).resolve().parent
if str(_RAIZ) not in sys.path:
    sys.path.insert(0, str(_RAIZ))

import tema  # noqa: E402
from config import (APP_VERSAO, MODO_BLOCO, MODO_LINHA,  # noqa: E402
                    PRIMEIRA_LINHA_DADOS_PADRAO, Config, ConfigAba,
                    normalizar_nome)
from core import composta, inspector, processor, writer  # noqa: E402
from core.formula import (intervalo_pi, montar_formula,  # noqa: E402
                          montar_formula_bloco, para_exibicao)
from core.timeline import (quantidade_carimbos,  # noqa: E402
                           quantidade_formulas, ultimo_carimbo)
from logger import Logger  # noqa: E402

DATA_PADRAO = date(2026, 9, 11)
INTERVALO_PADRAO = 5
# O nome tecnico (APP_NOME, em config.py) e todo em maiusculas e serve ao log.
# Na tela vale a grafia da marca: APP_NOME.title() escreveria "Pi Datalink".
NOME_EXIBIDO = "PI DataLink"
SUBTITULO = ("Preenche as formulas PIAdvCalcVal em lote. "
             "Os arquivos originais nunca sao alterados.")

# --- Limites que existem para a pagina nao cair -----------------------------
# Quantas linhas do registro ficam na tela. O log completo vai inteiro para
# dentro do .zip; o navegador e que nao aguenta desenhar dezenas de milhares.
LINHAS_VISIVEIS_REGISTRO = 400
# Intervalo minimo entre dois redesenhos do registro, em segundos.
INTERVALO_REDESENHO = 0.4
# Acima disto a planilha fica pesada e o servidor pode ficar sem memoria.
LIMITE_LINHAS_AVISO = 100_000
# Limite do proprio Excel: uma planilha nao tem mais linhas do que isto.
LIMITE_LINHAS_EXCEL = 1_048_576
# Pastas de sessoes paradas ha mais tempo que isto sao apagadas.
HORAS_ATE_A_FAXINA = 6


def numero(valor: int) -> str:
    """12345 -> '12.345' (separador de milhar como se escreve aqui)."""
    return "{:,}".format(int(valor)).replace(",", ".")


# ---------------------------------------------------------------------------
# Pasta de trabalho da sessao
# ---------------------------------------------------------------------------

@st.cache_resource(ttl=3600, show_spinner=False)
def faxina_de_sessoes_antigas() -> int:
    """Apaga as pastas temporarias que sobraram de sessoes ja encerradas.

    Quando alguem fecha a aba, o Streamlit nao avisa ninguem e a pasta fica la.
    Num servidor que vive dias no ar isso enche o disco - e disco cheio derruba
    a pagina do mesmo jeito que memoria cheia. O cache_resource faz esta funcao
    rodar uma vez por hora, nao a cada clique.
    """
    limite = time.time() - HORAS_ATE_A_FAXINA * 3600
    apagadas = 0
    for pasta in Path(tempfile.gettempdir()).glob("pi_datalink_*"):
        try:
            if pasta.is_dir() and pasta.stat().st_mtime < limite:
                shutil.rmtree(pasta, ignore_errors=True)
                apagadas += 1
        except OSError:
            continue  # pasta de outra sessao, em uso ou sem permissao
    return apagadas


def pasta_de_trabalho() -> Path:
    """Pasta temporaria propria desta sessao, com entrada/ e saida/ dentro."""
    guardada = st.session_state.get("pasta_trabalho")
    if guardada and Path(guardada).is_dir():
        return Path(guardada)

    # Primeira visita - ou a faxina acima levou a pasta de uma sessao que ficou
    # horas aberta sem uso. Nos dois casos, comeca de novo.
    pasta = Path(tempfile.mkdtemp(prefix="pi_datalink_"))
    st.session_state.pasta_trabalho = str(pasta)
    st.session_state.assinatura = None
    return pasta


def descartar_download() -> None:
    """Apaga o .zip da rodada anterior: ele nao serve mais e ocupa disco."""
    anterior = st.session_state.get("download")
    if anterior:
        try:
            Path(anterior[1]).unlink()
        except OSError:
            pass
    st.session_state.download = None


def gravar_enviados(enviados: list) -> tuple:
    """Grava os uploads em entrada/ e devolve a assinatura do conjunto.

    A gravacao acontece so quando o conjunto muda - o Streamlit reexecuta o
    script inteiro a cada clique, e regravar tudo toda vez seria desperdicio.
    """
    assinatura = tuple((arquivo.name, arquivo.size) for arquivo in enviados)
    entrada = pasta_de_trabalho() / "entrada"

    if st.session_state.get("assinatura") != assinatura:
        shutil.rmtree(entrada, ignore_errors=True)
        entrada.mkdir(parents=True, exist_ok=True)
        for arquivo in enviados:
            # Path(...).name descarta qualquer caminho vindo do navegador.
            (entrada / Path(arquivo.name).name).write_bytes(arquivo.getbuffer())
        st.session_state.assinatura = assinatura
        st.session_state.registro = []
        descartar_download()

    return assinatura


# ---------------------------------------------------------------------------
# Configuracao
# ---------------------------------------------------------------------------

def ler_data(texto: str, rotulo: str):
    """Aceita dd/mm/aaaa, aaaa-mm-dd e dd-mm-aaaa. Vazio devolve None."""
    texto = (texto or "").strip()
    if not texto:
        return None
    for formato in ("%d/%m/%Y", "%Y-%m-%d", "%d-%m-%Y"):
        try:
            return datetime.strptime(texto, formato).date()
        except ValueError:
            continue
    raise ValueError("data %s invalida: %r (use dd/mm/aaaa)" % (rotulo, texto))


def montar_config(entrada, saida, inicio, fim, intervalo, abas, por_aba,
                  rapido=True, calculo_manual=True, matriz_dinamica=True,
                  manter_contas=True, com_log=True) -> Config:
    """Monta o Config da execucao - mesmo conjunto de opcoes da janela antiga."""
    saida = Path(saida)
    return Config(
        pasta_entrada=Path(entrada),
        pasta_saida=saida,
        data_inicial=inicio,
        data_final=fim,
        intervalo_minutos=intervalo,
        abas_selecionadas=abas or None,
        configuracoes_por_aba=por_aba or {},
        modo_formula=MODO_BLOCO if rapido else MODO_LINHA,
        calculo_manual=calculo_manual,
        matriz_dinamica=matriz_dinamica,
        preservar_formulas_personalizadas=manter_contas,
        ignorar_colunas_auxiliares=manter_contas,
        pasta_log=saida / "logs" if com_log else None,
    )


# ---------------------------------------------------------------------------
# Leitura das planilhas - a parte cara, feita UMA vez por conjunto enviado
# ---------------------------------------------------------------------------

@st.cache_data(show_spinner="Lendo as planilhas...", max_entries=2, ttl=3600)
def resumo_dos_arquivos(assinatura: tuple, entrada: str,
                        manter_contas: bool) -> list:
    """Le cada planilha uma vez e guarda so os numeros que a pagina usa.

    Abrir um .xlsx significa descompactar o pacote inteiro na memoria, e a
    pagina precisava disso duas vezes: uma para listar as abas, outra para
    estimar as consultas - e a segunda se repetia a cada data trocada.

    Nada do que e lido aqui depende do periodo: as abas, as TAGs de cada aba e
    o formato das contas proprias sao os mesmos para qualquer data. Entao a
    leitura acontece uma vez e o resto da pagina trabalha em cima deste resumo:

        [{"nome": "Moenda_A.xlsx", "erro": "",
          "abas": {"Indicadores": ((consultas, conta_propria, cabe_em_bloco), ...)}}]
    """
    cfg = montar_config(entrada, Path(entrada).parent / "saida",
                        DATA_PADRAO, DATA_PADRAO, INTERVALO_PADRAO,
                        None, None, manter_contas=manter_contas)
    resumo = []
    for caminho in processor.listar_arquivos(entrada):
        info = inspector.inspecionar_arquivo(caminho, cfg)
        if info.erro:
            resumo.append({"nome": caminho.name, "erro": info.erro, "abas": {}})
            continue
        abas = {}
        for aba in info.abas_processaveis:
            abas[aba.nome] = tuple(
                (tag.consultas,
                 bool(tag.personalizada),
                 bool(tag.personalizada and tag.personalizada.cabe_em_bloco))
                for tag in aba.tags
            )
        resumo.append({"nome": caminho.name, "erro": "", "abas": abas})
    return resumo


def abas_com_tag(resumo: list) -> dict:
    """{nome da aba: em quantos arquivos ela tem TAG}."""
    contagem = {}
    for arquivo in resumo:
        for nome in arquivo["abas"]:
            contagem[nome] = contagem.get(nome, 0) + 1
    return contagem


def formulas_de_cada_aba(nomes, inicio, fim, intervalo, por_aba) -> dict:
    """Quantas formulas cada aba gera, ja com o ajuste proprio dela.

    So conta - nao monta a lista de datas (ver core/timeline.quantidade_carimbos).
    """
    ajustes = {normalizar_nome(nome): ajuste
               for nome, ajuste in (por_aba or {}).items()}
    quantidades = {}
    for nome in nomes:
        ajuste = ajustes.get(normalizar_nome(nome))
        comeco = (ajuste.data_inicial if ajuste and ajuste.data_inicial else inicio)
        termino = (ajuste.data_final if ajuste and ajuste.data_final else fim)
        passo = (ajuste.intervalo_minutos
                 if ajuste and ajuste.intervalo_minutos else intervalo)
        try:
            quantidades[nome] = quantidade_formulas(
                quantidade_carimbos(comeco, termino, passo, True), True)
        except ValueError:
            quantidades[nome] = 0
    return quantidades


def estimar_consultas(resumo: list, escolhidas, formulas: dict,
                      rapido: bool) -> tuple:
    """(consultas no modo bloco, consultas no modo linha) do arquivo mais pesado.

    E a mesma conta de core/processor.estimar_consultas, so que em cima do
    resumo ja lido: puro numero, sem tocar em disco.
    """
    selecao = {nome.strip().lower() for nome in escolhidas or []}
    pior_bloco = pior_linha = 0
    for arquivo in resumo:
        bloco = linha = 0
        for nome, tags in arquivo["abas"].items():
            if selecao and nome.strip().lower() not in selecao:
                continue
            total = formulas.get(nome, 0)
            for consultas, conta_propria, cabe_em_bloco in tags:
                # Uma coluna com conta propria consulta uma TAG por chamada; e,
                # quando a conta nao cabe em matriz, ela continua linha a linha
                # mesmo com o resto do arquivo em modo bloco.
                em_bloco = (cabe_em_bloco and rapido) if conta_propria else True
                bloco += consultas if em_bloco else consultas * total
                linha += consultas * total
        pior_bloco = max(pior_bloco, bloco)
        pior_linha = max(pior_linha, linha)
    return pior_bloco, pior_linha


# ---------------------------------------------------------------------------
# Registro na tela
# ---------------------------------------------------------------------------

class Registro:
    """Junta as linhas do log e redesenha a caixa de vez em quando.

    Antes cada linha do log redesenhava a caixa inteira, o que significava uma
    mensagem pelo websocket por linha, carregando todo o texto acumulado. Num
    lote grande era isso que travava a aba do navegador e derrubava a conexao.
    Agora a tela e atualizada no maximo a cada INTERVALO_REDESENHO segundos.
    """

    def __init__(self, caixa):
        self.caixa = caixa
        self.linhas = []
        self._ultimo_desenho = 0.0

    def __call__(self, linha: str) -> None:
        """Chamado pelo Logger a cada linha escrita."""
        self.linhas.append(linha)
        agora = time.monotonic()
        if agora - self._ultimo_desenho >= INTERVALO_REDESENHO:
            self.desenhar()

    def desenhar(self) -> None:
        self._ultimo_desenho = time.monotonic()
        self.caixa.code(texto_do_registro(self.linhas), language="text")


def texto_do_registro(linhas: list) -> str:
    """Mostra so o fim do registro - o log completo vai inteiro no .zip."""
    if not linhas:
        return "O registro desta execucao aparece aqui."
    if len(linhas) <= LINHAS_VISIVEIS_REGISTRO:
        return "\n".join(linhas)
    escondidas = len(linhas) - LINHAS_VISIVEIS_REGISTRO
    aviso = ("... %s linha(s) anteriores estao no log completo, dentro do .zip"
             % numero(escondidas))
    return "\n".join([aviso, ""] + linhas[-LINHAS_VISIVEIS_REGISTRO:])


# ---------------------------------------------------------------------------
# Modo de teste - nenhum arquivo e gravado
# ---------------------------------------------------------------------------

def relatorio_de_teste(caminho: Path, cfg: Config) -> list:
    """Mostra o que seria feito em UM arquivo, sem gravar nada em disco."""
    linhas = []
    registrar = linhas.append

    registrar("=" * 64)
    registrar("MODO DE TESTE - nenhum arquivo sera alterado ou criado")
    registrar("=" * 64)
    registrar("Arquivo analisado: %s" % caminho.name)

    info = inspector.inspecionar_arquivo(caminho, cfg)
    if info.erro:
        registrar("ERRO: %s" % info.erro)
        return linhas

    carimbos = quantidade_carimbos(cfg.data_inicial, cfg.data_final,
                                   cfg.intervalo_minutos, True)
    registrar("Linhas de data: %d   |   Linhas de formula: %d"
              % (carimbos, quantidade_formulas(carimbos, True)))
    registrar("")

    selecionadas = {n.strip().lower() for n in (cfg.abas_selecionadas or [])}
    for aba in info.abas:
        if not aba.processavel:
            registrar("[ ] %s - ignorada (%s)" % (aba.nome, aba.motivo_ignorada))
            continue
        if selecionadas and aba.nome.strip().lower() not in selecionadas:
            registrar("[ ] %s - fora da selecao" % aba.nome)
            continue

        inicio, fim, intervalo = cfg.periodo_da_aba(aba.nome)
        proprio = cfg.config_da_aba(aba.nome).personalizada
        registrar("[X] %s - %d TAG(s)" % (aba.nome, len(aba.tags)))
        registrar("    Periodo: %s a %s, a cada %d min%s"
                  % (inicio.strftime("%d/%m/%Y"), fim.strftime("%d/%m/%Y"),
                     intervalo, "  (ajuste proprio)" if proprio else ""))
        for tag in aba.tags:
            if tag.personalizada:
                registrar("      coluna %s - conta propria da planilha, %d TAG(s): %s"
                          % (tag.coluna, tag.personalizada.total_chamadas,
                             " + ".join(tag.fontes)))
            else:
                registrar("      %s%d = %s" % (tag.coluna, cfg.linha_tags, tag.texto))
        if aba.auxiliares:
            registrar("    Colunas de apoio mantidas intactas (so guardam TAG): %s"
                      % ", ".join(tag.coluna for tag in aba.auxiliares))

        carimbos_aba = quantidade_carimbos(inicio, fim, intervalo,
                                           cfg.carimbo_final_extra)
        formulas_aba = quantidade_formulas(carimbos_aba, cfg.carimbo_final_extra)
        primeira_linha = cfg.primeira_linha_dados
        ultima_data = primeira_linha + carimbos_aba - 1
        em_bloco = cfg.modo_formula == MODO_BLOCO

        comuns = [tag for tag in aba.tags if not tag.personalizada]
        if comuns:
            primeira = comuns[0]
            registrar("    Formulas que serao inseridas:")
            if em_bloco:
                texto = montar_formula_bloco(
                    primeira.coluna, cfg.linha_tags, primeira_linha,
                    ultima_data, cfg.coluna_data, cfg.funcao_pi_bloco,
                    intervalo, cfg.argumentos_fixos)
                registrar("      %s%d %s  (uma so, em matriz)"
                          % (primeira.coluna, primeira_linha, para_exibicao(texto)))
            else:
                for deslocamento in range(min(3, formulas_aba)):
                    linha = primeira_linha + deslocamento
                    texto = montar_formula(primeira.coluna, cfg.linha_tags, linha,
                                           cfg.coluna_data, cfg.funcao_pi,
                                           cfg.argumentos_fixos)
                    registrar("      %s%d %s"
                              % (primeira.coluna, linha, para_exibicao(texto)))
                if formulas_aba > 3:
                    registrar("      ... ate a linha %d"
                              % (primeira_linha + formulas_aba - 1))

        for tag in aba.personalizadas:
            registrar("    Conta preservada na coluna %s:" % tag.coluna)
            if writer.usar_bloco_na_coluna_propria(tag.personalizada, cfg):
                texto = composta.montar_bloco(
                    tag.personalizada, primeira_linha, ultima_data,
                    cfg.coluna_data, intervalo_pi(intervalo))
                sufixo = "  (uma so, em matriz)"
            else:
                texto = composta.montar_linha(
                    tag.personalizada, primeira_linha, cfg.coluna_data)
                travam = tag.personalizada.funcoes_que_impedem_bloco
                sufixo = ("  (uma por linha; %s resume a coluna inteira)"
                          % "/".join(travam) if travam else "  (uma por linha)")
            registrar("      %s%d %s%s"
                      % (tag.coluna, primeira_linha, para_exibicao(texto), sufixo))
        registrar("")

    registrar("Nada foi gravado em disco.")
    return linhas


# ---------------------------------------------------------------------------
# Processamento
# ---------------------------------------------------------------------------

def montar_zip(saida: Path, destino: Path) -> Path:
    """Empacota as planilhas processadas e a pasta logs/ num .zip EM DISCO.

    Antes o .zip era montado na memoria e ainda ficava guardado na sessao -
    duas copias do resultado inteiro em memoria, de pe ate a pessoa fechar a
    aba. Agora ele e escrito direto no disco e a sessao guarda so o caminho.
    """
    with zipfile.ZipFile(destino, "w", zipfile.ZIP_DEFLATED) as pacote:
        for caminho in sorted(saida.rglob("*")):
            if caminho.is_file():
                pacote.write(caminho, caminho.relative_to(saida).as_posix())
    return destino


def processar(cfg: Config, barra, status, registro: Registro):
    """Roda o lote atualizando barra, status e registro em tempo real."""

    def progresso(indice, total, nome):
        # Cada planilha e aberta inteira na memoria. Devolver a anterior antes
        # de abrir a proxima e o que segura o consumo no lugar num lote grande.
        gc.collect()
        barra.progress(min(1.0, (indice - 1) / float(max(1, total))))
        status.markdown(
            tema.texto_fraco("Arquivo atual: <b>%s</b> &nbsp;&middot;&nbsp; "
                             "%d de %d" % (nome, indice, total)),
            unsafe_allow_html=True,
        )

    with Logger(cfg.pasta_log, callback=registro) as logger:
        resumo = processor.processar_pasta(cfg, logger, progresso=progresso)

    barra.progress(1.0)
    registro.desenhar()   # garante que as ultimas linhas aparecam na tela
    return resumo


# ---------------------------------------------------------------------------
# Pagina
# ---------------------------------------------------------------------------

st.set_page_config(page_title="Automatizador %s" % NOME_EXIBIDO,
                   page_icon=":bar_chart:",
                   layout="wide", initial_sidebar_state="expanded")
st.markdown(tema.css(), unsafe_allow_html=True)

faxina_de_sessoes_antigas()
st.session_state.setdefault("registro", [])
st.session_state.setdefault("download", None)

# --- Barra lateral: marca, periodo e opcoes --------------------------------
with st.sidebar:
    st.markdown(tema.marca(NOME_EXIBIDO, "Automatizador de planilhas",
                           APP_VERSAO), unsafe_allow_html=True)

    st.markdown(tema.grupo("Periodo"), unsafe_allow_html=True)
    data_inicial = st.date_input("Data inicial", value=DATA_PADRAO,
                                 format="DD/MM/YYYY")
    data_final = st.date_input("Data final", value=DATA_PADRAO,
                               format="DD/MM/YYYY")
    intervalo = st.number_input("Intervalo (minutos)", min_value=1,
                                max_value=100000, value=INTERVALO_PADRAO,
                                step=1,
                                help="60 = 1 hora, 720 = 12 horas, 1440 = 1 dia")

    st.markdown("")
    st.markdown(tema.grupo("Opcoes"), unsafe_allow_html=True)
    rapido = st.checkbox(
        "Buscar tudo de uma vez", value=True,
        help="Muito mais rapido: uma formula PIAdvCalcDat em matriz por coluna, "
             "no lugar de uma PIAdvCalcVal por linha.",
    )
    calculo_manual = st.checkbox("Abrir sem calcular", value=True,
                                 help="Os valores vem com F9 / PI DataLink Refresh.")
    matriz_dinamica = st.checkbox(
        "Matriz dinamica", value=True,
        help="Recomendado - deixa as celulas editaveis no Excel.")
    manter_contas = st.checkbox(
        "Manter contas da planilha", value=True,
        help="Colunas que somam duas ou mais TAGs: reaproveita a formula que ja "
             "esta la e troca so as datas.",
    )

    st.markdown("---")
    st.markdown(
        tema.texto_fraco("Os arquivos enviados ficam numa pasta temporaria da "
                         "sessao. Os originais no seu computador nunca sao "
                         "tocados."),
        unsafe_allow_html=True,
    )

# --- Titulo ----------------------------------------------------------------
st.markdown(tema.titulo_pagina("Processar planilhas", SUBTITULO),
            unsafe_allow_html=True)

# --- 1. Arquivos -----------------------------------------------------------
with st.container(border=True):
    st.markdown(tema.titulo_secao("Arquivos", "1"), unsafe_allow_html=True)
    enviados = st.file_uploader(
        "Arraste aqui as planilhas, ou clique para escolher",
        type=["xlsx", "xlsm"],
        accept_multiple_files=True,
        help="Formatos .xls e .xlsb nao sao suportados - converta para .xlsx antes.",
    )
    enviados = enviados or []
    assinatura = gravar_enviados(enviados)

    entrada = pasta_de_trabalho() / "entrada"
    arquivos = processor.listar_arquivos(entrada) if entrada.is_dir() else []

    resumo = (resumo_dos_arquivos(assinatura, str(entrada), manter_contas)
              if arquivos else [])
    abas_encontradas = abas_com_tag(resumo)

    if arquivos:
        ilegiveis = [item["nome"] for item in resumo if item["erro"]]
        st.markdown(
            tema.texto_fraco("Prontos para processar: <b>%d</b> arquivo(s)%s"
                             % (len(arquivos) - len(ilegiveis),
                                (" &nbsp;&middot;&nbsp; sem leitura: %s"
                                 % ", ".join(ilegiveis)) if ilegiveis else "")),
            unsafe_allow_html=True,
        )

# --- Previsao, so com conta (nenhum arquivo e reaberto aqui) ---------------
periodo_invalido = ""
carimbos_totais = 0
linhas_previstas = 0
if not data_inicial or not data_final:
    # O campo de data aceita ficar vazio; sem isto a pagina quebrava ao desenhar.
    periodo_invalido = "preencha a data inicial e a data final"
else:
    try:
        carimbos_totais = quantidade_carimbos(data_inicial, data_final,
                                              int(intervalo), True)
        linhas_previstas = quantidade_formulas(carimbos_totais, True)
    except ValueError as exc:
        periodo_invalido = str(exc)

ultima_linha_excel = PRIMEIRA_LINHA_DADOS_PADRAO + carimbos_totais - 1
excede_o_excel = ultima_linha_excel > LIMITE_LINHAS_EXCEL

por_aba = {}
problemas_aba = []
abas_escolhidas = []

# --- 2. Abas ---------------------------------------------------------------
with st.container(border=True):
    st.markdown(tema.titulo_secao("Abas", "2"), unsafe_allow_html=True)

    todas_as_abas = st.checkbox("Processar todas as abas que tiverem TAG valida",
                                value=True)
    if not todas_as_abas:
        abas_escolhidas = st.multiselect(
            "Abas a processar", options=sorted(abas_encontradas),
            placeholder="Escolha pelo menos uma aba",
        )

    # A tabela de periodo por aba usa um componente pesado, e quase ninguem
    # precisa dela - por isso ela so e desenhada quando alguem pede.
    proprio_por_aba = st.checkbox(
        "Periodo proprio para alguma aba", value=False,
        disabled=not abas_encontradas,
        help="Ex.: a aba diaria de 1440 em 1440 min enquanto o resto do arquivo "
             "continua de 5 em 5.",
    )

    if proprio_por_aba and abas_encontradas:
        st.markdown(
            tema.texto_fraco("Preencha so o que muda. Campo em branco usa o "
                             "periodo geral da barra lateral."),
            unsafe_allow_html=True,
        )
        linhas_tabela = [{"Aba": nome, "Data inicial": "", "Data final": "",
                          "Intervalo (min)": ""}
                         for nome in sorted(abas_encontradas)]
        editadas = st.data_editor(
            linhas_tabela,
            key="editor_abas_%d" % hash(assinatura),
            hide_index=True,
            num_rows="fixed",
            use_container_width=True,
            disabled=["Aba"],
            column_config={
                "Aba": st.column_config.TextColumn("Aba", width="large"),
                "Data inicial": st.column_config.TextColumn(help="dd/mm/aaaa"),
                "Data final": st.column_config.TextColumn(help="dd/mm/aaaa"),
                "Intervalo (min)": st.column_config.TextColumn(
                    help="60 = 1 hora, 720 = 12 horas, 1440 = 1 dia"),
            },
        )
        # O data_editor devolve o mesmo tipo que recebeu; normalizar aqui evita
        # surpresa caso um dia ele passe a devolver um DataFrame.
        if hasattr(editadas, "to_dict"):
            editadas = editadas.to_dict("records")
        for linha in editadas:
            nome = linha["Aba"]
            try:
                ajuste = ConfigAba(
                    data_inicial=ler_data(linha["Data inicial"], "inicial"),
                    data_final=ler_data(linha["Data final"], "final"),
                )
                bruto = str(linha["Intervalo (min)"] or "").strip()
                if bruto:
                    try:
                        ajuste.intervalo_minutos = int(bruto)
                    except ValueError:
                        raise ValueError("intervalo invalido: %r (informe minutos)"
                                         % bruto)
            except ValueError as exc:
                problemas_aba.append("Aba %r: %s" % (nome, exc))
                continue
            if ajuste.personalizada:
                por_aba[nome] = ajuste

    if not abas_encontradas:
        st.markdown(
            tema.texto_fraco("Envie as planilhas no passo 1 para ver as abas."),
            unsafe_allow_html=True,
        )

    for problema in problemas_aba:
        st.warning(problema)

# --- Numeros da rodada -----------------------------------------------------
consultas_bloco = consultas_linha = 0
if resumo and not periodo_invalido:
    consultas_bloco, consultas_linha = estimar_consultas(
        resumo,
        None if todas_as_abas else abas_escolhidas,
        formulas_de_cada_aba(abas_encontradas, data_inicial, data_final,
                             int(intervalo), por_aba),
        rapido,
    )

consultas = consultas_bloco if rapido else consultas_linha
if periodo_invalido:
    detalhe_linhas = "periodo invalido"
    valor_linhas = "--"
else:
    detalhe_linhas = "%s carimbos de data" % numero(carimbos_totais)
    valor_linhas = numero(linhas_previstas)

st.markdown(tema.cartelas([
    ("Arquivos", numero(len(arquivos)), ".xlsx e .xlsm"),
    ("Abas com TAG", numero(len(abas_encontradas)), "no conjunto enviado"),
    ("Linhas por arquivo", valor_linhas, detalhe_linhas),
    ("Consultas ao PI", numero(consultas) if consultas else "--",
     ("sem o modo rapido: %s" % numero(consultas_linha)) if rapido and consultas
     else "no arquivo mais pesado"),
]), unsafe_allow_html=True)

if periodo_invalido:
    st.error("Periodo invalido: %s" % periodo_invalido)
elif excede_o_excel:
    st.error("Este periodo geraria %s linhas e o Excel so vai ate %s. "
             "Aumente o intervalo ou encurte o periodo."
             % (numero(ultima_linha_excel), numero(LIMITE_LINHAS_EXCEL)))
else:
    st.markdown(
        tema.texto_fraco(
            "De %s ate %s, a cada %d minuto(s)."
            % (datetime.combine(data_inicial, datetime.min.time())
               .strftime("%d/%m/%Y %H:%M"),
               ultimo_carimbo(data_inicial, data_final, int(intervalo), True)
               .strftime("%d/%m/%Y %H:%M"),
               int(intervalo))),
        unsafe_allow_html=True,
    )
    if linhas_previstas > LIMITE_LINHAS_AVISO:
        st.warning("Periodo grande: %s linhas por arquivo. A planilha fica "
                   "pesada e o processamento pode faltar memoria - se puder, "
                   "aumente o intervalo ou divida em rodadas."
                   % numero(linhas_previstas))

# --- 3. Executar -----------------------------------------------------------
with st.container(border=True):
    st.markdown(tema.titulo_secao("Executar", "3"), unsafe_allow_html=True)

    arquivo_de_teste = None
    if arquivos:
        arquivo_de_teste = st.selectbox(
            "Arquivo usado no teste",
            options=[caminho.name for caminho in arquivos],
            help="O teste apenas mostra o que seria feito; nada e gravado.",
        )

    # use_container_width e o unico jeito de esticar o botao que funciona em
    # todas as versoes do Streamlit: width="stretch" so existe das versoes novas
    # em diante, e nas antigas o parametro width do data_editor era em pixels.
    col_teste, col_processar = st.columns([1, 2])
    pedido_teste = col_teste.button("Testar 1 arquivo", disabled=not arquivos,
                                    use_container_width=True)
    pedido_processar = col_processar.button("PROCESSAR", type="primary",
                                            disabled=not arquivos,
                                            use_container_width=True)
    barra = st.empty()
    status = st.empty()

# --- Registro --------------------------------------------------------------
with st.container(border=True, key="registro"):
    st.markdown(tema.titulo_secao("Registro da execucao"), unsafe_allow_html=True)
    caixa_log = st.empty()


if pedido_teste or pedido_processar:
    saida = pasta_de_trabalho() / "saida"
    cfg = None
    problemas = []
    if not periodo_invalido:
        cfg = montar_config(
            entrada, saida, data_inicial, data_final, int(intervalo),
            None if todas_as_abas else abas_escolhidas, por_aba,
            rapido=rapido, calculo_manual=calculo_manual,
            matriz_dinamica=matriz_dinamica, manter_contas=manter_contas,
        )
        problemas = cfg.validar()

    if periodo_invalido:
        st.error("Periodo invalido: %s" % periodo_invalido)
    elif problemas_aba:
        st.error("Corrija os periodos por aba antes de continuar.")
    elif excede_o_excel:
        st.error("Periodo maior do que uma planilha do Excel comporta.")
    elif not todas_as_abas and not abas_escolhidas:
        st.error("Selecione pelo menos uma aba, ou marque 'Processar todas as abas'.")
    elif problemas:
        st.error("Verifique os campos:\n%s"
                 % "\n".join("- " + item for item in problemas))
    elif pedido_teste:
        caminho = next(c for c in arquivos if c.name == arquivo_de_teste)
        st.session_state.registro = relatorio_de_teste(caminho, cfg)
        descartar_download()
    else:
        # A saida e refeita a cada execucao: assim o .zip traz so o resultado
        # desta rodada e nada fica marcado como "ja existe na pasta de saida".
        descartar_download()
        shutil.rmtree(saida, ignore_errors=True)
        saida.mkdir(parents=True, exist_ok=True)
        try:
            registro = Registro(caixa_log)
            resultado = processar(cfg, barra.progress(0.0), status, registro)
            st.session_state.registro = registro.linhas

            nome_zip = "processados_%s.zip" % datetime.now().strftime("%Y%m%d_%H%M%S")
            caminho_zip = montar_zip(saida, pasta_de_trabalho() / nome_zip)
            # O .zip ja tem tudo; manter a pasta de saida so dobraria o disco.
            shutil.rmtree(saida, ignore_errors=True)
            st.session_state.download = (nome_zip, str(caminho_zip))

            status.markdown(
                tema.texto_fraco("%d com sucesso, %d com erro, %d ignorados (de %d)"
                                 % (resultado.sucesso, resultado.erros,
                                    resultado.ignorados,
                                    len(resultado.resultados))),
                unsafe_allow_html=True,
            )
            if resultado.erros:
                st.warning("Concluido com erros: %d arquivo(s) falharam - "
                           "veja o registro." % resultado.erros)
            elif resultado.sucesso:
                st.success("Concluido: %d arquivo(s) processado(s) com sucesso."
                           % resultado.sucesso)
            else:
                st.info("Nenhum arquivo foi processado - veja o registro.")
        except Exception:
            falha = traceback.format_exc()
            st.session_state.registro = falha.splitlines()
            descartar_download()
            st.error("Falha inesperada: %s" % falha.strip().splitlines()[-1])

# Mostra o registro guardado (o proprio st.code ja traz o botao de copiar).
caixa_log.code(texto_do_registro(st.session_state.registro), language="text")

if st.session_state.download:
    nome, caminho_zip = st.session_state.download
    caminho_zip = Path(caminho_zip)
    if caminho_zip.is_file():
        # O arquivo e lido do disco na hora de desenhar o botao: a sessao nao
        # fica segurando o resultado inteiro na memoria.
        with caminho_zip.open("rb") as pacote:
            st.download_button(
                "Baixar arquivos processados (.zip)",
                data=pacote, file_name=nome, mime="application/zip",
                type="primary", use_container_width=True,
                help="O .zip traz as copias processadas e a pasta logs/ com o registro.",
            )
    else:
        st.session_state.download = None
