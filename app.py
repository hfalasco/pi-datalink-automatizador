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
"""

from __future__ import annotations

import io
import shutil
import sys
import tempfile
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
from config import (APP_NOME, APP_VERSAO, MODO_BLOCO,  # noqa: E402
                    MODO_LINHA, Config, ConfigAba)
from core import composta, inspector, processor, writer  # noqa: E402
from core.formula import (intervalo_pi, montar_formula,  # noqa: E402
                          montar_formula_bloco, para_exibicao)
from core.timeline import gerar_carimbos, quantidade_formulas  # noqa: E402
from logger import Logger  # noqa: E402

DATA_PADRAO = date(2026, 9, 11)
INTERVALO_PADRAO = 5
SUBTITULO = ("Preenche as formulas PIAdvCalcVal em lote. "
             "Os arquivos originais nunca sao alterados.")


# ---------------------------------------------------------------------------
# Pasta de trabalho da sessao
# ---------------------------------------------------------------------------

def pasta_de_trabalho() -> Path:
    """Pasta temporaria propria desta sessao, com entrada/ e saida/ dentro."""
    if "pasta_trabalho" not in st.session_state:
        st.session_state.pasta_trabalho = tempfile.mkdtemp(prefix="pi_datalink_")
    return Path(st.session_state.pasta_trabalho)


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
        st.session_state.download = None

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


def _serializar_por_aba(por_aba: dict) -> tuple:
    """Forma hashavel dos ajustes por aba, para servir de chave de cache."""
    return tuple(sorted(
        (nome,
         ajuste.data_inicial.isoformat() if ajuste.data_inicial else "",
         ajuste.data_final.isoformat() if ajuste.data_final else "",
         ajuste.intervalo_minutos or 0)
        for nome, ajuste in (por_aba or {}).items()
    ))


def _desserializar_por_aba(serial: tuple) -> dict:
    ajustes = {}
    for nome, inicio, fim, intervalo in serial:
        ajustes[nome] = ConfigAba(
            data_inicial=date.fromisoformat(inicio) if inicio else None,
            data_final=date.fromisoformat(fim) if fim else None,
            intervalo_minutos=intervalo or None,
        )
    return ajustes


@st.cache_data(show_spinner="Lendo as abas dos arquivos...")
def abas_dos_arquivos(assinatura: tuple, entrada: str, manter_contas: bool) -> dict:
    """{nome da aba: em quantos arquivos ela tem TAG}. Nao depende do periodo."""
    cfg = montar_config(entrada, Path(entrada).parent / "saida",
                        DATA_PADRAO, DATA_PADRAO, INTERVALO_PADRAO,
                        None, None, manter_contas=manter_contas)
    return processor.abas_disponiveis(processor.listar_arquivos(entrada), cfg)


@st.cache_data(show_spinner=False)
def estimativa_de_consultas(assinatura: tuple, entrada: str, inicio: date,
                            fim: date, intervalo: int, abas: tuple,
                            por_aba_serial: tuple, manter_contas: bool,
                            rapido: bool) -> tuple:
    """(consultas no modo bloco, consultas no modo linha) do arquivo mais pesado."""
    cfg = montar_config(entrada, Path(entrada).parent / "saida", inicio, fim,
                        intervalo, list(abas),
                        _desserializar_por_aba(por_aba_serial),
                        rapido=rapido, manter_contas=manter_contas)
    try:
        return processor.estimar_consultas(processor.listar_arquivos(entrada), cfg)
    except Exception:
        return (0, 0)


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

    carimbos = gerar_carimbos(cfg.data_inicial, cfg.data_final,
                              cfg.intervalo_minutos, True)
    formulas = quantidade_formulas(carimbos, True)
    registrar("Linhas de data: %d   |   Linhas de formula: %d"
              % (len(carimbos), formulas))
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

        carimbos_aba = gerar_carimbos(inicio, fim, intervalo, cfg.carimbo_final_extra)
        formulas_aba = quantidade_formulas(carimbos_aba, cfg.carimbo_final_extra)
        primeira_linha = cfg.primeira_linha_dados
        ultima_data = primeira_linha + len(carimbos_aba) - 1
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

def montar_zip(saida: Path) -> bytes:
    """Empacota as planilhas processadas e a pasta logs/ em um unico .zip."""
    memoria = io.BytesIO()
    with zipfile.ZipFile(memoria, "w", zipfile.ZIP_DEFLATED) as pacote:
        for caminho in sorted(saida.rglob("*")):
            if caminho.is_file():
                pacote.write(caminho, caminho.relative_to(saida).as_posix())
    return memoria.getvalue()


def processar(cfg: Config, barra, status, caixa_log):
    """Roda o lote atualizando barra, status e registro em tempo real."""
    linhas = []

    def registrar(linha):
        linhas.append(linha)
        caixa_log.code("\n".join(linhas), language="text")

    def progresso(indice, total, nome):
        barra.progress(min(1.0, (indice - 1) / float(max(1, total))))
        status.markdown(
            tema.texto_fraco("Arquivo atual: <b>%s</b> &nbsp;&middot;&nbsp; "
                             "%d de %d" % (nome, indice, total)),
            unsafe_allow_html=True,
        )

    with Logger(cfg.pasta_log, callback=registrar) as logger:
        resumo = processor.processar_pasta(cfg, logger, progresso=progresso)

    barra.progress(1.0)
    return resumo, linhas


# ---------------------------------------------------------------------------
# Pagina
# ---------------------------------------------------------------------------

st.set_page_config(page_title=APP_NOME.title(), page_icon=":bar_chart:",
                   layout="centered")
st.markdown(tema.css(), unsafe_allow_html=True)
st.markdown(tema.cabecalho("%s  v%s" % (APP_NOME, APP_VERSAO), SUBTITULO),
            unsafe_allow_html=True)

st.session_state.setdefault("registro", [])
st.session_state.setdefault("download", None)

# --- Arquivos --------------------------------------------------------------
with st.container(border=True):
    st.markdown(tema.titulo_secao("Arquivos"), unsafe_allow_html=True)
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
    st.markdown(tema.texto_fraco("Arquivos enviados: <b>%d</b>" % len(arquivos)),
                unsafe_allow_html=True)

# --- Periodo ---------------------------------------------------------------
with st.container(border=True):
    st.markdown(tema.titulo_secao("Periodo"), unsafe_allow_html=True)
    col_inicio, col_fim, col_intervalo = st.columns(3)
    data_inicial = col_inicio.date_input("Data inicial", value=DATA_PADRAO,
                                         format="DD/MM/YYYY")
    data_final = col_fim.date_input("Data final", value=DATA_PADRAO,
                                    format="DD/MM/YYYY")
    intervalo = col_intervalo.number_input("Intervalo (minutos)", min_value=1,
                                           max_value=100000,
                                           value=INTERVALO_PADRAO, step=1)
    previsao = st.empty()

# --- Abas ------------------------------------------------------------------
with st.container(border=True):
    st.markdown(tema.titulo_secao("Abas"), unsafe_allow_html=True)
    manter_contas = st.checkbox(
        "Manter contas da planilha (colunas que somam duas ou mais TAGs)",
        value=True,
        help="Reaproveita a formula que ja esta na coluna e troca so as datas.",
    )

    abas_encontradas = (abas_dos_arquivos(assinatura, str(entrada), manter_contas)
                        if arquivos else {})

    todas_as_abas = st.checkbox("Processar todas as abas que tiverem TAG valida",
                                value=True)
    abas_escolhidas = []
    if not todas_as_abas:
        abas_escolhidas = st.multiselect(
            "Abas a processar", options=sorted(abas_encontradas),
            placeholder="Escolha pelo menos uma aba",
        )

    por_aba = {}
    problemas_aba = []
    if abas_encontradas:
        st.markdown(
            tema.texto_fraco("Cada aba pode ter periodo e intervalo proprios. "
                             "Deixe em branco para usar o periodo geral acima."),
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

    for problema in problemas_aba:
        st.warning(problema)

# --- Opcoes ----------------------------------------------------------------
with st.container(border=True):
    st.markdown(tema.titulo_secao("Opcoes"), unsafe_allow_html=True)
    rapido = st.checkbox(
        "Buscar tudo de uma vez (muito mais rapido)", value=True,
        help="Uma formula PIAdvCalcDat em matriz por coluna, no lugar de uma "
             "PIAdvCalcVal por linha.",
    )
    calculo_manual = st.checkbox("Abrir sem calcular (buscar com F9 no Excel)",
                                 value=True)
    matriz_dinamica = st.checkbox(
        "Matriz dinamica (recomendado - deixa as celulas editaveis no Excel)",
        value=True)

# --- Previsao (depende das opcoes acima) -----------------------------------
try:
    carimbos = gerar_carimbos(data_inicial, data_final, int(intervalo), True)
    formulas = quantidade_formulas(carimbos, True)
    texto_previsao = ("Serao geradas %d linhas (%d carimbos de data), de %s ate %s."
                      % (formulas, len(carimbos),
                         carimbos[0].strftime("%d/%m/%Y %H:%M"),
                         carimbos[-1].strftime("%d/%m/%Y %H:%M")))
    if arquivos:
        bloco, por_linha = estimativa_de_consultas(
            assinatura, str(entrada), data_inicial, data_final, int(intervalo),
            tuple(abas_escolhidas), _serializar_por_aba(por_aba),
            manter_contas, rapido)
        if bloco and rapido:
            texto_previsao += ("<br>Consultas ao PI no arquivo mais pesado: "
                               "<b>%d</b>  (sem o modo rapido seriam %d)."
                               % (bloco, por_linha))
        elif bloco:
            texto_previsao += ("<br>Consultas ao PI no arquivo mais pesado: "
                               "<b>%d</b>.  Marque 'Buscar tudo de uma vez' "
                               "para cair para %d." % (por_linha, bloco))
    previsao.markdown(tema.texto_fraco(texto_previsao), unsafe_allow_html=True)
except ValueError as exc:
    previsao.markdown(tema.texto_fraco("Periodo invalido: %s" % exc),
                      unsafe_allow_html=True)

# --- Acoes -----------------------------------------------------------------
col_teste, col_processar = st.columns([1, 2])
pedido_teste = col_teste.button("Testar 1 arquivo", disabled=not arquivos)
pedido_processar = col_processar.button("PROCESSAR", type="primary",
                                        disabled=not arquivos)

arquivo_de_teste = None
if arquivos:
    arquivo_de_teste = st.selectbox(
        "Arquivo usado no teste",
        options=[caminho.name for caminho in arquivos],
        help="O teste apenas mostra o que seria feito; nada e gravado.",
    )

# --- Progresso -------------------------------------------------------------
barra = st.empty()
status = st.empty()

# --- Registro --------------------------------------------------------------
with st.container(border=True, key="registro"):
    st.markdown(tema.titulo_secao("Registro da execucao"), unsafe_allow_html=True)
    caixa_log = st.empty()


if pedido_teste or pedido_processar:
    saida = pasta_de_trabalho() / "saida"
    cfg = montar_config(
        entrada, saida, data_inicial, data_final, int(intervalo),
        None if todas_as_abas else abas_escolhidas, por_aba,
        rapido=rapido, calculo_manual=calculo_manual,
        matriz_dinamica=matriz_dinamica, manter_contas=manter_contas,
    )
    problemas = cfg.validar()

    if problemas_aba:
        st.error("Corrija os periodos por aba antes de continuar.")
    elif not todas_as_abas and not abas_escolhidas:
        st.error("Selecione pelo menos uma aba, ou marque 'Processar todas as abas'.")
    elif problemas:
        st.error("Verifique os campos:\n%s"
                 % "\n".join("- " + item for item in problemas))
    elif pedido_teste:
        caminho = next(c for c in arquivos if c.name == arquivo_de_teste)
        st.session_state.registro = relatorio_de_teste(caminho, cfg)
        st.session_state.download = None
    else:
        # A saida e refeita a cada execucao: assim o .zip traz so o resultado
        # desta rodada e nada fica marcado como "ja existe na pasta de saida".
        shutil.rmtree(saida, ignore_errors=True)
        saida.mkdir(parents=True, exist_ok=True)
        try:
            resumo, linhas = processar(cfg, barra.progress(0.0), status, caixa_log)
            st.session_state.registro = linhas
            st.session_state.download = (
                "processados_%s.zip" % datetime.now().strftime("%Y%m%d_%H%M%S"),
                montar_zip(saida),
            )
            status.markdown(
                tema.texto_fraco("%d com sucesso, %d com erro, %d ignorados (de %d)"
                                 % (resumo.sucesso, resumo.erros, resumo.ignorados,
                                    len(resumo.resultados))),
                unsafe_allow_html=True,
            )
            if resumo.erros:
                st.warning("Concluido com erros: %d arquivo(s) falharam - "
                           "veja o registro." % resumo.erros)
            elif resumo.sucesso:
                st.success("Concluido: %d arquivo(s) processado(s) com sucesso."
                           % resumo.sucesso)
            else:
                st.info("Nenhum arquivo foi processado - veja o registro.")
        except Exception:
            falha = traceback.format_exc()
            st.session_state.registro = falha.splitlines()
            st.session_state.download = None
            st.error("Falha inesperada: %s" % falha.strip().splitlines()[-1])

# Mostra o registro guardado (o proprio st.code ja traz o botao de copiar).
caixa_log.code("\n".join(st.session_state.registro)
               or "O registro desta execucao aparece aqui.", language="text")

if st.session_state.download:
    nome, conteudo = st.session_state.download
    st.download_button(
        "Baixar arquivos processados (.zip)",
        data=conteudo, file_name=nome, mime="application/zip", type="primary",
        help="O .zip traz as copias processadas e a pasta logs/ com o registro.",
    )
