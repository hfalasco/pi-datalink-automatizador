# -*- coding: utf-8 -*-
"""Pagina Tirar formulas: troca as formulas pelos valores ja calculados.

E o antigo tirar_formulas.py em forma de pagina, no mesmo roteiro da pagina
Processar planilhas: a pessoa envia as planilhas, clica em TIRAR FORMULAS,
acompanha o registro e baixa o resultado - cada planilha com o sufixo
_SEM_FORMULAS no nome, so com os numeros. O trabalho pesado esta em
core/sem_formulas.py.

As planilhas desta pagina sao as ja calculadas pelo PI DataLink, de dezenas a
centenas de MB cada. Por isso:

  * o limite de envio e maior que o da outra pagina (LIMITE_ENVIO_MB);
  * os arquivos sao tratados um de cada vez, com as abas lidas em pedacos;
  * com um arquivo so, o resultado e a propria planilha; com varios, um .zip
    que guarda as planilhas sem compactar de novo (um .xlsx ja e compactado).
"""

from __future__ import annotations

import gc
import shutil
import traceback
from datetime import datetime

import streamlit as st

import comum
import tema
from comum import numero
from core import processor, sem_formulas
from logger import Logger

SUBTITULO = ("Troca cada formula pelo valor que o Excel ja calculou. "
             "Os arquivos originais nunca sao alterados.")
# Planilhas ja calculadas pelo PI DataLink chegam a centenas de MB. O teto do
# servidor (server.maxUploadSize, em .streamlit/config.toml) e este mesmo numero.
LIMITE_ENVIO_MB = 500

area = comum.Area("sem_formulas")


@st.cache_data(show_spinner=False, max_entries=2, ttl=3600)
def abas_dos_arquivos(assinatura: tuple, entrada: str) -> dict:
    """{nome do arquivo: quantidade de abas}; None quando nao abre.

    Le so o xl/workbook.xml de cada arquivo (poucos KB), uma vez por conjunto
    enviado - as abas em si so sao lidas quando a pessoa clica no botao.
    """
    contagem = {}
    for caminho in processor.listar_arquivos(entrada):
        try:
            contagem[caminho.name] = len(sem_formulas.nomes_das_abas(caminho))
        except Exception:
            contagem[caminho.name] = None
    return contagem


def registrar_resultado(logger: Logger, resultado) -> None:
    """As mesmas linhas que o script imprimia, no registro e no log."""
    logger.escrever("Arquivo: %s" % resultado.arquivo.name)
    if not resultado.sucesso:
        logger.escrever("  ERRO: %s" % resultado.erro)
        logger.escrever("Status: ERRO")
        logger.linha_em_branco()
        return

    for aba in resultado.abas:
        medidas = "%s linhas" % numero(aba.linhas)
        if aba.colunas:
            medidas += ", %s colunas" % numero(aba.colunas)
        logger.escrever("  Aba '%s': %s, %s formula(s) trocada(s) pelo valor"
                        % (aba.nome, medidas, numero(aba.formulas)))
    logger.escrever("  Gerado: %s" % resultado.saida.name)
    if resultado.sem_valor:
        logger.escrever("  ATENCAO: %s formula(s) ficaram VAZIAS - o arquivo "
                        "nao tinha o valor gravado." % numero(resultado.sem_valor))
        logger.escrever("  Abra o original no Excel, espere o PI DataLink "
                        "atualizar, salve (Ctrl+S) e envie de novo.")
    logger.escrever("Status: OK")
    logger.linha_em_branco()


def tirar_do_lote(arquivos: list, saida, barra, status,
                  registro: comum.Registro) -> list:
    """Converte os arquivos UM DE CADA VEZ, atualizando barra e registro.

    Um de cada vez de proposito, como no script: com arquivos deste tamanho,
    varios ao mesmo tempo so disputam memoria e disco.
    """
    resultados = []
    total = len(arquivos)
    nome_log = datetime.now().strftime("log_sem_formulas_%Y%m%d_%H%M%S.txt")

    with Logger(saida / "logs", callback=registro, nome_arquivo=nome_log) as logger:
        logger.escrever("CONVERSAO INICIADA")
        logger.escrever("Data: " + datetime.now().strftime("%d/%m/%Y %H:%M:%S"))
        logger.escrever("Arquivos: %d (um de cada vez)" % total)
        logger.linha_em_branco()

        for indice, caminho in enumerate(arquivos, start=1):
            gc.collect()
            status.markdown(
                tema.texto_fraco("Arquivo atual: <b>%s</b> &nbsp;&middot;&nbsp; "
                                 "%d de %d" % (caminho.name, indice, total)),
                unsafe_allow_html=True,
            )

            def progresso(fracao, indice=indice):
                barra.progress(min(1.0, (indice - 1 + fracao) / float(total)))

            destino = saida / sem_formulas.nome_de_saida(caminho)
            resultado = sem_formulas.tirar_formulas(caminho, destino, progresso)
            registrar_resultado(logger, resultado)
            resultados.append(resultado)

        certos = [item for item in resultados if item.sucesso]
        logger.escrever("CONVERSAO CONCLUIDA")
        logger.escrever("Arquivos convertidos com sucesso: %d/%d"
                        % (len(certos), total))
        falhas = [item for item in resultados if not item.sucesso]
        if falhas:
            logger.escrever("Arquivos com erro: %d" % len(falhas))
            for item in falhas:
                logger.escrever("  - %s: %s" % (item.arquivo.name, item.erro))
        logger.escrever("Termino: " + datetime.now().strftime("%d/%m/%Y %H:%M:%S"))

    barra.progress(1.0)
    registro.desenhar()   # garante que as ultimas linhas aparecam na tela
    return resultados


def guardar_resultado(resultados: list, saida) -> None:
    """Deixa pronto o que a pessoa vai baixar.

    Um arquivo so: a propria planilha, sem .zip no meio. Varios: um .zip com
    as planilhas e a pasta logs/.
    """
    certos = [item for item in resultados if item.sucesso]
    if not certos:
        return
    if len(resultados) == 1:
        planilha = certos[0].saida
        destino = area.pasta / planilha.name
        shutil.move(str(planilha), str(destino))
        area.download = (planilha.name, str(destino))
    else:
        nome_zip = datetime.now().strftime("sem_formulas_%Y%m%d_%H%M%S.zip")
        area.download = (nome_zip,
                         str(comum.montar_zip(saida, area.pasta / nome_zip)))
    # O resultado ja esta separado; manter a pasta de saida so dobraria o disco.
    shutil.rmtree(saida, ignore_errors=True)


# ---------------------------------------------------------------------------
# Pagina
# ---------------------------------------------------------------------------

# --- Barra lateral: o passo a passo (a marca e os modulos vem de app.py) ---
with st.sidebar:
    st.markdown("")
    st.markdown(tema.grupo("Antes de enviar"), unsafe_allow_html=True)
    st.markdown(tema.passos([
        "Abra a planilha no Excel e espere o PI DataLink atualizar.",
        "Salve (Ctrl+S) e feche o arquivo.",
        "Envie aqui e clique em <b>TIRAR FORMULAS</b>.",
    ]), unsafe_allow_html=True)

    st.markdown("---")
    st.markdown(
        tema.texto_fraco("Os arquivos enviados ficam numa pasta temporaria da "
                         "sessao. Os originais no seu computador nunca sao "
                         "tocados."),
        unsafe_allow_html=True,
    )

# --- Titulo ----------------------------------------------------------------
st.markdown(tema.titulo_pagina("Tirar formulas", SUBTITULO),
            unsafe_allow_html=True)

# --- 1. Arquivos -----------------------------------------------------------
with st.container(border=True):
    st.markdown(tema.titulo_secao("Arquivos", "1"), unsafe_allow_html=True)
    st.markdown(
        tema.aviso("Envie a planilha <b>ja calculada</b>: aberta no Excel, com o "
                   "PI DataLink atualizado e salva. O programa usa os valores "
                   "que o Excel gravou no arquivo - formula sem valor gravado "
                   "vira celula vazia."),
        unsafe_allow_html=True,
    )
    enviados = st.file_uploader(
        "Arraste aqui as planilhas, ou clique para escolher",
        type=["xlsx", "xlsm"],
        accept_multiple_files=True,
        max_upload_size=LIMITE_ENVIO_MB,
        help="Formatos .xls e .xlsb nao sao suportados - abra no Excel e salve "
             "como .xlsx antes.",
    )
    enviados = enviados or []
    assinatura = area.gravar_enviados(enviados)

    entrada = area.entrada
    arquivos = processor.listar_arquivos(entrada) if entrada.is_dir() else []
    abas = abas_dos_arquivos(assinatura, str(entrada)) if arquivos else {}

    if arquivos:
        ilegiveis = [nome for nome, quantidade in abas.items() if quantidade is None]
        st.markdown(
            tema.texto_fraco("Prontos para converter: <b>%d</b> arquivo(s)%s"
                             % (len(arquivos) - len(ilegiveis),
                                (" &nbsp;&middot;&nbsp; sem leitura: %s"
                                 % ", ".join(ilegiveis)) if ilegiveis else "")),
            unsafe_allow_html=True,
        )

# --- Numeros ---------------------------------------------------------------
def numeros() -> str:
    """As quatro cartelas. A ultima so vale para o conjunto enviado agora."""
    ultimo = area.lido("resultado")
    if ultimo and ultimo[0] == assinatura:
        _, trocadas, vazias = ultimo
        valor_trocadas = numero(trocadas)
        detalhe_trocadas = ("%s ficaram vazias" % numero(vazias) if vazias
                            else "todas com valor gravado")
    else:
        valor_trocadas, detalhe_trocadas = "--", "aparece depois de converter"
    return tema.cartelas([
        ("Arquivos", numero(len(arquivos)), ".xlsx e .xlsm"),
        ("Tamanho", comum.tamanho(sum(c.stat().st_size for c in arquivos))
         if arquivos else "--", "no conjunto enviado"),
        ("Abas", numero(sum(q for q in abas.values() if q)) if arquivos else "--",
         "todas sao convertidas"),
        ("Formulas trocadas", valor_trocadas, detalhe_trocadas),
    ])


painel = st.empty()
painel.markdown(numeros(), unsafe_allow_html=True)

# --- 2. Executar -----------------------------------------------------------
with st.container(border=True):
    st.markdown(tema.titulo_secao("Executar", "2"), unsafe_allow_html=True)
    st.markdown(
        tema.texto_fraco("Cada planilha volta com o sufixo "
                         "<b>%s</b> no nome. Cores, larguras, mesclagens e "
                         "graficos ficam como estavam; so as formulas saem."
                         % sem_formulas.SUFIXO_SAIDA),
        unsafe_allow_html=True,
    )
    pedido = st.button("TIRAR FORMULAS", type="primary", disabled=not arquivos,
                       use_container_width=True)
    barra = st.empty()
    status = st.empty()

# --- Registro --------------------------------------------------------------
caixa_log = comum.caixa_do_registro()

if pedido:
    saida = area.saida
    # A saida e refeita a cada execucao: o resultado traz so esta rodada.
    area.descartar_download()
    area.guardar("resultado", None)
    shutil.rmtree(saida, ignore_errors=True)
    saida.mkdir(parents=True, exist_ok=True)
    try:
        registro = comum.Registro(caixa_log)
        resultados = tirar_do_lote(arquivos, saida, barra.progress(0.0), status,
                                   registro)
        area.registro = registro.linhas
        guardar_resultado(resultados, saida)

        certos = [item for item in resultados if item.sucesso]
        erros = len(resultados) - len(certos)
        trocadas = sum(item.formulas for item in certos)
        vazias = sum(item.sem_valor for item in certos)
        area.guardar("resultado", (assinatura, trocadas, vazias))

        status.markdown(
            tema.texto_fraco("%d convertido(s), %d com erro (de %d)"
                             % (len(certos), erros, len(resultados))),
            unsafe_allow_html=True,
        )
        if erros:
            st.warning("Concluido com erros: %d arquivo(s) falharam - veja o "
                       "registro." % erros)
        if vazias:
            st.warning("%s formula(s) nao tinham valor gravado e ficaram vazias. "
                       "Abra o arquivo original no Excel, espere o PI DataLink "
                       "atualizar, salve (Ctrl+S) e envie de novo."
                       % numero(vazias))
        if certos and not erros and not vazias:
            st.success("Concluido: %d arquivo(s) sem formulas - %s formula(s) "
                       "trocada(s) pelo valor." % (len(certos), numero(trocadas)))
        painel.markdown(numeros(), unsafe_allow_html=True)
    except Exception:
        falha = traceback.format_exc()
        area.registro = falha.splitlines()
        area.descartar_download()
        st.error("Falha inesperada: %s" % falha.strip().splitlines()[-1])

# Mostra o registro guardado (o proprio st.code ja traz o botao de copiar).
caixa_log.code(comum.texto_do_registro(area.registro), language="text")

if area.download and area.download[0].lower().endswith(".zip"):
    comum.botao_de_download(
        area, "Baixar planilhas sem formulas (.zip)",
        "O .zip traz as planilhas sem formulas e a pasta logs/ com o registro.")
else:
    comum.botao_de_download(
        area, "Baixar planilha sem formulas",
        "A planilha so com os valores. O registro fica aqui na tela.")
