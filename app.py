# -*- coding: utf-8 -*-
"""Entrada do Automatizador PI DataLink.

Para abrir:

    streamlit run app.py

O programa tem duas paginas, escolhidas nos Modulos da barra lateral:

  * Processar planilhas (paginas/processar.py) - preenche as formulas do PI
    DataLink em lote;
  * Tirar formulas (paginas/tirar_formulas.py) - troca as formulas pelos
    valores ja calculados, deixando so os numeros.

Este arquivo roda a cada clique, em qualquer pagina. Ele desenha o que e
comum as duas - o tema e o topo da barra lateral, com a marca e os modulos -
e em seguida roda a pagina escolhida, que acrescenta o resto (os ajustes dela
na barra lateral e o miolo).
"""

from __future__ import annotations

import sys
from pathlib import Path

import streamlit as st

# Esta pasta e a raiz da aplicacao, e os modulos sao importados direto dela.
# Nada depende do NOME da pasta: ao publicar, o servidor clona o repositorio
# numa pasta com o nome do repositorio ("pi-datalink-automatizador", com
# hifens), que nao seria um nome de pacote valido.
_RAIZ = Path(__file__).resolve().parent
if str(_RAIZ) not in sys.path:
    sys.path.insert(0, str(_RAIZ))

import comum  # noqa: E402
import tema  # noqa: E402
from config import APP_VERSAO  # noqa: E402
from core import modelo  # noqa: E402

# O nome tecnico (APP_NOME, em config.py) e todo em maiusculas e serve ao log.
# Na tela vale a grafia da marca: APP_NOME.title() escreveria "Pi Datalink".
NOME_EXIBIDO = "PI DataLink"


@st.cache_data(show_spinner=False)
def planilha_base() -> bytes:
    """O .xlsx do modelo em branco - tem poucos KB e e sempre o mesmo."""
    return modelo.montar_planilha_base()


st.set_page_config(page_title="Automatizador %s" % NOME_EXIBIDO,
                   page_icon=str(tema.CAMINHO_LOGO),
                   layout="wide", initial_sidebar_state="expanded")
st.markdown(tema.css(), unsafe_allow_html=True)

comum.faxina_de_sessoes_antigas()
comum.manter_ajustes()

# O menu automatico do Streamlit fica escondido: os links das paginas entram
# na barra lateral, dentro de Modulos, logo abaixo da marca.
PROCESSAR = st.Page(str(_RAIZ / "paginas" / "processar.py"),
                    title="Processar planilhas", icon=":material/table_view:",
                    default=True)
TIRAR_FORMULAS = st.Page(str(_RAIZ / "paginas" / "tirar_formulas.py"),
                         title="Tirar formulas", icon=":material/function:",
                         url_path="tirar-formulas")
pagina = st.navigation([PROCESSAR, TIRAR_FORMULAS], position="hidden")

# --- Barra lateral: marca e modulos (o resto vem de cada pagina) -----------
with st.sidebar:
    st.image(str(tema.CAMINHO_LOGO), width="stretch")
    st.markdown(tema.marca(NOME_EXIBIDO, "Automatizador de planilhas",
                           APP_VERSAO), unsafe_allow_html=True)

    st.markdown(tema.grupo("Modulos"), unsafe_allow_html=True)
    # O Streamlit nao marca no HTML qual link e o da pagina aberta; a chave
    # do container e que deixa o tema pintar esse link de verde.
    with st.container(key="paginas", gap=None):
        for destino in (PROCESSAR, TIRAR_FORMULAS):
            aberta = destino.url_path == pagina.url_path
            with st.container(key="link_ativo" if aberta else None):
                st.page_link(destino, use_container_width=True)

    # Planilha base: o modelo em branco, com a estrutura que o programa
    # reconhece ja destacada (ver core/modelo.py).
    celulas = modelo.celulas_do_modelo()
    with st.container(border=True):
        st.markdown(
            tema.modulo("Planilha base",
                        "O modelo das nossas planilhas, com cada celula "
                        "destacada."),
            unsafe_allow_html=True,
        )
        st.markdown(
            tema.miniatura_planilha(modelo.ROTULO_DATA, modelo.ROTULO_NOME,
                                    modelo.ROTULO_TAG),
            unsafe_allow_html=True,
        )
        st.markdown(
            tema.texto_fraco(
                "Preencha <b>%s</b> com o nome do indicador e <b>%s</b> com a "
                "TAG do PI. Para mais indicadores, repita nas colunas C, D..."
                % (celulas["nome"], celulas["tag"])),
            unsafe_allow_html=True,
        )
        # on_click="ignore": baixar o modelo nao reexecuta a pagina inteira.
        st.download_button(
            "Baixar planilha base", data=planilha_base(),
            file_name=modelo.NOME_ARQUIVO,
            mime=comum.TIPOS_DOWNLOAD[".xlsx"],
            icon=":material/download:", on_click="ignore",
            use_container_width=True,
        )

pagina.run()
