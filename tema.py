# -*- coding: utf-8 -*-
"""Identidade visual da pagina - a mesma do AutoFill (template Sao Martinho).

O arranjo segue o do AutoFill: barra lateral em degrade verde -> azul-marinho
com o logotipo no topo e o texto em branco; no miolo, o cabecalho com as
faixas inclinadas e o filete amarelo, titulos de passo em verde e cartoes
brancos com a borda esquerda grossa na cor da marca.

Parte das cores tambem esta em .streamlit/config.toml - o Streamlit precisa
delas antes de o script rodar, para pintar a pagina inteira. O que o config.toml
nao alcanca (barra lateral, cabecalho, cartoes, cartelas de numeros, caixa do
registro) esta no CSS aqui embaixo.

Nada aqui baixa fonte ou icone da internet: a fonte e a que vem dentro do
proprio Streamlit e o logotipo esta em assets/. Numa rede de fabrica, sem
saida para a internet, a pagina abre igual.
"""

from __future__ import annotations

from pathlib import Path

# Paleta - a do AutoFill (autofill/view/tema.py), com os nomes de sempre.
COR_FUNDO = "#FFFFFF"          # fundo da pagina
COR_PAINEL = "#FFFFFF"         # cartoes
COR_BORDA = "#DCE7E0"          # bordas discretas, puxadas para o verde
COR_TEXTO = "#13263A"
COR_TEXTO_FRACO = "#4A5A6A"
COR_DESTAQUE = "#1F7E4E"       # verde da marca
COR_DESTAQUE_ATIVO = "#14603A"  # verde escuro
COR_DESTAQUE_MEDIO = "#8CD3A9"  # verde claro das faixas e do tracejado
COR_DESTAQUE_CLARO = "#EAF5EE"  # fundo de apoio no verde da marca
COR_CABECALHO = "#0B3A5B"      # azul-marinho: fim do degrade e valores
COR_AZUL = "#0B4F8A"           # azul das faixas e dos cartoes azuis
COR_LOG_FUNDO = "#0A2A42"      # variacao escura do azul-marinho
COR_LOG_TEXTO = "#DCE7F1"
COR_GOLD = "#E8C21B"           # amarelo do filete e dos numeros de passo
COR_GOLD_CLARO = "#FFF6DA"     # fundo dos avisos
COR_GOLD_TEXTO = "#8A6D00"     # texto sobre o amarelo claro

# A miniatura da planilha base imita o arquivo baixado, entao usa as cores de
# core/modelo.py - e nao a paleta da pagina.
COR_PLANILHA_DATA = "#1F4E79"
COR_PLANILHA_NOME = "#E8F3EC"
COR_PLANILHA_NOME_TEXTO = "#166139"
COR_PLANILHA_TAG = "#FBF1CC"
COR_PLANILHA_TAG_TEXTO = "#5C4A00"

# Logotipo do topo da barra lateral e icone da aba do navegador (app.py).
CAMINHO_LOGO = Path(__file__).resolve().parent / "assets" / "logo_sao_martinho.png"


# As cores entram so aqui, por substituicao. O resto do CSS e texto literal e
# usa var(--...), entao "100%" e "50%" podem ser escritos uma vez so.
_VARIAVEIS = """
<style>
  :root {
    --pi-fundo: %(fundo)s;
    --pi-painel: %(painel)s;
    --pi-borda: %(borda)s;
    --pi-texto: %(texto)s;
    --pi-fraco: %(fraco)s;
    --pi-verde: %(destaque)s;
    --pi-verde-escuro: %(destaque_ativo)s;
    --pi-verde-medio: %(destaque_medio)s;
    --pi-verde-claro: %(destaque_claro)s;
    --pi-marinho: %(cabecalho)s;
    --pi-azul: %(azul)s;
    --pi-log-fundo: %(log_fundo)s;
    --pi-log-texto: %(log_texto)s;
    --pi-amarelo: %(gold)s;
    --pi-amarelo-claro: %(gold_claro)s;
    --pi-amarelo-texto: %(gold_texto)s;
    --pi-planilha-data: %(planilha_data)s;
    --pi-planilha-nome: %(planilha_nome)s;
    --pi-planilha-nome-texto: %(planilha_nome_texto)s;
    --pi-planilha-tag: %(planilha_tag)s;
    --pi-planilha-tag-texto: %(planilha_tag_texto)s;
    --pi-raio: 12px;
    --pi-sombra: 0 2px 6px rgba(11, 58, 91, .06);
  }
"""

_REGRAS = """
  /* ----------------------------------------------------------------- */
  /* Estrutura                                                          */
  /* ----------------------------------------------------------------- */
  [data-testid="stAppViewContainer"] { background: var(--pi-fundo); }
  [data-testid="stHeader"] { background: transparent; }

  /* Enfeites do proprio Streamlit que nao dizem nada a quem usa a pagina.
     O botao "Deploy" so aparece para quem roda localmente. */
  [data-testid="stAppDeployButton"], footer { display: none; }

  .block-container {
    padding-top: 1.6rem;
    padding-bottom: 3rem;
    max-width: 1250px;
  }

  /* ----------------------------------------------------------------- */
  /* Barra lateral - degrade da marca, texto branco                     */
  /* ----------------------------------------------------------------- */
  section[data-testid="stSidebar"] > div {
    background: linear-gradient(180deg, var(--pi-verde) 0%, var(--pi-verde-escuro) 55%, var(--pi-marinho) 100%);
  }
  section[data-testid="stSidebar"] * { color: #FFFFFF; }
  section[data-testid="stSidebar"] hr { border-color: rgba(255, 255, 255, .25); }

  /* Campos da barra lateral: caixa branca com texto escuro, como no AutoFill. */
  section[data-testid="stSidebar"] [data-baseweb="select"] > div,
  section[data-testid="stSidebar"] [data-testid="stDateInputField"],
  section[data-testid="stSidebar"] [data-testid="stTextInput"] [data-baseweb="input"],
  section[data-testid="stSidebar"] [data-testid="stTextInput"] [data-baseweb="base-input"],
  section[data-testid="stSidebar"] [data-testid="stNumberInput"] [data-baseweb="input"],
  section[data-testid="stSidebar"] [data-testid="stNumberInputStepDown"],
  section[data-testid="stSidebar"] [data-testid="stNumberInputStepUp"] { background: #FFFFFF; }
  section[data-testid="stSidebar"] [data-baseweb="select"] *,
  section[data-testid="stSidebar"] [data-testid="stDateInputField"] *,
  section[data-testid="stSidebar"] [data-testid="stTextInput"] input,
  section[data-testid="stSidebar"] [data-testid="stNumberInput"] input,
  section[data-testid="stSidebar"] [data-testid="stNumberInputStepDown"] *,
  section[data-testid="stSidebar"] [data-testid="stNumberInputStepUp"] * {
    color: var(--pi-texto) !important;
    -webkit-text-fill-color: var(--pi-texto);
  }
  div[data-baseweb="popover"] * { color: var(--pi-texto); }

  /* Caixa de marcar: sobre o degrade, o quadrado marcado ganha borda branca. */
  section[data-testid="stSidebar"] [data-testid="stCheckbox"] label > span:first-child {
    border-color: #FFFFFF;
  }

  /* Botoes da barra lateral: brancos com texto verde. */
  section[data-testid="stSidebar"] [data-testid="stDownloadButton"] button,
  section[data-testid="stSidebar"] .stButton button {
    background: #FFFFFF;
    border: none;
  }
  section[data-testid="stSidebar"] [data-testid="stDownloadButton"] button *,
  section[data-testid="stSidebar"] .stButton button * { color: var(--pi-verde) !important; }
  section[data-testid="stSidebar"] [data-testid="stDownloadButton"] button:hover {
    background: var(--pi-verde-claro);
  }

  /* Nome do programa, logo abaixo do logotipo. */
  .pi-marca {
    background: rgba(255, 255, 255, .10);
    border: 1px solid rgba(255, 255, 255, .22);
    border-radius: var(--pi-raio);
    padding: 12px 16px;
    margin: 4px 0 18px;
  }
  .pi-marca .pi-marca-nome {
    font-size: 1.1rem;
    font-weight: 800;
    line-height: 1.25;
    letter-spacing: .2px;
  }
  .pi-marca .pi-marca-sub {
    font-size: .82rem;
    opacity: .88;
    margin-top: 4px;
  }
  .pi-versao {
    display: inline-block;
    margin-top: 8px;
    padding: 1px 9px;
    border-radius: 999px;
    background: var(--pi-amarelo);
    color: var(--pi-marinho) !important;
    font-size: .7rem;
    font-weight: 800;
    letter-spacing: .4px;
  }

  /* Titulo de cada grupo da barra lateral - como os "### 1. ..." do AutoFill. */
  .pi-grupo {
    font-size: 1.12rem;
    font-weight: 700;
    margin: 6px 0 4px;
  }

  /* Links das paginas, no topo dos Modulos. O da pagina aberta fica branco
     com o texto verde (app.py poe esse link num container com a chave
     "link_ativo"). */
  section[data-testid="stSidebar"] [data-testid="stPageLink-NavLink"] {
    border-radius: 9px;
    padding: 7px 10px;
  }
  .st-key-paginas { gap: 2px; margin-bottom: 8px; }
  section[data-testid="stSidebar"] [data-testid="stPageLink-NavLink"]:hover {
    background: rgba(255, 255, 255, .12);
  }
  .st-key-link_ativo [data-testid="stPageLink-NavLink"],
  .st-key-link_ativo [data-testid="stPageLink-NavLink"]:hover {
    background: #FFFFFF;
  }
  section[data-testid="stSidebar"] .st-key-link_ativo [data-testid="stPageLink-NavLink"] * {
    color: var(--pi-verde) !important;
    font-weight: 800;
  }

  /* O modulo da barra lateral (st.container(border=True) que comeca com
     tema.modulo) fica translucido, sem o cartao branco do miolo. */
  section[data-testid="stSidebar"] [data-testid="stVerticalBlock"]:has(> [data-testid="stElementContainer"] .pi-modulo-nome) {
    background: rgba(255, 255, 255, .08);
    border: 1px solid rgba(255, 255, 255, .30);
  }

  /* Modulos da barra lateral: titulo e linha de apoio de cada um. */
  .pi-modulo-nome { font-size: .95rem; font-weight: 700; }
  .pi-modulo-sub {
    font-size: .8rem;
    line-height: 1.45;
    opacity: .88;
    margin: 2px 0 10px;
  }

  /* Miniatura da planilha base: uma grade com cabecalho de colunas (A, B) e
     de linhas (1, 2, 3), nas mesmas cores das celulas do arquivo baixado.
     O fundo da grade e a cor da borda; o gap de 1px vira a linha entre as
     celulas. Feita com div, e nao com table, para o estilo de tabela do
     proprio Streamlit nao interferir. */
  .pi-modelo {
    display: grid;
    grid-template-columns: 22px 1fr 1.35fr;
    gap: 1px;
    background: #C9D6CF;
    border: 1px solid #C9D6CF;
    border-radius: 8px;
    overflow: hidden;
    font-size: .76rem;
    margin-bottom: 8px;
  }
  .pi-modelo > div {
    background: #FFFFFF;
    padding: 4px 7px;
    min-height: 24px;
    display: flex;
    align-items: center;
    white-space: nowrap;
    overflow: hidden;
  }
  section[data-testid="stSidebar"] .pi-modelo > div { color: var(--pi-texto); }
  section[data-testid="stSidebar"] .pi-modelo .pi-m-cab {
    background: #F2F5F7;
    color: var(--pi-fraco);
    font-size: .68rem;
    font-weight: 600;
    justify-content: center;
    padding: 2px 0;
    min-height: 0;
  }
  section[data-testid="stSidebar"] .pi-modelo .pi-m-data {
    background: var(--pi-planilha-data);
    color: #FFFFFF;
    font-weight: 700;
    justify-content: center;
  }
  section[data-testid="stSidebar"] .pi-modelo .pi-m-nome {
    background: var(--pi-planilha-nome);
    color: var(--pi-planilha-nome-texto);
    font-weight: 700;
  }
  section[data-testid="stSidebar"] .pi-modelo .pi-m-tag {
    background: var(--pi-planilha-tag);
    color: var(--pi-planilha-tag-texto);
    font-weight: 700;
  }
  section[data-testid="stSidebar"] .pi-modelo .pi-m-auto {
    color: var(--pi-fraco);
    font-style: italic;
  }

  /* Passo a passo numerado da barra lateral (pagina Tirar formulas). */
  .pi-passos > div {
    display: flex;
    gap: 9px;
    font-size: .85rem;
    line-height: 1.45;
    margin: 6px 0 8px;
  }

  /* Numero de passo: a bolinha amarela do AutoFill. */
  .pi-passo {
    display: inline-flex;
    align-items: center;
    justify-content: center;
    flex: 0 0 1.5rem;
    width: 1.5rem;
    height: 1.5rem;
    border-radius: 50%;
    background: var(--pi-amarelo);
    color: var(--pi-marinho) !important;
    font-size: .8rem;
    font-weight: 900;
  }

  /* ----------------------------------------------------------------- */
  /* Cabecalho da pagina - degrade com as faixas inclinadas             */
  /* ----------------------------------------------------------------- */
  .pi-titulo {
    position: relative;
    overflow: hidden;
    border-radius: 14px;
    padding: 26px 32px 24px 32px;
    background: linear-gradient(100deg, var(--pi-verde) 0%, #13705A 45%, var(--pi-marinho) 100%);
    color: #FFFFFF;
    margin-bottom: 1.2rem;
    border-bottom: 3px solid var(--pi-amarelo);
  }
  .pi-titulo h1 {
    position: relative;
    font-weight: 900;
    font-size: 1.9rem;
    margin: 0;
    padding: 0;
    color: #FFFFFF;
    letter-spacing: .3px;
    line-height: 1.25;
  }
  .pi-titulo p {
    position: relative;
    margin: 6px 0 0 0;
    font-size: .98rem;
    opacity: .92;
    color: #FFFFFF;
  }
  .pi-titulo .faixa { position: absolute; top: 0; height: 100%; width: 46px; transform: skewX(-32deg); }
  .pi-titulo .f1 { right: 150px; background: rgba(140, 211, 169, .35); }
  .pi-titulo .f2 { right: 80px; background: rgba(11, 79, 138, .85); }
  .pi-titulo .f3 { right: 10px; background: rgba(255, 255, 255, .12); }

  /* ----------------------------------------------------------------- */
  /* Cartoes - borda esquerda verde                                     */
  /* ----------------------------------------------------------------- */
  /* O Streamlit nao marca no HTML qual container tem borda. Cartao, aqui,
     e o st.container(border=True) que comeca com tema.titulo_secao. */
  [data-testid="stMain"] [data-testid="stVerticalBlock"]:has(> [data-testid="stElementContainer"] .pi-secao) {
    background: var(--pi-painel);
    border: 1px solid var(--pi-borda) !important;
    border-left: 6px solid var(--pi-verde) !important;
    border-radius: var(--pi-raio) !important;
    box-shadow: var(--pi-sombra);
  }

  /* Titulo de cada cartao: bolinha do passo e o nome em verde. */
  .pi-secao {
    display: flex;
    align-items: center;
    gap: 10px;
    margin-bottom: .6rem;
  }
  .pi-secao .pi-nome {
    color: var(--pi-verde);
    font-size: 1.25rem;
    font-weight: 800;
  }
  .texto-fraco { color: var(--pi-fraco); font-size: .85rem; line-height: 1.5; }
  section[data-testid="stSidebar"] .texto-fraco { opacity: .88; }

  /* Aviso que a pessoa precisa ler antes de enviar - raro, dai o amarelo. */
  .pi-aviso {
    background: var(--pi-amarelo-claro);
    border-left: 4px solid var(--pi-amarelo);
    border-radius: 8px;
    color: var(--pi-amarelo-texto);
    font-size: .85rem;
    line-height: 1.5;
    padding: 10px 13px;
    margin: 0 0 12px;
  }

  /* ----------------------------------------------------------------- */
  /* Cartelas de numeros - os cards do AutoFill                         */
  /* ----------------------------------------------------------------- */
  .pi-cartelas {
    display: grid;
    grid-template-columns: repeat(auto-fit, minmax(170px, 1fr));
    gap: 14px;
    margin: 14px 0 6px;
  }
  .pi-cartela {
    background: var(--pi-painel);
    border: 1px solid var(--pi-borda);
    border-left: 6px solid var(--pi-verde);
    border-radius: var(--pi-raio);
    box-shadow: var(--pi-sombra);
    padding: 16px 18px;
  }
  .pi-cartela:first-child { border-left-color: var(--pi-azul); }
  .pi-cartela .pi-rotulo {
    color: var(--pi-fraco);
    font-size: .74rem;
    font-weight: 600;
    letter-spacing: .6px;
    text-transform: uppercase;
  }
  .pi-cartela .pi-valor {
    color: var(--pi-marinho);
    font-size: 1.6rem;
    font-weight: 800;
    line-height: 1.25;
    margin-top: 4px;
  }
  .pi-cartela .pi-detalhe { color: var(--pi-fraco); font-size: .8rem; margin-top: 4px; }

  /* ----------------------------------------------------------------- */
  /* Campos e botoes                                                    */
  /* ----------------------------------------------------------------- */
  [data-testid="stFileUploaderDropzone"] {
    background: var(--pi-verde-claro);
    border: 1.5px dashed var(--pi-verde-medio);
    border-radius: var(--pi-raio);
  }

  /* A largura dos botoes nao fica aqui: quem encolhe e o stElementContainer,
     varios niveis acima, e o CSS no botao nao o alcanca. Quem resolve e
     use_container_width=True na chamada, em app.py. */
  .stButton button {
    border-radius: 10px;
    font-weight: 600;
  }
  .stButton button[kind="primary"] {
    background: var(--pi-verde);
    border-color: var(--pi-verde);
    color: #FFFFFF;
    font-weight: 800;
    letter-spacing: .5px;
  }
  .stButton button[kind="primary"]:hover,
  .stButton button[kind="primary"]:focus:not(:active) {
    background: var(--pi-verde-escuro);
    border-color: var(--pi-verde-escuro);
    color: #FFFFFF;
  }
  .stButton button[kind="secondary"] {
    border-color: var(--pi-borda);
    color: var(--pi-texto);
  }
  .stButton button[kind="secondary"]:hover:not(:disabled) {
    border-color: var(--pi-verde);
    color: var(--pi-verde);
  }
  [data-testid="stDownloadButton"] button { border-radius: 10px; font-weight: 700; }

  [data-testid="stProgress"] > div > div > div > div {
    background: linear-gradient(90deg, var(--pi-verde-medio), var(--pi-verde));
  }

  /* ----------------------------------------------------------------- */
  /* Registro da execucao                                               */
  /* ----------------------------------------------------------------- */
  [data-testid="stMain"] [data-testid="stVerticalBlock"].st-key-registro:has(> [data-testid="stElementContainer"] .pi-secao) {
    border-left-color: var(--pi-azul) !important;
  }
  .st-key-registro pre, .st-key-registro code {
    background: var(--pi-log-fundo) !important;
    color: var(--pi-log-texto) !important;
    font-size: .78rem !important;
    line-height: 1.5 !important;
  }
  .st-key-registro [data-testid="stCode"] { border-radius: 10px; }
  .st-key-registro pre { max-height: 360px; overflow-y: auto; }
  /* O botao de copiar do st.code fica invisivel sobre o fundo escuro. */
  .st-key-registro [data-testid="stCopyButton"] svg { stroke: var(--pi-log-texto); }

  /* Em telas estreitas o cartao encosta na borda; um respiro resolve. */
  @media (max-width: 640px) {
    .block-container { padding-left: 1rem; padding-right: 1rem; }
    .pi-titulo { padding: 20px 20px 18px; }
    .pi-titulo h1 { font-size: 1.5rem; }
    .pi-titulo .faixa { display: none; }
  }
</style>
"""


def css() -> str:
    """Bloco <style> com o que o tema do config.toml nao cobre."""
    return _VARIAVEIS % {
        "fundo": COR_FUNDO,
        "painel": COR_PAINEL,
        "borda": COR_BORDA,
        "texto": COR_TEXTO,
        "fraco": COR_TEXTO_FRACO,
        "destaque": COR_DESTAQUE,
        "destaque_ativo": COR_DESTAQUE_ATIVO,
        "destaque_medio": COR_DESTAQUE_MEDIO,
        "destaque_claro": COR_DESTAQUE_CLARO,
        "cabecalho": COR_CABECALHO,
        "azul": COR_AZUL,
        "log_fundo": COR_LOG_FUNDO,
        "log_texto": COR_LOG_TEXTO,
        "gold": COR_GOLD,
        "gold_claro": COR_GOLD_CLARO,
        "gold_texto": COR_GOLD_TEXTO,
        "planilha_data": COR_PLANILHA_DATA,
        "planilha_nome": COR_PLANILHA_NOME,
        "planilha_nome_texto": COR_PLANILHA_NOME_TEXTO,
        "planilha_tag": COR_PLANILHA_TAG,
        "planilha_tag_texto": COR_PLANILHA_TAG_TEXTO,
    } + _REGRAS


def marca(nome: str, descricao: str, versao: str) -> str:
    """Cartao do programa, logo abaixo do logotipo da barra lateral."""
    return (
        '<div class="pi-marca">'
        '<div class="pi-marca-nome">%s</div>'
        '<div class="pi-marca-sub">%s</div>'
        '<div class="pi-versao">v%s</div>'
        "</div>" % (nome, descricao, versao)
    )


def modulo(nome: str, descricao: str) -> str:
    """Titulo e linha de apoio de um modulo da barra lateral."""
    return ('<div class="pi-modulo-nome">%s</div>'
            '<div class="pi-modulo-sub">%s</div>' % (nome, descricao))


def miniatura_planilha(data: str, nome: str, tag: str) -> str:
    """Grade A/B x 1/2/3 mostrando onde vai cada coisa na planilha base.

    As celulas destacadas usam as mesmas cores do arquivo baixado; a linha 3
    mostra, em cinza, o que o programa preenche sozinho.
    """
    return (
        '<div class="pi-modelo">'
        '<div class="pi-m-cab"></div>'
        '<div class="pi-m-cab">A</div><div class="pi-m-cab">B</div>'
        '<div class="pi-m-cab">1</div>'
        '<div></div><div class="pi-m-nome">%s</div>'
        '<div class="pi-m-cab">2</div>'
        '<div class="pi-m-data">%s</div><div class="pi-m-tag">%s</div>'
        '<div class="pi-m-cab">3</div>'
        '<div class="pi-m-auto">datas...</div>'
        '<div class="pi-m-auto">valores...</div>'
        "</div>" % (nome, data, tag)
    )


def titulo_pagina(titulo: str, subtitulo: str) -> str:
    """Cabecalho em degrade do topo do miolo, com as faixas da marca."""
    return ('<div class="pi-titulo">'
            '<span class="faixa f1"></span><span class="faixa f2"></span>'
            '<span class="faixa f3"></span>'
            '<h1>%s</h1><p>%s</p></div>' % (titulo, subtitulo))


def grupo(texto: str) -> str:
    """Rotulo de um grupo de ajustes da barra lateral."""
    return '<div class="pi-grupo">%s</div>' % texto


def titulo_secao(texto: str, passo: str = "") -> str:
    """Titulo de um cartao, com o numero do passo quando houver."""
    marcador = '<span class="pi-passo">%s</span>' % passo if passo else ""
    return ('<div class="pi-secao">%s<span class="pi-nome">%s</span></div>'
            % (marcador, texto))


def texto_fraco(texto: str) -> str:
    """Texto de apoio em cinza."""
    return '<div class="texto-fraco">%s</div>' % texto


def aviso(texto: str) -> str:
    """Faixa amarela para o que nao pode passar despercebido."""
    return '<div class="pi-aviso">%s</div>' % texto


def passos(itens) -> str:
    """Lista numerada curta, no estilo dos numeros de passo dos cartoes."""
    corpo = "".join(
        '<div><span class="pi-passo">%d</span><span>%s</span></div>'
        % (numero, texto) for numero, texto in enumerate(itens, start=1)
    )
    return '<div class="pi-passos">%s</div>' % corpo


def cartelas(itens) -> str:
    """Faixa de numeros do topo: [(rotulo, valor, detalhe), ...]."""
    corpo = "".join(
        '<div class="pi-cartela">'
        '<div class="pi-rotulo">%s</div>'
        '<div class="pi-valor">%s</div>'
        '<div class="pi-detalhe">%s</div>'
        "</div>" % (rotulo, valor, detalhe)
        for rotulo, valor, detalhe in itens
    )
    return '<div class="pi-cartelas">%s</div>' % corpo
