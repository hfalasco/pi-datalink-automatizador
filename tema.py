# -*- coding: utf-8 -*-
"""Identidade visual da pagina.

A paleta continua sendo a da marca - verde de destaque, azul no cabecalho,
cinza no texto de apoio - com os mesmos nomes de antes. O que mudou foi o
arranjo: a pagina agora tem barra lateral com a marca e os ajustes, e o miolo
virou uma sequencia de cartoes brancos sobre fundo claro.

Parte das cores tambem esta em .streamlit/config.toml - o Streamlit precisa
delas antes de o script rodar, para pintar a pagina inteira. O que o config.toml
nao alcanca (barra lateral, cartoes, cartelas de numeros, caixa do registro)
esta no CSS aqui embaixo.

Nada aqui baixa fonte ou icone da internet: tudo usa fonte do proprio sistema.
Numa rede de fabrica, sem saida para a internet, a pagina abre igual.
"""

from __future__ import annotations

# Paleta - tons sobrios, bom contraste, sem depender de tema do sistema.
COR_FUNDO = "#F2F5F7"          # fundo da pagina
COR_PAINEL = "#FFFFFF"         # cartoes e barra lateral
COR_BORDA = "#E3E9EE"          # bordas discretas
COR_TEXTO = "#1B2733"
COR_TEXTO_FRACO = "#64748B"    # sm-gray
COR_DESTAQUE = "#1E7E4A"       # sm-green
COR_DESTAQUE_ATIVO = "#166139"  # sm-green-dk
COR_DESTAQUE_CLARO = "#E8F3EC"  # fundo de apoio no verde da marca
COR_CABECALHO = "#1F4E79"      # sm-blue
COR_CABECALHO_ESCURO = "#16324A"
COR_SUBTITULO = "#B9C6D4"      # texto secundario sobre o azul da marca
COR_LOG_FUNDO = "#0F2436"      # variacao escura de sm-blue
COR_LOG_TEXTO = "#DCE7F1"
COR_TEAL = "#1A7A7A"           # sm-teal - reservado para acentos pontuais
COR_GOLD = "#C8A000"           # sm-gold - reservado para avisos/destaques raros
COR_GOLD_CLARO = "#FBF1CC"     # fundo da celula de TAG na planilha base
COR_GOLD_TEXTO = "#5C4A00"     # texto sobre o dourado claro

FAMILIA_FONTE = '"Segoe UI", system-ui, -apple-system, Roboto, Arial, sans-serif'
FAMILIA_MONO = '"Cascadia Mono", Consolas, "Courier New", monospace'


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
    --pi-verde-claro: %(destaque_claro)s;
    --pi-azul: %(cabecalho)s;
    --pi-azul-escuro: %(cabecalho_escuro)s;
    --pi-azul-claro: %(subtitulo)s;
    --pi-log-fundo: %(log_fundo)s;
    --pi-log-texto: %(log_texto)s;
    --pi-dourado: %(gold)s;
    --pi-dourado-claro: %(gold_claro)s;
    --pi-dourado-texto: %(gold_texto)s;
    --pi-fonte: %(fonte)s;
    --pi-mono: %(mono)s;
    --pi-raio: 12px;
    --pi-sombra: 0 1px 2px rgba(16, 40, 64, .06), 0 6px 16px rgba(16, 40, 64, .05);
  }
"""

_REGRAS = """
  html, body, [class*="st-"], button, input, textarea, select {
    font-family: var(--pi-fonte);
  }

  /* Excecao obrigatoria: os icones do Streamlit sao LIGADURAS. O elemento
     contem o texto "upload" e a fonte de icones o transforma no desenho.
     A regra acima alcanca esses elementos (a classe deles comeca com "st-")
     e, sem a fonte certa, o navegador escreve a palavra na tela - era o que
     fazia aparecer "upload" escrito na caixa de envio de arquivos. */
  [data-testid="stIconMaterial"],
  [class*="material-symbols"],
  .material-icons {
    font-family: "Material Symbols Rounded", "Material Icons" !important;
  }

  /* ----------------------------------------------------------------- */
  /* Estrutura                                                          */
  /* ----------------------------------------------------------------- */
  [data-testid="stAppViewContainer"] { background: var(--pi-fundo); }
  [data-testid="stHeader"] { background: transparent; }

  /* Enfeites do proprio Streamlit que nao dizem nada a quem usa a pagina.
     O botao "Deploy" so aparece para quem roda localmente. */
  [data-testid="stAppDeployButton"], footer { display: none; }

  /* A pagina usa layout largo por causa da barra lateral, mas o texto fica
     mais facil de ler numa coluna estreita - dai o limite de largura. */
  .block-container {
    padding-top: 2.6rem;
    padding-bottom: 3rem;
    max-width: 1060px;
  }

  /* ----------------------------------------------------------------- */
  /* Barra lateral                                                      */
  /* ----------------------------------------------------------------- */
  [data-testid="stSidebar"] {
    background: var(--pi-painel);
    border-right: 1px solid var(--pi-borda);
  }
  [data-testid="stSidebar"] [data-testid="stSidebarUserContent"],
  [data-testid="stSidebarContent"] { padding-top: .6rem; }

  .pi-marca {
    background: linear-gradient(135deg, var(--pi-azul) 0%, var(--pi-azul-escuro) 100%);
    border-radius: var(--pi-raio);
    padding: 16px 18px;
    margin-bottom: 18px;
  }
  .pi-marca .pi-marca-nome {
    color: #FFFFFF;
    font-size: 1.05rem;
    font-weight: 700;
    line-height: 1.25;
    letter-spacing: .2px;
  }
  .pi-marca .pi-marca-sub {
    color: var(--pi-azul-claro);
    font-size: .78rem;
    margin-top: 6px;
  }
  .pi-versao {
    display: inline-block;
    margin-top: 10px;
    padding: 2px 9px;
    border-radius: 999px;
    background: rgba(255, 255, 255, .14);
    color: #FFFFFF;
    font-size: .7rem;
    font-weight: 600;
    letter-spacing: .4px;
  }

  /* Titulo de cada grupo da barra lateral. */
  .pi-grupo {
    color: var(--pi-fraco);
    font-size: .72rem;
    font-weight: 700;
    letter-spacing: .9px;
    text-transform: uppercase;
    margin: 4px 0 2px;
  }

  /* Links das paginas, no topo dos Modulos. O da pagina aberta fica verde
     (app.py poe esse link num container com a chave "link_ativo"). */
  [data-testid="stSidebar"] [data-testid="stPageLink-NavLink"] {
    border-radius: 9px;
    padding: 7px 10px;
  }
  .st-key-paginas { gap: 2px; margin-bottom: 6px; }
  [data-testid="stSidebar"] [data-testid="stPageLink-NavLink"]:hover {
    background: var(--pi-fundo);
  }
  .st-key-link_ativo [data-testid="stPageLink-NavLink"],
  .st-key-link_ativo [data-testid="stPageLink-NavLink"]:hover {
    background: var(--pi-verde-claro);
  }
  .st-key-link_ativo [data-testid="stPageLink-NavLink"] span,
  .st-key-link_ativo [data-testid="stPageLink-NavLink"] p {
    color: var(--pi-verde-escuro) !important;
    font-weight: 700;
  }

  /* Modulos da barra lateral: titulo e linha de apoio de cada um. */
  .pi-modulo-nome {
    color: var(--pi-texto);
    font-size: .92rem;
    font-weight: 700;
  }
  .pi-modulo-sub {
    color: var(--pi-fraco);
    font-size: .78rem;
    line-height: 1.45;
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
    background: var(--pi-borda);
    border: 1px solid var(--pi-borda);
    border-radius: 8px;
    overflow: hidden;
    font-size: .76rem;
    margin-bottom: 8px;
  }
  .pi-modelo > div {
    background: var(--pi-painel);
    padding: 4px 7px;
    min-height: 24px;
    display: flex;
    align-items: center;
    white-space: nowrap;
    overflow: hidden;
  }
  .pi-modelo .pi-m-cab {
    background: var(--pi-fundo);
    color: var(--pi-fraco);
    font-size: .68rem;
    font-weight: 600;
    justify-content: center;
    padding: 2px 0;
    min-height: 0;
  }
  .pi-modelo .pi-m-data {
    background: var(--pi-azul);
    color: #FFFFFF;
    font-weight: 700;
    justify-content: center;
  }
  .pi-modelo .pi-m-nome {
    background: var(--pi-verde-claro);
    color: var(--pi-verde-escuro);
    font-weight: 700;
  }
  .pi-modelo .pi-m-tag {
    background: var(--pi-dourado-claro);
    color: var(--pi-dourado-texto);
    font-weight: 700;
  }
  .pi-modelo .pi-m-auto { color: var(--pi-fraco); font-style: italic; }

  /* Passo a passo numerado da barra lateral (pagina Tirar formulas). */
  .pi-passos > div {
    display: flex;
    gap: 9px;
    color: var(--pi-fraco);
    font-size: .82rem;
    line-height: 1.45;
    margin: 6px 0 8px;
  }
  .pi-passos .pi-passo {
    display: inline-flex;
    align-items: center;
    justify-content: center;
    flex: 0 0 20px;
    height: 20px;
    border-radius: 6px;
    background: var(--pi-verde-claro);
    color: var(--pi-verde);
    font-size: .72rem;
    font-weight: 700;
  }

  /* ----------------------------------------------------------------- */
  /* Titulo da pagina                                                   */
  /* ----------------------------------------------------------------- */
  .pi-titulo h1 {
    font-size: 1.6rem;
    font-weight: 700;
    color: var(--pi-texto);
    margin: 0;
    padding: 0;
    line-height: 1.25;
  }
  .pi-titulo p {
    color: var(--pi-fraco);
    font-size: .92rem;
    margin: 6px 0 18px;
  }

  /* ----------------------------------------------------------------- */
  /* Cartoes (st.container(border=True))                                */
  /* ----------------------------------------------------------------- */
  div[data-testid="stVerticalBlockBorderWrapper"] {
    background: var(--pi-painel);
    border: 1px solid var(--pi-borda) !important;
    border-radius: var(--pi-raio) !important;
    box-shadow: var(--pi-sombra);
  }
  /* A barra lateral tambem usa containers; la eles ficam sem cartao. */
  [data-testid="stSidebar"] div[data-testid="stVerticalBlockBorderWrapper"] {
    box-shadow: none;
    background: transparent;
  }

  /* Titulo de cada cartao, com o numero do passo em verde. */
  .pi-secao {
    display: flex;
    align-items: center;
    gap: 9px;
    margin-bottom: .6rem;
  }
  .pi-secao .pi-passo {
    display: inline-flex;
    align-items: center;
    justify-content: center;
    width: 22px;
    height: 22px;
    border-radius: 7px;
    background: var(--pi-verde-claro);
    color: var(--pi-verde);
    font-size: .75rem;
    font-weight: 700;
    flex: 0 0 auto;
  }
  .pi-secao .pi-nome {
    color: var(--pi-texto);
    font-size: 1rem;
    font-weight: 700;
  }
  .texto-fraco { color: var(--pi-fraco); font-size: .85rem; line-height: 1.5; }

  /* Aviso que a pessoa precisa ler antes de enviar - raro, dai o dourado. */
  .pi-aviso {
    background: var(--pi-dourado-claro);
    border-left: 3px solid var(--pi-dourado);
    border-radius: 8px;
    color: var(--pi-dourado-texto);
    font-size: .85rem;
    line-height: 1.5;
    padding: 10px 13px;
    margin: 0 0 12px;
  }

  /* ----------------------------------------------------------------- */
  /* Cartelas de numeros                                                */
  /* ----------------------------------------------------------------- */
  .pi-cartelas {
    display: grid;
    grid-template-columns: repeat(auto-fit, minmax(150px, 1fr));
    gap: 12px;
    margin: 14px 0 4px;
  }
  .pi-cartela {
    background: var(--pi-painel);
    border: 1px solid var(--pi-borda);
    border-radius: var(--pi-raio);
    box-shadow: var(--pi-sombra);
    padding: 13px 15px;
  }
  .pi-cartela .pi-rotulo {
    color: var(--pi-fraco);
    font-size: .72rem;
    font-weight: 600;
    letter-spacing: .6px;
    text-transform: uppercase;
  }
  .pi-cartela .pi-valor {
    color: var(--pi-texto);
    font-size: 1.5rem;
    font-weight: 700;
    line-height: 1.3;
    margin-top: 2px;
  }
  .pi-cartela .pi-detalhe { color: var(--pi-fraco); font-size: .76rem; }

  /* ----------------------------------------------------------------- */
  /* Campos e botoes                                                    */
  /* ----------------------------------------------------------------- */
  [data-testid="stFileUploaderDropzone"] {
    background: var(--pi-verde-claro);
    border: 1px dashed var(--pi-verde);
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
    font-weight: 700;
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

  [data-testid="stProgress"] > div > div > div > div { background: var(--pi-verde); }

  /* ----------------------------------------------------------------- */
  /* Registro da execucao                                               */
  /* ----------------------------------------------------------------- */
  .st-key-registro pre, .st-key-registro code {
    background: var(--pi-log-fundo) !important;
    color: var(--pi-log-texto) !important;
    font-family: var(--pi-mono) !important;
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
        "destaque_claro": COR_DESTAQUE_CLARO,
        "cabecalho": COR_CABECALHO,
        "cabecalho_escuro": COR_CABECALHO_ESCURO,
        "subtitulo": COR_SUBTITULO,
        "log_fundo": COR_LOG_FUNDO,
        "log_texto": COR_LOG_TEXTO,
        "gold": COR_GOLD,
        "gold_claro": COR_GOLD_CLARO,
        "gold_texto": COR_GOLD_TEXTO,
        "fonte": FAMILIA_FONTE,
        "mono": FAMILIA_MONO,
    } + _REGRAS


def marca(nome: str, descricao: str, versao: str) -> str:
    """Cartao azul do topo da barra lateral."""
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
    """Titulo e linha de apoio no topo do miolo."""
    return ('<div class="pi-titulo"><h1>%s</h1><p>%s</p></div>'
            % (titulo, subtitulo))


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
    """Faixa dourada para o que nao pode passar despercebido."""
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
