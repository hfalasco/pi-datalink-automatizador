# -*- coding: utf-8 -*-
"""Identidade visual da pagina.

A paleta veio da antiga janela Tkinter (gui/aparencia.py) e continua com os
mesmos nomes: verde como cor de destaque, azul no cabecalho, cinza para texto
secundario. Teal e dourado ficam disponiveis para usos pontuais (links, avisos)
sem poluir o visual minimalista.

Parte das cores tambem esta em .streamlit/config.toml - o Streamlit precisa
delas antes de o script rodar, para pintar a pagina inteira. O que o config.toml
nao alcanca (faixa do cabecalho, caixa do registro, bordas, hover do botao)
esta no CSS aqui embaixo.
"""

from __future__ import annotations

# Paleta - tons sobrios, bom contraste, sem depender de tema do sistema.
COR_FUNDO = "#EAF3EE"          # sm-green-lt
COR_PAINEL = "#ffffff"
COR_BORDA = "#CFE3D6"          # tom claro derivado do verde, para bordas discretas
COR_TEXTO = "#262626"
COR_TEXTO_FRACO = "#595959"    # sm-gray
COR_DESTAQUE = "#1E7E4A"       # sm-green
COR_DESTAQUE_ATIVO = "#0F4C2A"  # sm-green-dk
COR_CABECALHO = "#1F4E79"      # sm-blue
COR_SUBTITULO = "#b9c6d4"      # texto secundario sobre o azul do cabecalho
COR_LOG_FUNDO = "#16324A"      # variacao escura de sm-blue
COR_LOG_TEXTO = "#dfe7ef"
COR_TEAL = "#1A7A7A"           # sm-teal - reservado para acentos pontuais
COR_GOLD = "#C8A000"           # sm-gold - reservado para avisos/destaques raros

FAMILIA_FONTE = "Segoe UI"
FAMILIA_MONO = "Consolas"


def css() -> str:
    """Bloco <style> com o que o tema do config.toml nao cobre.

    O texto passa por formatacao com %, entao todo sinal de porcentagem
    literal do CSS (100%%, 50%%...) precisa vir dobrado.
    """
    return """
<style>
  html, body, [class*="st-"], button, input, textarea, select {
    font-family: "%(fonte)s", "Segoe UI", system-ui, sans-serif;
  }

  /* A faixa azul faz as vezes de cabecalho; a folga no topo evita que a
     barra de ferramentas do Streamlit corte o titulo. */
  .block-container { padding-top: 3.5rem; max-width: 1100px; }

  /* Faixa do cabecalho, equivalente ao Cabecalho.TFrame da janela antiga. */
  .faixa-cabecalho {
    background: %(cabecalho)s;
    padding: 14px 20px;
    border-radius: 6px;
    margin-bottom: 18px;
  }
  .faixa-cabecalho .titulo {
    color: #ffffff;
    font-size: 1.35rem;
    font-weight: 700;
    letter-spacing: .3px;
  }
  .faixa-cabecalho .subtitulo {
    color: %(subtitulo)s;
    font-size: .85rem;
    margin-top: 3px;
  }

  /* Bordas discretas nos blocos - o TLabelframe de antes. */
  [data-testid="stVerticalBlockBorderWrapper"] > div > [data-testid="stVerticalBlock"] {
    border-radius: 6px;
  }
  div[data-testid="stVerticalBlockBorderWrapper"] {
    border-color: %(borda)s !important;
  }

  /* Titulo de cada bloco, no verde de destaque, como o TLabelframe.Label. */
  .titulo-secao {
    color: %(destaque)s;
    font-weight: 700;
    font-size: .95rem;
    margin-bottom: .4rem;
  }
  .texto-fraco { color: %(fraco)s; font-size: .85rem; }

  /* Botoes ocupam a largura da coluna - evita o botao solto no meio da linha.
     Feito no CSS, e nao por parametro, porque o nome do parametro que faz isso
     mudou entre versoes do Streamlit. */
  .stButton button, .stDownloadButton button { width: 100%%; }
  /* Um botao com tooltip (help=) ganha um involucro que encolhe ate o texto;
     sem isto o width acima mede 100%% de uma caixa estreita. */
  .stButton [data-testid="stTooltipHoverTarget"],
  .stDownloadButton [data-testid="stTooltipHoverTarget"] {
    width: 100%% !important;
    flex: 1 1 auto !important;
  }

  /* Botao PROCESSAR: verde solido, escurecendo no hover. */
  .stButton button[kind="primary"] {
    background: %(destaque)s;
    border-color: %(destaque)s;
    font-weight: 700;
    letter-spacing: .5px;
  }
  .stButton button[kind="primary"]:hover,
  .stButton button[kind="primary"]:focus:not(:active) {
    background: %(destaque_ativo)s;
    border-color: %(destaque_ativo)s;
    color: #ffffff;
  }

  /* Registro da execucao: fundo escuro e fonte fixa, como o tk.Text de antes. */
  .st-key-registro pre, .st-key-registro code {
    background: %(log_fundo)s !important;
    color: %(log_texto)s !important;
    font-family: "%(mono)s", "Courier New", monospace !important;
    font-size: .8rem !important;
    line-height: 1.45 !important;
  }
  .st-key-registro [data-testid="stCode"] { border-radius: 6px; }
  .st-key-registro pre { max-height: 420px; overflow-y: auto; }
  /* O botao de copiar do st.code fica invisivel sobre o fundo escuro. */
  .st-key-registro [data-testid="stCopyButton"] svg { stroke: %(log_texto)s; }
</style>
""" % {
        "fonte": FAMILIA_FONTE,
        "mono": FAMILIA_MONO,
        "cabecalho": COR_CABECALHO,
        "subtitulo": COR_SUBTITULO,
        "borda": COR_BORDA,
        "destaque": COR_DESTAQUE,
        "destaque_ativo": COR_DESTAQUE_ATIVO,
        "fraco": COR_TEXTO_FRACO,
        "log_fundo": COR_LOG_FUNDO,
        "log_texto": COR_LOG_TEXTO,
    }


def cabecalho(titulo: str, subtitulo: str) -> str:
    """Faixa azul do topo, com titulo branco e subtitulo em tom claro."""
    return (
        '<div class="faixa-cabecalho">'
        '<div class="titulo">%s</div>'
        '<div class="subtitulo">%s</div>'
        "</div>" % (titulo, subtitulo)
    )


def titulo_secao(texto: str) -> str:
    """Rotulo verde no topo de cada bloco (era o TLabelframe.Label)."""
    return '<div class="titulo-secao">%s</div>' % texto


def texto_fraco(texto: str) -> str:
    """Texto de apoio em cinza (era o estilo Fraco.TLabel)."""
    return '<div class="texto-fraco">%s</div>' % texto
