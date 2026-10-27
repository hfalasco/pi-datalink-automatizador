# Design System — Streamlit São Martinho

Identidade visual comum aos apps Streamlit internos (AutoFill, PI DataLink
Automatizador e os próximos). Este arquivo é a referência única: cores,
tipografia, componentes, o `tema.py` pronto para copiar e as armadilhas do
Streamlit que já custaram tempo.

> Testado com **Streamlit 1.64**. Os seletores CSS dependem do HTML que o
> Streamlit gera, que muda entre versões — veja [Armadilhas](#8-armadilhas-do-streamlit).

---

## 1. Princípios

1. **Marca primeiro, enfeite depois.** Verde é a cor de ação e destaque; azul-marinho
   dá peso aos números; amarelo é exceção (passo, filete, aviso).
2. **A barra lateral é o "painel de entrada".** Tudo o que a pessoa configura
   ou envia fica nela, em passos numerados. O miolo mostra resultado e ações.
3. **Leitura em passos.** O miolo é uma sequência de seções numeradas
   (1. Arquivos, 2. Conferência, 3. Baixar...).
4. **Funciona offline.** Nada de fonte, ícone ou script da internet: a fonte é a
   que vem dentro do Streamlit e o logotipo fica em `assets/`. Numa rede de
   fábrica a página abre igual.
5. **Apresentação separada da regra de negócio.** Cores e CSS só em `tema.py`
   e `.streamlit/config.toml`; o resto do código chama funções do tema.

---

## 2. Cores (tokens)

| Token | Hex | Uso |
|---|---|---|
| `VERDE` | `#1F7E4E` | Cor primária: botões, títulos de seção, borda dos cartões, início do degradê |
| `VERDE_ESCURO` | `#14603A` | Hover de botão, meio do degradê da barra lateral |
| `VERDE_CLARO` | `#8CD3A9` | Tracejado da área de envio, faixa clara do cabeçalho, início das barras |
| `VERDE_FUNDO` | `#EAF5EE` | Fundo de apoio: área de envio, estado vazio, trilho das barras, campos do miolo |
| `AZUL_MARINHO` | `#0B3A5B` | Fim do degradê, valores dos cartões, subtítulos |
| `AZUL` | `#0B4F8A` | Faixa do cabeçalho, cartão "azul" (informação), borda do registro |
| `AMARELO` | `#E8C21B` | Filete do cabeçalho, bolinha de passo, selo de versão, borda de aviso |
| `AMARELO_FUNDO` | `#FFF6DA` | Fundo de aviso / célula "Conferir" |
| `AMARELO_TEXTO` | `#8A6D00` | Texto sobre o amarelo claro |
| `VERMELHO` | `#C0392B` | Erro / "Não aceitável" (uso raro) |
| `TEXTO` | `#13263A` | Texto principal |
| `CINZA_TEXTO` | `#4A5A6A` | Texto de apoio, rótulos dos cartões |
| `BORDA` | `#DCE7E0` | Bordas discretas (puxadas para o verde) |
| `LOG_FUNDO` / `LOG_TEXTO` | `#0A2A42` / `#DCE7F1` | Caixa escura de registro/log |

**Degradês**

- Barra lateral (vertical): `linear-gradient(180deg, VERDE 0%, VERDE_ESCURO 55%, AZUL_MARINHO 100%)`
- Cabeçalho (diagonal): `linear-gradient(100deg, VERDE 0%, #13705A 45%, AZUL_MARINHO 100%)`
- Barra de progresso / percentual: `linear-gradient(90deg, VERDE_CLARO, VERDE)`

**Situação em tabelas** (Styler do pandas)

| Situação | Fundo | Texto |
|---|---|---|
| OK / Aceitável | `#E3F4EA` | `#14603A` |
| Erro / Não aceitável | `#FBE7E4` | `#A93226` |
| Atenção / Conferir | `#FFF6DA` | `#8A6D00` |

---

## 3. Tipografia e medidas

- **Fonte:** a padrão do Streamlit (Source Sans, embutida). Não trocar — ver
  [Armadilhas](#8-armadilhas-do-streamlit), item dos ícones.
- **Pesos:** 900 no título do cabeçalho; 800 em títulos de seção e valores;
  600–700 em rótulos.
- **Escala:** cabeçalho `1.9rem` · título de seção `1.25rem` · valor de cartão
  `1.22rem` (grande: `2rem`) · subtítulo `.92rem` · rótulo de cartão `.74rem`
  em MAIÚSCULAS com `letter-spacing: .6px`.
- **Raios:** cabeçalho `14px` · cartões `12px` · botões `10px` · avisos `8px`.
- **Sombra dos cartões:** `0 2px 6px rgba(11, 58, 91, .06)`.
- **Largura do miolo:** `max-width: 1250px`, `padding-top: 1.6rem`.

---

## 4. Configuração base

`.streamlit/config.toml` — o Streamlit lê **antes** do script rodar, então as
cores de base precisam estar aqui:

```toml
[theme]
base = "light"
primaryColor = "#1F7E4E"
backgroundColor = "#FFFFFF"
secondaryBackgroundColor = "#EAF5EE"
textColor = "#13263A"

[server]
maxUploadSize = 50

[browser]
gatherUsageStats = false
```

Início do `app.py`:

```python
import streamlit as st
import tema

st.set_page_config(
    page_title="Nome do App | São Martinho",
    page_icon=str(tema.CAMINHO_LOGO),
    layout="wide",
    initial_sidebar_state="expanded",
)
tema.aplicar()

with st.sidebar:
    st.image(str(tema.CAMINHO_LOGO), width="stretch")
    tema.marca("Nome do App", "O que ele faz, em uma linha", "1.0.0")
```

Estrutura de pastas mínima:

```
projeto/
    app.py
    tema.py                  # seção 7
    assets/
        logo_sao_martinho.png
    .streamlit/
        config.toml
```

---

## 5. Layout da página

```
┌──────────────┬──────────────────────────────────────────────────────┐
│  [logotipo]  │ ╔══════════════════════════════════════════╱╱╱╗      │
│ ┌──────────┐ │ ║ Título do app (900, branco)             ╱╱╱ ║      │
│ │Nome  v1.0│ │ ║ Subtítulo                              ╱╱╱  ║      │
│ └──────────┘ │ ╚══════════════════════════════════════════════╝      │
│ 1. Entrada   │   ▔▔▔▔▔ filete amarelo                               │
│ [ upload  ]  │ ▌RÓTULO  ▌RÓTULO  ▌RÓTULO  ▌RÓTULO   ← cartões        │
│ 2. Ajustes   │ ▌valor   ▌valor   ▌valor   ▌valor                     │
│ [campo    ]  │                                                      │
│ 3. Período   │ ① Título da seção (verde)                            │
│ [mês][ano ]  │   tabela / prévia / botões                           │
│ ──────────   │ ② Próxima seção                                       │
│ rodapé       │                                                      │
└──────────────┴──────────────────────────────────────────────────────┘
 degradê verde → azul-marinho           fundo branco
```

- **Barra lateral:** logotipo → cartão translúcido com nome/versão → grupos
  numerados (`tema.grupo("1. Planilha de análises")`) → rodapé com origem/destino.
- **Miolo:** `tema.cabecalho(...)` → cartões de resumo → seções numeradas.
- **Sem dados ainda:** `tema.estado_vazio(...)` explicando os passos, e `st.stop()`.

---

## 6. Componentes

Todos estão no `tema.py` da seção 7.

| Função | O que desenha | Quando usar |
|---|---|---|
| `aplicar()` | Injeta o CSS | Uma vez, logo após `set_page_config` |
| `marca(nome, descricao, versao)` | Cartão translúcido com nome e selo amarelo de versão | Topo da barra lateral, abaixo do logotipo |
| `grupo(texto)` | Título de grupo da barra lateral (branco, 1.12rem) | "1. Planilha", "2. Opções"... |
| `cabecalho(titulo, subtitulo)` | Faixa em degradê com 3 faixas inclinadas e filete amarelo | Topo do miolo, uma vez por página |
| `titulo(texto, subtitulo="", passo="")` | Título de seção verde, com bolinha amarela numerada opcional | Início de cada seção |
| `cartao(chave, azul=False)` | `st.container(border=True)` branco com borda esquerda de 6px | Agrupar uma seção (opcional) |
| `card(rotulo, valor, detalhe, estilo, grande)` | Cartão de número (KPI) | Resumo no topo; `estilo` = `""`, `"azul"` ou `"amarelo"` |
| `cards(itens)` | Faixa responsiva de cartões `[(rotulo, valor, detalhe, estilo), ...]` | Quando não quiser usar `st.columns` |
| `aviso(texto)` | Faixa amarela com borda à esquerda | O que não pode passar despercebido |
| `barra(nome, pct, destaque=False)` | Barra horizontal de percentual | Rankings, conformidade |
| `estado_vazio(titulo, corpo_html)` | Caixa tracejada verde-clara centralizada | Antes de a pessoa enviar dados |
| `passo(n)` | Bolinha amarela com número (HTML inline) | Dentro de listas/textos |
| `rodape(texto)` | Texto pequeno do fim da barra lateral | Origem → destino, versão das regras |
| `COR_SITUACAO` + `estilo_situacao` | Cores de célula por situação | `df.style.map(tema.estilo_situacao, subset=[...])` |

**Cores dos cartões de número:** verde (padrão) para o principal, `azul` para
contexto (período, totais lidos) e `amarelo` para o que precisa de conferência.
Use `grande=True` em um único cartão — o número mais importante da tela.

**Botões:** a ação principal é `type="primary"` (verde, `width="stretch"`); as
demais são secundárias. Download do resultado final também é `primary`.

**Registro/log:** `with st.container(border=True, key="registro"): st.code(texto, language="text")`
→ caixa escura com fonte monoespaçada (`language="text"` evita que números
saiam coloridos pelo realce de sintaxe).

---

## 7. `tema.py` pronto para copiar

```python
"""Identidade visual São Martinho para apps Streamlit.

Copie este arquivo como tema.py e o logotipo para assets/. Referência:
DESIGN_SYSTEM.md. Nada aqui baixa recurso da internet.
"""

from __future__ import annotations

import html
from pathlib import Path

import streamlit as st

VERDE = "#1F7E4E"
VERDE_ESCURO = "#14603A"
VERDE_CLARO = "#8CD3A9"
VERDE_FUNDO = "#EAF5EE"
AZUL_MARINHO = "#0B3A5B"
AZUL = "#0B4F8A"
AMARELO = "#E8C21B"
AMARELO_FUNDO = "#FFF6DA"
AMARELO_TEXTO = "#8A6D00"
VERMELHO = "#C0392B"
TEXTO = "#13263A"
CINZA_TEXTO = "#4A5A6A"
BORDA = "#DCE7E0"
LOG_FUNDO = "#0A2A42"
LOG_TEXTO = "#DCE7F1"

CAMINHO_LOGO = Path(__file__).resolve().parent / "assets" / "logo_sao_martinho.png"

COR_SITUACAO = {
    "ok": "background-color: #E3F4EA; color: #14603A; font-weight: 600",
    "erro": "background-color: #FBE7E4; color: #A93226; font-weight: 600",
    "atencao": "background-color: #FFF6DA; color: #8A6D00; font-weight: 600",
}
# Textos de célula -> situação. Acrescente os do seu projeto.
SITUACAO_POR_TEXTO = {
    "Aceitável": "ok", "OK": "ok",
    "Não aceitável": "erro", "Erro": "erro",
    "Conferir": "atencao", "Atenção": "atencao",
}

_TOKENS = {
    "verde": VERDE, "verde-escuro": VERDE_ESCURO, "verde-claro": VERDE_CLARO,
    "verde-fundo": VERDE_FUNDO, "marinho": AZUL_MARINHO, "azul": AZUL,
    "amarelo": AMARELO, "amarelo-fundo": AMARELO_FUNDO, "amarelo-texto": AMARELO_TEXTO,
    "texto": TEXTO, "cinza": CINZA_TEXTO, "borda": BORDA,
    "log-fundo": LOG_FUNDO, "log-texto": LOG_TEXTO,
}

# CSS literal (sem f-string): as cores entram por var(--sm-...).
_REGRAS = """
[data-testid="stHeader"] { background: transparent; }
[data-testid="stAppDeployButton"], footer { display: none; }
.block-container { padding-top: 1.6rem; padding-bottom: 3rem; max-width: 1250px; }

/* ---------- Barra lateral ---------- */
section[data-testid="stSidebar"] > div {
  background: linear-gradient(180deg, var(--sm-verde) 0%, var(--sm-verde-escuro) 55%, var(--sm-marinho) 100%);
}
section[data-testid="stSidebar"] * { color: #FFFFFF; }
section[data-testid="stSidebar"] hr { border-color: rgba(255, 255, 255, .25); }

section[data-testid="stSidebar"] [data-testid="stFileUploaderDropzone"] {
  background: rgba(255, 255, 255, .10); border: 1.5px dashed rgba(255, 255, 255, .55);
}
section[data-testid="stSidebar"] [data-testid="stFileUploaderDropzone"] button { background: #FFFFFF; border: none; }
section[data-testid="stSidebar"] [data-testid="stFileUploaderDropzone"] button * { color: var(--sm-verde) !important; }

/* Campos: caixa branca, texto escuro (o "*" acima pintaria tudo de branco). */
section[data-testid="stSidebar"] [data-baseweb="select"] > div,
section[data-testid="stSidebar"] [data-testid="stDateInputField"],
section[data-testid="stSidebar"] [data-testid="stTextInput"] [data-baseweb="input"],
section[data-testid="stSidebar"] [data-testid="stTextInput"] [data-baseweb="base-input"],
section[data-testid="stSidebar"] [data-testid="stNumberInput"] [data-baseweb="input"],
section[data-testid="stSidebar"] [data-testid="stTextArea"] [data-baseweb="textarea"],
section[data-testid="stSidebar"] [data-testid="stNumberInputStepDown"],
section[data-testid="stSidebar"] [data-testid="stNumberInputStepUp"] { background: #FFFFFF; }
section[data-testid="stSidebar"] [data-baseweb="select"] *,
section[data-testid="stSidebar"] [data-testid="stSelectbox"] input,
section[data-testid="stSidebar"] [data-testid="stDateInputField"] *,
section[data-testid="stSidebar"] [data-testid="stTextInput"] input,
section[data-testid="stSidebar"] [data-testid="stNumberInput"] input,
section[data-testid="stSidebar"] [data-testid="stTextArea"] textarea,
section[data-testid="stSidebar"] [data-testid="stNumberInputStepDown"] *,
section[data-testid="stSidebar"] [data-testid="stNumberInputStepUp"] *,
section[data-testid="stSidebar"] [data-testid="stFileUploaderFile"] *,
section[data-testid="stSidebar"] [data-testid="stFileChip"] * {
  color: var(--sm-texto) !important; -webkit-text-fill-color: var(--sm-texto);
}
section[data-testid="stSidebar"] input::placeholder,
section[data-testid="stSidebar"] textarea::placeholder { color: #6B7A89 !important; -webkit-text-fill-color: #6B7A89; }
section[data-testid="stSidebar"] input { caret-color: var(--sm-texto); }
div[data-baseweb="popover"] * { color: var(--sm-texto); }

/* Botões da barra lateral: brancos com texto verde. */
section[data-testid="stSidebar"] .stButton button,
section[data-testid="stSidebar"] [data-testid="stDownloadButton"] button { background: #FFFFFF; border: none; }
section[data-testid="stSidebar"] .stButton button *,
section[data-testid="stSidebar"] [data-testid="stDownloadButton"] button * { color: var(--sm-verde) !important; }

/* Links de página (st.page_link): o da página aberta vai num container key="link_ativo". */
section[data-testid="stSidebar"] [data-testid="stPageLink-NavLink"] { border-radius: 9px; padding: 7px 10px; }
section[data-testid="stSidebar"] [data-testid="stPageLink-NavLink"]:hover { background: rgba(255, 255, 255, .12); }
section[data-testid="stSidebar"] .st-key-link_ativo [data-testid="stPageLink-NavLink"] { background: #FFFFFF; }
section[data-testid="stSidebar"] .st-key-link_ativo [data-testid="stPageLink-NavLink"] * {
  color: var(--sm-verde) !important; font-weight: 800;
}

.sm-marca {
  background: rgba(255, 255, 255, .10); border: 1px solid rgba(255, 255, 255, .22);
  border-radius: 12px; padding: 12px 16px; margin: 4px 0 18px;
}
.sm-marca .nome { font-size: 1.1rem; font-weight: 800; line-height: 1.25; }
.sm-marca .sub { font-size: .82rem; opacity: .88; margin-top: 4px; }
.sm-versao {
  display: inline-block; margin-top: 8px; padding: 1px 9px; border-radius: 999px;
  background: var(--sm-amarelo); color: var(--sm-marinho) !important; font-size: .7rem; font-weight: 800;
}
.sm-grupo { font-size: 1.12rem; font-weight: 700; margin: 10px 0 4px; }
.sm-rodape { font-size: .75rem; opacity: .85; margin-top: 1rem; }

/* ---------- Cabeçalho ---------- */
.sm-header {
  position: relative; overflow: hidden; border-radius: 14px; padding: 26px 32px 24px 32px;
  background: linear-gradient(100deg, var(--sm-verde) 0%, #13705A 45%, var(--sm-marinho) 100%);
  color: #FFFFFF; margin-bottom: 1.2rem; border-bottom: 3px solid var(--sm-amarelo);
}
.sm-header h1 {
  position: relative; font-weight: 900; font-size: 1.9rem; margin: 0; padding: 0;
  color: #FFFFFF; letter-spacing: .3px; line-height: 1.25;
}
.sm-header p { position: relative; margin: 6px 0 0 0; font-size: .98rem; opacity: .92; color: #FFFFFF; }
.sm-header .faixa { position: absolute; top: 0; height: 100%; width: 46px; transform: skewX(-32deg); }
.sm-header .f1 { right: 150px; background: rgba(140, 211, 169, .35); }
.sm-header .f2 { right: 80px; background: rgba(11, 79, 138, .85); }
.sm-header .f3 { right: 10px; background: rgba(255, 255, 255, .12); }

/* ---------- Títulos de seção ---------- */
.sm-titulo {
  display: flex; align-items: center; gap: 10px;
  color: var(--sm-verde); font-weight: 800; font-size: 1.25rem; margin: 1.4rem 0 .2rem 0;
}
.sm-subtitulo { color: var(--sm-marinho); font-weight: 600; font-size: .92rem; margin-bottom: .7rem; }
.sm-passo {
  display: inline-flex; align-items: center; justify-content: center; flex: 0 0 1.5rem;
  width: 1.5rem; height: 1.5rem; border-radius: 50%; background: var(--sm-amarelo);
  color: var(--sm-marinho) !important; font-size: .8rem; font-weight: 900;
}

/* ---------- Cartões de seção: st.container(border=True, key="sm_cartao_...") ---------- */
[data-testid="stMain"] [class*="st-key-sm_cartao"] {
  background: #FFFFFF; border: 1px solid var(--sm-borda) !important;
  border-left: 6px solid var(--sm-verde) !important; border-radius: 12px !important;
  box-shadow: 0 2px 6px rgba(11, 58, 91, .06);
}
[data-testid="stMain"] [class*="st-key-sm_cartao_azul"] { border-left-color: var(--sm-azul) !important; }
[class*="st-key-sm_cartao"] .sm-titulo { margin-top: 0; }
section[data-testid="stSidebar"] [class*="st-key-sm_cartao"] {
  background: rgba(255, 255, 255, .08); border: 1px solid rgba(255, 255, 255, .30);
}

/* ---------- Cartões de número ---------- */
.sm-cards { display: grid; grid-template-columns: repeat(auto-fit, minmax(170px, 1fr)); gap: 14px; margin: 6px 0 10px; }
.sm-card {
  background: #FFFFFF; border-radius: 12px; padding: 16px 18px; height: 100%;
  border: 1px solid var(--sm-borda); border-left: 6px solid var(--sm-verde);
  box-shadow: 0 2px 6px rgba(11, 58, 91, .06);
}
.sm-card.azul { border-left-color: var(--sm-azul); }
.sm-card.amarelo { border-left-color: var(--sm-amarelo); }
.sm-card .rotulo { font-size: .74rem; text-transform: uppercase; letter-spacing: .6px; color: var(--sm-cinza); font-weight: 600; }
.sm-card .valor { font-size: 1.22rem; font-weight: 800; color: var(--sm-marinho); margin-top: 4px; line-height: 1.25; }
.sm-card .valor.grande { font-size: 2rem; color: var(--sm-verde); }
.sm-card .detalhe { font-size: .8rem; color: var(--sm-cinza); margin-top: 4px; }

/* ---------- Aviso, estado vazio, barras ---------- */
.sm-aviso {
  background: var(--sm-amarelo-fundo); border-left: 4px solid var(--sm-amarelo); border-radius: 8px;
  color: var(--sm-amarelo-texto); font-size: .88rem; line-height: 1.5; padding: 10px 13px; margin: 0 0 12px;
}
.sm-vazio {
  border: 2px dashed var(--sm-verde-claro); border-radius: 14px; padding: 34px; text-align: center;
  background: var(--sm-verde-fundo); color: var(--sm-marinho);
}
.sm-vazio h3 { color: var(--sm-verde); font-weight: 800; margin-bottom: 6px; }
.sm-vazio ol { text-align: left; display: inline-block; margin: .6rem auto 0 auto; }

.sm-barra { display: flex; align-items: center; gap: 12px; margin: 7px 0; }
.sm-barra .nome { width: 290px; font-size: .88rem; color: var(--sm-marinho); font-weight: 600; }
.sm-barra .trilho { flex: 1; background: var(--sm-verde-fundo); border-radius: 6px; height: 18px; overflow: hidden; }
.sm-barra .preenchido { height: 100%; border-radius: 6px; background: linear-gradient(90deg, var(--sm-verde-claro), var(--sm-verde)); }
.sm-barra.destaque .preenchido { background: linear-gradient(90deg, var(--sm-verde), var(--sm-azul)); }
.sm-barra.destaque .nome { color: var(--sm-verde); font-weight: 800; }
.sm-barra .pct { width: 70px; text-align: right; font-weight: 800; color: var(--sm-marinho); font-size: .92rem; }

/* ---------- Campos e botões do miolo ---------- */
[data-testid="stMain"] [data-testid="stFileUploaderDropzone"] {
  background: var(--sm-verde-fundo); border: 1.5px dashed var(--sm-verde-claro); border-radius: 12px;
}
.stButton button, [data-testid="stDownloadButton"] button { border-radius: 10px; font-weight: 700; }
[data-testid="stMain"] .stButton button[kind="primary"]:hover,
[data-testid="stMain"] [data-testid="stDownloadButton"] button[kind="primary"]:hover {
  background: var(--sm-verde-escuro); border-color: var(--sm-verde-escuro); color: #FFFFFF;
}
[data-testid="stProgress"] > div > div > div > div {
  background: linear-gradient(90deg, var(--sm-verde-claro), var(--sm-verde));
}

/* ---------- Registro / log: st.container(border=True, key="registro") ---------- */
.st-key-registro pre, .st-key-registro code {
  background: var(--sm-log-fundo) !important; color: var(--sm-log-texto) !important;
  font-size: .78rem !important; line-height: 1.5 !important;
}
.st-key-registro pre { max-height: 360px; overflow-y: auto; }
.st-key-registro [data-testid="stCopyButton"] svg { stroke: var(--sm-log-texto); }

@media (max-width: 700px) {
  .block-container { padding-left: 1rem; padding-right: 1rem; }
  .sm-header { padding: 20px 20px 18px; }
  .sm-header h1 { font-size: 1.5rem; }
  .sm-header .faixa { display: none; }
  .sm-barra .nome { width: 140px; font-size: .78rem; }
}
"""


def _md(conteudo: str) -> None:
    st.markdown(conteudo, unsafe_allow_html=True)


def aplicar() -> None:
    """Injeta o CSS. Chame uma vez, logo depois de st.set_page_config."""
    variaveis = "".join("--sm-%s: %s;" % item for item in _TOKENS.items())
    _md("<style>:root {%s}%s</style>" % (variaveis, _REGRAS))


def marca(nome: str, descricao: str, versao: str = "") -> None:
    selo = '<div class="sm-versao">v%s</div>' % html.escape(versao) if versao else ""
    _md('<div class="sm-marca"><div class="nome">%s</div><div class="sub">%s</div>%s</div>'
        % (html.escape(nome), html.escape(descricao), selo))


def grupo(texto: str) -> None:
    _md('<div class="sm-grupo">%s</div>' % html.escape(texto))


def rodape(texto_html: str) -> None:
    _md('<div class="sm-rodape">%s</div>' % texto_html)


def cabecalho(titulo: str, subtitulo: str = "") -> None:
    _md('<div class="sm-header"><span class="faixa f1"></span><span class="faixa f2"></span>'
        '<span class="faixa f3"></span><h1>%s</h1><p>%s</p></div>'
        % (html.escape(titulo), html.escape(subtitulo)))


def passo(numero) -> str:
    return '<span class="sm-passo">%s</span>' % html.escape(str(numero))


def titulo(texto: str, subtitulo: str = "", numero="") -> None:
    marcador = passo(numero) if numero != "" else ""
    _md('<div class="sm-titulo">%s<span>%s</span></div>' % (marcador, html.escape(texto)))
    if subtitulo:
        _md('<div class="sm-subtitulo">%s</div>' % html.escape(subtitulo))


def cartao(chave: str, azul: bool = False):
    """Container de seção com borda esquerda. Use: with tema.cartao("arquivos"): ..."""
    prefixo = "sm_cartao_azul_" if azul else "sm_cartao_"
    return st.container(border=True, key=prefixo + chave)


def _card_html(rotulo: str, valor: str, detalhe: str = "", estilo: str = "", grande: bool = False) -> str:
    return ('<div class="sm-card %s"><div class="rotulo">%s</div>'
            '<div class="valor %s">%s</div><div class="detalhe">%s</div></div>'
            % (estilo, html.escape(rotulo), "grande" if grande else "",
               html.escape(str(valor)), html.escape(detalhe)))


def card(rotulo: str, valor, detalhe: str = "", estilo: str = "", grande: bool = False) -> None:
    """Cartão de número. estilo: "" (verde), "azul" ou "amarelo"."""
    _md(_card_html(rotulo, valor, detalhe, estilo, grande))


def cards(itens) -> None:
    """Faixa responsiva: [(rotulo, valor, detalhe, estilo), ...] (estilo opcional)."""
    _md('<div class="sm-cards">%s</div>' % "".join(_card_html(*item) for item in itens))


def aviso(texto_html: str) -> None:
    _md('<div class="sm-aviso">%s</div>' % texto_html)


def estado_vazio(titulo_vazio: str, corpo_html: str) -> None:
    _md('<div class="sm-vazio"><h3>%s</h3>%s</div>' % (html.escape(titulo_vazio), corpo_html))


def barra(nome: str, pct: float, destaque: bool = False) -> None:
    """Barra de percentual; pct de 0 a 100."""
    largura = max(0.0, min(100.0, pct))
    _md('<div class="sm-barra %s"><div class="nome">%s</div><div class="trilho">'
        '<div class="preenchido" style="width:%.1f%%"></div></div><div class="pct">%s%%</div></div>'
        % ("destaque" if destaque else "", html.escape(nome), largura,
           ("%.1f" % pct).replace(".", ",")))


def estilo_situacao(valor) -> str:
    """Para df.style.map(tema.estilo_situacao, subset=["Situação"])."""
    return COR_SITUACAO.get(SITUACAO_POR_TEXTO.get(str(valor), ""), "")
```

### Exemplo de uso

```python
import pandas as pd
import streamlit as st
import tema

st.set_page_config(page_title="Exemplo | São Martinho", page_icon=str(tema.CAMINHO_LOGO),
                   layout="wide", initial_sidebar_state="expanded")
tema.aplicar()

with st.sidebar:
    st.image(str(tema.CAMINHO_LOGO), width="stretch")
    tema.marca("Exemplo", "Modelo de app São Martinho", "1.0.0")
    tema.grupo("1. Planilha")
    arquivo = st.file_uploader("Planilha de entrada", type=["xlsx"])
    tema.grupo("2. Período")
    c1, c2 = st.columns([3, 2])
    c1.selectbox("Mês", ["Janeiro", "Fevereiro"])
    c2.number_input("Ano", value=2026)
    tema.rodape("Origem: FO.XXX → Destino: FO.YYY")

tema.cabecalho("Título do app", "O que ele faz, em uma linha")

if arquivo is None:
    tema.estado_vazio("Carregue a planilha",
                      "Em dois passos, pela barra lateral:<ol><li>a planilha;</li><li>o período.</li></ol>")
    st.stop()

tema.cards([
    ("Período", "Jan/2026", "120 linhas lidas", "azul"),
    ("Prontos", "18 de 20", "com dados no mês", ""),
    ("Para conferir", "2", "sem dado ou fora do limite", "amarelo"),
])

with tema.cartao("conferencia"):
    tema.titulo("Conferência", "Cada linha é um item.", numero=1)
    df = pd.DataFrame({"Ponto": ["417", "631"], "Situação": ["Aceitável", "Conferir"]})
    st.dataframe(df.style.map(tema.estilo_situacao, subset=["Situação"]), hide_index=True)

with tema.cartao("baixar"):
    tema.titulo("Baixar", numero=2)
    st.download_button("Baixar tudo (.zip)", b"...", "resultado.zip", type="primary", width="stretch")
```

---

## 8. Armadilhas do Streamlit

1. **Container com borda não tem marca no HTML (1.64).** O antigo
   `[data-testid="stVerticalBlockBorderWrapper"]` não existe mais: o container
   é um `stVerticalBlock` qualquer. Para estilizar um cartão, dê uma `key` —
   ela vira a classe `st-key-<key>` — e mire nela (`tema.cartao` faz isso). A
   alternativa é `:has(> [data-testid="stElementContainer"] .classe-do-titulo)`.
2. **`section[data-testid="stSidebar"] *` pinta tudo de branco**, inclusive o
   texto digitado nos campos. Cada tipo de campo precisa da regra de texto
   escuro. O `st.selectbox` mostra o valor num `<input>` (dá para digitar e
   filtrar), e o `st.date_input` em `<span>`s dentro de `stDateInputField`
   — por isso as duas regras são diferentes.
3. **Botão de ajuda (?) também é `<button>`.** Pintar "todos os botões do
   number input" deixa o ícone de ajuda num quadrado branco. Mire em
   `stNumberInputStepDown` / `stNumberInputStepUp`.
4. **Trocar a fonte quebra os ícones.** Os ícones do Streamlit são ligaduras
   da fonte Material Symbols; uma regra de `font-family` em `[class*="st-"]`
   faz aparecer a palavra "upload" no lugar do ícone. Se trocar a fonte,
   devolva `font-family: "Material Symbols Rounded"` a
   `[data-testid="stIconMaterial"]`.
5. **Logotipo: `st.image`, não base64 no CSS/HTML.** O `st.image` serve o
   arquivo uma vez; base64 no `st.markdown` reenvia ~150 KB a cada clique.
6. **`config.toml` só é lido na partida** — mudou cor ali, reinicie o
   `streamlit run`. Alterações em módulos importados (`tema.py`) às vezes
   também só aparecem depois de reiniciar.
7. **Largura de botão:** CSS não alcança o container que encolhe o botão; use
   `width="stretch"` (ou `use_container_width=True` em versões antigas).
8. **Links de página ativos:** o Streamlit não marca o link da página aberta.
   Coloque o `st.page_link` da página atual num `st.container(key="link_ativo")`.

---

## 9. Checklist para um app novo

- [ ] `.streamlit/config.toml` com o `[theme]` da seção 4
- [ ] `tema.py` da seção 7 e `assets/logo_sao_martinho.png`
- [ ] `set_page_config(layout="wide", page_icon=logo)` + `tema.aplicar()`
- [ ] Barra lateral: logotipo → `marca` → grupos numerados → `rodape`
- [ ] Miolo: `cabecalho` → `estado_vazio` (sem dados) ou cartões + seções numeradas
- [ ] Ação principal `type="primary"`, `width="stretch"`
- [ ] Conferir em largura de celular (~420 px): cabeçalho sem faixas, cartões empilhados
