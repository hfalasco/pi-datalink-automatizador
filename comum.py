# -*- coding: utf-8 -*-
"""O que as duas paginas usam igual: pasta da sessao, envio, registro e .zip.

As duas paginas seguem o mesmo roteiro - a pessoa envia planilhas, clica num
botao, acompanha o registro e baixa o resultado - e as duas precisam das
mesmas protecoes para a pagina nao cair:

  * o registro e redesenhado algumas vezes por segundo, nao a cada linha -
    uma mensagem por linha entupia a conexao e derrubava a pagina;
  * o resultado e montado em disco e fica em disco; a memoria da sessao
    guarda so o caminho dele;
  * as pastas de sessoes abandonadas sao apagadas de tempos em tempos.

Cada pagina tem a sua Area: uma subpasta propria dentro da pasta da sessao e
as suas proprias chaves no session_state. Assim o que foi enviado numa pagina
nunca aparece na outra.
"""

from __future__ import annotations

import shutil
import tempfile
import time
import zipfile
from pathlib import Path

import streamlit as st

import tema

# Quantas linhas do registro ficam na tela. O log completo vai inteiro para
# dentro do resultado; o navegador e que nao aguenta desenhar dezenas de milhares.
LINHAS_VISIVEIS_REGISTRO = 400
# Intervalo minimo entre dois redesenhos do registro, em segundos.
INTERVALO_REDESENHO = 0.4
# Pastas de sessoes paradas ha mais tempo que isto sao apagadas.
HORAS_ATE_A_FAXINA = 6

# Tipo de cada resultado que as paginas entregam para download.
TIPOS_DOWNLOAD = {
    ".zip": "application/zip",
    ".xlsx": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    ".xlsm": "application/vnd.ms-excel.sheet.macroEnabled.12",
}
# Um .xlsx ja e um ZIP compactado: compactar de novo dentro do .zip do
# resultado so gasta tempo - com planilhas de centenas de MB, muito tempo.
JA_COMPACTADOS = (".xlsx", ".xlsm", ".zip")


def numero(valor: int) -> str:
    """12345 -> '12.345' (separador de milhar como se escreve aqui)."""
    return "{:,}".format(int(valor)).replace(",", ".")


def tamanho(bytes_: int) -> str:
    """Tamanho de arquivo como se le: '840 KB', '12,3 MB'."""
    if bytes_ < 1024 * 1024:
        return "%d KB" % max(1, round(bytes_ / 1024.0))
    return ("%.1f MB" % (bytes_ / 1024.0 / 1024.0)).replace(".", ",")


_LEMBRADOS = "_campos_lembrados"


def lembrar(chave: str, padrao) -> None:
    """Valor inicial de um campo que deve sobreviver a troca de pagina.

    O campo e criado com key=chave e SEM value= - o valor vem daqui.
    """
    st.session_state.setdefault(chave, padrao)
    st.session_state.setdefault(_LEMBRADOS, set()).add(chave)


def manter_ajustes() -> None:
    """Faz os campos lembrados sobreviverem a ida para a outra pagina.

    O Streamlit apaga o valor de um campo quando ele deixa de ser desenhado -
    e trocar de pagina e exatamente isso. Regravar o valor a cada execucao,
    em qualquer pagina (app.py chama esta funcao), e o jeito indicado na
    documentacao do Streamlit para o campo voltar como estava.
    """
    for chave in st.session_state.get(_LEMBRADOS, ()):
        if chave in st.session_state:
            st.session_state[chave] = st.session_state[chave]


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
    """Pasta temporaria propria desta sessao; cada pagina usa uma subpasta."""
    guardada = st.session_state.get("pasta_trabalho")
    if guardada and Path(guardada).is_dir():
        return Path(guardada)

    # Primeira visita - ou a faxina acima levou a pasta de uma sessao que ficou
    # horas aberta sem uso. Nos dois casos, comeca de novo.
    pasta = Path(tempfile.mkdtemp(prefix="pi_datalink_"))
    st.session_state.pasta_trabalho = str(pasta)
    return pasta


class Area:
    """Pasta e estado de UMA pagina dentro da sessao.

    area.entrada e area.saida sao as pastas dos arquivos enviados e do
    resultado; area.registro e area.download ficam no session_state com o nome
    da area na frente, para as paginas nao pisarem uma na outra.
    """

    def __init__(self, nome: str):
        self.nome = nome

    def _chave(self, item: str) -> str:
        return "%s_%s" % (self.nome, item)

    @property
    def pasta(self) -> Path:
        pasta = pasta_de_trabalho() / self.nome
        pasta.mkdir(exist_ok=True)
        return pasta

    @property
    def entrada(self) -> Path:
        return self.pasta / "entrada"

    @property
    def saida(self) -> Path:
        return self.pasta / "saida"

    def lido(self, item: str, padrao=None):
        """Valor guardado por esta pagina no session_state."""
        return st.session_state.get(self._chave(item), padrao)

    def guardar(self, item: str, valor) -> None:
        st.session_state[self._chave(item)] = valor

    # --- registro --------------------------------------------------------
    @property
    def registro(self) -> list:
        return self.lido("registro", [])

    @registro.setter
    def registro(self, linhas: list) -> None:
        self.guardar("registro", linhas)

    # --- resultado para baixar ------------------------------------------
    @property
    def download(self):
        """(nome do arquivo, caminho em disco) do resultado, ou None."""
        return self.lido("download")

    @download.setter
    def download(self, valor) -> None:
        self.guardar("download", valor)

    def descartar_download(self) -> None:
        """Apaga o resultado da rodada anterior: ele nao serve mais e ocupa disco."""
        anterior = self.download
        if anterior:
            try:
                Path(anterior[1]).unlink()
            except OSError:
                pass
        self.download = None

    # --- envio -----------------------------------------------------------
    def gravar_enviados(self, enviados: list) -> tuple:
        """Grava os uploads em entrada/ e devolve a assinatura do conjunto.

        A gravacao acontece so quando o conjunto muda - o Streamlit reexecuta o
        script inteiro a cada clique, e regravar tudo toda vez seria desperdicio.
        Se a pasta sumiu (faxina de sessao parada), grava de novo.
        """
        assinatura = tuple((arquivo.name, arquivo.size) for arquivo in enviados)
        entrada = self.entrada

        if self.lido("assinatura") != assinatura or not entrada.is_dir():
            shutil.rmtree(entrada, ignore_errors=True)
            entrada.mkdir(parents=True, exist_ok=True)
            for arquivo in enviados:
                # Path(...).name descarta qualquer caminho vindo do navegador.
                (entrada / Path(arquivo.name).name).write_bytes(arquivo.getbuffer())
            self.guardar("assinatura", assinatura)
            self.registro = []
            self.descartar_download()

        return assinatura


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
    """Mostra so o fim do registro - o log completo vai inteiro no resultado."""
    if not linhas:
        return "O registro desta execucao aparece aqui."
    if len(linhas) <= LINHAS_VISIVEIS_REGISTRO:
        return "\n".join(linhas)
    escondidas = len(linhas) - LINHAS_VISIVEIS_REGISTRO
    aviso = ("... %s linha(s) anteriores estao no log completo, dentro do .zip"
             % numero(escondidas))
    return "\n".join([aviso, ""] + linhas[-LINHAS_VISIVEIS_REGISTRO:])


def caixa_do_registro():
    """Cartao escuro do registro; devolve o espaco onde o texto e desenhado."""
    with st.container(border=True, key="registro"):
        st.markdown(tema.titulo_secao("Registro da execucao"),
                    unsafe_allow_html=True)
        return st.empty()


# ---------------------------------------------------------------------------
# Resultado
# ---------------------------------------------------------------------------

def montar_zip(saida: Path, destino: Path) -> Path:
    """Empacota tudo o que esta em saida/ num .zip EM DISCO.

    Antes o .zip era montado na memoria e ainda ficava guardado na sessao -
    duas copias do resultado inteiro em memoria, de pe ate a pessoa fechar a
    aba. Agora ele e escrito direto no disco e a sessao guarda so o caminho.
    """
    with zipfile.ZipFile(destino, "w", zipfile.ZIP_DEFLATED) as pacote:
        for caminho in sorted(saida.rglob("*")):
            if caminho.is_file():
                modo = (zipfile.ZIP_STORED
                        if caminho.suffix.lower() in JA_COMPACTADOS
                        else zipfile.ZIP_DEFLATED)
                pacote.write(caminho, caminho.relative_to(saida).as_posix(),
                             compress_type=modo)
    return destino


def botao_de_download(area: Area, rotulo: str, ajuda: str) -> None:
    """Botao do resultado da ultima rodada, se ainda existir em disco."""
    if not area.download:
        return
    nome, caminho = area.download
    caminho = Path(caminho)
    if not caminho.is_file():
        area.download = None
        return
    # O arquivo so e lido quando a pessoa clica: a pagina nao carrega o
    # resultado inteiro na memoria a cada redesenho. on_click="ignore" evita
    # que baixar reexecute a pagina.
    st.download_button(
        rotulo, data=caminho.read_bytes, file_name=nome,
        mime=TIPOS_DOWNLOAD.get(caminho.suffix.lower(), "application/octet-stream"),
        type="primary", icon=":material/download:", on_click="ignore",
        use_container_width=True, help=ajuda,
    )
