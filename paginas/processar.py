# -*- coding: utf-8 -*-
"""Pagina Processar planilhas: preenche as formulas do PI DataLink em lote.

Como a pagina roda no navegador, nao existe o seletor de pastas do Windows:
o usuario envia os arquivos e recebe de volta um .zip com as copias
processadas e o log da execucao. Os arquivos enviados ficam em uma pasta
temporaria propria de cada sessao - os originais no computador do usuario
nunca sao tocados.

Todo o motor (core/, config.py, logger.py) e o mesmo de antes; esta pagina e
so a casca que substitui a antiga janela Tkinter. A marca e os modulos da
barra lateral sao desenhados por app.py, antes desta pagina.

O Streamlit roda este arquivo INTEIRO a cada clique, tecla ou data trocada.
Por isso tudo que e caro fica fora do caminho de um clique comum:

  * as planilhas sao lidas UMA vez por conjunto enviado (resumo_dos_arquivos);
    trocar data, intervalo ou opcao nao reabre arquivo nenhum;
  * a previsao de linhas e feita por conta, sem montar a lista de datas;
  * o registro e o .zip seguem as regras de comum.py (redesenho espacado,
    resultado em disco).
"""

from __future__ import annotations

import gc
import shutil
import traceback
from datetime import date, datetime
from pathlib import Path

import streamlit as st

import comum
import tema
from comum import numero
from config import (MODO_BLOCO, MODO_LINHA, PRIMEIRA_LINHA_DADOS_PADRAO,
                    Config, ConfigAba, normalizar_nome)
from core import composta, inspector, processor, writer
from core.formula import (intervalo_pi, montar_formula, montar_formula_bloco,
                          para_exibicao)
from core.timeline import (quantidade_carimbos, quantidade_formulas,
                           ultimo_carimbo)
from logger import Logger

DATA_PADRAO = date(2026, 9, 11)
INTERVALO_PADRAO = 5
SUBTITULO = ("Preenche as formulas PIAdvCalcVal em lote. "
             "Os arquivos originais nunca sao alterados.")

# --- Limites que existem para a pagina nao cair -----------------------------
# Acima disto a planilha fica pesada e o servidor pode ficar sem memoria.
LIMITE_LINHAS_AVISO = 100_000
# Limite do proprio Excel: uma planilha nao tem mais linhas do que isto.
LIMITE_LINHAS_EXCEL = 1_048_576
# Planilhas de template tem alguns KB - 50 MB por arquivo ja e muita folga.
# O servidor aceita mais (server.maxUploadSize) por causa da pagina Tirar
# formulas; aqui o limite continua o de antes.
LIMITE_ENVIO_MB = 50

area = comum.Area("processar")


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


def processar(cfg: Config, barra, status, registro: comum.Registro):
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

# --- Barra lateral: periodo e opcoes (a marca e os modulos vem de app.py) ---
# Os campos tem chave e valor inicial lembrado: assim continuam como estavam
# quando a pessoa vai para a outra pagina e volta (ver comum.lembrar).
for chave, padrao in (("data_inicial", DATA_PADRAO), ("data_final", DATA_PADRAO),
                      ("intervalo", INTERVALO_PADRAO), ("rapido", True),
                      ("calculo_manual", True), ("matriz_dinamica", True),
                      ("manter_contas", True)):
    comum.lembrar(chave, padrao)

with st.sidebar:
    st.markdown("")
    st.markdown(tema.grupo("Periodo"), unsafe_allow_html=True)
    data_inicial = st.date_input("Data inicial", key="data_inicial",
                                 format="DD/MM/YYYY")
    data_final = st.date_input("Data final", key="data_final",
                               format="DD/MM/YYYY")
    intervalo = st.number_input("Intervalo (minutos)", min_value=1,
                                max_value=100000, step=1, key="intervalo",
                                help="60 = 1 hora, 720 = 12 horas, 1440 = 1 dia")

    st.markdown("")
    st.markdown(tema.grupo("Opcoes"), unsafe_allow_html=True)
    rapido = st.checkbox(
        "Buscar tudo de uma vez", key="rapido",
        help="Muito mais rapido: uma formula PIAdvCalcDat em matriz por coluna, "
             "no lugar de uma PIAdvCalcVal por linha.",
    )
    calculo_manual = st.checkbox("Abrir sem calcular", key="calculo_manual",
                                 help="Os valores vem com F9 / PI DataLink Refresh.")
    matriz_dinamica = st.checkbox(
        "Matriz dinamica", key="matriz_dinamica",
        help="Recomendado - deixa as celulas editaveis no Excel.")
    manter_contas = st.checkbox(
        "Manter contas da planilha", key="manter_contas",
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
        max_upload_size=LIMITE_ENVIO_MB,
        help="Formatos .xls e .xlsb nao sao suportados - converta para .xlsx antes.",
    )
    enviados = enviados or []
    assinatura = area.gravar_enviados(enviados)

    entrada = area.entrada
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
caixa_log = comum.caixa_do_registro()


if pedido_teste or pedido_processar:
    saida = area.saida
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
        area.registro = relatorio_de_teste(caminho, cfg)
        area.descartar_download()
    else:
        # A saida e refeita a cada execucao: assim o .zip traz so o resultado
        # desta rodada e nada fica marcado como "ja existe na pasta de saida".
        area.descartar_download()
        shutil.rmtree(saida, ignore_errors=True)
        saida.mkdir(parents=True, exist_ok=True)
        try:
            registro = comum.Registro(caixa_log)
            resultado = processar(cfg, barra.progress(0.0), status, registro)
            area.registro = registro.linhas

            nome_zip = "processados_%s.zip" % datetime.now().strftime("%Y%m%d_%H%M%S")
            caminho_zip = comum.montar_zip(saida, area.pasta / nome_zip)
            # O .zip ja tem tudo; manter a pasta de saida so dobraria o disco.
            shutil.rmtree(saida, ignore_errors=True)
            area.download = (nome_zip, str(caminho_zip))

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
            area.registro = falha.splitlines()
            area.descartar_download()
            st.error("Falha inesperada: %s" % falha.strip().splitlines()[-1])

# Mostra o registro guardado (o proprio st.code ja traz o botao de copiar).
caixa_log.code(comum.texto_do_registro(area.registro), language="text")

comum.botao_de_download(
    area, "Baixar arquivos processados (.zip)",
    "O .zip traz as copias processadas e a pasta logs/ com o registro.")
