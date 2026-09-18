# -*- coding: utf-8 -*-
"""Registro do processamento em arquivo de texto e na interface.

O log segue o formato pedido:

    PROCESSAMENTO INICIADO
    Data: 14/09/2026 10:30
    ...
    Arquivo: Moenda_A.xlsx
    Abas processadas: Indicadores, Dados
    TAGs encontradas: 18
    Linhas geradas: 288
    Status: OK
    ...
    PROCESSAMENTO FINALIZADO
    Sucesso: 2
    Erros: 1
"""

from __future__ import annotations

from datetime import datetime
from pathlib import Path

SEPARADOR = "-" * 60


class Logger:
    """Escreve no arquivo de log e repassa cada linha para a interface."""

    def __init__(self, pasta_log=None, callback=None, nome_arquivo=None):
        self.callback = callback
        self.caminho = None
        self._arquivo = None

        if pasta_log:
            pasta = Path(pasta_log)
            pasta.mkdir(parents=True, exist_ok=True)
            nome = nome_arquivo or datetime.now().strftime("log_%Y%m%d_%H%M%S.txt")
            self.caminho = pasta / nome
            self._arquivo = open(self.caminho, "w", encoding="utf-8")

    # ------------------------------------------------------------------
    def escrever(self, texto: str = "") -> None:
        if self._arquivo:
            self._arquivo.write(texto + "\n")
            self._arquivo.flush()
        if self.callback:
            try:
                self.callback(texto)
            except Exception:
                pass  # a interface nunca pode derrubar o processamento

    def linha_em_branco(self) -> None:
        self.escrever("")

    def separador(self) -> None:
        self.escrever(SEPARADOR)

    # ------------------------------------------------------------------
    def inicio(self, cfg, total_arquivos: int) -> None:
        self.escrever("PROCESSAMENTO INICIADO")
        self.escrever("Data: " + datetime.now().strftime("%d/%m/%Y %H:%M:%S"))
        self.escrever("Pasta de entrada: %s" % cfg.pasta_entrada)
        self.escrever("Pasta de saida:   %s" % cfg.pasta_saida)
        self.escrever(
            "Periodo: %s a %s  |  intervalo: %d min"
            % (
                cfg.data_inicial.strftime("%d/%m/%Y"),
                cfg.data_final.strftime("%d/%m/%Y"),
                cfg.intervalo_minutos,
            )
        )
        abas = cfg.abas_selecionadas or []
        self.escrever("Abas: %s" % (", ".join(abas) if abas else "todas com TAG valida"))
        self.escrever("Arquivos encontrados: %d" % total_arquivos)
        self.linha_em_branco()

    def arquivo_ok(self, nome: str, abas: list, tags: int, linhas: int,
                   avisos: list = None, consultas: int = 0, modo: str = "") -> None:
        self.escrever("Arquivo: %s" % nome)
        self.escrever("Abas processadas: %s" % ", ".join(abas))
        self.escrever("TAGs encontradas: %d" % tags)
        self.escrever("Linhas geradas: %d" % linhas)
        if modo:
            self.escrever("Modo de formula: %s" % modo)
        if consultas:
            self.escrever("Consultas ao PI: %d" % consultas)
        for aviso in avisos or []:
            self.escrever("Aviso: %s" % aviso)
        self.escrever("Status: OK")
        self.linha_em_branco()

    def arquivo_erro(self, nome: str, mensagem: str) -> None:
        self.escrever("Arquivo: %s" % nome)
        self.escrever("Erro: %s" % mensagem)
        self.escrever("Status: ERRO")
        self.linha_em_branco()

    def arquivo_ignorado(self, nome: str, motivo: str) -> None:
        self.escrever("Arquivo: %s" % nome)
        self.escrever("Motivo: %s" % motivo)
        self.escrever("Status: IGNORADO")
        self.linha_em_branco()

    def fim(self, sucesso: int, erros: int, ignorados: int = 0,
            cancelado: bool = False) -> None:
        self.escrever(
            "PROCESSAMENTO CANCELADO" if cancelado else "PROCESSAMENTO FINALIZADO"
        )
        self.escrever("Sucesso: %d" % sucesso)
        self.escrever("Erros: %d" % erros)
        if ignorados:
            self.escrever("Ignorados: %d" % ignorados)
        self.escrever("Termino: " + datetime.now().strftime("%d/%m/%Y %H:%M:%S"))
        if self.caminho:
            self.escrever("Log salvo em: %s" % self.caminho)

    # ------------------------------------------------------------------
    def fechar(self) -> None:
        if self._arquivo:
            self._arquivo.close()
            self._arquivo = None

    def __enter__(self) -> "Logger":
        return self

    def __exit__(self, *_) -> None:
        self.fechar()
