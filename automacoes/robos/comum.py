"""Configuração e rotinas de planilha compartilhadas pelos robôs.

A configuração fica em config_local.ini, na raiz do repositório (fora do Git; modelo em
config_local.exemplo.ini). Senhas NUNCA ficam no arquivo: vêm de variáveis de ambiente.
"""
from __future__ import annotations

import configparser
import datetime as dt
import logging
import os
import shutil
import time
from pathlib import Path

from automacoes import cnj

RAIZ = Path(__file__).resolve().parents[2]
CONFIG_PADRAO = RAIZ / "config_local.ini"
ABAS_TRIBUNAIS = ["TJPB", "TJRN", "TJCE", "TJPE", "TJSP", "TJPR"]
LINHA_INICIAL = 4
COL_PROC_1, COL_DATA_1 = 2, 4    # B e D: 1ª instância
COL_PROC_2, COL_DATA_2 = 11, 13  # K e M: 2ª instância

log = logging.getLogger("robos")


def configurar_log() -> None:
    logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(message)s")


def carregar_config(caminho: Path | str | None = None) -> configparser.ConfigParser:
    caminho = Path(caminho or os.environ.get("AUTOMACOES_CONFIG", CONFIG_PADRAO))
    if not caminho.exists():
        raise FileNotFoundError(f"Configuração não encontrada: {caminho}. Copie config_local.exemplo.ini.")
    cfg = configparser.ConfigParser(interpolation=None)
    cfg.read(caminho, encoding="utf-8")
    return cfg


def segredo(nome: str) -> str:
    valor = os.environ.get(nome, "").strip()
    if not valor:
        raise RuntimeError(f"Defina a variável de ambiente {nome} (no Windows: setx {nome} \"valor\").")
    return valor


def aguardar_arquivo(caminho: Path, tentativas: int = 10, intervalo: int = 30) -> None:
    """O Google Drive pode demorar a montar a unidade G: após o logon."""
    for i in range(tentativas):
        if caminho.exists():
            return
        log.warning("Planilha ainda indisponível (%s/%s): %s", i + 1, tentativas, caminho)
        time.sleep(intervalo)
    raise FileNotFoundError(f"Planilha não encontrada: {caminho}")


def verificar_livre(caminho: Path) -> None:
    """No Windows, o Excel bloqueia o arquivo aberto; abrir para escrita falha."""
    try:
        with caminho.open("r+b"):
            pass
    except PermissionError as e:
        raise RuntimeError(f"A planilha está aberta no Excel. Feche-a: {caminho}") from e


def backup(caminho: Path, prefixo: str, manter: int = 10) -> Path:
    pasta = caminho.parent / "backups"
    pasta.mkdir(exist_ok=True)
    destino = pasta / f"{prefixo}_{dt.datetime.now():%Y%m%d_%H%M%S}_{caminho.name}"
    shutil.copy2(caminho, destino)
    antigos = sorted(pasta.glob(f"{prefixo}_*_{caminho.name}"))[:-manter]
    for velho in antigos:
        velho.unlink()
    log.info("Backup: %s (%s antigos removidos)", destino.name, len(antigos))
    return destino


def mapear_processos(wb) -> dict[str, list[dict]]:
    """CNJ (20 dígitos) -> ocorrências nas abas dos tribunais."""
    mapa: dict[str, list[dict]] = {}
    for aba in ABAS_TRIBUNAIS:
        if aba not in wb.sheetnames:
            continue
        ws = wb[aba]
        for r in range(LINHA_INICIAL, ws.max_row + 1):
            for col_proc, col_data, inst in ((COL_PROC_1, COL_DATA_1, "1ª"), (COL_PROC_2, COL_DATA_2, "2ª")):
                d = cnj.digitos(str(ws.cell(row=r, column=col_proc).value or ""))
                if len(d) != 20:
                    continue
                if motivo := cnj.diagnosticar(d):
                    log.warning("%s linha %s: CNJ %s com problema (%s)", aba, r, cnj.formatar(d), motivo)
                mapa.setdefault(d, []).append({"aba": aba, "linha": r, "col_data": col_data, "instancia": inst})
    return mapa
