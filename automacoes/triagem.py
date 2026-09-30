"""Pré-triagem determinística de publicações, para gastar IA só no que exige leitura.

Entrada: CSV (vírgula ou ponto e vírgula; exportação do Sheets ou do robô) com as colunas
    data_disponibilizacao (AAAA-MM-DD ou DD/MM/AAAA), processo, texto

O Python resolve sozinho: validade do CNJ, tribunal, comarca, tipo de ato por palavra-chave,
prazo expresso ("prazo de 15 dias", "em 5 (cinco) dias"), vencimento, termo legal e
criticidade. O trecho com o comando judicial é recortado da própria publicação.

Vão para revisão (IA ou advogado) só as linhas com: prazo não encontrado, mais de um
prazo, prazo em horas ou meses, menção a dobra ou a audiência, ou CNJ inválido.
Essas linhas saem também em um arquivo enxuto (--para-ia) para colar no agente.

Uso:
    python -m automacoes.triagem triagem_diaria.csv -o prazos_processuais.csv --para-ia revisar.txt
"""
from __future__ import annotations

import argparse
import csv
import datetime as dt
import re
from pathlib import Path

from automacoes import cnj
from automacoes.prazos import CSV_PADRAO, Calendario, calcular_prazo

RE_PRAZO = re.compile(
    r"(?:prazo\s+(?:comum\s+|sucessivo\s+|legal\s+)?de|no\s+prazo\s+de|em|dentro\s+de)\s+"
    r"(\d{1,3})\s*(?:\([^)]{1,30}\)\s*)?(dias?\s+úteis|dias?|horas?|meses|mês)",
    re.IGNORECASE,
)
TIPOS = [  # ordem importa: primeiro que casar
    ("Sentença", re.compile(r"\bjulgo\s+(?:procedente|improcedente|parcialmente)|\bsentença\b", re.I)),
    ("Intimação de pauta/audiência", re.compile(r"\baudiência\b", re.I)),
    ("Expedição de alvará/RPV", re.compile(r"\balvará\b|\bRPV\b|requisição de pequeno valor", re.I)),
    ("Decisão interlocutória", re.compile(r"\b(?:defiro|indefiro|decido|tutela)\b", re.I)),
    ("Despacho", re.compile(r"\b(?:intime-se|intimem-se|manifeste-se|cite-se|vista)\b", re.I)),
]
RE_PRAZO_EXTENSO = re.compile(
    r"prazo\s+(?:comum\s+|sucessivo\s+)?de\s+(cinco|dez|quinze|trinta|sessenta)\s+dias", re.IGNORECASE
)
EXTENSO = {"cinco": 5, "dez": 10, "quinze": 15, "trinta": 30, "sessenta": 60}
RE_DOBRO = re.compile(r"\bem\s+dobro\b|\bart\.?\s*229\b", re.I)
COLUNAS = [
    "id", "processo", "tribunal", "comarca", "tipo_ato", "prazo_dias", "inicio_contagem",
    "vencimento_interno", "termo_legal", "dias_uteis_restantes", "criticidade", "cor",
    "situacao", "motivos_revisao", "trecho_comando", "avisos",
]


def ler_data(s: str) -> dt.date:
    s = s.strip()
    return dt.datetime.strptime(s, "%d/%m/%Y").date() if "/" in s else dt.date.fromisoformat(s)


def trecho(texto: str, inicio: int, fim: int, margem: int = 160) -> str:
    """Frase em volta do comando, limitada, sem quebra de linha."""
    a = texto.rfind(".", 0, inicio) + 1
    if inicio - a > margem:
        a = inicio - margem
    b = texto.find(".", fim)
    b = len(texto) if b == -1 or b - fim > margem else b + 1
    return " ".join(texto[a:b].split())


def tipo_ato(texto: str) -> str:
    return next((nome for nome, rx in TIPOS if rx.search(texto)), "Outro")


def detectar_prazo(texto: str) -> tuple[int | None, list[str], list[re.Match]]:
    """Prazo em dias só quando é único e inequívoco; do contrário, None e os motivos."""
    achados = list(RE_PRAZO.finditer(texto))
    dias = {int(m.group(1)) for m in achados if m.group(2).lower().startswith("dia")}
    dias |= {EXTENSO[m.group(1).lower()] for m in RE_PRAZO_EXTENSO.finditer(texto)}
    motivos = []
    if any(not m.group(2).lower().startswith("dia") for m in achados):
        motivos.append("prazo em horas ou meses")
    if not dias:
        motivos.append("prazo expresso não encontrado (art. 218, § 3º? ciência pura?)")
    elif len(dias) > 1:
        motivos.append(f"mais de um prazo: {sorted(dias)}")
    if RE_DOBRO.search(texto):
        motivos.append("menção a prazo em dobro")
    prazo = next(iter(dias)) if len(dias) == 1 and not motivos else None
    return prazo, motivos, achados


def triar_linha(i: int, linha: dict, calendarios: dict, hoje: dt.date) -> dict:
    texto = linha.get("texto", "") or ""
    numero = (linha.get("processo") or "").strip()
    motivos: list[str] = []

    erro_cnj = cnj.diagnosticar(numero) if numero else "número ausente"
    trib = cnj.tribunal(numero) if not erro_cnj else None
    if erro_cnj:
        motivos.append(f"CNJ: {erro_cnj}")

    tipo = tipo_ato(texto)
    prazo, motivos_prazo, achados = detectar_prazo(texto)
    motivos += motivos_prazo
    if tipo == "Intimação de pauta/audiência":
        motivos.append("audiência: conferir data/hora e agenda")

    saida = {c: "" for c in COLUNAS}
    saida.update(id=i, processo=cnj.formatar(numero) if not erro_cnj else numero,
                 tribunal=trib or "", comarca=cnj.comarca(numero) if trib else "", tipo_ato=tipo)
    if achados:
        saida["trecho_comando"] = trecho(texto, achados[0].start(), achados[0].end())

    if prazo and trib:
        chave = (trib, saida["comarca"])
        if chave not in calendarios:
            calendarios[chave] = Calendario(CSV_PADRAO, trib, saida["comarca"])
        r = calcular_prazo(ler_data(linha["data_disponibilizacao"]), prazo, calendarios[chave], hoje)
        saida.update(prazo_dias=r.dias_prazo, inicio_contagem=r.inicio_contagem,
                     vencimento_interno=r.vencimento_interno, termo_legal=r.termo_legal,
                     dias_uteis_restantes=r.dias_uteis_restantes, criticidade=r.criticidade,
                     cor=r.cor, avisos=" | ".join(r.avisos))
    saida["situacao"] = "REVISAR" if motivos else "OK"
    saida["motivos_revisao"] = "; ".join(motivos)
    return saida


def ler_csv(caminho: Path) -> list[dict]:
    bruto = caminho.read_text(encoding="utf-8-sig")
    delim = ";" if bruto.split("\n", 1)[0].count(";") > bruto.split("\n", 1)[0].count(",") else ","
    return list(csv.DictReader(bruto.splitlines(), delimiter=delim))


def main(argv: list[str] | None = None) -> int:
    p = argparse.ArgumentParser(description="Pré-triagem determinística de publicações.")
    p.add_argument("entrada", type=Path)
    p.add_argument("-o", "--saida", type=Path, default=Path("prazos_processuais.csv"))
    p.add_argument("--para-ia", type=Path, help="arquivo só com as linhas que exigem leitura")
    p.add_argument("--hoje", type=dt.date.fromisoformat, default=dt.date.today())
    a = p.parse_args(argv)

    linhas = ler_csv(a.entrada)
    calendarios: dict = {}
    resultado = [triar_linha(i, l, calendarios, a.hoje) for i, l in enumerate(linhas, start=1)]
    with a.saida.open("w", encoding="utf-8-sig", newline="") as f:
        w = csv.DictWriter(f, fieldnames=COLUNAS, delimiter=";")
        w.writeheader()
        w.writerows(resultado)

    revisar = [r for r in resultado if r["situacao"] == "REVISAR"]
    if a.para_ia and revisar:
        blocos = [
            f"#{r['id']} | {r['processo']} | disp. {linhas[r['id'] - 1]['data_disponibilizacao']} | {r['motivos_revisao']}\n"
            f"{' '.join(linhas[r['id'] - 1]['texto'].split())}"
            for r in revisar
        ]
        a.para_ia.write_text(
            "Para cada publicação, responda só: #id | tipo de ato | providência | prazo em dias úteis "
            "(ou 'sem prazo') | fundamento. Não calcule datas.\n\n" + "\n\n".join(blocos) + "\n",
            encoding="utf-8",
        )
    print(f"{len(resultado)} publicações | {len(resultado) - len(revisar)} resolvidas em Python | "
          f"{len(revisar)} para revisão -> {a.saida}" + (f", {a.para_ia}" if a.para_ia and revisar else ""))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
