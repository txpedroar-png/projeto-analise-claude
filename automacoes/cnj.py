"""Validação do número único de processo (Res. CNJ 65/2008) e auditoria de minutas.

Formato: NNNNNNN-DD.AAAA.J.TR.OOOO. O dígito DD satisfaz
    int(NNNNNNN AAAA J TR OOOO DD) mod 97 == 1.

Limite: passar no teste prova que o número é BEM FORMADO, não que o processo existe.
Um número inventado ao acaso passa em cerca de 1 a cada 97 casos. A existência
se confere no DataJud ou no PJe.

Uso (CLI): código de saída 1 se houver número inválido, para uso em lote.
    python -m automacoes.cnj minuta.docx
    python -m automacoes.cnj minuta.txt
"""
from __future__ import annotations

import argparse
import datetime as dt
import json
import re
import sys
import zipfile
from pathlib import Path

# Segmento J=8 (Justiça Estadual): código TR -> tribunal.
TJ = {
    "01": "TJAC", "02": "TJAL", "03": "TJAP", "04": "TJAM", "05": "TJBA", "06": "TJCE", "07": "TJDFT",
    "08": "TJES", "09": "TJGO", "10": "TJMA", "11": "TJMT", "12": "TJMS", "13": "TJMG", "14": "TJPA",
    "15": "TJPB", "16": "TJPR", "17": "TJPE", "18": "TJPI", "19": "TJRJ", "20": "TJRN", "21": "TJRS",
    "22": "TJRO", "23": "TJRR", "24": "TJSC", "25": "TJSE", "26": "TJSP", "27": "TJTO",
}

# Com pontuação padrão ou 20 dígitos corridos; delimitado para não casar dentro de números maiores.
PADRAO = re.compile(r"(?<!\d)(\d{7}-\d{2}\.\d{4}\.\d\.\d{2}\.\d{4}|\d{20})(?!\d)")


def digitos(numero: str) -> str:
    return re.sub(r"\D", "", numero)


def formatar(numero: str) -> str:
    d = digitos(numero)
    return f"{d[:7]}-{d[7:9]}.{d[9:13]}.{d[13]}.{d[14:16]}.{d[16:]}"


def calcular_dv(sequencial: str, ano: str, j: str, tr: str, origem: str) -> str:
    return f"{98 - int(sequencial + ano + j + tr + origem) * 100 % 97:02d}"


def diagnosticar(numero: str, ano_max: int | None = None) -> str | None:
    """None se válido; senão, o motivo."""
    d = digitos(numero)
    if len(d) != 20:
        return "não tem 20 dígitos"
    ano, j = int(d[9:13]), d[13]
    if j == "0":
        return "segmento de justiça (J) inexistente"
    if not 1900 <= ano <= (ano_max or dt.date.today().year + 1):
        return f"ano {ano} implausível"
    if j == "8" and d[14:16] not in TJ:
        return f"código de tribunal estadual {d[14:16]} inexistente"
    if int(d[:7] + d[9:] + d[7:9]) % 97 != 1:
        return f"dígito verificador errado (esperado {calcular_dv(d[:7], d[9:13], j, d[14:16], d[16:])})"
    return None


def validar(numero: str) -> bool:
    return diagnosticar(numero) is None


def tribunal(numero: str) -> str | None:
    d = digitos(numero)
    if len(d) != 20:
        return None
    j, tr = d[13], d[14:16]
    if j == "8":
        return TJ.get(tr)
    if j == "4":
        return f"TRF{int(tr)}"
    return f"J{j}-TR{tr}"


def comarca(numero: str) -> str:
    return digitos(numero)[16:]


def auditar_texto(texto: str) -> dict:
    vistos: dict[str, str | None] = {}
    for m in PADRAO.finditer(texto):
        canon = formatar(m.group(1))
        vistos.setdefault(canon, diagnosticar(canon))
    invalidos = {n: motivo for n, motivo in vistos.items() if motivo}
    return {
        "total": len(vistos),
        "bem_formados": sorted(n for n, motivo in vistos.items() if not motivo),
        "invalidos": invalidos,
        "aprovado": not invalidos,
        "observacao": "bem formado não prova existência: conferir no DataJud/PJe",
    }


def ler_texto(caminho: Path) -> str:
    if caminho.suffix.lower() == ".docx":
        with zipfile.ZipFile(caminho) as z:
            xml = z.read("word/document.xml").decode("utf-8")
        xml = re.sub(r"</w:p>", "\n", xml)
        return re.sub(r"<[^>]+>", "", xml)
    return caminho.read_text(encoding="utf-8", errors="replace")


def main(argv: list[str] | None = None) -> int:
    p = argparse.ArgumentParser(description="Audita números CNJ citados em uma minuta (.docx ou texto).")
    p.add_argument("arquivo", nargs="?", help="omitido: lê da entrada padrão")
    a = p.parse_args(argv)
    texto = ler_texto(Path(a.arquivo)) if a.arquivo else sys.stdin.read()
    r = auditar_texto(texto)
    print(json.dumps(r, ensure_ascii=False, indent=1))
    return 0 if r["aprovado"] else 1


if __name__ == "__main__":
    raise SystemExit(main())
