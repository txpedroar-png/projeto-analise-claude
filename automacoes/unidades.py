"""Unidade jurisdicional atual de um processo.

O código de origem (4 últimos dígitos do CNJ) é fixado na distribuição e não muda quando a unidade é
unificada ou extinta (ex.: TJPB, Bayeux 0751 e Santa Rita 0121 unificadas). Por isso a unidade atual é
resolvida nesta ordem:
1. nome do órgão julgador (vem do DJEN e reflete a unidade de hoje), quando for de 1º grau e citar um
   município da UF do tribunal (validado na lista do IBGE) ou a Capital, ou um Núcleo 4.0 de saúde;
2. aba Comarcas: "Unidade atual" do código de origem, se preenchida; senão, a "Comarca" do código;
3. "Cód. NNNN", para a equipe completar.
Em qualquer caso, a coluna "Unidade atual" da aba Comarcas funciona como apelido: se a unidade
encontrada for o nome antigo de uma unidade unificada, vale o nome novo.
"""
from __future__ import annotations

import csv
import re
import unicodedata
from functools import lru_cache
from pathlib import Path

MUNICIPIOS_CSV = Path(__file__).with_name("municipios_ibge.csv")
UF_DO_TRIBUNAL = {"TJPB": "PB", "TJRN": "RN", "TJPE": "PE", "TJCE": "CE", "TJSP": "SP", "TJPR": "PR",
                  "TJAL": "AL", "TJBA": "BA", "TJRJ": "RJ", "TJMG": "MG", "TJRS": "RS", "TJSC": "SC"}
CAPITAIS = {"TJPB": ("João Pessoa", "João Pessoa (Capital)"), "TJRN": ("Natal", "Natal (Capital)"),
            "TJPE": ("Recife", "Recife (Capital)"), "TJCE": ("Fortaleza", "Fortaleza (Capital)"),
            "TJSP": ("São Paulo", "São Paulo (Capital)"), "TJPR": ("Curitiba", "Curitiba (Capital)")}
NUCLEO_SAUDE = "Núcleo 4.0 – Saúde Suplementar"
CONECTIVOS = {"de", "da", "do", "das", "dos"}
RE_SEGUNDO_GRAU = re.compile(r"Câmara|Gabinete|\bGab\.|Turma|Tribunal Pleno|Presidência|Vice-Presid", re.I)
RE_NUCLEO_SAUDE = re.compile(r"Núcleo.{0,25}4\.0.{0,30}Sa[úu]de", re.I)
RE_CAPITAL = re.compile(r"\bda Capital\b", re.I)
RE_CANDIDATO = re.compile(r"\b(?:de|do|da)\s+((?:[A-ZÁÉÍÓÚÂÊÔÃÕÇ][\wÀ-ÿ'’-]*)(?:\s+(?:d[aeo]s?\s+)?[A-ZÁÉÍÓÚÂÊÔÃÕÇ][\wÀ-ÿ'’-]*)*)")


def normalizar(texto: str) -> str:
    sem = unicodedata.normalize("NFKD", str(texto or "")).encode("ascii", "ignore").decode()
    return re.sub(r"\s+", " ", sem).strip().upper()


@lru_cache(maxsize=None)
def municipios(uf: str) -> dict[str, str]:
    """Nome normalizado -> nome oficial, para a UF."""
    with MUNICIPIOS_CSV.open(encoding="utf-8") as f:
        return {normalizar(l["municipio"]): l["municipio"] for l in csv.DictReader(f) if l["uf"] == uf}


def municipio_do_orgao(orgao: str, tribunal: str) -> str | None:
    """Último município da UF citado no nome do órgão (ex.: '1ª Vara de Família de Campina Grande')."""
    uf = UF_DO_TRIBUNAL.get(tribunal)
    if not uf:
        return None
    validos = municipios(uf)
    achados = []
    for candidato in RE_CANDIDATO.findall(orgao or ""):
        # "Comarca de Canguaretama", "Família de Campina Grande": testa os finais, do mais longo ao mais curto,
        # para manter nomes compostos inteiros ("Baía da Traição", "São José de Piranhas").
        palavras = candidato.split()
        for i in range(len(palavras)):
            if palavras[i].lower() in CONECTIVOS:
                continue
            final = normalizar(" ".join(palavras[i:]))
            if final in validos:
                achados.append(validos[final])
                break
    return achados[-1] if achados else None


def unidade_pelo_orgao(orgao: str, tribunal: str) -> str | None:
    if not orgao or RE_SEGUNDO_GRAU.search(orgao):
        return None
    if RE_NUCLEO_SAUDE.search(orgao):
        return NUCLEO_SAUDE
    capital = CAPITAIS.get(tribunal)
    if RE_CAPITAL.search(orgao) and capital:
        return capital[1]
    municipio = municipio_do_orgao(orgao, tribunal)
    if municipio and capital and municipio == capital[0]:
        return capital[1]
    return municipio


class Unidades:
    """Resolve a unidade atual a partir das linhas da aba Comarcas
    (Tribunal | Código de origem | Comarca | Conferido | Observação | Unidade atual)."""

    def __init__(self, linhas: list[dict]):
        self.por_codigo: dict[tuple[str, str], dict] = {}
        self.apelido: dict[tuple[str, str], str] = {}
        for l in linhas:
            trib = str(l.get("Tribunal") or "").strip()
            if not trib:
                continue
            cod = str(l.get("Código de origem") or "").strip().split(".")[0].zfill(4)
            comarca = str(l.get("Comarca") or "").strip()
            atual = str(l.get("Unidade atual") or "").strip()
            self.por_codigo[(trib, cod)] = {"comarca": comarca, "atual": atual}
            if comarca and atual:
                self.apelido[(trib, normalizar(comarca))] = atual

    def resolver(self, tribunal: str, codigo: str, orgao: str = "") -> str:
        nome = unidade_pelo_orgao(orgao, tribunal)
        if not nome:
            reg = self.por_codigo.get((tribunal, codigo))
            nome = (reg["atual"] or reg["comarca"]) if reg else ""
        if not nome:
            return f"Cód. {codigo}"
        return self.apelido.get((tribunal, normalizar(nome)), nome)

    def conhece(self, tribunal: str, codigo: str) -> bool:
        return (tribunal, codigo) in self.por_codigo
