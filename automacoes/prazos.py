"""Motor de prazos processuais em dias úteis (arts. 219, 220 e 224 do CPC).

REGRA CONSERVADORA DO ESCRITÓRIO (intencional):
    A contagem começa no primeiro dia útil seguinte à DISPONIBILIZAÇÃO, e não à
    publicação. O "vencimento_interno" fica, portanto, um dia útil antes do termo
    legal (art. 224, §§ 2º e 3º, do CPC c/c art. 4º, §§ 3º e 4º, da Lei 11.419/2006).
    O "termo_legal" também é devolvido, para sustentar tempestividade.

Fora do escopo por decisão do escritório: prazos em dobro (arts. 180, 183, 186 e 229
do CPC). Os prazos do escritório não se dobram em autos eletrônicos.

Feriados vêm de feriados_forenses.csv. Só entram na contagem as linhas com
conferido=S. As linhas não conferidas não adiam o vencimento; se caírem dentro da
contagem, geram aviso. Assim, um feriado não confirmado nunca empurra o prazo
para depois do termo real.

Uso (CLI):
    python -m automacoes.prazos 2026-12-17 5
    python -m automacoes.prazos 2026-10-01 15 --tribunal TJPB --comarca 0231 --hoje 2026-10-05
"""
from __future__ import annotations

import argparse
import csv
import datetime as dt
import json
from dataclasses import asdict, dataclass, field
from pathlib import Path

CSV_PADRAO = Path(__file__).with_name("feriados_forenses.csv")
UM_DIA = dt.timedelta(days=1)

# Paleta do Google Sheets usada na aba "Prazos Processuais" (manual, seção 11).
CORES = {"NORMAL": "#d9ead3", "URGENTE": "#fff2cc", "CRITICO": "#f4cccc", "VENCIDO": "#f4cccc"}


@dataclass(frozen=True)
class Evento:
    data: dt.date
    descricao: str
    abrangencia: str
    tipo: str
    conferido: bool


def _data(v) -> dt.date:
    if isinstance(v, dt.datetime):
        return v.date()
    if isinstance(v, dt.date):
        return v
    s = str(v).strip()
    return dt.datetime.strptime(s, "%d/%m/%Y").date() if "/" in s else dt.date.fromisoformat(s)


class Calendario:
    """Calendário forense de um tribunal (e, opcionalmente, de uma comarca)."""

    def __init__(self, caminho_csv: Path | str = CSV_PADRAO, tribunal: str = "TJPB", comarca: str | None = None):
        caminho = Path(caminho_csv)
        if not caminho.exists():
            # Falha explícita: calendário ausente não pode virar "sem feriados" em silêncio.
            raise FileNotFoundError(f"Calendário de feriados não encontrado: {caminho}")
        with caminho.open(encoding="utf-8-sig", newline="") as f:
            self._carregar(list(csv.DictReader(f)), str(caminho), tribunal, comarca)

    @classmethod
    def de_registros(cls, registros: list[dict], tribunal: str = "TJPB", comarca: str | None = None,
                     origem: str = "aba Feriados", unidade: str | None = None) -> "Calendario":
        """Mesmo calendário a partir de dicionários (ex.: linhas da aba Feriados da planilha).
        'data' aceita date/datetime ou texto AAAA-MM-DD / DD/MM/AAAA."""
        cal = cls.__new__(cls)
        cal.unidade = unidade
        cal._carregar(registros, origem, tribunal, comarca)
        return cal

    def _carregar(self, registros, origem: str, tribunal: str, comarca: str | None) -> None:
        self.tribunal = tribunal.upper()
        self.comarca = comarca
        self.eventos: dict[dt.date, list[Evento]] = {}
        for n, linha in enumerate(registros, start=2):
            try:
                ev = Evento(
                    data=_data(linha["data"]),
                    descricao=str(linha["descricao"]).strip(),
                    abrangencia=str(linha["abrangencia"]).strip().upper(),
                    tipo=str(linha["tipo"]).strip().upper(),
                    conferido=str(linha["conferido"]).strip().upper() == "S",
                )
            except (KeyError, ValueError, AttributeError, TypeError) as e:
                raise ValueError(f"{origem}, linha {n}: registro inválido ({e})") from e
            if self._aplica(ev.abrangencia):
                self.eventos.setdefault(ev.data, []).append(ev)
        self.anos_cobertos = {d.year for d in self.eventos}

    def _aplica(self, abrangencia: str) -> bool:
        if abrangencia == "NACIONAL" or abrangencia == self.tribunal:
            return True
        # Feriado municipal: "TJPB:0231" (código de origem = 4 últimos dígitos do CNJ) ou
        # "TJPB:SANTA RITA/BAYEUX" (unidade atual, para unidades unificadas).
        ids = {self.comarca, (getattr(self, "unidade", None) or "").upper()} - {None, ""}
        return any(abrangencia == f"{self.tribunal}:{i}" for i in ids)

    @staticmethod
    def em_recesso(d: dt.date) -> bool:
        """Art. 220 do CPC: prazos suspensos de 20/12 a 20/01, inclusive."""
        return (d.month == 12 and d.day >= 20) or (d.month == 1 and d.day <= 20)

    def feriado_conferido(self, d: dt.date) -> bool:
        return any(e.tipo == "FERIADO" and e.conferido for e in self.eventos.get(d, ()))

    def eh_dia_util(self, d: dt.date) -> bool:
        return d.weekday() < 5 and not self.em_recesso(d) and not self.feriado_conferido(d)

    def proximo_dia_util(self, d: dt.date) -> dt.date:
        while not self.eh_dia_util(d):
            d += UM_DIA
        return d

    def somar_dias_uteis(self, inicio_exclusivo: dt.date, n: int) -> dt.date:
        """Devolve o n-ésimo dia útil posterior a inicio_exclusivo."""
        d = inicio_exclusivo
        for _ in range(n):
            d = self.proximo_dia_util(d + UM_DIA)
        return d

    def dias_uteis_entre(self, hoje: dt.date, alvo: dt.date) -> int:
        """Dias úteis de hoje (exclusive) até alvo (inclusive); negativo se alvo já passou."""
        sinal, a, b = (1, hoje, alvo) if alvo >= hoje else (-1, alvo, hoje)
        total, d = 0, a + UM_DIA
        while d <= b:
            total += self.eh_dia_util(d)
            d += UM_DIA
        return sinal * total


@dataclass
class ResultadoPrazo:
    disponibilizacao: str
    dias_prazo: int
    inicio_contagem: str
    vencimento_interno: str
    termo_legal: str
    dias_uteis_restantes: int
    criticidade: str
    cor: str
    avisos: list[str] = field(default_factory=list)


def criticidade(dias_uteis_restantes: int) -> str:
    if dias_uteis_restantes > 5:
        return "NORMAL"
    if dias_uteis_restantes >= 1:
        return "URGENTE"
    return "CRITICO" if dias_uteis_restantes == 0 else "VENCIDO"


def calcular_prazo(
    disponibilizacao: dt.date,
    dias_prazo: int,
    calendario: Calendario | None = None,
    hoje: dt.date | None = None,
) -> ResultadoPrazo:
    if not isinstance(dias_prazo, int) or isinstance(dias_prazo, bool) or dias_prazo < 1:
        raise ValueError(f"dias_prazo deve ser inteiro >= 1 (recebido: {dias_prazo!r})")
    cal = calendario or Calendario()
    hoje = hoje or dt.date.today()

    vencimento = cal.somar_dias_uteis(disponibilizacao, dias_prazo)
    inicio = cal.somar_dias_uteis(disponibilizacao, 1)
    termo_legal = cal.somar_dias_uteis(vencimento, 1)
    restantes = cal.dias_uteis_entre(hoje, vencimento)
    nivel = criticidade(restantes)

    avisos = []
    for ano in sorted(set(range(disponibilizacao.year, termo_legal.year + 1)) - cal.anos_cobertos):
        avisos.append(f"Sem feriados cadastrados para {ano}: vencimento pode estar antecipado.")
    d = disponibilizacao + UM_DIA
    while d <= termo_legal:
        for ev in cal.eventos.get(d, ()):
            if not ev.conferido and d.weekday() < 5 and not cal.em_recesso(d):
                avisos.append(f"{d:%d/%m/%Y} {ev.descricao} ({ev.tipo}) NÃO CONFERIDO: se confirmado, o prazo avança.")
            elif ev.tipo == "EXPEDIENTE_REDUZIDO" and d in (vencimento, termo_legal):
                avisos.append(f"{d:%d/%m/%Y} expediente reduzido no vencimento: conferir prorrogação (art. 224, § 1º, CPC).")
        d += UM_DIA
    if any(cal.em_recesso(disponibilizacao + UM_DIA * i) for i in range((termo_legal - disponibilizacao).days + 1)):
        avisos.append("Contagem atravessa o recesso do art. 220 do CPC (20/12 a 20/01): suspensão aplicada.")

    return ResultadoPrazo(
        disponibilizacao=disponibilizacao.isoformat(),
        dias_prazo=dias_prazo,
        inicio_contagem=inicio.isoformat(),
        vencimento_interno=vencimento.isoformat(),
        termo_legal=termo_legal.isoformat(),
        dias_uteis_restantes=restantes,
        criticidade=nivel,
        cor=CORES[nivel],
        avisos=avisos,
    )


def main(argv: list[str] | None = None) -> int:
    p = argparse.ArgumentParser(description="Calcula prazo processual em dias úteis (regra conservadora).")
    p.add_argument("disponibilizacao", type=dt.date.fromisoformat, help="AAAA-MM-DD")
    p.add_argument("dias", type=int)
    p.add_argument("--tribunal", default="TJPB")
    p.add_argument("--comarca", help="código de origem (4 últimos dígitos do CNJ)")
    p.add_argument("--hoje", type=dt.date.fromisoformat)
    p.add_argument("--csv", default=CSV_PADRAO)
    a = p.parse_args(argv)
    r = calcular_prazo(a.disponibilizacao, a.dias, Calendario(a.csv, a.tribunal, a.comarca), a.hoje)
    print(json.dumps(asdict(r), ensure_ascii=False, indent=1))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
