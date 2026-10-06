"""Diagnóstico do HTTP 403 do Robô Diário na API Comunica PJe/DJEN.

Hipótese (relatos de terceiros, não confirmada pelo CNJ): a API recusa IPs de fora
do Brasil, como os do Google Colab. Este script faz uma única consulta pequena e
registra o resultado em diagnostico_djen.log. Rode-o no Colab e num computador do
escritório: 403 no primeiro e 200 no segundo confirmam a hipótese.

Só usa a biblioteca padrão. Não grava credenciais.

Uso:
    python -m automacoes.diagnostico_djen --oab 29573 --uf PB
"""
from __future__ import annotations

import argparse
import datetime as dt
import json
import socket
import urllib.error
import urllib.parse
import urllib.request
from pathlib import Path

ENDPOINT = "https://comunicaapi.pje.jus.br/api/v1/comunicacao"
LOG = Path("diagnostico_djen.log")


def consultar(oab: str, uf: str, dias: int = 7, timeout: int = 30) -> dict:
    fim = dt.date.today()
    params = {
        "numeroOab": oab,
        "ufOab": uf.upper(),
        "dataDisponibilizacaoInicio": (fim - dt.timedelta(days=dias)).isoformat(),
        "dataDisponibilizacaoFim": fim.isoformat(),
        "pagina": 1,
        "itensPorPagina": 5,
    }
    url = f"{ENDPOINT}?{urllib.parse.urlencode(params)}"
    registro = {"quando": dt.datetime.now().isoformat(timespec="seconds"), "maquina": socket.gethostname(), "url": url}
    req = urllib.request.Request(url, headers={"Accept": "application/json", "User-Agent": "diagnostico-djen/1.0"})
    try:
        with urllib.request.urlopen(req, timeout=timeout) as resp:
            corpo = json.loads(resp.read().decode("utf-8"))
            registro.update(status=resp.status, total=corpo.get("count"), chaves=sorted(corpo)[:10])
    except urllib.error.HTTPError as e:
        registro.update(
            status=e.code,
            servidor=e.headers.get("Server"),
            cache=e.headers.get("X-Cache"),
            corpo=e.read(300).decode("utf-8", "replace"),
        )
    except (urllib.error.URLError, TimeoutError, OSError) as e:
        registro.update(status=None, erro=repr(e))
    return registro


def main(argv: list[str] | None = None) -> int:
    p = argparse.ArgumentParser(description="Testa o acesso à API do DJEN a partir desta máquina.")
    p.add_argument("--oab", required=True)
    p.add_argument("--uf", required=True)
    p.add_argument("--dias", type=int, default=7)
    a = p.parse_args(argv)
    r = consultar(a.oab, a.uf, a.dias)
    with LOG.open("a", encoding="utf-8") as f:
        f.write(json.dumps(r, ensure_ascii=False) + "\n")
    print(json.dumps(r, ensure_ascii=False, indent=1))
    if r.get("status") == 403:
        print("\n403: recusa do servidor. Se o mesmo comando der 200 na rede do escritório, a causa é a origem do IP.")
    return 0 if r.get("status") == 200 else 1


if __name__ == "__main__":
    raise SystemExit(main())
