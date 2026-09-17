"""Coleta de processos do TJPB via API pública do DataJud (CNJ).

Filtra por classe processual + assunto (no servidor) e, dentro disso,
por órgão julgador (1ª Câmara Cível) no cliente, já que o nome do órgão
varia de tribunal para tribunal e não há garantia de um código estável
conhecido de antemão.

Classe 198 (Apelação Cível) é a classe de 2º grau/recursal -- Câmaras
Cíveis julgam recursos, não processos de origem (ex.: "Procedimento
Comum Cível", classe 7, é 1º grau e praticamente não coexiste com
órgão colegiado tipo Câmara).

Há campos que não pude confirmar com o manual público do DataJud no
momento (indisponível/bloqueado): em especial, um indicador de
"prioridade de tramitação (idoso)". Em vez de arriscar um nome de
campo, use `--diagnostico` (ver main()) para inspecionar o _source
bruto de processos reais e descobrir os nomes/códigos exatos.
"""

import argparse
import json
import os
import re
import time
import unicodedata

import pandas as pd
import requests

URL = "https://api-publica.datajud.cnj.jus.br/api_publica_tjpb/_search"

# Chave pública documentada pelo CNJ para uso geral da API (não é um
# segredo pessoal), mas fica sobrescrevível por variável de ambiente.
API_KEY = os.environ.get(
    "DATAJUD_API_KEY",
    "cDZHYzlZa0JadVREZDJCendQbXY6SkJlTzNjLV9TRENyQk1RdnFKZGRQdw==",
)

HEADERS = {
    "Authorization": f"APIKey {API_KEY}",
    "Content-Type": "application/json",
}

CLASSE_CODIGO = 198  # Apelação Cível
ASSUNTOS_CODIGOS = [7779, 7780, 11807]  # Dano Moral, Dano Material, Tarifas (qualquer um)

# Opcional: se a inspeção via --diagnostico confirmar um campo de grau
# (ex.: "grau": "G2") no schema atual, adicione aqui para reduzir o
# volume no servidor -- Câmaras só existem em 2º grau. Não incluí por
# não ter conseguido confirmar o nome/valor exatos agora.
# FILTRO_GRAU = {"term": {"grau": "G2"}}

TAMANHO_LOTE = 100
MAX_TENTATIVAS = 5  # por página, em caso de erro de rede/HTTP
TIMEOUT = 45

# Regex aplicado ao nome do órgão já normalizado (maiúsculo, sem acento).
# \b1\b exige "1" isolado, evitando falso positivo com "11ª", "21ª" etc.
# O [AO]? opcional absorve o resíduo da decomposição NFKD de "ª"/"º"
# (que vira a letra "a"/"o", não uma marca combinante removível).
PADRAO_1_CAMARA_CIVEL = re.compile(r"\b1\s*[AO]?\s*CAMARA\b.*\bCIVEL\b")


def normalizar(texto: str) -> str:
    """Maiúsculas, sem acento, sem pontuação -- para comparação robusta."""
    sem_acento = unicodedata.normalize("NFKD", texto)
    sem_acento = "".join(c for c in sem_acento if not unicodedata.combining(c))
    sem_acento = re.sub(r"[^A-Za-z0-9]+", " ", sem_acento)
    return sem_acento.upper().strip()


def eh_primeira_camara_civel(orgao_nome: str) -> bool:
    return bool(PADRAO_1_CAMARA_CIVEL.search(normalizar(orgao_nome)))


def _post_com_retentativa(payload: dict) -> dict:
    """POST com backoff exponencial. Nunca avança a paginação em caso de
    falha -- só retorna (ou lança) depois de esgotar as tentativas, para
    não perder lotes de dados silenciosamente."""
    espera = 1.0
    for tentativa in range(1, MAX_TENTATIVAS + 1):
        try:
            resposta = requests.post(
                URL, headers=HEADERS, json=payload, timeout=TIMEOUT
            )
        except requests.exceptions.RequestException as exc:
            if tentativa == MAX_TENTATIVAS:
                raise
            print(f"⚠️ Falha de conexão ({exc}). Nova tentativa em {espera:.0f}s...")
            time.sleep(espera)
            espera *= 2
            continue

        if resposta.status_code == 200:
            return resposta.json()

        if resposta.status_code == 429 or resposta.status_code >= 500:
            if tentativa == MAX_TENTATIVAS:
                resposta.raise_for_status()
            retry_after = resposta.headers.get("Retry-After")
            espera_efetiva = float(retry_after) if retry_after else espera
            print(
                f"⚠️ HTTP {resposta.status_code}. Nova tentativa em"
                f" {espera_efetiva:.0f}s..."
            )
            time.sleep(espera_efetiva)
            espera *= 2
            continue

        # Erro que não se resolve tentando de novo (400, 401, 403...).
        raise requests.exceptions.HTTPError(
            f"HTTP {resposta.status_code}: {resposta.text}"
        )

    raise RuntimeError("Número máximo de tentativas excedido.")


def montar_query() -> dict:
    return {
        "bool": {
            "must": [
                {"term": {"classe.codigo": CLASSE_CODIGO}},
                {"terms": {"assuntos.codigo": ASSUNTOS_CODIGOS}},
                # FILTRO_GRAU,  # ver comentário acima
            ]
        }
    }


def inspecionar_amostra(tamanho: int = 3) -> None:
    """Imprime o _source bruto de alguns processos reais que casam com
    classe+assuntos, sem nenhum filtro de órgão. Use para descobrir, na
    prática, o nome/código exato do órgão colegiado e de qualquer campo
    de prioridade -- sem depender do manual do CNJ."""
    payload = {"query": montar_query(), "size": tamanho}
    dados = _post_com_retentativa(payload)
    hits = dados.get("hits", {}).get("hits", [])
    if not hits:
        print("Nenhum resultado para classe/assuntos atuais -- nada para inspecionar.")
        return
    for hit in hits:
        print(json.dumps(hit.get("_source", {}), indent=2, ensure_ascii=False))
        print("-" * 60)


def buscar_processos_1_camara_civel() -> pd.DataFrame:
    processos_1_camara = []
    search_after = None
    total_geral = None
    pagina = 0

    print(
        "=== PROJETO: ANÁLISE ESCRITÓRIO DE JURISMETRIA ==="
        "\nIniciando varredura no DataJud (TJPB) com filtro exclusivo para a 1ª"
        " Câmara Cível..."
    )

    while True:
        payload = {
            "query": montar_query(),
            "size": TAMANHO_LOTE,
            # sort explícito é necessário para paginação estável e para o
            # search_after funcionar; usa @timestamp + _id como par único
            # para não deixar de fora nem duplicar registros entre páginas.
            "sort": [{"@timestamp": "asc"}, {"_id": "asc"}],
        }
        if search_after is not None:
            payload["search_after"] = search_after

        dados = _post_com_retentativa(payload)
        hits = dados.get("hits", {}).get("hits", [])

        if total_geral is None:
            total_geral = dados.get("hits", {}).get("total", {}).get("value", 0)
            print(f"Total bruto de registros na base do tribunal: {total_geral}\n")

        if not hits:
            break

        for item in hits:
            source = item.get("_source", {})
            orgao = source.get("orgaoJulgador", {}).get("nome", "")

            if eh_primeira_camara_civel(orgao):
                assuntos = ", ".join(
                    a.get("nome", "") for a in source.get("assuntos", [])
                )
                data_ajuizamento = source.get("dataAjuizamento") or "N/D"

                processos_1_camara.append({
                    "Número CNJ": source.get("numeroProcesso"),
                    "Classe": source.get("classe", {}).get("nome"),
                    "Órgão Julgador": orgao,
                    "Assuntos": assuntos,
                    "Data Ajuizamento": data_ajuizamento[:10],
                })

        pagina += 1
        print(f"Progresso: página {pagina} ({pagina * TAMANHO_LOTE} lidos)...")

        if len(hits) < TAMANHO_LOTE:
            break

        search_after = hits[-1]["sort"]
        time.sleep(0.5)  # margem de segurança para limite de requisições

    return pd.DataFrame(processos_1_camara)


def _exibir(df: pd.DataFrame) -> None:
    try:
        from IPython.display import display as _ipython_display

        _ipython_display(df)
    except ImportError:
        print(df.to_string(index=False))


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument(
        "--diagnostico",
        action="store_true",
        help=(
            "Em vez da coleta completa, imprime o _source bruto de alguns"
            " processos (classe+assuntos, sem filtro de órgão) para você"
            " inspecionar nomes/códigos de campos reais."
        ),
    )
    args = parser.parse_args()

    if args.diagnostico:
        inspecionar_amostra()
        return

    df_resultado = buscar_processos_1_camara_civel()

    print("\n" + "=" * 60)
    if not df_resultado.empty:
        print(f"✅ Apuração concluída! Processos vinculados à 1ª Câmara Cível: {len(df_resultado)}\n")
        _exibir(df_resultado)
        caminho_csv = "processos_1_camara_civel_tjpb.csv"
        df_resultado.to_csv(caminho_csv, index=False, encoding="utf-8-sig")
        print(f"\n💾 Resultado salvo em {caminho_csv}")
    else:
        print(
            "❌ Nenhum processo correspondente foi isolado para a 1ª Câmara Cível"
            " com os parâmetros atuais."
        )


if __name__ == "__main__":
    main()
