"""Coleta de processos do TJPB via API pública do DataJud (CNJ).

Por padrão, busca só por assunto (ex.: Tarifas, código 11807) -- sem
exigir classe processual nem órgão julgador -- porque empilhar vários
filtros ao mesmo tempo (classe + assunto + órgão) tende a zerar o
resultado sem deixar claro qual filtro foi o culpado. Classe e o
filtro de "1ª Câmara Cível" ficam disponíveis como opções (--classe,
--somente-1-camara) para quem quiser restringir de novo depois de
inspecionar os dados reais com --diagnostico.

Há campos que não pude confirmar com o manual público do DataJud
(indisponível/bloqueado a partir deste ambiente): em especial, um
indicador de "prioridade de tramitação (idoso)". Use --diagnostico
para inspecionar o _source bruto de processos reais e descobrir
nomes/códigos exatos, em vez de arriscar um palpite.
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

ASSUNTO_TARIFAS = 11807
ASSUNTO_DANO_MORAL = 7779
ASSUNTO_DANO_MATERIAL = 7780
DEFAULT_ASSUNTOS = [ASSUNTO_TARIFAS]

CLASSE_APELACAO_CIVEL = 198
CLASSE_PROCEDIMENTO_COMUM_CIVEL = 7

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


def montar_query(assuntos_codigos: list, classe_codigo: int | None) -> dict:
    must = [{"terms": {"assuntos.codigo": assuntos_codigos}}]
    if classe_codigo is not None:
        must.append({"term": {"classe.codigo": classe_codigo}})
    return {"bool": {"must": must}}


def inspecionar_amostra(
    assuntos_codigos: list, classe_codigo: int | None, tamanho: int = 3
) -> None:
    """Imprime o _source bruto de alguns processos reais que casam com o
    filtro atual, sem nenhum filtro de órgão. Use para descobrir, na
    prática, o nome/código exato do órgão colegiado e de qualquer campo
    de prioridade -- sem depender do manual do CNJ."""
    payload = {"query": montar_query(assuntos_codigos, classe_codigo), "size": tamanho}
    dados = _post_com_retentativa(payload)
    hits = dados.get("hits", {}).get("hits", [])
    if not hits:
        print("Nenhum resultado para os parâmetros atuais -- nada para inspecionar.")
        return
    for hit in hits:
        print(json.dumps(hit.get("_source", {}), indent=2, ensure_ascii=False))
        print("-" * 60)


def _construir_sort(usar_tiebreak: bool) -> list:
    # NÃO use "_id" aqui: o Elasticsearch do DataJud proíbe fielddata
    # sobre o campo meta "_id" (HTTP 400 "Fielddata access on the _id
    # field is disallowed"). numeroProcesso serve de desempate real
    # para @timestamp, evitando perder/duplicar registros em empates.
    sort = [{"@timestamp": "asc"}]
    if usar_tiebreak:
        sort.append({"numeroProcesso": "asc"})
    return sort


def buscar_processos(
    assuntos_codigos: list,
    classe_codigo: int | None,
    somente_1_camara: bool = False,
) -> pd.DataFrame:
    processos = []
    search_after = None
    total_geral = None
    pagina = 0
    usar_tiebreak = True

    print(
        "=== PROJETO: ANÁLISE ESCRITÓRIO DE JURISMETRIA ==="
        f"\nAssuntos: {assuntos_codigos} | Classe: {classe_codigo or 'qualquer'} |"
        f" Filtro 1ª Câmara Cível: {'sim' if somente_1_camara else 'não'}"
    )

    while True:
        payload = {
            "query": montar_query(assuntos_codigos, classe_codigo),
            "size": TAMANHO_LOTE,
            "sort": _construir_sort(usar_tiebreak),
        }
        if search_after is not None:
            payload["search_after"] = search_after

        try:
            dados = _post_com_retentativa(payload)
        except requests.exceptions.HTTPError as exc:
            # Se o campo de desempate também não puder ser ordenado
            # neste cluster, recua para @timestamp isolado em vez de
            # travar -- não há como confirmar de antemão quais campos
            # o índice público aceita para "sort".
            if usar_tiebreak and "fielddata" in str(exc).lower():
                print(
                    "⚠️ numeroProcesso não é ordenável neste cluster;"
                    " prosseguindo apenas com @timestamp..."
                )
                usar_tiebreak = False
                continue
            raise
        hits = dados.get("hits", {}).get("hits", [])

        if total_geral is None:
            total_geral = dados.get("hits", {}).get("total", {}).get("value", 0)
            print(f"Total bruto de registros na base do tribunal: {total_geral}\n")

        if not hits:
            break

        for item in hits:
            source = item.get("_source", {})
            orgao = source.get("orgaoJulgador", {}).get("nome", "")

            if somente_1_camara and not eh_primeira_camara_civel(orgao):
                continue

            assuntos = ", ".join(
                a.get("nome", "") for a in source.get("assuntos", [])
            )
            data_ajuizamento = source.get("dataAjuizamento") or "N/D"

            processos.append({
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

    return pd.DataFrame(processos)


def _exibir(df: pd.DataFrame) -> None:
    try:
        from IPython.display import display as _ipython_display

        _ipython_display(df)
    except ImportError:
        print(df.to_string(index=False))


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument(
        "--assunto",
        type=int,
        action="append",
        dest="assuntos",
        help=(
            "Código de assunto a incluir (repita para OR entre vários)."
            f" Padrão: {DEFAULT_ASSUNTOS} (Tarifas)."
        ),
    )
    parser.add_argument(
        "--classe",
        type=int,
        default=None,
        help="Código de classe processual para restringir a busca (padrão: nenhum).",
    )
    parser.add_argument(
        "--somente-1-camara",
        action="store_true",
        help="Filtra, no cliente, apenas processos da 1ª Câmara Cível.",
    )
    parser.add_argument(
        "--diagnostico",
        action="store_true",
        help=(
            "Em vez da coleta completa, imprime o _source bruto de alguns"
            " processos (mesmo filtro de assunto/classe, sem filtro de"
            " órgão) para você inspecionar nomes/códigos de campos reais."
        ),
    )
    args = parser.parse_args()

    assuntos_codigos = args.assuntos or DEFAULT_ASSUNTOS

    if args.diagnostico:
        inspecionar_amostra(assuntos_codigos, args.classe)
        return

    df_resultado = buscar_processos(assuntos_codigos, args.classe, args.somente_1_camara)

    print("\n" + "=" * 60)
    if not df_resultado.empty:
        print(f"✅ Apuração concluída! Processos encontrados: {len(df_resultado)}\n")
        _exibir(df_resultado)
        caminho_csv = "processos_datajud_tjpb.csv"
        df_resultado.to_csv(caminho_csv, index=False, encoding="utf-8-sig")
        print(f"\n💾 Resultado salvo em {caminho_csv}")
    else:
        print("❌ Nenhum processo correspondente foi encontrado com os parâmetros atuais.")


if __name__ == "__main__":
    main()
