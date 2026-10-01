import csv
import datetime as dt

from automacoes import cnj, triagem

N1 = f"0000001-{cnj.calcular_dv('0000001', '2026', '8', '15', '0231')}.2026.8.15.0231"
N2 = f"0000002-{cnj.calcular_dv('0000002', '2026', '8', '15', '0231')}.2026.8.15.0231"


def test_triagem_resolve_o_simples_e_separa_o_resto(tmp_path):
    entrada = tmp_path / "t.csv"
    entrada.write_text(
        "data_disponibilizacao;processo;texto\n"
        f"09/10/2026;{N1};Intime-se a parte autora para, no prazo de 5 (cinco) dias, manifestar-se sobre a contestação.\n"
        f"09/10/2026;{N2};Ciência às partes do retorno dos autos.\n"
        f"09/10/2026;{N1};Designo audiência de conciliação para 10/11/2026.\n"
        "09/10/2026;0000001-35.2020.8.15.0001;Manifeste-se em 15 dias.\n",
        encoding="utf-8",
    )
    saida, para_ia = tmp_path / "s.csv", tmp_path / "ia.txt"
    triagem.main([str(entrada), "-o", str(saida), "--para-ia", str(para_ia), "--hoje", "2026-10-09"])
    linhas = list(csv.DictReader(saida.read_text(encoding="utf-8-sig").splitlines(), delimiter=";"))

    ok = linhas[0]
    assert (ok["situacao"], ok["tribunal"], ok["prazo_dias"], ok["vencimento_interno"]) == ("OK", "TJPB", "5", "2026-10-19")
    assert "prazo de 5 (cinco) dias" in ok["trecho_comando"]
    assert [l["situacao"] for l in linhas[1:]] == ["REVISAR"] * 3
    assert "dígito verificador" in linhas[3]["motivos_revisao"]

    enviado = para_ia.read_text(encoding="utf-8")
    assert "#1 " not in enviado and all(f"#{i} " in enviado for i in (2, 3, 4))


def test_regex_de_prazo():
    casos = {
        "no prazo de 15 (quinze) dias úteis": ("15", "dias úteis"),
        "em 48 horas": ("48", "horas"),
        "prazo comum de 10 dias": ("10", "dias"),
    }
    for texto, esperado in casos.items():
        m = triagem.RE_PRAZO.search(texto)
        assert (m.group(1), m.group(2)) == esperado, texto
