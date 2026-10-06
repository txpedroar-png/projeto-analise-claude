import zipfile

import pytest

from automacoes import cnj

VALIDO = "0000001-90.2020.8.15.0001"  # DV calculado pela fórmula da Res. CNJ 65/2008


def test_numero_valido():
    assert cnj.validar(VALIDO)
    assert cnj.validar(cnj.digitos(VALIDO))
    assert cnj.tribunal(VALIDO) == "TJPB" and cnj.comarca(VALIDO) == "0001"


def test_exemplo_do_gemini_e_invalido():
    assert cnj.diagnosticar("0000001-35.2020.8.15.0001") == "dígito verificador errado (esperado 90)"


@pytest.mark.parametrize("numero, motivo", [
    ("123", "20 dígitos"),
    ("0000001-90.2020.0.15.0001", "segmento"),
    ("0000001-90.1800.8.15.0001", "ano"),
    ("0000001-90.2020.8.99.0001", "tribunal estadual"),
])
def test_estrutura_invalida(numero, motivo):
    assert motivo in cnj.diagnosticar(numero)


def test_calcular_dv_bate_com_validacao():
    for seq in ("0000001", "0803256", "1234567"):
        dv = cnj.calcular_dv(seq, "2025", "8", "15", "0231")
        assert cnj.validar(f"{seq}-{dv}.2025.8.15.0231")


def test_auditoria_deduplica_e_separa():
    texto = f"Autos {VALIDO} e {cnj.digitos(VALIDO)}; citado 0000001-35.2020.8.15.0001; ignorar 123456789012345678901."
    r = cnj.auditar_texto(texto)
    assert r["total"] == 2
    assert r["bem_formados"] == [VALIDO]
    assert list(r["invalidos"]) == ["0000001-35.2020.8.15.0001"]
    assert r["aprovado"] is False


def test_leitura_docx_com_runs_partidos(tmp_path):
    p = tmp_path / "minuta.docx"
    xml = ('<w:document><w:body><w:p><w:r><w:t>Processo 0000001-</w:t></w:r>'
           '<w:r><w:t>90.2020.8.15.0001</w:t></w:r></w:p></w:body></w:document>')
    with zipfile.ZipFile(p, "w") as z:
        z.writestr("word/document.xml", xml)
    assert cnj.main([str(p)]) == 0
