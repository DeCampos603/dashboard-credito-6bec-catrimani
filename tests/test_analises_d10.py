"""Passos 10.1–10.4: fila de ação, recolhimentos, idade do saldo e metas do COTER."""
import datetime
import json
import os
import sys

import pytest

RAIZ = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, RAIZ)

import analises_d10 as a  # noqa: E402
from gerar_dashboard import etl  # noqa: E402

HOJE = datetime.date(2026, 9, 29)
FIX = os.path.join(RAIZ, "tests", "fixtures", "credito_disp_11set2026.xlsx")


@pytest.fixture(scope="module")
def catr():
    return etl(FIX, hoje=HOJE)[3]


def test_fila_ordenada_por_urgencia(catr):
    fila = a.fila_acao(catr["linhas"])
    assert fila and all(i["cred"] > 0.005 for i in fila)
    ordem = [a._ORDEM[i["situacao_prazo"]] for i in fila]
    assert ordem == sorted(ordem)
    assert abs(sum(i["cred"] for i in fila) - catr["totais"]["cred"]) < 0.01


def test_texto_cobranca_so_com_dados_da_fonte(catr):
    fila = [i for i in a.fila_acao(catr["linhas"]) if i["ug"] == "160352"]
    txt = a.texto_cobranca("7º BIS", "160352", fila, HOJE)
    assert "UG 160352" in txt and "R$" in txt and "PRAZO VENCIDO" in txt


def test_recolhimentos_so_negativos(catr):
    rc = a.recolhimentos(catr["linhas"])
    assert rc["total"] >= 0 and rc["n"] >= 0
    assert abs(sum(rc["por_mes"].values()) - rc["total"]) < 0.01


def test_idade_ponderada(catr):
    idade = a.idade_ociosa(catr["linhas"], HOJE)
    assert idade
    for u in idade.values():
        assert u["max"] >= u["media"] >= 0


def test_metas_parametro_e_desvio(tmp_path, catr):
    assert a.carregar_metas(str(tmp_path / "nao_existe.json"))["marcos"] == []
    arq = tmp_path / "m.json"
    arq.write_text(json.dumps({"fonte": "teste", "marcos": [
        {"rotulo": "T", "de": "2026-09-01", "ate": "2026-10-31", "emp": 90.0, "liq": 50.0}]}), encoding="utf-8")
    metas = a.carregar_metas(str(arq))
    meta = a.meta_vigente(metas, HOJE)
    assert meta and meta["emp"] == 90.0
    d = a.desvio_metas(catr["por_ug"], meta)
    u = catr["por_ug"][0]
    assert abs(d[u["cod"]]["d_emp"] - (d[u["cod"]]["pct_emp"] - 90.0)) < 1e-9
    assert a.meta_vigente(metas, datetime.date(2027, 1, 1)) is None
    assert a.desvio_metas(catr["por_ug"], None)[u["cod"]]["d_emp"] is None
