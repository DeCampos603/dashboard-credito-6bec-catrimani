"""Prazos de empenho, trava anti-dado-velho e estado 'sem atualização' do e-mail."""
import datetime
import json
import os
import sys

import pytest

RAIZ = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, RAIZ)

import regras_nc  # noqa: E402
import gerar_dashboard as g  # noqa: E402
import relatorio_email_6bec as em  # noqa: E402

HOJE = datetime.date(2026, 9, 29)


@pytest.mark.parametrize("obj,data,tipo", [
    ("EMPENHO ATE 30 SET 26", datetime.date(2026, 9, 30), "DATA"),
    ("PRAZO DE EMPENHO: 30 DE JULHO DE 2026", datetime.date(2026, 7, 30), "DATA"),
    ("EMPENHO ATE O DIA 05 OUT 2026", datetime.date(2026, 10, 5), "DATA"),
    ("EMPENHO IMEDIATO", None, "IMEDIATO"),
    ("MATERIAL DE CONSUMO", None, "SEM_PRAZO"),
    (None, None, "SEM_PRAZO"),
])
def test_prazo_empenho(obj, data, tipo):
    assert regras_nc.prazo_empenho(obj) == (data, tipo)


@pytest.mark.parametrize("prazo,tipo,esperado", [
    (datetime.date(2026, 9, 28), "DATA", regras_nc.SIT_VENCIDO),
    (datetime.date(2026, 9, 29), "DATA", regras_nc.SIT_7D),
    (datetime.date(2026, 10, 6), "DATA", regras_nc.SIT_7D),
    (datetime.date(2026, 10, 7), "DATA", regras_nc.SIT_30D),
    (datetime.date(2026, 10, 29), "DATA", regras_nc.SIT_30D),
    (datetime.date(2026, 10, 30), "DATA", regras_nc.SIT_LONGO),
    (None, "IMEDIATO", regras_nc.SIT_IMEDIATO),
    (None, "SEM_PRAZO", regras_nc.SIT_SEM),
])
def test_situacao_prazo(prazo, tipo, esperado):
    assert regras_nc.situacao_prazo(prazo, tipo, HOJE) == esperado


def test_trava():
    assert regras_nc.tem_trava("NAO DEVE ALTERAR ND/UGR")
    assert not regras_nc.tem_trava("MATERIAL DE CONSUMO")


def test_baixar_sem_fonte_falha(monkeypatch):
    monkeypatch.delenv("SHEETS_CSV_URL", raising=False)
    with pytest.raises(SystemExit):
        g.baixar(None)
    with pytest.raises(SystemExit):
        em.baixar(None)


def test_baixar_com_falha_nao_cai_em_planilha_antiga(monkeypatch):
    def boom(*a, **k):
        raise OSError("rede fora")
    monkeypatch.setattr(g.urllib.request, "urlopen", boom)
    with pytest.raises(SystemExit):
        g.baixar("https://exemplo.invalid/x.csv")


def test_variacao_so_entre_posicoes_distintas():
    h = [{"data": "2026-09-16", "posicao": "2026-09-15", "total": {"cred": 10}},
         {"data": "2026-09-17", "posicao": "2026-09-15", "total": {"cred": 10}},
         {"data": "2026-09-18", "posicao": "2026-09-15", "total": {"cred": 10}}]
    assert g.snapshot_anterior(h) is None
    h.append({"data": "2026-09-19", "posicao": "2026-09-18", "total": {"cred": 7}})
    assert g.snapshot_anterior(h)["data"] == "2026-09-18"


def test_trava_de_posicao_regressiva(tmp_path, monkeypatch):
    fix = os.path.join(RAIZ, "tests", "fixtures", "credito_disp_11set2026.xlsx")
    hist = tmp_path / "h.json"
    hist.write_text(json.dumps([{"data": "2026-09-25", "posicao": "2026-09-25", "total": {"cred": 1}}]), encoding="utf-8")
    monkeypatch.setattr(sys, "argv", ["gerar_dashboard.py", "--local", fix, "--date", "2026-09-29", "--hist-file", str(hist)])
    with pytest.raises(SystemExit) as e:
        g.main()
    assert "anterior à já publicada" in str(e.value)


def test_sem_atualizacao_desde(tmp_path):
    hist = tmp_path / "h.json"
    hist.write_text(json.dumps([
        {"data": "2026-09-25", "posicao": "2026-09-24"},
        {"data": "2026-09-26", "posicao": "2026-09-25"},
        {"data": "2026-09-27", "posicao": "2026-09-25"},
    ]), encoding="utf-8")
    assert em.desde_quando_sem_atualizar("2026-09-25", datetime.date(2026, 9, 28), str(hist)) == "2026-09-26"
    assert em.desde_quando_sem_atualizar("2026-09-28", datetime.date(2026, 9, 28), str(hist)) is None
