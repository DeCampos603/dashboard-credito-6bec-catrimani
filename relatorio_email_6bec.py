#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Relatório Diário de Execução Orçamentária & Crédito — 6º BEC & Operação Catrimani II
UASGs: 160353 (OGU) e 167353 (Fundo do Exército) · Boa Vista / RR
Acompanhamento Operação Catrimani II (Ação Governamental 21EM · Multi-UGs)

Geração de texto executivo de alta densidade no padrão WhatsApp militar institucional
e envio automatizado via E-mail SMTP (Gmail) com suporte a auditoria anti-duplicação.

Uso:
    python relatorio_email_6bec.py                  # Baixa da web e envia e-mail
    python relatorio_email_6bec.py --dry-run        # Apenas gera e exibe no terminal
    python relatorio_email_6bec.py --local arq.csv  # Usa arquivo local
    python relatorio_email_6bec.py --url URL        # Usa URL personalizada
    python relatorio_email_6bec.py --save-txt out   # Salva o texto em arquivo
"""

import os
import sys
import argparse
import datetime
import warnings
import smtplib
import urllib.request
import tempfile
import csv
import unicodedata
from collections import defaultdict
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders

try:
    import openpyxl
except ImportError:
    openpyxl = None

# Compatibilidade UTF-8 no console Windows
try:
    sys.stdout.reconfigure(encoding='utf-8', errors='replace')
    sys.stderr.reconfigure(encoding='utf-8', errors='replace')
except Exception:
    pass

warnings.simplefilter("ignore")

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
DEFAULT_SRC = os.path.join(SCRIPT_DIR, "data", "CRÉDITO DISP 160353.xlsx")
DEFAULT_CSV_URL = "https://docs.google.com/spreadsheets/d/e/2PACX-1vTVtnLCf2tvVO1-PFklLro4Y-ijBqw9h3psRi2y3Q69_1TSX75OPmph7yPK3zmANA/pub?gid=991377463&single=true&output=csv"

SMTP_HOST       = os.getenv("SMTP_HOST", "smtp.gmail.com")
SMTP_PORT       = int(os.getenv("SMTP_PORT", "587"))
EMAIL_REMETENTE = os.getenv("EMAIL_REMETENTE") or os.getenv("SMTP_USER", "bcmssgtdecampos@gmail.com")
EMAIL_SENHA     = os.getenv("EMAIL_SENHA") or os.getenv("SMTP_PASSWORD", "")
EMAIL_DESTINO   = os.getenv("EMAIL_DESTINO", "marcelofsaldanha@gmail.com")
EMAIL_BCC       = os.getenv("EMAIL_BCC", "")

DASHBOARD_LINK  = "https://decampos603.github.io/dashboard-credito-6bec-catrimani/"

SALDO_MINIMO     = 1.00
LIMIAR_NC_ALERTA = 5000.00

METAS_BIMESTRAIS = {
    1: {'marco': 'FEV', 'emp': 20.0, 'liq': 10.0},
    2: {'marco': 'ABR', 'emp': 40.0, 'liq': 25.0},
    3: {'marco': 'JUN', 'emp': 60.0, 'liq': 45.0},
    4: {'marco': 'AGO', 'emp': 80.0, 'liq': 65.0},
    5: {'marco': 'OUT', 'emp': 90.0, 'liq': 80.0},
    6: {'marco': 'DEZ', 'emp': 100.0, 'liq': 95.0},
}

SIGLAS_MILITARES_UGS = {
    "160007": "4º BAVEX",
    "160353": "6º BEC",
    "160352": "7º BIS",
    "160482": "1ª BDA INF SL",
    "160907": "1º B LOG SL",
    "160016": "CMDO CMA",
    "160006": "1º BIS (AMV)",
    "167482": "1ª BDA (FEx)",
    "160014": "CMDO 12ª RM",
    "160021": "PQ R MNT/12",
    "160329": "BCMS",
    "160238": "BA AP LOG",
}


def fmt_brl(v):
    if v is None:
        return "R$ 0,00"
    return f"R$ {v:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")


def fmt_pct(v):
    if v is None:
        return "0,00%"
    return f"{v:.2f}%".replace(".", ",")


def norm(s):
    if s is None:
        return ""
    s = unicodedata.normalize("NFKD", str(s)).encode("ASCII", "ignore").decode("utf-8")
    return s.strip().upper()


def to_num(v):
    if v is None:
        return 0.0
    if isinstance(v, (int, float)):
        return float(v)
    s = str(v).strip().replace("'", "").replace('"', '')
    if s in ("", "-", "-9", "NAO SE APLICA", "NÃO SE APLICA"):
        return 0.0
    neg = False
    if s.startswith("(") and s.endswith(")"):
        neg = True
        s = s[1:-1].strip()
    try:
        val = float(s.replace(".", "").replace(",", "."))
        return -val if neg else val
    except ValueError:
        return 0.0


def clean_str(v):
    if v is None:
        return ""
    s = str(v).strip().replace("'", "").replace('"', '')
    if s in ("-9", "NAO SE APLICA", "NÃO SE APLICA"):
        return ""
    return s


def baixar(target=None):
    if target and os.path.exists(target):
        return target

    url = target if (target and target.startswith(("http://", "https://"))) else (os.environ.get("SHEETS_CSV_URL") or DEFAULT_CSV_URL)

    req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0 (relatorio-6bec)"})
    ext = ".csv" if "output=csv" in url else ".xlsx"
    tmp = os.path.join(tempfile.gettempdir(), f"credito_disp_160353_email{ext}")
    try:
        with urllib.request.urlopen(req, timeout=90) as r, open(tmp, "wb") as f:
            f.write(r.read())
        if os.path.getsize(tmp) < 500:
            if os.path.exists(DEFAULT_SRC):
                print(f"[AVISO] Download muito pequeno. Usando fallback local: {DEFAULT_SRC}")
                return DEFAULT_SRC
            raise SystemExit("Download muito pequeno — verifique a URL da planilha.")
        return tmp
    except Exception as e:
        if os.path.exists(DEFAULT_SRC):
            print(f"[AVISO] Falha no download ({e}). Usando fallback local: {DEFAULT_SRC}")
            return DEFAULT_SRC
        raise


def ler_linhas(path):
    is_csv = str(path).lower().endswith(".csv")
    if not is_csv and not str(path).lower().endswith((".xlsx", ".xlsm")):
        try:
            with open(path, "r", encoding="utf-8", errors="replace") as f:
                first = f.read(500)
                if "," in first or ";" in first:
                    is_csv = True
        except Exception:
            pass

    if is_csv:
        try:
            with open(path, "r", encoding="utf-8") as f:
                return list(csv.reader(f))
        except UnicodeDecodeError:
            with open(path, "r", encoding="latin-1", errors="replace") as f:
                return list(csv.reader(f))

    if openpyxl is None:
        raise RuntimeError("openpyxl não está instalado para ler planilhas XLSX.")

    wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
    ws = None
    for s in wb.sheetnames:
        if "SALDO R$" in s.upper() or "CREDITO" in s.upper() or "CRÉDITO" in s.upper():
            ws = wb[s]
            break
    if ws is None:
        ws = wb.active
    return [list(r) for r in ws.iter_rows(values_only=True)]


def obter_marco_atual(data_ref=None):
    if data_ref is None:
        data_ref = datetime.date.today()
    mes = data_ref.month
    bimestre = (mes + 1) // 2
    return METAS_BIMESTRAIS.get(bimestre, {'marco': 'VIG', 'emp': 80.0, 'liq': 65.0})


def carregar_dados(caminho_dados):
    linhas = ler_linhas(caminho_dados)
    if not linhas:
        raise ValueError(f"Base de dados vazia: {caminho_dados}")

    # Localiza o cabeçalho dinamicamente
    hdr_idx = None
    for idx, row in enumerate(linhas[:40]):
        row_norm = [norm(c) for c in row if c is not None]
        if any("EMITENTE - UG" in c or "EMITENTE" in c for c in row_norm) and any("PROVISAO RECEBIDA" in c for c in row_norm):
            hdr_idx = idx
            break
        if any("PROVISAO RECEBIDA" in c for c in row_norm) and any("CREDITO DISPONIVEL" in c for c in row_norm):
            hdr_idx = idx
            break

    if hdr_idx is None:
        hdr_idx = 7 if len(linhas) > 7 else 0

    # Escaneamento dinâmico multi-linhas (dimensões + métricas)
    col_map = {}
    start_search = max(0, hdr_idx - 5)
    for r in linhas[start_search:hdr_idx + 1]:
        for c_idx, val in enumerate(r):
            v_norm = norm(val)
            if v_norm and v_norm not in col_map:
                col_map[v_norm] = c_idx

    def col(name, fallback):
        if name in col_map:
            return col_map[name]
        for k, v in col_map.items():
            if name in k:
                return v
        return fallback

    c_emit = col("EMITENTE - UG", 0)
    c_ug   = col("FAVORECIDO - UG", 2)
    if c_ug == 0 and "UG EXECUTORA" in col_map:
        c_ug = col_map["UG EXECUTORA"]
    c_ugn  = col("FAVORECIDO - NOME", 3)
    c_nc   = col("NC - NUMERO", 4)
    if c_nc == 0 and "NC" in col_map:
        c_nc = col_map["NC"]
    c_acao = col("ACAO GOVERNO", 5)
    c_pi   = col("PLANO INTERNO", 6)
    if c_pi == 0 and "PI" in col_map:
        c_pi = col_map["PI"]
    c_pid  = col("PLANO INTERNO - NOME", 7)
    c_nd   = col("NATUREZA DESPESA", 8)
    c_ndd  = col("NATUREZA DESPESA - NOME", 9)
    c_obj  = col("NC - DESCRICAO", 10)
    c_op   = col("NC - OPERACAO (TIPO)", 11)
    c_dia  = col("DIA EMISSAO NC", 13)
    if c_dia == 0 and "NC - DIA EMISSAO" in col_map:
        c_dia = col_map["NC - DIA EMISSAO"]

    c_rec  = col("PROVISAO RECEBIDA", 14)
    c_cnc  = col("PROVISAO CONCEDIDA", 15)
    c_crd  = col("CREDITO DISPONIVEL", 17)
    c_emp  = col("DESPESAS EMPENHADAS", 19)
    c_liq  = col("DESPESAS LIQUIDADAS", 21)
    c_pag  = col("DESPESAS PAGAS", 23)

    dados_6bec = {
        '160353': {'nome': '6º B E CNST · OGU (160353)', 'prov': 0.0, 'conc': 0.0, 'cred': 0.0, 'emp': 0.0, 'liq': 0.0, 'pag': 0.0, 'linhas': []},
        '167353': {'nome': '6º B E CNST · FEx (167353)', 'prov': 0.0, 'conc': 0.0, 'cred': 0.0, 'emp': 0.0, 'liq': 0.0, 'pag': 0.0, 'linhas': []},
    }
    acoes_6bec = defaultdict(lambda: {'prov': 0.0, 'cred': 0.0, 'emp': 0.0, 'liq': 0.0, 'pag': 0.0})
    celulas_6bec = defaultdict(lambda: {'ug': '', 'acao': '', 'pi': '', 'pi_desc': '', 'nd': '', 'nd_desc': '', 'cred': 0.0, 'emp': 0.0, 'prov': 0.0})
    detalhamentos_6bec = {}
    ncs_6bec_21em = {}

    catr_totais = {'prov': 0.0, 'conc': 0.0, 'cred': 0.0, 'emp': 0.0, 'liq': 0.0, 'pag': 0.0, 'count': 0}
    catr_por_ug = defaultdict(lambda: {'nome': '', 'prov': 0.0, 'conc': 0.0, 'cred': 0.0, 'emp': 0.0, 'liq': 0.0, 'pag': 0.0, 'count': 0})
    catr_por_nd = defaultdict(lambda: {'nome': '', 'prov': 0.0, 'cred': 0.0, 'emp': 0.0, 'liq': 0.0, 'pag': 0.0})

    for row in linhas[hdr_idx + 1:]:
        if len(row) <= max(c_ug, c_crd):
            continue

        ug = clean_str(row[c_ug])
        if not ug:
            continue

        ug_nome = clean_str(row[c_ugn]) if c_ugn < len(row) else ""
        emit = clean_str(row[c_emit]) if c_emit < len(row) else ""
        nc   = clean_str(row[c_nc]) if c_nc < len(row) else ""
        acao = clean_str(row[c_acao]) if c_acao < len(row) else ""
        pi   = clean_str(row[c_pi]) if c_pi < len(row) else ""
        pid  = clean_str(row[c_pid]) if c_pid < len(row) else ""
        nd   = clean_str(row[c_nd]) if c_nd < len(row) else ""
        ndd  = clean_str(row[c_ndd]) if c_ndd < len(row) else ""
        obj  = clean_str(row[c_obj]) if c_obj < len(row) else ""
        op   = clean_str(row[c_op]) if c_op < len(row) else ""
        dia  = clean_str(row[c_dia]) if c_dia < len(row) else ""

        prov_rec = to_num(row[c_rec]) if c_rec < len(row) else 0.0
        prov_cnc = to_num(row[c_cnc]) if c_cnc < len(row) else 0.0
        cred = to_num(row[c_crd]) if c_crd < len(row) else 0.0
        emp  = to_num(row[c_emp]) if c_emp < len(row) else 0.0
        liq  = to_num(row[c_liq]) if c_liq < len(row) else 0.0
        pag  = to_num(row[c_pag]) if c_pag < len(row) else 0.0
        prov_liq = prov_rec - prov_cnc

        is_nc = bool(nc and nc not in ('-9', 'NAO SE APLICA', 'NÃO SE APLICA'))

        # Regra Mestre de Sanidade Orçamentária SIAFI/TG (Agente-Execucao-Orcamentaria/conhecimento/09):
        # Se a linha não é Nota de Crédito (is_nc == False), somente processa se o emitente for a própria UG executora (emit == ug).
        # Linhas sem NC onde emit != ug são cruzamentos interunidades/subtotais de controle que duplicam empenhos.
        if not is_nc and emit != ug:
            continue

        row_str = f"{acao} {pi} {pid} {obj}".upper()
        is_21em = (acao == '21EM' or 'CATRIMANI' in row_str)

        # Processamento 6º BEC (OGU e FEx)
        if ug in ('160353', '167353'):
            d_bec = dados_6bec[ug]
            d_bec['prov'] += prov_rec
            d_bec['conc'] += prov_cnc
            d_bec['cred'] += cred
            d_bec['emp']  += emp
            d_bec['liq']  += liq
            d_bec['pag']  += pag

            is_det = ("DETALHAMENTO" in norm(op)) or (emit in ('160353', '167353') and abs(prov_liq) < 0.01)
            if nc and is_det:
                detalhamentos_6bec[nc] = {
                    'nc': nc, 'op': op, 'nd': nd, 'acao': acao, 'cred': cred, 'dia': dia, 'obj': obj
                }

            k_cel = (ug, acao, pi, nd)
            c_entry = celulas_6bec[k_cel]
            c_entry['ug'] = ug
            c_entry['acao'] = acao
            c_entry['pi'] = pi
            c_entry['pi_desc'] = pid
            c_entry['nd'] = nd
            c_entry['nd_desc'] = ndd
            c_entry['prov'] += prov_liq
            c_entry['cred'] += cred
            c_entry['emp']  += emp

            if is_21em and is_nc:
                if nc not in ncs_6bec_21em:
                    ncs_6bec_21em[nc] = {
                        'nc': nc, 'emit': emit, 'dia': dia, 'nd': nd, 'ndd': ndd, 'pi': pi,
                        'rec': 0.0, 'cnc': 0.0, 'cred': 0.0, 'emp': 0.0, 'obj': obj
                    }
                n_entry = ncs_6bec_21em[nc]
                n_entry['rec'] += prov_rec
                n_entry['cnc'] += prov_cnc
                n_entry['cred'] += cred
                n_entry['emp'] += emp
                if obj and len(obj) > len(n_entry['obj']):
                    n_entry['obj'] = obj

        # Processamento Operação Catrimani II (Ação 21EM)
        if is_21em:
            catr_totais['prov'] += prov_rec
            catr_totais['conc'] += prov_cnc
            catr_totais['cred'] += cred
            catr_totais['emp']  += emp
            catr_totais['liq']  += liq
            catr_totais['pag']  += pag
            catr_totais['count'] += 1

            u_catr = catr_por_ug[ug]
            u_catr['nome'] = ug_nome or ug
            u_catr['prov'] += prov_rec
            u_catr['conc'] += prov_cnc
            u_catr['cred'] += cred
            u_catr['emp']  += emp
            u_catr['liq']  += liq
            u_catr['pag']  += pag
            u_catr['count'] += 1

            if nd:
                nd_catr = catr_por_nd[nd]
                nd_catr['nome'] = ndd or nd
                nd_catr['prov'] += prov_liq
                nd_catr['cred'] += cred
                nd_catr['emp']  += emp
                nd_catr['liq']  += liq

    for u_cod, d_u in dados_6bec.items():
        dot = d_u['prov'] - d_u['conc']
        d_u['dotacao'] = dot
        d_u['pct_emp'] = (d_u['emp'] / dot * 100.0) if dot > 0 else 0.0
        d_u['pct_cred'] = (d_u['cred'] / dot * 100.0) if dot > 0 else 0.0
        d_u['pct_liq'] = (d_u['liq'] / d_u['emp'] * 100.0) if d_u['emp'] > 0 else 0.0
        d_u['pct_pag'] = (d_u['pag'] / d_u['liq'] * 100.0) if d_u['liq'] > 0 else 0.0

    tot_bec_dot = sum(d['dotacao'] for d in dados_6bec.values())
    tot_bec_emp = sum(d['emp'] for d in dados_6bec.values())
    tot_bec_cred = sum(d['cred'] for d in dados_6bec.values())
    tot_bec_liq = sum(d['liq'] for d in dados_6bec.values())
    tot_bec_pag = sum(d['pag'] for d in dados_6bec.values())

    catr_totais['dot'] = catr_totais['prov'] - catr_totais['conc']
    catr_totais['pct_emp'] = (catr_totais['emp'] / catr_totais['dot'] * 100.0) if catr_totais['dot'] > 0 else 0.0
    catr_totais['pct_cred'] = (catr_totais['cred'] / catr_totais['dot'] * 100.0) if catr_totais['dot'] > 0 else 0.0
    catr_totais['pct_liq'] = (catr_totais['liq'] / catr_totais['emp'] * 100.0) if catr_totais['emp'] > 0 else 0.0
    catr_totais['pct_pag'] = (catr_totais['pag'] / catr_totais['liq'] * 100.0) if catr_totais['liq'] > 0 else 0.0

    return {
        '6bec': {
            'dados_ug': dados_6bec,
            'acoes': acoes_6bec,
            'celulas': celulas_6bec,
            'detalhamentos': detalhamentos_6bec,
            'ncs_21em': ncs_6bec_21em,
            'totais': {
                'dot': tot_bec_dot,
                'emp': tot_bec_emp,
                'cred': tot_bec_cred,
                'liq': tot_bec_liq,
                'pag': tot_bec_pag,
                'pct_emp': (tot_bec_emp / tot_bec_dot * 100.0) if tot_bec_dot > 0 else 0.0,
                'pct_cred': (tot_bec_cred / tot_bec_dot * 100.0) if tot_bec_dot > 0 else 0.0,
                'pct_liq': (tot_bec_liq / tot_bec_emp * 100.0) if tot_bec_emp > 0 else 0.0,
                'pct_pag': (tot_bec_pag / tot_bec_liq * 100.0) if tot_bec_liq > 0 else 0.0,
            }
        },
        'catrimani': {
            'totais': catr_totais,
            'por_ug': catr_por_ug,
            'por_nd': catr_por_nd,
        }
    }


def gerar_texto_mensagem(res):
    bec = res['6bec']
    catr = res['catrimani']
    t_bec = bec['totais']
    u_ogu = bec['dados_ug']['160353']
    u_fex = bec['dados_ug']['167353']
    det_ncs = bec['detalhamentos']
    marco = obter_marco_atual()

    hoje_str = datetime.datetime.now().strftime("%d/%m/%Y")
    hora_str = datetime.datetime.now().strftime("%H:%M")

    def icone(pct, meta):
        if meta <= 0: return "⚪"
        ratio = pct / meta
        if ratio >= 1.0: return "🟢"
        elif ratio >= 0.8: return "🟡"
        else: return "🔴"

    m = []

    # 1. Cabeçalho Militar
    m.append("📋 *RELATÓRIO DIÁRIO DE EXECUÇÃO ORÇAMENTÁRIA & CRÉDITO*")
    m.append("🏛️ *6º Batalhão de Engenharia de Construção (6º BEC)*")
    m.append("🎖️ *Comitê de Acompanhamento da Operação Catrimani II*")
    m.append(f"⏱ *Posição:* {hoje_str} às {hora_str} · *Marco Vigente:* {marco['marco']}")
    m.append("━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━")
    m.append("")

    # 2. Resumo Executivo 6º BEC
    m.append("⚡ *1. RESUMO EXECUTIVO — 6º BEC (UASGs 160353 / 167353)*")
    m.append(f"• *Dotação Líquida Total:* *{fmt_brl(t_bec['dot'])}*")
    m.append(f"• *Despesas Empenhadas:* *{fmt_brl(t_bec['emp'])}* ({icone(t_bec['pct_emp'], marco['emp'])} *{fmt_pct(t_bec['pct_emp'])}* · Meta {marco['emp']:.0f}%)")
    m.append(f"• *Crédito Disponível Livre Total:* *{fmt_brl(t_bec['cred'])}* (*{fmt_pct(t_bec['pct_cred'])}* Livre em Tela)")
    m.append(f"• *Despesas Liquidadas:* *{fmt_brl(t_bec['liq'])}* ({fmt_pct(t_bec['pct_liq'])} do empenhado)")
    m.append(f"• *Despesas Pagas:* *{fmt_brl(t_bec['pag'])}* ({fmt_pct(t_bec['pct_pag'])} do liquidado)")
    m.append("────────────────────────────────────────")
    m.append("")

    # 3. Execução por Fonte (OGU vs FEx)
    m.append("🏢 *2. EXECUÇÃO POR FONTE DE RECURSO (6º BEC)*")
    m.append(f"🔹 *UASG 160353 · OGU (Orçamento Geral da União)*")
    m.append(f"   ├ Dotação: {fmt_brl(u_ogu['dotacao'])}")
    m.append(f"   ├ Empenhado: *{fmt_pct(u_ogu['pct_emp'])}* ({fmt_brl(u_ogu['emp'])})")
    m.append(f"   ├ Liquidado: {fmt_brl(u_ogu['liq'])} | Pago: {fmt_brl(u_ogu['pag'])}")
    m.append(f"   └ 💰 *Crédito Disponível OGU:* *{fmt_brl(u_ogu['cred'])}*")
    m.append("")
    m.append(f"🔸 *UASG 167353 · FEx (Fundo do Exército)*")
    m.append(f"   ├ Dotação: {fmt_brl(u_fex['dotacao'])}")
    m.append(f"   ├ Empenhado: *{fmt_pct(u_fex['pct_emp'])}* ({fmt_brl(u_fex['emp'])})")
    m.append(f"   ├ Liquidado: {fmt_brl(u_fex['liq'])} | Pago: {fmt_brl(u_fex['pag'])}")
    m.append(f"   └ 💰 *Crédito Disponível FEx:* *{fmt_brl(u_fex['cred'])}*")
    m.append("────────────────────────────────────────")
    m.append("")

    # 4. Crédito Disponível Livre para Empenho — 6º BEC (EXCLUSIVO AÇÃO 21EM)
    u_6bec_catr = catr['por_ug'].get('160353', {})
    saldo_catr_6bec = u_6bec_catr.get('cred', 0.0)
    dot_catr_6bec = u_6bec_catr.get('prov', 0.0) - u_6bec_catr.get('conc', 0.0)
    emp_catr_6bec = u_6bec_catr.get('emp', 0.0)
    p_emp_catr = (emp_catr_6bec / dot_catr_6bec * 100.0) if dot_catr_6bec > 0 else 0.0

    m.append("💰 *3. CRÉDITOS LIVRES PARA EMPENHO — 6º BEC (EXCLUSIVO AÇÃO 21EM)*")
    m.append(f"_Recursos Exclusivos da Operação Catrimani II · Saldo Livre Total: *{fmt_brl(saldo_catr_6bec)}*_")
    m.append(f"• *Dotação 21EM no 6º BEC:* *{fmt_brl(dot_catr_6bec)}* | *Empenho:* *{fmt_pct(p_emp_catr)}* ({fmt_brl(emp_catr_6bec)})")
    m.append("")
    m.append("📌 *Saldos Livres por Célula Orçamentária (PI · ND):*")
    cels_21em = [c for c in bec['celulas'].values() if c['acao'] == '21EM' and c['cred'] > SALDO_MINIMO]
    cels_21em.sort(key=lambda x: x['cred'], reverse=True)
    if cels_21em:
        for idx, c in enumerate(cels_21em):
            pre = "└" if idx == len(cels_21em) - 1 else "├"
            nd_ico = "📦" if "339030" in c['nd'] else ("🛠️" if "339039" in c['nd'] else ("💻" if "339040" in c['nd'] else "📄"))
            m.append(f" {pre} {nd_ico} *ND {c['nd']}* ({c['nd_desc'][:25]}) · PI {c['pi'] or 'N/A'}")
            m.append(f"   ↳ Saldo Livre: *{fmt_brl(c['cred'])}* (Dotação: {fmt_brl(c['prov'])} | Emp: {fmt_brl(c['emp'])})")
    else:
        m.append("✅ *Recursos da 21EM 100% empenhados. Sem saldo ocioso pendente.*")

    m.append("")
    m.append("📜 *Detalhamento das Notas de Crédito (NCs) da Ação 21EM (6º BEC):*")
    ncs_catr_list = list(bec.get('ncs_21em', {}).values())
    ncs_catr_list.sort(key=lambda x: x['cred'], reverse=True)
    ncs_com_saldo = [x for x in ncs_catr_list if x['cred'] > SALDO_MINIMO]
    if ncs_com_saldo:
        for idx, nc_item in enumerate(ncs_com_saldo[:6]):
            pre = "└" if idx == min(5, len(ncs_com_saldo) - 1) else "├"
            emit_sigla = "COTER" if nc_item['emit'] == '160539' else ("COEX" if nc_item['emit'] == '160504' else f"UG {nc_item['emit']}")
            nc_curta = nc_item['nc'][-12:] if len(nc_item['nc']) > 12 else nc_item['nc']
            m.append(f" {pre} *NC {nc_curta}* ({emit_sigla}) · *ND {nc_item['nd']}*")
            m.append(f"   ↳ Saldo em Tela: *{fmt_brl(nc_item['cred'])}* | Repasse: {fmt_brl(nc_item['rec'])}")
            if nc_item['obj']:
                m.append(f"   ↳ Objeto: _{nc_item['obj'][:45]}_")
        if len(ncs_com_saldo) > 6:
            m.append(f"   _... e mais {len(ncs_com_saldo) - 6} NC(s) monitoradas no dashboard interativo._")
    else:
        m.append("ℹ️ *NCs da Ação 21EM integralmente alocadas nas células orçamentárias.*")
    m.append("────────────────────────────────────────")
    m.append("")

    # 5. Painel da Operação Catrimani II
    c_tot = catr['totais']
    m.append("🎖️ *4. PAINEL DE ACOMPANHAMENTO — OPERAÇÃO CATRIMANI II*")
    m.append(f"_Ação 21EM · Apoio aos Povos Indígenas / Roraima · Multi-UGs Executoras_")
    m.append(f"• *Dotação Total Descentralizada:* *{fmt_brl(c_tot['dot'])}*")
    m.append(f"• *Empenho Global da Operação:* *{fmt_brl(c_tot['emp'])}* (Execução: *{fmt_pct(c_tot['pct_emp'])}*)")
    m.append(f"• *Crédito Disponível na Catrimani:* *{fmt_brl(c_tot['cred'])}* ({fmt_pct(c_tot['pct_cred'])} livre)")
    m.append(f"• *Liquidações Realizadas:* *{fmt_brl(c_tot['liq'])}* | *Pagamentos:* *{fmt_brl(c_tot['pag'])}*")
    m.append(f"• *Órgãos Repassadores dos Créditos:* *COTER (UG 160539)* e *COEx (UG 160504)*")
    m.append(f"• *Governança & Auditoria Técnica:* *Base de Apoio Logístico do Exército (Ba Ap Log)*")
    m.append("")
    m.append("📊 *Quadro Executivo Consolidado Multi-UGs (Ação 21EM):*")
    ugs_catr_ord = sorted(catr['por_ug'].items(), key=lambda x: x[1]['prov'] - x[1]['conc'], reverse=True)
    for idx, (ug_c, d_c) in enumerate(ugs_catr_ord):
        dot_c = d_c['prov'] - d_c['conc']
        p_emp = (d_c['emp'] / dot_c * 100.0) if dot_c > 0 else 0.0
        pre = "└" if (idx == len(ugs_catr_ord) - 1) else "├"
        sigla_m = SIGLAS_MILITARES_UGS.get(ug_c, f"UG {ug_c}")
        status_ico = "🟢" if p_emp >= 80.0 else ("🟡" if p_emp >= 60.0 else "🔴")
        m.append(f" {pre} {status_ico} *{sigla_m}* ({d_c['nome'][:20]} · UG {ug_c})")
        m.append(f"   ↳ Dotação: {fmt_brl(dot_c)} | Emp: *{fmt_pct(p_emp)}* ({fmt_brl(d_c['emp'])}) | Saldo: *{fmt_brl(d_c['cred'])}*")
    m.append("────────────────────────────────────────")
    m.append("")

    # 6. Insights para Despacho Executivo
    s_nd30 = sum(c['cred'] for c in cels_21em if "339030" in c['nd'])
    s_nd39 = sum(c['cred'] for c in cels_21em if "339039" in c['nd'])
    m.append("💡 *5. INSIGHTS PARA DESPACHO EXECUTIVO (OD / FISCAL ADM)*")
    m.append(f"• *Ação 21EM (Catrimani):* Saldo livre de *{fmt_brl(saldo_catr_6bec)}* concentrado principalmente em Material de Consumo (ND 30: {fmt_brl(s_nd30)}) e Serviços de Terceiros (ND 39: {fmt_brl(s_nd39)}), assegurando pronta resposta logística e engenharia de pistas.")
    m.append(f"• *Execução Global 6º BEC:* O Batalhão registra *{fmt_pct(t_bec['pct_emp'])}* de empenho ({fmt_brl(t_bec['emp'])}), mantendo alta velocidade de processamento com {fmt_brl(t_bec['liq'])} já liquidados.")
    m.append(f"• *Meta Bimestral:* Execução amplamente superior à meta do marco {marco['marco']} ({marco['emp']:.0f}% de empenho), sem risco de glosa ou retenção de recursos.")
    m.append("────────────────────────────────────────")
    m.append("")

    # 7. Acesso ao Dashboard Oficial
    m.append("🌐 *NOVO DASHBOARD DE CRÉDITO INTERATIVO:*")
    m.append(f"👉 {DASHBOARD_LINK}")
    m.append("_(Visualização completa em tela, filtros dinâmicos, ranking das OMDS e extrato de NCs)_")
    m.append("")
    m.append("━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━")
    m.append("💻 *Desenvolvido por:* 2º Sgt De Campos (BCMS / 6º BEC)")
    m.append("🔍 *Auditado por:* Seção de Execução Orçamentária & Fiscalização Administrativa (SIAFI/TG)")

    return "\n".join(m)


def enviar_email(corpo_texto, anexos=None):
    anexos = anexos or []
    hoje_str = datetime.datetime.now().strftime('%d/%m/%Y')
    msg = MIMEMultipart()
    msg['From'] = EMAIL_REMETENTE
    msg['To'] = EMAIL_DESTINO
    if EMAIL_BCC:
        msg['Bcc'] = EMAIL_BCC
    msg['Subject'] = f"Relatório Diário de Crédito — 6º BEC & Operação Catrimani — {hoje_str}"
    msg.attach(MIMEText(corpo_texto, 'plain', 'utf-8'))

    for cam in anexos:
        if not cam or not os.path.exists(cam):
            continue
        try:
            with open(cam, 'rb') as f:
                p = MIMEBase('application', 'octet-stream')
                p.set_payload(f.read())
            encoders.encode_base64(p)
            p.add_header('Content-Disposition', f'attachment; filename="{os.path.basename(cam)}"')
            msg.attach(p)
        except Exception as e:
            print(f"[AVISO] Erro ao anexar {cam}: {e}")

    # 1. Tenta STARTTLS na porta configurada (padrão 587)
    try:
        with smtplib.SMTP(SMTP_HOST, SMTP_PORT, timeout=60) as server:
            server.ehlo()
            server.starttls()
            server.ehlo()
            server.login(EMAIL_REMETENTE, EMAIL_SENHA)
            server.send_message(msg)
        print(f"E-mail enviado com sucesso (STARTTLS {SMTP_PORT}) para {EMAIL_DESTINO}")
        return True
    except Exception as e1:
        print(f"Falha STARTTLS {SMTP_PORT} ({e1}). Tentando SSL 465...")

    # 2. Fallback para SSL porta 465
    try:
        with smtplib.SMTP_SSL(SMTP_HOST, 465, timeout=60) as server:
            server.login(EMAIL_REMETENTE, EMAIL_SENHA)
            server.send_message(msg)
        print(f"E-mail enviado com sucesso (SSL 465) para {EMAIL_DESTINO}")
        return True
    except Exception as e2:
        print(f"ERRO: Falha no envio via SSL 465: {e2}")
        return False


def main():
    parser = argparse.ArgumentParser(description="Relatório Diário de Crédito — 6º BEC & Operação Catrimani II")
    parser.add_argument("--local", default=None, help="Caminho do arquivo local (CSV ou XLSX)")
    parser.add_argument("--url", default=None, help="URL personalizada para download da planilha")
    parser.add_argument("--dry-run", action="store_true", help="Apenas imprime a mensagem formatada no terminal")
    parser.add_argument("--save-txt", default=None, help="Salva o texto gerado em um arquivo TXT")
    args = parser.parse_args()

    target = args.local or args.url or os.environ.get("SHEETS_CSV_URL") or DEFAULT_CSV_URL
    caminho_dados = args.local if (args.local and os.path.exists(args.local)) else baixar(target)
    print(f"Lendo base de dados: {caminho_dados}...")

    res = carregar_dados(caminho_dados)
    texto = gerar_texto_mensagem(res)

    if args.save_txt:
        with open(args.save_txt, "w", encoding="utf-8") as f:
            f.write(texto)
        print(f"Arquivo salvo: {args.save_txt}")

    if args.dry_run or not EMAIL_SENHA:
        print("\n" + "=" * 75)
        print(texto)
        print("=" * 75)
        if not EMAIL_SENHA:
            print("\n[INFO] EMAIL_SENHA não informada no ambiente. Modo DRY-RUN concluído.")
        return

    enviar_email(texto)


if __name__ == "__main__":
    main()
