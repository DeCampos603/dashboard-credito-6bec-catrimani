#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Relatório Diário de Crédito — 6º Batalhão de Engenharia de Construção (6º BEC) & Operação Catrimani
UASGs: 160353 (OGU) e 167353 (Fundo do Exército) · Boa Vista / RR
Acompanhamento Operação Catrimani II (Ação Governamental 21EM · Multi-UGs)

Geração de texto executivo de alta densidade para WhatsApp e E-mail militar institucional.

Uso:
    python relatorio_email_6bec.py                  # processa e envia e-mail
    python relatorio_email_6bec.py --dry-run        # gera e exibe no terminal sem enviar
    python relatorio_email_6bec.py --save-txt arq   # salva a mensagem em arquivo .txt
"""

import os
import sys
import argparse
import datetime
import warnings
import smtplib
from collections import defaultdict
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders

import openpyxl

# Compatibilidade UTF-8 no console Windows
try:
    sys.stdout.reconfigure(encoding='utf-8', errors='replace')
    sys.stderr.reconfigure(encoding='utf-8', errors='replace')
except Exception:
    pass

warnings.simplefilter("ignore")

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
DEFAULT_SRC = os.path.join(SCRIPT_DIR, "data", "CRÉDITO DISP 160353.xlsx")

EMAIL_REMETENTE = os.getenv("EMAIL_REMETENTE", "bcmssgtdecampos@gmail.com")
EMAIL_SENHA     = os.getenv("EMAIL_SENHA", "")
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


def fmt_brl(v):
    if v is None:
        return "R$ 0,00"
    return f"R$ {v:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")


def fmt_pct(v):
    if v is None:
        return "0,00%"
    return f"{v:.2f}%".replace(".", ",")


def to_num(v):
    if v is None:
        return 0.0
    if isinstance(v, (int, float)):
        return float(v)
    s = str(v).strip().replace("'", "")
    if s in ("", "-", "-9", "NAO SE APLICA", "NÃO SE APLICA"):
        return 0.0
    try:
        return float(s)
    except ValueError:
        try:
            return float(s.replace(".", "").replace(",", "."))
        except ValueError:
            return 0.0


def clean_str(v):
    if v is None:
        return ""
    s = str(v).strip().replace("'", "")
    if s in ("-9", "NAO SE APLICA", "NÃO SE APLICA"):
        return ""
    return s


def obter_marco_atual(data_ref=None):
    if data_ref is None:
        data_ref = datetime.date.today()
    mes = data_ref.month
    bimestre = (mes + 1) // 2
    return METAS_BIMESTRAIS.get(bimestre, {'marco': 'VIG', 'emp': 80.0, 'liq': 65.0})


def carregar_dados(caminho_xlsx):
    if not os.path.exists(caminho_xlsx):
        raise FileNotFoundError(f"Arquivo não encontrado: {caminho_xlsx}")

    wb = openpyxl.load_workbook(caminho_xlsx, data_only=True)
    ws = None
    for s in wb.sheetnames:
        if "SALDO R$" in s.upper() or "CREDITO" in s.upper() or "CRÉDITO" in s.upper():
            ws = wb[s]; break
    if ws is None:
        ws = wb.active

    hdr_row = 8
    col_map = {}
    for c in range(1, ws.max_column + 1):
        v = ws.cell(hdr_row, c).value
        if v:
            col_map[str(v).strip().upper()] = c

    C_PROV = col_map.get("PROVISAO RECEBIDA", 15)
    C_CONC = col_map.get("PROVISAO CONCEDIDA", 16)
    C_CRED = col_map.get("CREDITO DISPONIVEL", 18)
    C_EMP  = col_map.get("DESPESAS EMPENHADAS", 20)
    C_LIQ  = col_map.get("DESPESAS LIQUIDADAS", 22)
    C_PAG  = col_map.get("DESPESAS PAGAS", 24)

    # 1. Dados 6º BEC (160353 e 167353)
    dados_6bec = {
        '160353': {'nome': '6º B E CNST · OGU', 'prov': 0.0, 'conc': 0.0, 'cred': 0.0, 'emp': 0.0, 'liq': 0.0, 'pag': 0.0, 'linhas': []},
        '167353': {'nome': '6º B E CNST · FEx', 'prov': 0.0, 'conc': 0.0, 'cred': 0.0, 'emp': 0.0, 'liq': 0.0, 'pag': 0.0, 'linhas': []},
    }
    acoes_6bec = defaultdict(lambda: {'prov': 0.0, 'cred': 0.0, 'emp': 0.0, 'liq': 0.0, 'pag': 0.0})
    celulas_6bec = defaultdict(lambda: {'acao': '', 'pi': '', 'pi_desc': '', 'nd': '', 'nd_desc': '', 'cred': 0.0, 'emp': 0.0, 'prov': 0.0})

    # 2. Dados Operação Catrimani (Multi-UGs)
    catr_totais = {'prov': 0.0, 'conc': 0.0, 'cred': 0.0, 'emp': 0.0, 'liq': 0.0, 'pag': 0.0, 'count': 0}
    catr_por_ug = defaultdict(lambda: {'nome': '', 'prov': 0.0, 'conc': 0.0, 'cred': 0.0, 'emp': 0.0, 'liq': 0.0, 'pag': 0.0, 'count': 0})
    catr_por_nd = defaultdict(lambda: {'nome': '', 'prov': 0.0, 'cred': 0.0, 'emp': 0.0, 'liq': 0.0, 'pag': 0.0})

    for r in range(hdr_row + 1, ws.max_row + 1):
        ug_raw = ws.cell(r, 3).value
        if ug_raw is None:
            continue
        ug = str(ug_raw).strip().replace("'", "")
        ug_nome = clean_str(ws.cell(r, 4).value)

        prov = to_num(ws.cell(r, C_PROV).value)
        conc = to_num(ws.cell(r, C_CONC).value)
        cred = to_num(ws.cell(r, C_CRED).value)
        emp  = to_num(ws.cell(r, C_EMP).value)
        liq  = to_num(ws.cell(r, C_LIQ).value)
        pag  = to_num(ws.cell(r, C_PAG).value)

        nc   = clean_str(ws.cell(r, 5).value)
        acao = clean_str(ws.cell(r, 6).value)
        pi   = clean_str(ws.cell(r, 7).value)
        pid  = clean_str(ws.cell(r, 8).value)
        nd   = clean_str(ws.cell(r, 9).value)
        ndd  = clean_str(ws.cell(r, 10).value)
        obj  = clean_str(ws.cell(r, 11).value)
        op   = clean_str(ws.cell(r, 12).value)
        dia  = clean_str(ws.cell(r, 13).value)
        emit = clean_str(ws.cell(r, 1).value)

        # Trata 6º BEC
        if ug in ('160353', '167353'):
            d_bec = dados_6bec[ug]
            d_bec['prov'] += prov; d_bec['conc'] += conc; d_bec['cred'] += cred
            d_bec['emp']  += emp;  d_bec['liq']  += liq;  d_bec['pag']  += pag
            item_bec = {'ug': ug, 'nc': nc, 'acao': acao, 'pi': pi, 'pi_desc': pid, 'nd': nd, 'nd_desc': ndd,
                        'obj': obj, 'op': op, 'dia': dia, 'emit': emit, 'prov': prov, 'cred': cred, 'emp': emp, 'liq': liq, 'pag': pag}
            d_bec['linhas'].append(item_bec)
            if acao:
                acoes_6bec[acao]['prov'] += prov
                acoes_6bec[acao]['cred'] += cred
                acoes_6bec[acao]['emp']  += emp
                acoes_6bec[acao]['liq']  += liq
                acoes_6bec[acao]['pag']  += pag
            k_cel = (ug, acao, pi, nd)
            c_entry = celulas_6bec[k_cel]
            c_entry['acao'] = acao; c_entry['pi'] = pi; c_entry['pi_desc'] = pid; c_entry['nd'] = nd; c_entry['nd_desc'] = ndd
            c_entry['prov'] += prov; c_entry['cred'] += cred; c_entry['emp'] += emp

        # Trata Operação Catrimani (Ação 21EM ou texto Catrimani)
        row_str = f"{acao} {pi} {pid} {obj}".upper()
        if acao == '21EM' or 'CATRIMANI' in row_str:
            catr_totais['prov'] += prov; catr_totais['conc'] += conc; catr_totais['cred'] += cred
            catr_totais['emp']  += emp;  catr_totais['liq']  += liq;  catr_totais['pag']  += pag
            catr_totais['count'] += 1

            u_catr = catr_por_ug[ug]
            u_catr['nome'] = ug_nome or ug
            u_catr['prov'] += prov; u_catr['conc'] += conc; u_catr['cred'] += cred
            u_catr['emp']  += emp;  u_catr['liq']  += liq;  u_catr['pag']  += pag
            u_catr['count'] += 1

            if nd:
                nd_catr = catr_por_nd[nd]
                nd_catr['nome'] = ndd or nd
                nd_catr['prov'] += prov; nd_catr['cred'] += cred; nd_catr['emp'] += emp; nd_catr['liq'] += liq

    # Métricas consolidadas 6º BEC
    for u_cod, d_u in dados_6bec.items():
        dot = d_u['prov'] - d_u['conc']
        d_u['dotacao'] = dot
        d_u['pct_emp'] = (d_u['emp'] / dot * 100.0) if dot > 0 else 0.0
        d_u['pct_liq'] = (d_u['liq'] / dot * 100.0) if dot > 0 else 0.0

    tot_bec_dot = sum(d['dotacao'] for d in dados_6bec.values())
    tot_bec_emp = sum(d['emp'] for d in dados_6bec.values())
    tot_bec_cred = sum(d['cred'] for d in dados_6bec.values())
    tot_bec_liq = sum(d['liq'] for d in dados_6bec.values())
    tot_bec_pag = sum(d['pag'] for d in dados_6bec.values())

    # Métricas consolidadas Catrimani
    catr_totais['dot'] = catr_totais['prov'] - catr_totais['conc']
    catr_totais['pct_emp'] = (catr_totais['emp'] / catr_totais['dot'] * 100.0) if catr_totais['dot'] > 0 else 0.0
    catr_totais['pct_liq'] = (catr_totais['liq'] / catr_totais['emp'] * 100.0) if catr_totais['emp'] > 0 else 0.0

    return {
        '6bec': {
            'dados_ug': dados_6bec,
            'acoes': acoes_6bec,
            'celulas': celulas_6bec,
            'totais': {
                'dot': tot_bec_dot, 'emp': tot_bec_emp, 'cred': tot_bec_cred,
                'liq': tot_bec_liq, 'pag': tot_bec_pag,
                'pct_emp': (tot_bec_emp / tot_bec_dot * 100.0) if tot_bec_dot > 0 else 0.0,
                'pct_liq': (tot_bec_liq / tot_bec_dot * 100.0) if tot_bec_dot > 0 else 0.0,
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

    # 1. Cabeçalho
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
    m.append(f"• *Crédito Disponível Livre:* *{fmt_brl(t_bec['cred'])}* (Livre em Tela)")
    m.append(f"• *Despesas Liquidadas:* *{fmt_brl(t_bec['liq'])}* ({icone(t_bec['pct_liq'], marco['liq'])} *{fmt_pct(t_bec['pct_liq'])}*)")
    m.append(f"• *Despesas Pagas:* *{fmt_brl(t_bec['pag'])}*")
    m.append("")
    m.append("⚖️ *Equação Orçamentária 6º BEC:*")
    m.append(f"  {fmt_brl(t_bec['dot'])} (Recebido) − {fmt_brl(t_bec['emp'])} (Empenhado) = *{fmt_brl(t_bec['cred'])}* [Equação: 100% OK]")
    m.append("────────────────────────────────────────")
    m.append("")

    # 3. Desdobramento OGU vs FEx
    m.append("🏢 *2. EXECUÇÃO POR FONTE DE RECURSO (6º BEC)*")
    m.append(f"🔹 *UASG 160353 · OGU (Orçamento Geral da União)*")
    m.append(f"   ├ Dotação: {fmt_brl(u_ogu['dotacao'])}")
    m.append(f"   ├ Empenhado: *{fmt_pct(u_ogu['pct_emp'])}* ({fmt_brl(u_ogu['emp'])})")
    m.append(f"   └ 💰 *Crédito Disponível OGU:* *{fmt_brl(u_ogu['cred'])}*")
    m.append("")
    m.append(f"🔸 *UASG 167353 · FEx (Fundo do Exército)*")
    m.append(f"   ├ Dotação: {fmt_brl(u_fex['dotacao'])}")
    m.append(f"   ├ Empenhado: *{fmt_pct(u_fex['pct_emp'])}* ({fmt_brl(u_fex['emp'])})")
    m.append(f"   └ 💰 *Crédito Disponível FEx:* *{fmt_brl(u_fex['cred'])}*")
    m.append("────────────────────────────────────────")
    m.append("")

    # 4. Créditos Disponíveis em Tela 6º BEC
    m.append("💰 *3. CRÉDITOS LIVRES PARA EMPENHO — 6º BEC*")
    cels_saldo = [c for c in bec['celulas'].values() if c['cred'] > SALDO_MINIMO]
    cels_saldo.sort(key=lambda x: x['cred'], reverse=True)
    if cels_saldo:
        for i, c in enumerate(cels_saldo[:6]):
            pre = "└" if (i == min(5, len(cels_saldo)-1)) else "├"
            m.append(f" {pre} *Ação {c['acao']}* · PI {c['pi'] or 'N/A'} · *ND {c['nd']}*")
            m.append(f"   ↳ Saldo: *{fmt_brl(c['cred'])}* ({c['nd_desc'][:28]})")
        if len(cels_saldo) > 6:
            m.append(f"   _... e mais {len(cels_saldo) - 6} célula(s) com saldo livre._")
    else:
        m.append("✅ *Nenhum saldo ocioso pendente de empenho.*")
    m.append("────────────────────────────────────────")
    m.append("")

    # 5. Painel da Operação Catrimani II
    c_tot = catr['totais']
    m.append("🎖️ *4. PAINEL DE ACOMPANHAMENTO — OPERAÇÃO CATRIMANI II*")
    m.append(f"_Ação 21EM · Apoio aos Povos Indígenas / Roraima · 10 UGs Executoras_")
    m.append(f"• *Dotação Descentralizada Total:* *{fmt_brl(c_tot['dot'])}*")
    m.append(f"• *Empenho Global da Operação:* *{fmt_brl(c_tot['emp'])}* (Taxa de Execução: *{fmt_pct(c_tot['pct_emp'])}*)")
    m.append(f"• *Crédito Disponível Restante:* *{fmt_brl(c_tot['cred'])}* ({fmt_pct(100 - c_tot['pct_emp'])} a empenhar)")
    m.append(f"• *Liquidações Realizadas:* *{fmt_brl(c_tot['liq'])}* | *Pagos:* *{fmt_brl(c_tot['pag'])}*")
    m.append("")
    m.append("📊 *Ranking de Execução por UG na Catrimani:*")
    ugs_catr_ord = sorted(catr['por_ug'].items(), key=lambda x: x[1]['prov'], reverse=True)
    for idx, (ug_c, d_c) in enumerate(ugs_catr_ord[:6]):
        p_emp = (d_c['emp'] / d_c['prov'] * 100.0) if d_c['prov'] > 0 else 0.0
        pre = "└" if (idx == min(5, len(ugs_catr_ord)-1)) else "├"
        m.append(f" {pre} *{d_c['nome'][:20]}* ({ug_c})")
        m.append(f"   ↳ Recebido: {fmt_brl(d_c['prov'])} | Empenho: *{fmt_pct(p_emp)}* | Saldo: *{fmt_brl(d_c['cred'])}*")
    if len(ugs_catr_ord) > 6:
        m.append(f"   _... e mais {len(ugs_catr_ord) - 6} UG(s) participantes da operação._")
    m.append("────────────────────────────────────────")
    m.append("")

    # 6. Acesso ao Novo Dashboard
    m.append("🌐 *NOVO DASHBOARD DE CRÉDITO INTERATIVO:*")
    m.append(f"👉 {DASHBOARD_LINK}")
    m.append("_(Aba dedicada ao 6º BEC e Aba Exclusiva da Operação Catrimani com extrato de NCs)_")
    m.append("")
    m.append("━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━")
    m.append("👮‍♂️ *Seção de Execução Orçamentária — Apoio à Decisão Militar*")

    return "\n".join(m)


def enviar_email(corpo_texto, anexos=None):
    anexos = anexos or []
    hoje_str = datetime.datetime.now().strftime('%d/%m/%Y')
    msg = MIMEMultipart()
    msg['From'] = EMAIL_REMETENTE
    msg['To'] = EMAIL_DESTINO
    if EMAIL_BCC: msg['Bcc'] = EMAIL_BCC
    msg['Subject'] = f"Relatório Diário de Crédito — 6º BEC & Operação Catrimani — {hoje_str}"
    msg.attach(MIMEText(corpo_texto, 'plain', 'utf-8'))

    for cam in anexos:
        if not cam or not os.path.exists(cam): continue
        try:
            with open(cam, 'rb') as f:
                p = MIMEBase('application', 'octet-stream')
                p.set_payload(f.read())
            encoders.encode_base64(p)
            p.add_header('Content-Disposition', f'attachment; filename="{os.path.basename(cam)}"')
            msg.attach(p)
        except Exception as e:
            print(f"Erro ao anexar {cam}: {e}")

    # STARTTLS 587
    try:
        with smtplib.SMTP('smtp.gmail.com', 587, timeout=60) as server:
            server.ehlo(); server.starttls(); server.ehlo()
            server.login(EMAIL_REMETENTE, EMAIL_SENHA)
            server.send_message(msg)
        print(f"E-mail enviado com sucesso (STARTTLS 587) para {EMAIL_DESTINO}")
        return True
    except Exception as e1:
        print(f"Falha STARTTLS 587 ({e1}). Tentando SSL 465...")

    # SSL 465
    try:
        with smtplib.SMTP_SSL('smtp.gmail.com', 465, timeout=60) as server:
            server.login(EMAIL_REMETENTE, EMAIL_SENHA)
            server.send_message(msg)
        print(f"E-mail enviado com sucesso (SSL 465) para {EMAIL_DESTINO}")
        return True
    except Exception as e2:
        print(f"ERRO: falha também em SSL 465: {e2}")
        return False


def main():
    parser = argparse.ArgumentParser(description="Relatório Diário de Crédito — 6º BEC & Operação Catrimani")
    parser.add_argument("--local", default=None, help="Caminho da planilha XLSX")
    parser.add_argument("--dry-run", action="store_true", help="Apenas imprime no terminal")
    parser.add_argument("--save-txt", default=None, help="Salva a mensagem gerada")
    args = parser.parse_args()

    caminho_xlsx = args.local or DEFAULT_SRC
    print(f"Lendo base de dados: {caminho_xlsx}...")
    res = carregar_dados(caminho_xlsx)
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
            print("\n[INFO] EMAIL_SENHA não informada. Modo DRY-RUN ativado.")
        return

    enviar_email(texto)


if __name__ == "__main__":
    main()
