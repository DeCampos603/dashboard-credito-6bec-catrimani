# 🎖️ Dashboard de Crédito — 6º BEC & Operação Catrimani II

> **Sistema Independente de Monitoramento e Acompanhamento Orçamentário**  
> **6º Batalhão de Engenharia de Construção (6º BEC)** — Boa Vista / RR (UASGs **160353** e **167353**)  
> **12ª Região Militar (12ª RM)** · Comando Militar da Amazônia (CMA)  
> **Operação Catrimani II** — Ação Orçamentária **21EM** (Consolidado Multi-UGs)  
> 🌐 **URL Pública**: [https://decampos603.github.io/dashboard-credito-6bec-catrimani/](https://decampos603.github.io/dashboard-credito-6bec-catrimani/)

---

## 📌 Visão Geral

Este repositório foi construído de forma **100% autônoma e isolada**, utilizando as melhores práticas de inteligência de dados, UI militar de alta densidade e conciliação matemática anti-falha do SIAFI / Tesouro Gerencial.

### 🏛️ Estrutura de Monitoramento

1. **6º Batalhão de Engenharia de Construção (6º BEC)**:
   - **UASG 160353 (OGU)**: Recursos federais ordinários.
   - **UASG 167353 (FEx)**: Fundo do Exército.
   - Detalhamento de Provisão Recebida, Concedida, Empenhada, Liquidada, Paga e **Crédito Disponível Líquido**.
   - Conciliação exata por célula orçamentária (`Ação` · `Plano Interno - PI` · `Natureza de Despesa - ND`).
   - Rastreamento de remanejamentos de despesa (trocas de ND via notas de crédito anula/reforço).

2. **🎖️ Operação Catrimani II (Ação 21EM / Terra Indígena Yanomami)**:
   - Painel de controle inter-unidades com as **10 UGs executoras**:
     - CMDO FRON RR / 7º BIS (160352)
     - CMDO 1ª BDA INF SL (160482)
     - 6º B E CNST (160353)
     - CMDO C M A (160016)
     - 1º B LOG SL (160907)
     - 4º B AV EX (160007)
     - 1º B I S (AMV) (160006)
     - CMDO 1ª BDA INF SL - FEx (167482)
     - CMDO 12ª RM (160014)
     - PQ R MNT/12 (160021)
   - Gráfico de pizza/distribuição por Natureza de Despesa (NDs).
   - Tabela integral com **788 Notas de Crédito**, busca em tempo real por UG, ND, Favorecido, Emitente, Objeto e **exportação direta para Excel (.xlsx)**.

---

## 🗂️ Estrutura do Repositório

```text
Dashboard-Credito-6BEC-Catrimani/
├── .github/
│   └── workflows/
│       └── deploy.yml              # Pipeline CI/CD GitHub Actions (GitHub Pages diário)
├── assets/
│   └── logos/                      # Brasões heráldicos em alta resolução
│       ├── 12RM.png                # Brasão da 12ª Região Militar (Região Mendonça Furtado)
│       ├── 6BEC.png                # Brasão do 6º Batalhão de Engenharia de Construção
│       └── CATRIMANI.png           # Distintivo oficial da Operação Catrimani
├── data/
│   ├── CRÉDITO DISP 160353.xlsx    # Base de dados oficial do Tesouro Gerencial
│   └── history.json                # Histórico temporal consolidado dia a dia
├── site/
│   ├── assets/logos/               # Assets estáticos para publicação web
│   ├── data/history.json           # Espelho do histórico para consumo no front-end
│   └── index.html                  # Dashboard web autocontido (WCAG AAA + Dark/Light Mode)
├── gerar_dashboard.py              # Motor ETL ultra-rápido (openpyxl read_only ~2s) e gerador do site
├── relatorio_email_6bec.py         # Gerador de mensagens executivas (WhatsApp + E-mail HTML)
├── requirements.txt                # Dependências Python mínimas (openpyxl)
└── README.md                       # Documentação técnica completa
```

---

## 🚀 Como Executar

### 1. Pré-requisitos
- Python 3.10+
- Dependências:
  ```bash
  pip install -r requirements.txt
  ```

### 2. Gerar Mensagem Executiva Diária (E-mail e WhatsApp)
Gera o texto de alta densidade no terminal com o resumo da tomada de decisão e cria o arquivo HTML pronto para envio institucional:
```bash
python relatorio_email_6bec.py --dry-run
```
*(Opcional) Para envio direto via SMTP, informe as variáveis ou argumentos `--smtp-host`, `--user`, `--password`.*

### 3. Gerar o Dashboard Web
Lê a planilha local ou faz o download seguro via Google Drive / Google Sheets:
```bash
python gerar_dashboard.py --local "data/CRÉDITO DISP 160353.xlsx"
```
O arquivo final é salvo em `site/index.html` e pode ser aberto em qualquer navegador moderno.

---

## 📐 Regra de Negócio & Metodologia de Cálculo

$$\text{Dotação Atual} = \text{Provisão Recebida} - \text{Provisão Concedida}$$

$$\text{Crédito Disponível} = \text{Dotação Atual} - \text{Despesas Empenhadas}$$

- **Validação Anti-Falha**: O sistema confere se $\text{Recebido} - \text{Concedido} - \text{Empenhado} \equiv \text{Crédito Disponível}$ em nível de linha, célula e OM.
- **Percentual de Empenho**: $\frac{\text{Empenhado}}{\text{Dotação Atual}} \times 100\%$
- **Percentual de Liquidação**: $\frac{\text{Liquidado}}{\text{Empenhado}} \times 100\%$

---

## 🛡️ Segurança e Integridade
- Este repositório é completamente isolado dos demais painéis (`Dashboard-Credito-BCMS` e `Agente-Execucao-Orcamentaria`), não alterando qualquer dependência ou rotina já existente.
