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

## ❓ Perguntas Frequentes

### "O 6º BEC diz que não há crédito disponível da Catrimani, mas o site mostra um saldo alto. Por quê?"

O número que o site/e-mail mostra em **"Créditos Livres para Empenho — 6º BEC (Exclusivo Ação 21EM)"** é **exclusivo da Operação Catrimani II (Ação 21EM)** — não é o orçamento geral do Batalhão, e não pode ser usado para qualquer finalidade.

Pense nesse saldo como um **conjunto de vale-presentes**, não como dinheiro solto:

- Cada **Nota de Crédito (NC)** que compõe o saldo é um "vale" enviado pelo COTER ou pela COEx com uma finalidade escrita no próprio Objeto (ex.: *"AQS SV DEDETIZAÇÃO E LIMPEZA DE FOSSA"*, *"CONTRATAÇÃO SV INTERNET SATELITAL"*). O texto **"ESSA UG NÃO DEVE ALTERAR ND/UGR"** significa que o 6º BEC não pode gastar aquele vale em outra coisa sem autorização do órgão que o enviou.
- Muitas dessas NCs também trazem um **prazo de empenho** embutido no texto (ex.: *"EMPENHO ATÉ 5 SET 26"*). Depois desse prazo, o valor continua aparecendo no Tesouro Gerencial até ser formalmente recolhido, mas na prática **já não pode mais ser empenhado** pela unidade.
- O relatório diário (`relatorio_email_6bec.py`) agora identifica automaticamente esses dois casos e mostra um aviso ⚠️ logo abaixo do saldo total, com o valor que está **vencido** (🔴) e o valor **travado por finalidade** (🔒).

**Resumo:** o saldo total é real e está corretamente contabilizado no SIAFI, mas ele é fatiado em pedaços pequenos e amarrados a um uso específico — por isso a unidade pode, com razão, dizer que "não há crédito disponível" para uma necessidade que não se encaixe em nenhum desses vales, mesmo o total exibido sendo alto.

### "O saldo de crédito disponível do site é o mesmo em todas as telas?"

Não — o site mostra **números diferentes com escopos diferentes**, e é fácil confundi-los:

| Onde aparece | O que é | Escopo |
|---|---|---|
| "Crédito Disponível em Tela" (topo do Resumo) | Saldo livre de **todas** as Ações Orçamentárias do 6º BEC (OGU + FEx) | 6º BEC inteiro |
| "Créditos Livres — Exclusivo Ação 21EM" (e-mail, Seção 3) | Saldo livre **só** da Operação Catrimani II dentro do 6º BEC | 6º BEC · só 21EM |
| "Crédito Disponível na Catrimani" (Painel Catrimani, Seção 4) | Saldo livre da 21EM somando as **10 UGs executoras**, não só o 6º BEC | Multi-UG · só 21EM |

Ao comparar um número citado verbalmente pela unidade com o site, confirme sempre **qual dessas três linhas** está sendo referida.

---

## 🛡️ Segurança e Integridade
- Este repositório é completamente isolado dos demais painéis (`Dashboard-Credito-BCMS` e `Agente-Execucao-Orcamentaria`), não alterando qualquer dependência ou rotina já existente.
