# 📋 Skill `demonstracoes-cpc06` — Proposta de Arquitetura

> Tritura contratos de arrendamento (PDF/Word/Excel) e gera tabela de amortização + cronograma do direito de uso + lançamentos contábeis conforme **CPC 06 (R2) / IFRS 16**.

---

## 🎯 Posicionamento vs Skills Existentes

| Skill | Input | Output | Complexidade |
|---|---|---|---|
| `demonstracoes-cpc26` (existente) | Balancete .xlsx estruturado | DRE + BP + DFC + DMPL | **Baixa** — formato fixo |
| **`demonstracoes-cpc06` (nova)** | Contrato .pdf / .docx / .xlsx | Schedule passivo + RoU + JEs + notas | **Alta** — NLP + cálculo financeiro |

A diferença-chave: cpc-26 transforma dados **estruturados** em demonstrações. cpc-06 transforma **texto livre de contrato** em demonstrações.

---

## 🧠 O que a skill faz (em uma frase)

> "Joga o PDF do contrato de leasing na pasta. Sai a tabela de amortização do passivo + cronograma de depreciação do RoU + lançamentos mês a mês conforme CPC 06 (R2) em 60 segundos."

---

## 📥 Inputs aceitos (3 modos)

### Modo 1 — **Contrato individual** (.pdf ou .docx)
```
input/
└── contrato-leasing-frota-veiculos.pdf
```
A skill extrai os termos do contrato via NLP/regex e calcula sozinha.

### Modo 2 — **Lista em batch** (.xlsx)
```
input/
└── contratos-2026.xlsx   # 50 contratos por linha
```
Colunas esperadas (qualquer ordem):
- `identificacao_bem` · `prazo_meses` · `valor_mensal` · `data_inicio` · `taxa_juros_aa` · `tipo_pagamento` (antecipado/postecipado) · `opcao_compra_valor` · `opcao_compra_provavel` (S/N) · `valor_residual_garantido` · `incentivos_recebidos` · `custos_diretos_iniciais` · `provisao_restauracao`

### Modo 3 — **Híbrido** — Excel + PDFs anexos
```
input/
├── contratos-2026.xlsx        # tabela com dados extraídos
└── pdfs/
    ├── contrato-001.pdf       # PDF como evidência
    ├── contrato-002.pdf
    └── ...
```

---

## 📤 Outputs gerados (5 arquivos por contrato OU 1 consolidado)

### 1. `schedule_passivo.xlsx`
Tabela de amortização do passivo de arrendamento, linha por mês:

| Mês | Data | Saldo Inicial | Juros | Pagamento | Principal | Saldo Final |
|---|---|---|---|---|---|---|
| 0 | 01/jul/26 | 0,00 | 0,00 | 0,00 | 0,00 | 210.618,00 |
| 1 | 01/ago/26 | 210.618,00 | 1.053,09 | 5.000,00 | 3.946,91 | 206.671,09 |
| 2 | 01/set/26 | 206.671,09 | 1.033,36 | 5.000,00 | 3.966,64 | 202.704,45 |
| ... | ... | ... | ... | ... | ... | ... |

### 2. `schedule_rou.xlsx`
Cronograma de depreciação do Ativo de Direito de Uso (RoU):

| Mês | Valor RoU | Depreciação | Acumulado | Saldo Líquido |
|---|---|---|---|---|
| 0 | 210.618,00 | 0,00 | 0,00 | 210.618,00 |
| 1 | 210.618,00 | 3.510,30 | 3.510,30 | 207.107,70 |
| ... | ... | ... | ... | ... |

### 3. `lancamentos.xlsx`
Lançamentos contábeis (JEs) prontos pra ERP:

| Data | Conta Débito | Conta Crédito | Histórico | Valor |
|---|---|---|---|---|
| 01/jul/26 | 1.2.04.001 RoU | 2.2.05.001 Passivo Arrend. | Reconh. inicial CPC 06 | 210.618,00 |
| 31/jul/26 | 4.1.05.001 Juros Arrend. | 2.2.05.001 Passivo Arrend. | Juros mês 1 | 1.053,09 |
| 31/jul/26 | 2.2.05.001 Passivo Arrend. | 1.1.01.001 Caixa | Pagamento mês 1 | 5.000,00 |
| 31/jul/26 | 4.1.06.001 Depr. RoU | 1.2.04.002 Depr. Acum. RoU | Depreciação RoU mês 1 | 3.510,30 |

### 4. `memo_decisoes.md`
Documenta os **julgamentos significativos** que a skill tomou:
- Prazo do arrendamento adotado (incluindo opções "razoavelmente certas")
- Taxa de desconto utilizada (implícita ou IBR)
- Aplicação de isenções (curto prazo / baixo valor)
- Tratamento de pagamentos variáveis indexados
- Inclusão/exclusão de custos diretos iniciais e restauração

### 5. `notas_explicativas.md`
Texto preenchível pra incluir no BP/DRE conforme CPC 06 (R2) §51-60.

---

## 🔧 Arquitetura Técnica

```
demonstracoes-cpc06-skill/
├── SKILL.md                          # Frontmatter Claude Code
├── README.md
├── install.sh
├── hooks/
│   └── post-tool-use.sh              # Auto-disparo se detectar contrato-*.pdf/.docx/.xlsx
├── scripts/
│   ├── extract_pdf.py                # PyPDF2 + regex pra extrair termos do PDF
│   ├── extract_docx.py               # python-docx + regex
│   ├── extract_xlsx.py               # openpyxl pra modo batch
│   ├── normalize_terms.py            # Padroniza termos extraídos numa estrutura única
│   ├── calc_lease_liability.py       # PV dos pagamentos (np.npv ou implementação manual)
│   ├── calc_rou_asset.py             # RoU = Liability + IDC + Restauração - Incentivos
│   ├── gen_schedule_passivo.py       # Tabela amortização mês a mês
│   ├── gen_schedule_rou.py           # Cronograma depreciação RoU
│   ├── gen_lancamentos.py            # JEs auto-gerados com plano de contas BR
│   ├── gen_notas.py                  # Notas explicativas CPC 06 R2 §51-60
│   ├── validations.py                # Validações: PV math, sum check, etc
│   └── gerar_cpc06.py                # ORQUESTRADOR (entry point)
├── templates/
│   ├── modelo-schedule-passivo.xlsx
│   ├── modelo-schedule-rou.xlsx
│   └── modelo-lancamentos.xlsx
├── examples/
│   ├── contrato-leasing-veiculos.pdf    # exemplo PDF
│   ├── contrato-aluguel-imovel.docx     # exemplo Word
│   └── contratos-batch-50.xlsx          # exemplo batch
└── tests/
    ├── test_pv_calc.py               # asserções financeiras
    ├── test_pdf_extraction.py
    └── test_validations.py
```

---

## 🔬 Extração de termos do contrato — como funciona

### Estratégia multi-camada
1. **Layer 1 — Regex/Pattern matching:** captura padrões óbvios
   - `R\$\s?[\d.,]+` → valores monetários
   - `\d+\s?(?:meses|anos)` → prazos
   - `taxa de \d+(?:,\d+)?%` → taxas de juros
   
2. **Layer 2 — Keyword anchoring:** identifica seções por palavras-chave
   - "Cláusula" / "Prazo" / "Contraprestação" / "Opção de compra" / "Valor residual"
   
3. **Layer 3 — Claude AI fallback:** se layers 1+2 não trazem o suficiente, chama Anthropic API com o texto do contrato e prompt específico:
   > "Extraia desta cláusula contratual: prazo, valor mensal, juros, opção de compra, garantia residual, incentivos. Devolva JSON estruturado."

### Confiabilidade
- Cada termo extraído tem um campo **`confidence`** (0.0-1.0)
- Termos com confidence < 0.7 vão pro `memo_decisoes.md` flaggados pra revisão humana
- Skill **nunca chuta** — se não consegue extrair, pergunta ao user ou para com erro claro

---

## 💰 Cálculo financeiro — Conformidade CPC 06 (R2)

### Mensuração Inicial do Passivo (`CPC 06 R2 §26`)
```
Liability_0 = Σ [P_n / (1 + i)^n]  para n = 1 a N
```
Onde:
- `P_n` = pagamento no mês n
- `i` = taxa de desconto mensal (implícita do contrato OU IBR)
- `N` = número total de pagamentos

### Mensuração Inicial do RoU (`CPC 06 R2 §24`)
```
RoU_0 = Liability_0 
      + Pagamentos antecipados
      + Custos diretos iniciais
      - Incentivos recebidos do arrendador
      + Estimativa de custos de desmontagem/restauração
```

### Mensuração Subsequente
- **Passivo:** método da taxa efetiva de juros (custo amortizado)
  ```
  Juros_n = Saldo_(n-1) × i
  Principal_n = Pagamento_n - Juros_n
  Saldo_n = Saldo_(n-1) - Principal_n
  ```
- **RoU:** depreciação linear sobre o prazo do arrendamento OU vida útil (o que for menor)
  ```
  Depreciação_mensal = RoU_0 / N
  ```

### Validações automáticas (FAIL FAST)
- ✓ `Σ Juros + Σ Principal = Σ Pagamentos`
- ✓ `Saldo_N = 0` (passivo amortizado completamente)
- ✓ `Σ Depreciação = RoU_0` (RoU totalmente amortizado)
- ✓ `Liability_0 calculado = Σ PV(pagamentos)` (PV math)

---

## 🎯 Triggers da skill (frontmatter SKILL.md)

A skill ativa automaticamente quando Claude detectar no input do usuário:
- "CPC 06" ou "CPC 06 R2"
- "IFRS 16"
- "leasing" ou "arrendamento mercantil"
- "right-of-use" ou "RoU"
- "lease liability" ou "passivo de arrendamento"
- "contrato de aluguel" (com contexto B2B)
- "lease accounting"

Ou se o file watcher (hook post-tool-use.sh) detectar:
- `contrato-*.pdf`
- `contrato-*.docx`
- `contratos-*.xlsx`

---

## 🇧🇷 Adaptações brasileiras vs IFRS 16 padrão

| Item | IFRS 16 (global) | CPC 06 R2 (Brasil) |
|---|---|---|
| Isenção baixo valor | ~US$ 5.000 quando novo | Equivalente em R$ (avaliação por bem) |
| Taxa IBR | Incremental Borrowing Rate genérica | Recomendação: Selic + spread por porte |
| Tributação | Não aplicável | **Importante:** IRPJ/CSLL — Lei 12.973/2014 (RTT eliminado), neutralidade fiscal específica |
| Plano de contas | Genérico | Adaptado: 1.2.04 (RoU), 2.2.05 (Passivo Arrend.), 4.1.05 (Juros), 4.1.06 (Depr.) |
| Notas explicativas | IFRS 16 §51-60 | CPC 06 R2 §51-60 + ITG 02 (transição) se aplicável |
| Sublocação | Operacional/financeira por referência ao RoU | Mesma regra, mas observar Lei do Inquilinato (Lei 8.245/91) |

A skill incluirá flag `--brasil-fiscal` que ativa **memorando tributário** explicando impactos IRPJ/CSLL específicos.

---

## 📅 Plano de construção (4-6 horas de dev)

| Fase | Tempo | Entregável |
|---|---|---|
| **F1 — Estrutura base** | 30min | Frontmatter SKILL.md + install.sh + hooks |
| **F2 — Parser PDF** | 1h | `extract_pdf.py` com PyPDF2 + regex + fallback Claude |
| **F3 — Parser DOCX + XLSX** | 30min | `extract_docx.py` + `extract_xlsx.py` |
| **F4 — Engine financeiro** | 1h | `calc_lease_liability.py` + `calc_rou_asset.py` + validations |
| **F5 — Schedules + JEs** | 1h | `gen_schedule_passivo.py` + `gen_schedule_rou.py` + `gen_lancamentos.py` |
| **F6 — Notas + memo + orquestrador** | 30min | `gen_notas.py` + `gen_memo_decisoes.py` + `gerar_cpc06.py` |
| **F7 — Exemplos + testes** | 30min | 3 exemplos (PDF, Word, batch) + 5 testes unitários |
| **F8 — README + push GitHub** | 15min | README completo + commit `SimplesficaErp/demonstracoes-cpc06-skill` |

**Total: ~5h de implementação real.**

---

## 🚀 ROI esperado pro contador

| Métrica | Antes (manual) | Depois (skill) | Economia |
|---|---|---|---|
| Tempo por contrato de leasing | 2-3 horas | 60 segundos | **99%** |
| Probabilidade de erro de cálculo | ~15% | < 1% | **94%** |
| Audit trail | Manual (planilha) | Automatico (memo + validações) | **100% rastreabilidade** |
| Conformidade CPC 06 R2 | Variável | Determinística | **garantida** |
| Reconciliação BS/DRE | Manual | Automática | **horas economizadas/mês** |

Pra escritório com **50 contratos de leasing ativos**: ~100h/mês economizadas. Equivale a R$ 8-15k/mês em custo de mão de obra (junior + revisão sênior).

---

## ⏭️ Próximos passos — escolhe

1. **"vai"** → começo a construir AGORA, em ordem F1 → F8, te entregando cada fase pra revisar
2. **"ajusta proposta"** → me diz o que mudar (escopo, formato, prioridade) e refaço
3. **"só estrutura"** → eu só monto SKILL.md + estrutura de pastas + README (sem código, ~30min)
4. **"foca no parser"** → começa pelo parser PDF que é a parte mais difícil, valida com 1 contrato real teu

Saldo Nano Banana atual: ~388 créditos (suficiente pra gerar arts do post de divulgação no GitHub depois).

**Pra rodar:** preciso saber se tens **1-3 contratos de leasing reais** (anonimizados) que eu possa usar como teste de extração. Sem contrato real, dev fica em ambiente sintético e parser pode falhar em casos de borda.
