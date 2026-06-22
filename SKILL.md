---
name: demonstracoes-cpc06
description: Tritura contratos de arrendamento (PDF, Word, Excel) e gera tabela de amortização do passivo, cronograma do direito de uso (RoU), créditos PIS/COFINS (Lucro Real até 31/12/2026), créditos CBS/IBS (a partir de 01/2027 — Reforma Tributária) e lançamentos contábeis prontos pra ERP, conforme CPC 06 (R2) / IFRS 16. Use SEMPRE que receber arquivo contrato*.pdf/.docx/.xlsx, ou quando o usuário pedir "calcular leasing", "schedule de arrendamento", "RoU", "passivo de arrendamento", "lançamentos IFRS 16", "CPC 06", "tritura esse contrato". Aplica taxa Selic acumulada 12m (BACEN SGS série 4390) automaticamente como discount rate quando não informada.
---

# Demonstrações CPC 06 — Skill

Recebe contrato de arrendamento em PDF, Word ou Excel e gera o pacote completo de cálculos contábeis no padrão **CPC 06 (R2) / IFRS 16**, com fórmulas vivas e lançamentos prontos pra ERP.

## Regra Zero — Ler ANTES de executar

```
OBRIGATÓRIO antes de calcular:
  1. Identificar tipo de input (PDF, DOCX, XLSX)
  2. Validar campos mínimos do contrato:
     - prazo_meses (int > 0)
     - valor_mensal (decimal > 0)
     - data_inicio (date)
     - taxa_desconto_aa (decimal) — se ausente, usar Selic 12m via BACEN
     - regime_tributario (LUCRO_REAL | SIMPLES | PRESUMIDO)
  3. Se algum campo crítico ausente → PARAR e pedir manualmente
```

## O que esta skill gera

| Output | Norma | Conteúdo |
|---|---|---|
| **Schedule Passivo** | CPC 06 R2 §26-39 | Saldo, Juros, Pagamento, Principal, Saldo Final mês a mês |
| **Schedule RoU** | CPC 06 R2 §22-25, §31 | Valor RoU, Depreciação linear, Acumulado, Saldo Líquido |
| **Créditos PIS/COFINS** | Lei 10.637/10.833 | Aplicado se Regime = Lucro Real e data ≤ 31/12/2026 |
| **Créditos CBS/IBS** | LC 214/2025 + Reforma | Aplicado a partir de 01/01/2027 — todas empresas, transição até 2033 |
| **Lançamentos Contábeis** | CPC 06 R2 §47-51 | JEs prontos pra ERP (D/C com históricos) |
| **Relatório Consolidado** | — | Excel multi-aba com CAPA + Contratos + RESUMO |

Todos com:
- Fórmulas vivas no Excel (não valores estáticos)
- Formato brasileiro `1.234.567,89`
- Logo Newledger oficial (se `logo-newledger.png` presente)
- Cabeçalho com data de geração + taxa Selic capturada
- Notas explicativas com referência ao parágrafo do CPC 06

## Pipeline de execução

```
1. extract_contract($input)         → dict com campos do contrato
2. fetch_selic_12m_bacen()          → taxa de desconto (fallback yfinance)
3. calc_pv_passivo()                → VP = Σ [P / (1+i)^n]
4. gen_schedule_passivo()           → tabela amortização
5. gen_schedule_rou()                → depreciação linear
6. apply_pis_cofins(regime, data)    → créditos até 31/12/2026
7. apply_cbs_ibs(data)               → créditos a partir de 01/2027
8. gen_lancamentos_contabeis()       → JEs prontos
9. compile_relatorio_excel()         → output multi-aba
```

## Tabela de alíquotas CBS/IBS (Reforma Tributária)

| Ano | CBS | IBS | Total | Status |
|---|---|---|---|---|
| 2026 | 0,9% | 0,1% | 1,0% | Teste (compensável) |
| 2027 | 0,88% | 1,77% | 2,65% | Início vigência |
| 2028 | 2,64% | 5,31% | 7,95% | Transição |
| 2029 | 4,40% | 8,85% | 13,25% | Transição |
| 2030 | 6,16% | 12,39% | 18,55% | Transição |
| 2031+ | 8,80% | 17,70% | 26,50% | Regime pleno |

Fonte: LC 214/2025 + projeções da Receita Federal.

## Mensuração inicial (CPC 06 R2 §22-28)

```
Passivo de Arrendamento = VP dos pagamentos futuros
  → VP = Σ [Pagamento_n / (1 + taxa_mensal)^n]
  → taxa_mensal = (1 + taxa_aa)^(1/12) - 1

Ativo Direito de Uso (RoU) = Passivo de Arrendamento
                           + Pagamentos antecipados
                           + Custos diretos iniciais
                           + Estimativa de restauração
                           - Incentivos recebidos
```

## Lançamentos contábeis padrão

### Reconhecimento inicial
```
D | 1.2.04.001 — Ativo Direito de Uso (RoU)      [valor RoU]
C | 2.2.05.001 — Passivo de Arrendamento         [valor RoU]
```

### Mensal — pagamento
```
D | 2.2.05.001 — Passivo de Arrendamento         [principal]
D | 4.1.05.001 — Despesa Juros Arrendamento      [juros]
C | 1.1.01.001 — Banco/Caixa                     [pagamento total]
```

### Mensal — depreciação RoU
```
D | 4.1.06.001 — Despesa Depreciação RoU         [valor mês]
C | 1.2.04.901 — Depreciação Acumulada RoU       [valor mês]
```

## Validações obrigatórias (FAIL FAST)

| Check | Falha → |
|---|---|
| Valor mensal informado e > 0 | PARAR — pedir dado |
| Prazo em meses informado e > 0 | PARAR — pedir dado |
| Data início válida | PARAR — pedir dado |
| Soma do schedule = valor total a pagar | PARAR — fórmula quebrada |
| Saldo final do passivo no mês N = 0 | PARAR — schedule não fecha |

## Como rodar

```bash
# Modo 1 — Contrato Excel (mais comum)
python scripts/gerar_relatorio_cpc06.py path/to/contratos.xlsx

# Modo 2 — Output customizado
python scripts/gerar_relatorio_cpc06.py contratos.xlsx output-cliente-x.xlsx

# Modo 3 — Via hook (automático quando arquivo contrato-*.pdf é salvo)
.claude/hooks/post-tool-use.sh path/to/contrato.pdf
```

## Output esperado

```
relatorio-cpc06-newledger.xlsx
├── 📋 CAPA              (parâmetros + Selic + sumário)
├── Tabela Aliquotas    (PIS/COFINS + CBS/IBS por ano)
├── Contrato 1          (inputs + schedule + JEs)
├── Contrato 2          (...)
├── ...
└── 📑 RESUMO            (consolidado todos contratos)
```

## Limites conhecidos

- Extração de PDF assume layout legível (não OCR de scan ruim)
- Não trata sub-leasing (CPC 06 R2 §B58)
- Não calcula contratos de curto prazo (<12 meses) — esses ficam fora do CPC 06
- Não trata moeda estrangeira (sem conversão)
- BACEN inacessível em sandbox → usa fallback de Selic 15% a.a.

## Próximas versões

- [ ] OCR de PDFs escaneados (Tesseract)
- [ ] Sub-leasing
- [ ] Multi-moeda
- [ ] Reavaliação por mudança de prazo (CPC 06 R2 §40-43)
- [ ] Integração SPED Contábil + ECD
- [ ] Notas explicativas auto-geradas

## Quem fez

[Misael Holanda Neto](https://newledger.com.br) — Contador há 23 anos, dev. Mesma mão da skill `demonstracoes-cpc26`.

## Licença

MIT.
