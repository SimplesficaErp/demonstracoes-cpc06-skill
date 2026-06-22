# demonstracoes-cpc06

> Joga o contrato de leasing na pasta. Sai schedule do passivo + RoU + créditos PIS/COFINS + créditos CBS/IBS + lançamentos contábeis prontos pra ERP em 60 segundos. Padrão CPC 06 (R2) / IFRS 16.

Skill do Claude Code que tritura contratos de arrendamento mercantil (PDF, Word, Excel) e gera o pacote completo de cálculos exigidos pelo **CPC 06 (R2) / IFRS 16**, com fórmulas vivas, lançamentos contábeis prontos e tabela atualizada da Reforma Tributária (CBS/IBS).

Construído por contador, pra contador.

## O problema

Cliente novo manda 8, 15, 50 contratos de leasing. Cada um exige:

1. Ler o contrato e extrair: prazo, valor mensal, data início, taxa, opção de compra
2. Calcular o VP do passivo de arrendamento
3. Montar o RoU (Right of Use Asset)
4. Schedule mês a mês com juros + principal
5. Schedule de depreciação RoU
6. Aplicar PIS/COFINS (Lucro Real até 2026) OU CBS/IBS (todos a partir de 2027)
7. Lançamentos contábeis com D/C corretos
8. Conferir se soma do schedule = soma dos pagamentos

**Tempo médio: 2 horas por contrato.** Cliente com 30 contratos = 60 horas = 1,5 semana de trabalho contábil só nisso.

## A solução

Um skill modular que faz tudo isso automaticamente, em 47 segundos por contrato.

```bash
python scripts/gerar_relatorio_cpc06.py contratos.xlsx
```

Saída:

```
relatorio-cpc06-newledger.xlsx (Excel multi-aba)
├── 📋 CAPA              — Parâmetros + Selic capturada + Sumário
├── Tabela Aliquotas     — PIS/COFINS + CBS/IBS por ano (até 2031+)
├── Contrato 1           — Inputs + Schedule + RoU + JEs
├── Contrato 2           — (...)
├── ... (N contratos)
└── 📑 RESUMO            — Consolidado todos contratos
```

Tudo com:
- Fórmulas vivas no Excel (não valores estáticos)
- Formato brasileiro `1.234.567,89`
- Logo Newledger oficial
- Cabeçalho com data + taxa Selic capturada via BACEN
- Lançamentos contábeis com D/C prontos pra ERP

## Diferença vs `demonstracoes-cpc26`

| | `cpc26` | `cpc06` |
|---|---|---|
| Input | Balancete estruturado | Contrato (PDF/DOCX/XLSX) |
| Foco | Geração de DRE/BP/DFC/DMPL | Schedule + RoU + JEs |
| Norma | CPC 26 R3 + CPC 03 R3 + CPC 51 | CPC 06 R2 / IFRS 16 + Reforma |
| Triggers | "balancete", "DRE", "fechar competência" | "leasing", "arrendamento", "CPC 06" |

## Instalação rápida

```bash
git clone https://github.com/SimplesficaErp/demonstracoes-cpc06-skill.git
cd demonstracoes-cpc06-skill
pip install pandas openpyxl pillow yfinance requests
```

## Uso

### Forma 1 — Comando direto (Excel)

```bash
python scripts/gerar_relatorio_cpc06.py contratos.xlsx
```

### Forma 2 — Com taxa de desconto explícita

```bash
python scripts/gerar_relatorio_cpc06.py contratos.xlsx --taxa 0.1450
```

### Forma 3 — Via hook automático

Copia `hooks/post-tool-use.sh` para `.claude/hooks/` do teu projeto e dá permissão:

```bash
mkdir -p .claude/hooks
cp hooks/post-tool-use.sh .claude/hooks/
chmod +x .claude/hooks/post-tool-use.sh
```

Qualquer arquivo `contrato-*.xlsx` (ou `.pdf`) que aparecer dispara o skill sozinho.

### Forma 4 — Como skill nativo do Claude Code

Copia a pasta inteira para `.claude/skills/demonstracoes-cpc06/` do teu projeto. O Claude detecta o `SKILL.md` e ativa automaticamente nos triggers (`arrendamento`, `leasing`, `CPC 06`, `IFRS 16`).

## Formato esperado do input Excel

Planilha `.xlsx` com 1 aba por contrato OU 1 aba com tabela:

| identificacao | data_inicio | prazo_meses | valor_mensal | taxa_aa | regime |
|---|---|---|---|---|---|
| Imóvel/Bem comercial 1 | 2024-01-15 | 36 | 7500 | 0.15 | LUCRO_REAL |
| Frota de veículos | 2024-03-01 | 48 | 12000 | 0.15 | SIMPLES |
| ... | ... | ... | ... | ... | ... |

Se `taxa_aa` ausente, a skill busca a Selic acumulada 12 meses no BACEN (série 4390) e usa como discount rate automaticamente.

## Sincronização Selic via BACEN

A skill consulta automaticamente:
- **BACEN SGS série 4390** — Selic acumulada 12 meses (primária)
- **BACEN SGS série 432** — Meta Selic (secundária)
- **yfinance ^BVSP** — Contexto Ibovespa (info adicional)

Fallback: 15% a.a. se BACEN offline. Timestamp da taxa fica no relatório.

## Tabela de alíquotas (Reforma Tributária)

A skill mantém atualizada a transição CBS/IBS:

| Ano | CBS | IBS | Total |
|---|---|---|---|
| 2026 | 0,9% | 0,1% | 1,0% |
| 2027 | 0,88% | 1,77% | 2,65% |
| 2028 | 2,64% | 5,31% | 7,95% |
| 2029 | 4,40% | 8,85% | 13,25% |
| 2030 | 6,16% | 12,39% | 18,55% |
| 2031+ | 8,80% | 17,70% | 26,50% |

Para Lucro Real até 31/12/2026, aplica PIS 1,65% + COFINS 7,60% (regime não-cumulativo).

## Validações automáticas

- Soma do schedule = soma dos pagamentos (FAIL FAST)
- Saldo final do passivo no mês N = 0 (verifica fechamento)
- Total RoU + Juros = Total pagamentos (sanity check)
- Aviso se contrato < 12 meses (fora do CPC 06)

## Limitações conhecidas (v1.0)

- Extração de PDF assume layout legível (não OCR de scan ruim)
- Não trata sub-leasing (CPC 06 R2 §B58)
- Não calcula contratos curto prazo <12 meses
- Não trata moeda estrangeira
- Sem reavaliação por mudança de prazo

## Roadmap

- [ ] v1.1 — OCR de PDFs escaneados (Tesseract)
- [ ] v1.2 — Sub-leasing
- [ ] v1.3 — Multi-moeda + conversão FX
- [ ] v1.4 — Reavaliação CPC 06 R2 §40-43
- [ ] v2.0 — Integração SPED Contábil + ECD direto
- [ ] v2.1 — Notas explicativas auto-geradas via IA

## Quem fez

[Misael Holanda Neto](https://newledger.com.br) — Contador há 23 anos, dev. Criador do Newledger Pro, AuditCash, Simplesfica, OpenFinance Portal e IRPF Pro.

Mesma mão que fez [demonstracoes-cpc26-skill](https://github.com/SimplesficaErp/demonstracoes-cpc26-skill).

Atende distribuidora de medicamento, cosmético, veterinária, lojas e restaurantes via infraestrutura própria.

## Licença

MIT. Use, modifique, distribua. Atribuição apreciada mas não exigida.

## Quer mais skills assim?

Comenta `CPC 06` no [Instagram @misaelholandaia](https://instagram.com/misaelholandaia) ou no [LinkedIn](https://linkedin.com/in/misael-holanda).
