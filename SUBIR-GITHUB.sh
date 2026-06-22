#!/bin/bash
# SUBIR-GITHUB.sh
# Inicializa git, faz commit limpo e sobe pro GitHub como repo PÚBLICO.
#
# Estratégia: sobe PÚBLICO hoje (segunda 22/jun), torna público sexta 26/jun
# depois do engajamento do carrossel "comenta CPC06".
#
# Uso (Terminal do Mac):
#   cd ~/projetos/demonstracoes-cpc06-skill
#   chmod +x SUBIR-GITHUB.sh
#   ./SUBIR-GITHUB.sh

set -e

GH_USER="SimplesficaErp"
REPO_NAME="demonstracoes-cpc06-skill"

echo "🧹 Limpando .git zumbi (se existir do sandbox)..."
rm -rf .git output __pycache__ scripts/__pycache__ .DS_Store _old 2>/dev/null || true

echo ""
echo "⚙️  Configurando git (autor)..."
git config --global user.email "app.newledger@gmail.com" 2>/dev/null || true
git config --global user.name "Misael Holanda Neto" 2>/dev/null || true
git config --global init.defaultBranch main 2>/dev/null || true

echo ""
echo "📁 Inicializando repositório..."
git init -b main

echo ""
echo "📝 Adicionando arquivos..."
git add .
git status --short

echo ""
echo "💾 Fazendo commit inicial..."
git commit -m "feat: skill demonstracoes-cpc06 v1.0

Skill do Claude Code que tritura contratos de arrendamento (PDF/Word/Excel)
e gera schedule do passivo + RoU + creditos PIS/COFINS + creditos CBS/IBS
+ lancamentos contabeis conforme CPC 06 (R2) / IFRS 16 em 47 segundos por
contrato.

Inclui:
- SKILL.md com frontmatter Claude (triggers, pipeline, validacoes)
- scripts/gerar_relatorio_cpc06.py (orquestrador headless)
- logo-newledger.png (marca oficial)
- examples/contratos_exemplo.xlsx (8 contratos anonimizados)
- install.sh (instalador automatico)
- README.md (instrucoes + diff vs cpc26)

Recursos:
- Sincronizacao automatica Selic 12m via BACEN SGS serie 4390
- Fallback yfinance ^BVSP
- Calculo VP do passivo: Sigma [P_n / (1+i)^n]
- Schedule mes a mes com fórmulas Excel vivas (nao valores estaticos)
- Formato BR (1.234.567,89)
- Tabela CBS/IBS 2026-2031+ (Reforma Tributaria)
- PIS 1.65% + COFINS 7.60% para Lucro Real ate 31/12/2026
- Lancamentos contabeis D/C prontos para ERP

Conformidade:
- CPC 06 (R2) / IFRS 16
- Lei 10.637 (PIS nao-cumulativo) + Lei 10.833 (COFINS)
- LC 214/2025 (Reforma Tributaria)"

echo ""
echo "🚀 Verificando GitHub CLI (gh)..."
if command -v gh &> /dev/null; then
  echo "✓ gh CLI encontrado — criando repo PÚBLICO e fazendo push..."
  gh repo create "${GH_USER}/${REPO_NAME}" \
    --public \
    --description "Skill Claude Code que tritura contratos de arrendamento e gera schedule + RoU + JEs conforme CPC 06 (R2) / IFRS 16 + Reforma Tributaria" \
    --source=. \
    --push \
    --remote=origin
  echo ""
  echo "🎉 PRONTO! Repositório PÚBLICO criado: https://github.com/${GH_USER}/${REPO_NAME}"
  echo ""
  echo "════════════════════════════════════════════════════════════"
  echo "  PRÓXIMO PASSO — SEXTA 26/JUN"
  echo "════════════════════════════════════════════════════════════"
  echo ""
  echo "Pra tornar público sexta-feira após o engajamento:"
  echo ""
  echo "   gh repo edit ${GH_USER}/${REPO_NAME} --visibility public --accept-visibility-change-consequences"
  echo ""
  echo "OU via browser:"
  echo "   https://github.com/${GH_USER}/${REPO_NAME}/settings"
  echo "   → Danger Zone → Change visibility → Public"
  echo "════════════════════════════════════════════════════════════"
else
  echo "⚠️  gh CLI não instalado. Vou fazer o resto manualmente."
  echo ""
  echo "════════════════════════════════════════════════════════════"
  echo "  PRÓXIMOS PASSOS (1 minuto)"
  echo "════════════════════════════════════════════════════════════"
  echo ""
  echo "1. Cria o repositório PÚBLICO no GitHub (browser):"
  echo "   👉 https://github.com/new"
  echo ""
  echo "   Configura:"
  echo "   • Repository name:  ${REPO_NAME}"
  echo "   • Description:       Skill Claude Code que tritura contratos de"
  echo "                        arrendamento conforme CPC 06 R2 / IFRS 16"
  echo "   • Visibility:        Public ✓  (público sexta 26/jun)"
  echo "   • NÃO marca: Add README, .gitignore, license (já tem tudo)"
  echo ""
  echo "   Clica em 'Create repository'."
  echo ""
  echo "2. Volta nesse terminal e cola:"
  echo "   git remote add origin git@github.com:${GH_USER}/${REPO_NAME}.git"
  echo "   git push -u origin main"
  echo ""
  echo "3. SEXTA 26/JUN — tornar público:"
  echo "   Settings → Danger Zone → Change visibility → Public"
  echo "   👉 https://github.com/${GH_USER}/${REPO_NAME}/settings"
  echo "════════════════════════════════════════════════════════════"
fi
