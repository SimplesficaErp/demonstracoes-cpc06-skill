#!/bin/bash
# install.sh — Instala o skill demonstracoes-cpc06 em um projeto Claude Code.
#
# Uso:
#   ./install.sh                      # instala no projeto atual (.claude/skills/)
#   ./install.sh ~/.claude/skills     # instala no Claude global
#   ./install.sh /caminho/customizado

set -euo pipefail

DEST="${1:-.claude/skills}"
SKILL_NAME="demonstracoes-cpc06"

SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"

echo "📦 Instalando skill $SKILL_NAME"
echo "   Origem: $SCRIPT_DIR"
echo "   Destino: $DEST/$SKILL_NAME"

# Cria estrutura
mkdir -p "$DEST/$SKILL_NAME/scripts"
mkdir -p "$DEST/$SKILL_NAME/templates"
mkdir -p "$DEST/$SKILL_NAME/examples"

# Copia arquivos
cp "$SCRIPT_DIR/SKILL.md" "$DEST/$SKILL_NAME/"
cp "$SCRIPT_DIR/README.md" "$DEST/$SKILL_NAME/"
cp "$SCRIPT_DIR/scripts/"*.py "$DEST/$SKILL_NAME/scripts/"
cp "$SCRIPT_DIR/logo-newledger.png" "$DEST/$SKILL_NAME/" 2>/dev/null || true
cp "$SCRIPT_DIR/examples/"*.xlsx "$DEST/$SKILL_NAME/examples/" 2>/dev/null || true

# Hook (opcional)
read -p "Instalar hook automático em .claude/hooks/? (s/n): " RESP
if [[ "$RESP" =~ ^[Ss]$ ]]; then
    mkdir -p "$(dirname "$DEST")/hooks"
    cp "$SCRIPT_DIR/hooks/post-tool-use.sh" "$(dirname "$DEST")/hooks/" 2>/dev/null || true
    chmod +x "$(dirname "$DEST")/hooks/post-tool-use.sh" 2>/dev/null || true
    echo "✓ Hook instalado em $(dirname "$DEST")/hooks/post-tool-use.sh"
fi

# Verifica dependências
echo ""
echo "🔍 Verificando dependências Python..."
python3 -c "import pandas, openpyxl, PIL" 2>/dev/null && echo "✓ pandas + openpyxl + Pillow OK" || {
    echo "⚠ Faltam dependências. Instale com:"
    echo "    pip install pandas openpyxl pillow yfinance requests"
}

echo ""
echo "✅ Skill instalado com sucesso!"
echo ""
echo "Testar com contrato de exemplo:"
echo "  python3 $DEST/$SKILL_NAME/scripts/gerar_relatorio_cpc06.py \\"
echo "    $DEST/$SKILL_NAME/examples/contratos_exemplo.xlsx"
