#!/bin/bash
# post-tool-use.sh — Auto-dispatch do skill demonstracoes-cpc06
#
# Dispara quando qualquer arquivo contrato-*.xlsx ou contrato-*.pdf é salvo
# pelo Claude Code via Edit/Write tool.
#
# Instalar: copia para .claude/hooks/ do projeto e dá chmod +x
#   mkdir -p .claude/hooks
#   cp hooks/post-tool-use.sh .claude/hooks/
#   chmod +x .claude/hooks/post-tool-use.sh

set -e

FILE_PATH="$1"

if [[ -z "$FILE_PATH" ]]; then
    exit 0
fi

# Só dispara em arquivos contrato-*
case "$FILE_PATH" in
    *contrato-*.xlsx|*contrato-*.pdf|*contratos-*.xlsx)
        ;;
    *)
        exit 0
        ;;
esac

# Localiza skill instalada
SKILL_DIR="${CLAUDE_PROJECT_DIR:-$PWD}/.claude/skills/demonstracoes-cpc06"

if [[ ! -d "$SKILL_DIR" ]]; then
    echo "⚠ Skill demonstracoes-cpc06 não instalada em $SKILL_DIR"
    exit 0
fi

echo "🤖 Dispatching demonstracoes-cpc06 em: $FILE_PATH"
python3 "$SKILL_DIR/scripts/gerar_relatorio_cpc06.py" "$FILE_PATH"

echo "✅ Schedule + RoU + JEs gerados conforme CPC 06 (R2)"
