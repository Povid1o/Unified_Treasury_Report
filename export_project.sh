#!/usr/bin/env bash
# ═══════════════════════════════════════════════════════════════════════════
#  export_project.sh — экспорт проекта в zip с паролем для переноса по почте.
#
#  Пакует весь проект, КРОМЕ venv / .git / кэшей / output / logs / .claude,
#  в защищённый паролем архив (пароль см. PASSWORD ниже).
#
#  Запуск:
#     ./export_project.sh                 # архив рядом с папкой проекта
#     ./export_project.sh ~/Desktop       # архив в указанную папку
#
#  Требуется утилита `zip` (в macOS есть из коробки).
# ═══════════════════════════════════════════════════════════════════════════
set -euo pipefail

PASSWORD="Treasury"

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
PROJECT_NAME="$(basename "$SCRIPT_DIR")"
PARENT_DIR="$(dirname "$SCRIPT_DIR")"

# Куда класть архив: 1-й аргумент или родительская папка проекта
OUT_DIR_ARG="${1:-$PARENT_DIR}"
mkdir -p "$OUT_DIR_ARG"
OUT_DIR="$(cd "$OUT_DIR_ARG" && pwd)"

STAMP="$(date +%Y%m%d_%H%M)"
OUT_FILE="$OUT_DIR/${PROJECT_NAME}_${STAMP}.zip"

command -v zip >/dev/null 2>&1 || { echo "✗ Утилита 'zip' не найдена. Установите её и повторите."; exit 1; }

rm -f "$OUT_FILE"
cd "$PARENT_DIR"

echo "Пакую '$PROJECT_NAME' → $OUT_FILE ..."
zip -r -X -P "$PASSWORD" "$OUT_FILE" "$PROJECT_NAME" \
  -x "*/.venv/*" \
     "$PROJECT_NAME/.git/*" \
     "*/__pycache__/*" \
     "*.pyc" \
     "*/output/*" \
     "*/logs/*" \
     "*.DS_Store" \
     "$PROJECT_NAME/.DS_Store" \
     "$PROJECT_NAME/.claude/*" \
     "$PROJECT_NAME"/*.zip \
  >/dev/null

FILES="$(unzip -l "$OUT_FILE" | tail -1 | awk '{print $2}')"
SIZE="$(du -h "$OUT_FILE" | cut -f1)"

echo ""
echo "✓ Готово."
echo "  Архив : $OUT_FILE"
echo "  Размер: $SIZE  |  Файлов: $FILES  |  Пароль: $PASSWORD"
echo ""
echo "  Если почтовый шлюз режёт .zip — переименуйте в .zip.txt, на месте верните .zip."
