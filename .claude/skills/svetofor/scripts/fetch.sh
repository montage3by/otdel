#!/usr/bin/env bash
# Скачивает «Светофор» и все дневники проектов (ссылки — из notes/proekty-opt.md) в xlsx.
# Использование: fetch.sh <пустая папка вне репозитория>
# В дневниках есть пароли и телефоны: папку удалить после разбора, в репозиторий не класть.
set -u
OUT=${1:?нужна папка для выгрузки}
REPO=$(cd "$(dirname "$0")/../../../.." && pwd)
mkdir -p "$OUT"

curl -sSL -o "$OUT/svetofor.xlsx" \
  "https://docs.google.com/spreadsheets/d/12416tJFmz2MGFfDQlAPZ96R7D4ZXQYk-pjI-FsiIAa4/export?format=xlsx"

# Каждый раздел проекта: "## N. Название" и строка "**Ссылки:** [дневник проекта](...)"
awk '/^## [0-9]+\./ {sub(/^## [0-9]+\. /, ""); name=$0}
     /\[дневник проекта\]/ && name {match($0, /spreadsheets\/d\/[A-Za-z0-9_-]+/); id=substr($0, RSTART+15, RLENGTH-15); print id "\t" name; name=""}' \
  "$REPO/notes/proekty-opt.md" |
while IFS=$'\t' read -r id name; do
  safe=$(printf '%s' "$name" | sed 's/[^A-Za-zА-Яа-яЁё0-9]\+/_/g' | cut -c1-40)
  f="$OUT/$safe.xlsx"
  curl -sSL -o "$f" "https://docs.google.com/spreadsheets/d/$id/export?format=xlsx"
  if file -b "$f" | grep -q 'Excel'; then
    echo "OK     $name"
  else
    echo "CLOSED $name (нужен доступ по ссылке)"; rm -f "$f"
  fi
done
