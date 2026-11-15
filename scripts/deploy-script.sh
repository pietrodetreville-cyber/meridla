#!/usr/bin/env bash
# Nahraje kód z apps-script/ do Google Apps Scriptu a přenasadí webovou aplikaci (adresa .../exec zůstává).
# Použití:  scripts/deploy-script.sh test     (testovací tabulka)
#           scripts/deploy-script.sh prod     (ostrá tabulka)
set -euo pipefail
env="${1:-}"
cd "$(dirname "$0")/../apps-script"
here="$(cygpath -w "$(pwd)" 2>/dev/null || pwd)"

case "$env" in
  test) cfg=".clasp.test.json"; dep="AKfycbxs9B0HfgFENA7_BO4VSgoAY8gHmhxAC32D5nKJf902OufOF6bqKSR-T9N1ZInRk_Hm" ;;
  prod) cfg=".clasp.prod.json"; dep="AKfycbyx3UjK_0rqmvNLEMBZPWUpW4dcsu8N3Ovxy3eNb5d4YVUj1zM64mrr8xvn3_uQigSI" ;;
  *) echo "Použití: $0 test|prod" >&2; exit 2 ;;
esac

if [ ! -f "$cfg" ] || [ -z "$dep" ]; then
  echo "Prostředí '$env' zatím není nastavené (chybí $cfg nebo ID nasazení)." >&2
  exit 3
fi

desc="$(git log -1 --format='%h %s' | cut -c1-70)"
echo "== Nahrávám kód do prostředí: $env"
clasp push --force --project "$here\\$cfg"
echo "== Vytvářím novou verzi a přenasazuji (${desc})"
clasp deploy --project "$here\\$cfg" --deploymentId "$dep" --description "$desc"
echo "== Hotovo. Ověření: ?action=version na adrese nasazení."
