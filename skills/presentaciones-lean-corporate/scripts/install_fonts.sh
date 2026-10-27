#!/usr/bin/env bash
# Instala IBM Plex Sans (licencia SIL OFL, incluida en fonts/) para que la
# previsualización con LibreOffice y la medición de texto sean fieles a la plantilla.
# Uso: bash scripts/install_fonts.sh
set -euo pipefail
HERE="$(cd "$(dirname "$0")/.." && pwd)"
if fc-list 2>/dev/null | grep -qi "IBM Plex Sans"; then
  echo "IBM Plex Sans ya está instalada."; exit 0
fi
case "$(uname -s)" in
  Darwin) DEST="$HOME/Library/Fonts" ;;
  *)      DEST="$HOME/.fonts/ibm-plex" ;;
esac
mkdir -p "$DEST"
cp "$HERE"/fonts/IBMPlexSans-*.ttf "$DEST"/
command -v fc-cache >/dev/null && fc-cache -f "$DEST" >/dev/null || true
echo "IBM Plex Sans instalada en $DEST"
