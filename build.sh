#!/usr/bin/env bash
# Rebuild the manuscript from the JRM Word template.
# usage: ./build.sh [output.docx]   (default: krms_actexam_build.docx)
set -euo pipefail
cd "$(dirname "$0")"
OUT="${1:-krms_actexam_build.docx}"
WORK="$(mktemp -d)"
trap 'rm -rf "$WORK"' EXIT
unzip -q template/jrm_template.docx -d "$WORK/doc"
python3 src/manuscript.py "$WORK/doc"
(cd "$WORK/doc" && zip -qXr "$WORK/out.docx" .)
mv "$WORK/out.docx" "$OUT"
echo "built $OUT"
