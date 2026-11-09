#!/usr/bin/env bash
# Builds the fact sheet and white paper from figures.json and places the PDFs at the repository root.
set -euo pipefail
cd "$(dirname "$0")"
node docs2.js
soffice --headless --convert-to pdf --outdir out "Belpointe OZ White Paper.docx" "Belpointe OZ Fact Sheet.docx" >/dev/null 2>&1
cp "out/Belpointe OZ White Paper.pdf" ../white-paper.pdf
cp "out/Belpointe OZ Fact Sheet.pdf" ../fact-sheet.pdf
pdfinfo ../white-paper.pdf | grep Pages
pdfinfo ../fact-sheet.pdf | grep Pages
