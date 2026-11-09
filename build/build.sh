#!/usr/bin/env bash
# Builds the fact sheet, white paper and investor presentation from figures.json and places the PDFs at the repository root.
set -euo pipefail
cd "$(dirname "$0")"
mkdir -p out
node docs2.js
node deck.js
soffice --headless --convert-to pdf --outdir out "Belpointe OZ White Paper.docx" "Belpointe OZ Fact Sheet.docx" "Belpointe OZ Investor Presentation.pptx" >/dev/null 2>&1
python3 compress.py "out/Belpointe OZ Investor Presentation.pdf" ../investor-presentation.pdf
cp "out/Belpointe OZ White Paper.pdf" ../white-paper.pdf
cp "out/Belpointe OZ Fact Sheet.pdf" ../fact-sheet.pdf
for f in white-paper fact-sheet investor-presentation; do printf "%s: " $f; pdfinfo ../$f.pdf | grep Pages; done
ls -la ../*.pdf
