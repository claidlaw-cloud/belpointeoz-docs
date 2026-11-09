# Belpointe OZ investor documents

Published PDFs, served through GitHub Pages at fixed URLs (docs.belpointeoz.com):

- fact-sheet.pdf
- white-paper.pdf
- investor-presentation.pdf (maintained separately; replace the file to update)

## Regenerating the fact sheet and white paper

All figures live in `build/figures.json` (NAV, NAV date and source, total NAV, units outstanding, unit price and date,
leasing percentages and as-of date, latest periodic report and accession number). `build/docs2.js` builds both documents
from it with the `docx` npm package; `build/build.sh` converts them to PDF with LibreOffice and copies the PDFs to the root.

Automated update (weekly scheduled task): fetch `<relay>?type=facts` from the Belpointe OZ relay, run
`node build/update-figures.js facts.json`; if it prints `changed`, run `build/build.sh`, review page counts (white paper 9, fact sheet 2),
commit and push. Press release boilerplate and property descriptions are text in `docs2.js` and change only by hand.
