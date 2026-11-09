# Belpointe OZ investor documents

Published PDFs, served through GitHub Pages at fixed URLs (docs.belpointeoz.com):

- fact-sheet.pdf
- white-paper.pdf
- investor-presentation.pdf

## Regenerating the documents

All figures live in `build/figures.json` (NAV, NAV date and source, total NAV, units outstanding, unit price and date,
leasing percentages and as-of date, latest periodic report and accession number). `build/docs2.js` builds the white paper and fact sheet with the `docx` npm package and `build/deck.js` builds the investor presentation with `pptxgenjs`;
`build/build.sh` converts all three to PDF with LibreOffice, compresses the deck's images (`compress.py`, needs pikepdf and Pillow) and copies the PDFs to the root.

Debt figures in the deck (`debtAsOf`, `totalDebt`, `alDebt*`, `vivDebt*` in figures.json) come from the 10-Q debt note and are not produced by the relay;
when the report period changes, the scheduled task flags them for manual confirmation and keeps the previous values until updated.

Automated update (weekly scheduled task): fetch `<relay>?type=facts` from the Belpointe OZ relay, run
`node build/update-figures.js facts.json`; if it prints `changed`, run `build/build.sh`, review page counts (white paper 9, fact sheet 2, presentation 16),
commit and push. Press release boilerplate and property descriptions are text in `docs2.js` and change only by hand.
