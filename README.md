# Belpointe OZ investor documents

Published files (served through GitHub Pages at fixed URLs):

- fact-sheet.pdf
- white-paper.pdf
- investor-presentation.pdf

`build/docs2.js` generates the white paper and fact sheet (.docx) with the `docx` npm package; convert to PDF with LibreOffice
(`soffice --headless --convert-to pdf`). Figures (NAV, price, leasing, filing dates) are set in the script and must be updated
from the Company's latest NAV announcement and Form 10-Q/10-K before each regeneration. The investor presentation is maintained separately.
