const pptxgen = require("pptxgenjs");
const path = require("path");
const { applyTheme } = require(path.join(__dirname, "apply_theme.js"));
const FIG = require(path.join(__dirname, "figures.json"));
const money2 = (n) => "$" + Number(n).toFixed(2);
const millWords = (n) => "$" + Math.round(n / 1e6) + " million";
const D = { nav: money2(FIG.nav), price: money2(FIG.price), discount: Math.round((1 - FIG.price / FIG.nav) * 100) + "%", totalNavWords: millWords(FIG.totalNav), marketCapWords: millWords(FIG.price * FIG.units), units: Number(FIG.units).toLocaleString("en-US"), totalNavM: Math.round(FIG.totalNav / 1e5) / 10, marketCapM: Math.round(FIG.price * FIG.units / 1e5) / 10, report: FIG.reportForm + " for " + FIG.reportPeriod };

const IMG = (f) => path.join(__dirname, "img", f);
const THEME = { name: "Belpointe OZ", headFontFace: "Inter", bodyFontFace: "Inter", colors: { dk1: "1B2B3A", lt1: "FFFFFF", dk2: "1D4A3A", lt2: "F5F3EF", accent1: "2E5E4E", accent2: "5CA684", accent3: "4A8B6F", accent4: "C8956A", accent5: "E8DFD0", accent6: "3A6B8C", hlink: "2E5E4E", folHlink: "4A8B6F" } };
const NAVY = "1B2B3A", FOREST = "2E5E4E", SAGE = "4A8B6F", LGREEN = "5CA684", GOLD = "C8956A", CREAM = "F5F3EF", LINEN = "E8DFD0", WHITE = "FFFFFF", LINE = "DADDE0", N70 = "5F6A74", N50 = "8D949B", N30 = "BBC0C4";

const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE";
pres.author = "Cody Laidlaw"; pres.company = "Belpointe PREP, LLC"; pres.title = "Belpointe OZ Investor Presentation"; pres.subject = "Belpointe OZ Investor Presentation";
pres.theme = { headFontFace: "Inter", bodyFontFace: "Inter" };

const FOOT = "Belpointe PREP, LLC  ·  NYSE American: OZ  ·  " + FIG.month;
const mkLayout = (name, bg, dark) => pres.defineSlideMaster({
  title: name, background: { color: bg },
  objects: [
    { image: { x: 0.7, y: 6.9, w: 1.2, h: 0.26, path: IMG(dark ? "logo_dark.png" : "logo.png") } },
    { text: { text: FOOT, options: { x: 2.1, y: 6.88, w: 9, h: 0.3, fontSize: 9, color: dark ? "8C98A4" : N50, margin: 0, valign: "middle" } } },
    { placeholder: { options: { name: "label", type: "body", x: 0.7, y: 0.55, w: 10, h: 0.3, fontSize: 10, bold: true, color: dark ? LGREEN : FOREST, charSpacing: 2, margin: 0, valign: "top" } } },
    { placeholder: { options: { name: "title", type: "title", x: 0.7, y: 0.85, w: 11.9, h: 0.9, fontSize: 32, bold: true, color: dark ? WHITE : NAVY, margin: 0, valign: "top", align: "left" } } },
  ],
  slideNumber: { x: 12.2, y: 6.88, w: 0.45, h: 0.3, fontSize: 9, color: dark ? "8C98A4" : N50, align: "right" },
});
mkLayout("WHITE", WHITE, false); mkLayout("CREAM", CREAM, false); mkLayout("NAVY", NAVY, true);
pres.defineSlideMaster({ title: "COVER", background: { color: NAVY }, objects: [] });

const T = (s, text, o) => s.addText(text, Object.assign({ isTextBox: true, margin: 0, fontSize: 13, color: N70, valign: "top" }, o));
const label = (s, text) => s.addText(text.toUpperCase(), { placeholder: "label" });
const title = (s, text) => s.addText(text, { placeholder: "title" });
const src = (s, text, dark) => T(s, text, { x: 0.7, y: 6.28, w: 11.9, h: 0.52, fontSize: 8.5, color: dark ? "8C98A4" : N50 });
// site-style stat: forest left rule, big value, caption
const stat = (s, x, y, w, value, cap, dark, big = 30) => {
  s.addShape(pres.ShapeType.rect, { x, y, w: 0.04, h: 1.15, fill: { color: dark ? LGREEN : FOREST }, line: { color: dark ? LGREEN : FOREST } });
  T(s, value, { x: x + 0.22, y, w: w - 0.22, h: 0.7, fontSize: big, bold: true, color: dark ? WHITE : NAVY });
  T(s, cap, { x: x + 0.22, y: y + 0.72, w: w - 0.22, h: 0.45, fontSize: 10.5, color: dark ? "C9D1D8" : N70 });
};
const pill = (s, x, y, text, color, w = 1.9) => { s.addShape(pres.ShapeType.roundRect, { x, y, w, h: 0.3, rectRadius: 0.04, fill: { color }, line: { color } }); T(s, text.toUpperCase(), { x, y, w, h: 0.3, fontSize: 8.5, bold: true, color: WHITE, align: "center", valign: "middle", charSpacing: 1 }); };
const photo = (s, f, x, y, w, h) => s.addImage({ path: IMG(f), x, y, w, h, sizing: { type: "cover", w, h } });
const cap = (s, text, x, y, w) => T(s, text, { x, y, w, h: 0.3, fontSize: 9, color: N50 });

// status colors
const ST = { leasing: FOREST, operating: SAGE, dev: NAVY, pre: "3F5567", sale: GOLD, contract: "A8763F" };

// ================= 1 Cover =================
pres.addSection({ title: "Introduction" });
let s = pres.addSlide({ masterName: "COVER", sectionTitle: "Introduction" });
photo(s, "nyse.jpg", 0, 0, 13.33, 4.6);
s.addShape(pres.ShapeType.rect, { x: 0, y: 0, w: 13.33, h: 4.6, fill: { color: NAVY, transparency: 35 }, line: { color: NAVY, transparency: 100 } });
s.addImage({ path: IMG("logo_dark.png"), x: 0.7, y: 0.6, w: 2.3, h: 0.5 });
T(s, "Investor Presentation", { x: 0.7, y: 4.95, w: 8, h: 0.8, fontSize: 40, bold: true, color: WHITE });
T(s, "The only Qualified Opportunity Fund listed on a national securities exchange", { x: 0.7, y: 5.8, w: 8, h: 0.5, fontSize: 16, color: "C9D1D8" });
T(s, "NYSE American: OZ", { x: 9.3, y: 5.0, w: 3.4, h: 0.5, fontSize: 20, bold: true, color: LGREEN, align: "right" });
T(s, FIG.month, { x: 9.3, y: 5.5, w: 3.4, h: 0.4, fontSize: 13, color: "C9D1D8", align: "right" });
T(s, "Belpointe PREP, LLC (“Belpointe OZ”)", { x: 0.7, y: 6.6, w: 8, h: 0.3, fontSize: 10, color: "8C98A4" });
s.addNotes("Cover photo: the NYSE floor on October 18, 2021, the day Class A units began trading.");

// ================= 2 Important information =================
s = pres.addSlide({ masterName: "WHITE", sectionTitle: "Introduction" });
label(s, "Please read"); title(s, "Important information");
const para = (h, b) => [{ text: h + " ", options: { bold: true, color: NAVY } }, { text: b, options: { breakLine: true } }, { text: " ", options: { breakLine: true, fontSize: 6 } }];
T(s, [
  ...para("Not an offer.", "This presentation is for informational purposes only and does not constitute an offer to sell or the solicitation of an offer to buy any securities of Belpointe PREP, LLC (“Belpointe OZ” or the “Company”). Class A units trade on the NYSE American under the symbol OZ. Investing in Class A units involves a high degree of risk, including the possible loss of the entire investment. Before investing, read the Company’s filings with the Securities and Exchange Commission, including the risk factors in its Annual Report on Form 10-K and Quarterly Reports on Form 10-Q, available at www.sec.gov."),
  ...para("Forward-looking statements.", "This presentation contains forward-looking statements within the meaning of Section 27A of the Securities Act of 1933 and Section 21E of the Securities Exchange Act of 1934, including statements about leasing, stabilization, development plans, dispositions, net asset value and the Opportunity Zone program. Words such as “anticipate,” “estimate,” “expect,” “project,” “plan,” “intend,” “believe,” “may” and “will” identify forward-looking statements. They reflect current expectations and are subject to the risks and uncertainties described in the Company’s SEC filings. Actual results may differ materially, and the Company undertakes no obligation to update any forward-looking statement except as required by law."),
  ...para("Net asset value.", "NAV per Class A unit is calculated quarterly by the Company’s Manager under the valuation policy described in the Company’s SEC filings. NAV is an estimate, not a guarantee of value. The market price of Class A units may be above or below NAV and may never reach it. Market price data is delayed and provided for information only. Figures in this document are updated automatically from public sources, including the Company’s filings with the SEC and third-party market data, and may contain errors, omissions or delays. They are provided for informational purposes only. The Company’s filings with the SEC, available at www.sec.gov, are the official source and control if there is any difference."),
  { text: "Tax matters. ", options: { bold: true, color: NAVY } }, { text: "Nothing in this presentation is tax advice. Opportunity Zone benefits depend on each investor’s circumstances and on current law, which may change. Consult your own tax adviser. Belpointe OZ is managed by Belpointe PREP Manager, LLC (the “Manager”) and sponsored by Belpointe, LLC (the “Sponsor”). Property information is from the Company’s " + D.report + " unless otherwise stated; leasing percentages are the Company’s estimates as of " + FIG.leasingAsOf + " and change over time." },
], { x: 0.7, y: 1.9, w: 11.9, h: 4.6, fontSize: 11, color: N70, paraSpaceAfter: 0 });

// ================= 3 At a glance =================
pres.addSection({ title: "Overview" });
s = pres.addSlide({ masterName: "CREAM", sectionTitle: "Overview" });
label(s, "Belpointe OZ"); title(s, "Built properties, a public company, and a material discount to net asset value");
stat(s, 0.7, 2.2, 3.8, D.price, "Class A unit price, NYSE American close, " + FIG.priceDate + " (delayed)");
stat(s, 4.8, 2.2, 3.8, D.nav, "Net asset value per Class A unit as of " + FIG.navAsOf + ", announced " + FIG.navAnnounced);
stat(s, 8.9, 2.2, 3.8, D.discount + " below NAV", "Market price on " + FIG.priceDate + " divided by NAV as of " + FIG.navAsOf + ". NAV is an estimate, not a guarantee of value.");
stat(s, 0.7, 3.8, 3.8, "693", "Completed residences at Aster & Links (424) and VIV (269)");
stat(s, 4.8, 3.8, 3.8, D.totalNavWords, "Net asset value as of " + FIG.navAsOf + " (" + FIG.navSource + ")");
stat(s, 8.9, 3.8, 3.8, D.marketCapWords, "Market capitalization at " + D.price + " per unit (" + FIG.priceDate + ") on " + D.units + " Class A units outstanding at " + FIG.unitsAsOf);
src(s, "Sources: " + D.report + "; " + FIG.navSource + " (NAV and total NAV as of " + FIG.navAsOf + "). Market capitalization uses the " + FIG.priceDate + " closing price of " + D.price + " and " + D.units + " Class A units outstanding at " + FIG.unitsAsOf + " (" + FIG.unitsSource + "). NAV is calculated quarterly by the Manager under the Company’s valuation policy and is an estimate, not a guarantee of value.");

// ================= 4 Why now =================
s = pres.addSlide({ masterName: "WHITE", sectionTitle: "Overview" });
label(s, "The case"); title(s, "Why Belpointe OZ now");
const reasons = [
  ["A material discount to NAV", "Class A units have recently traded at a material discount to the Company’s most recently announced net asset value. A buyer on the exchange acquires the underlying real estate below the Company’s own estimate of value."],
  ["Completed, income-producing properties", "Aster & Links and VIV are built and leasing: " + FIG.alQualifier + " " + FIG.alPct + "% and " + FIG.vivPct + "% leased as of " + FIG.leasingAsOf + ". Most Opportunity Zone funds ask investors to commit capital to properties that do not exist yet, and the Company already owns assets in OZ 2.0 zones, including rural zones."],
  ["Construction complete at Aster & Links and VIV", "Construction risk on the two completed properties is behind us. Lease-up risk at each is substantially reduced, though not eliminated, until the property is stabilized. The development pipeline carries its own construction risk."],
  ["The only listed Qualified Opportunity Fund", "Buy through any brokerage account at the market price. No subscription process, no accreditation requirement, no capital calls, and no lock-up when an investor wants to sell. Liquidity depends on trading volume and market conditions."],
];
reasons.forEach((r, i) => {
  const x = 0.7 + (i % 2) * 6.1, y = 2.0 + Math.floor(i / 2) * 2.15;
  s.addShape(pres.ShapeType.rect, { x, y, w: 0.04, h: 1.85, fill: { color: FOREST }, line: { color: FOREST } });
  T(s, String(i + 1).padStart(2, "0"), { x: x + 0.25, y, w: 0.6, h: 0.4, fontSize: 12, bold: true, color: SAGE });
  T(s, r[0], { x: x + 0.25, y: y + 0.35, w: 5.5, h: 0.45, fontSize: 17, bold: true, color: NAVY });
  T(s, r[1], { x: x + 0.25, y: y + 0.85, w: 5.5, h: 1.0, fontSize: 11.5, color: N70 });
});
src(s, "NAV is an estimate, not a guarantee of value, and the market price may never reach it. Leasing percentages are Company estimates as of " + FIG.leasingAsOf + ". Comparisons describe general features of private Opportunity Zone funds and are not a description of any specific fund.");

// ================= 5 Price vs NAV =================
s = pres.addSlide({ masterName: "WHITE", sectionTitle: "Overview" });
label(s, "The contrarian entry point"); title(s, "Market price vs. net asset value");
s.addChart(pres.ChartType.bar, [{ name: "Per Class A unit", labels: ["NAV per unit, " + FIG.navAsOf, "Market price, " + FIG.priceDate], values: [FIG.nav, FIG.price] }], {
  x: 0.7, y: 2.0, w: 5.9, h: 4.2, barDir: "bar", chartColors: [FOREST, NAVY], showLegend: false, showTitle: true, title: "Per Class A unit", titleFontSize: 11, titleColor: NAVY, titleFontFace: "+mn-lt", titleAlign: "left",
  showValue: true, dataLabelPosition: "outEnd", dataLabelFormatCode: "$#,##0.00", dataLabelFontSize: 11, dataLabelColor: NAVY, dataLabelFontFace: "+mn-lt",
  catAxisLabelColor: N70, catAxisLabelFontSize: 10, catAxisLabelFontFace: "+mn-lt", valAxisHidden: true, valGridLine: { style: "none" }, catGridLine: { style: "none" }, valAxisMinVal: 0, valAxisMaxVal: 150, barGapWidthPct: 55,
});
s.addChart(pres.ChartType.bar, [{ name: "Whole company", labels: ["Net asset value, " + FIG.navAsOf, "Market capitalization, " + FIG.priceDate], values: [D.totalNavM, D.marketCapM] }], {
  x: 6.9, y: 2.0, w: 5.7, h: 4.2, barDir: "bar", chartColors: [FOREST, NAVY], showLegend: false, showTitle: true, title: "Whole company, $ millions", titleFontSize: 11, titleColor: NAVY, titleFontFace: "+mn-lt", titleAlign: "left",
  showValue: true, dataLabelPosition: "outEnd", dataLabelFormatCode: "$#,##0", dataLabelFontSize: 11, dataLabelColor: NAVY, dataLabelFontFace: "+mn-lt",
  catAxisLabelColor: N70, catAxisLabelFontSize: 10, catAxisLabelFontFace: "+mn-lt", valAxisHidden: true, valGridLine: { style: "none" }, catGridLine: { style: "none" }, valAxisMinVal: 0, valAxisMaxVal: 580, barGapWidthPct: 55,
});
src(s, "Market capitalization: " + D.price + " closing price on " + FIG.priceDate + " multiplied by " + D.units + " Class A units outstanding at " + FIG.unitsAsOf + " (" + FIG.unitsSource + "). NAV is calculated quarterly by the Manager under the Company’s valuation policy; it is an estimate, not a guarantee of value, and the market price may never reach it.");

// ================= 6 Portfolio ledger =================
pres.addSection({ title: "Portfolio" });
s = pres.addSlide({ masterName: "CREAM", sectionTitle: "Portfolio" });
label(s, "Portfolio"); title(s, "Every asset: status, purchase price and debt");
const cols = [["Asset", 0.7, 2.75], ["Location", 3.5, 1.1], ["Status", 4.65, 1.75], ["Zone", 6.45, 1.55], ["Purchase price", 8.05, 2.35], ["Debt at " + FIG.debtAsOf, 10.45, 2.15]];
const assets = [
  ["Aster & Links, 1991 Main Street and 1900 Fruitville Road", "Sarasota, FL", "Completed · Leasing (>" + FIG.alPct + "%)", ST.leasing, "Grandfathered OZ 1.0", "$30.3M: $20.7M site, $4.9M adjacent, $4.7M Fruitville", FIG.alDebtShort],
  ["VIV, 1000 First Avenue North, and 900 First Avenue North", "St. Petersburg, FL", "Completed · Leasing (>" + FIG.vivPct + "%)", ST.leasing, "OZ 2.0 zone", "$14.6M: $12.1M site, $2.5M 900 First", FIG.vivDebtShort],
  ["901-909 Central Avenue North", "St. Petersburg, FL", "Operating · Fully leased", ST.operating, "OZ 2.0 zone", "$2.6M", "No debt"],
  ["1700 Main Street", "Sarasota, FL", "Predevelopment", ST.pre, "Grandfathered OZ 1.0", "$6.9M", "No debt"],
  ["1701 and 1710 Ringling Boulevard", "Sarasota, FL", "Development · Office renovation", ST.dev, "Grandfathered OZ 1.0", "$7.0M", "No debt"],
  ["1702 Ringling Boulevard", "Sarasota, FL", "Held for future development", ST.pre, "Grandfathered OZ 1.0", "$1.5M", "No debt"],
  ["497-501 Middle Turnpike and Cedar Swamp Road", "Storrs, CT", "Predevelopment", ST.pre, "Rural OZ 2.0 zone", "$5.4M ($3.8M + $1.6M); $0.3M Cedar Swamp", "No debt"],
  ["1750 Storrs Road", "Storrs, CT", "Predevelopment", ST.pre, "Rural OZ 2.0 zone", "$5.5M", "No debt"],
  ["Storrs Road", "Storrs, CT", "Predevelopment", ST.pre, "Rural OZ 2.0 zone", "$0.1M", "No debt"],
  ["900 8th Avenue South", "Nashville, TN", "Under contract to be sold", ST.contract, "OZ 2.0 zone", "$19.7M", "$8.5M land loan, 9.50%, July 2027"],
  ["690/1106, 1130 and 1400 Davidson Street", "Nashville, TN", "Intended for sale", ST.sale, "OZ 2.0 zones", "$39.5M: $21.0M, $2.1M, $16.4M", "No debt"],
];
let y0 = 1.8;
cols.forEach((c) => T(s, c[0].toUpperCase(), { x: c[1], y: y0, w: c[2], h: 0.3, fontSize: 8.5, bold: true, color: FOREST, charSpacing: 1 }));
s.addShape(pres.ShapeType.line, { x: 0.7, y: y0 + 0.32, w: 11.9, h: 0, line: { color: NAVY, width: 1 } });
const rowH = 0.33;
assets.forEach((a, i) => {
  const y = y0 + 0.4 + i * rowH;
  T(s, a[0], { x: cols[0][1], y, w: cols[0][2], h: rowH, fontSize: 9.5, bold: true, color: NAVY, valign: "middle" });
  T(s, a[1], { x: cols[1][1], y, w: cols[1][2], h: rowH, fontSize: 9.5, color: N70, valign: "middle" });
  T(s, a[2], { x: cols[2][1], y, w: cols[2][2], h: rowH, fontSize: 9, bold: true, color: a[3], valign: "middle" });
  T(s, a[4], { x: cols[3][1], y, w: cols[3][2], h: rowH, fontSize: 9, color: a[4].startsWith("Rural") ? SAGE : a[4].startsWith("OZ 2.0") ? FOREST : N70, valign: "middle" });
  T(s, a[5], { x: cols[4][1], y, w: cols[4][2], h: rowH, fontSize: 9, color: N70, valign: "middle" });
  T(s, a[6], { x: cols[5][1], y, w: cols[5][2], h: rowH, fontSize: 9, color: a[6] === "No debt" ? N50 : N70, valign: "middle" });
  s.addShape(pres.ShapeType.line, { x: 0.7, y: y + rowH, w: 11.9, h: 0, line: { color: LINE, width: 0.5 } });
});
src(s, "Source: " + D.report + ", Our Investments and Debt, Net. Purchase prices include transaction costs and exclude construction and other capitalized costs. Leasing percentages are Company estimates as of " + FIG.leasingAsOf + ". Total debt " + FIG.totalDebt + " as of " + FIG.debtAsOf + ". Zone designations reflect the Company’s assessment under current law and guidance (OZ 2.0 zones designated under the One Big Beautiful Bill Act of 2025; grandfathered OZ 1.0 sites designated in 2018) and are subject to final Treasury and IRS guidance.");

// ================= 7 Aster & Links =================
s = pres.addSlide({ masterName: "WHITE", sectionTitle: "Portfolio" });
label(s, "Completed · Leasing"); title(s, "Aster & Links, Sarasota, Florida");
photo(s, "aster1.jpg", 0.7, 1.95, 7.3, 4.1);
cap(s, "Aster & Links, downtown Sarasota.", 0.7, 6.08, 7.3);
pill(s, 8.4, 1.95, "Leasing · >" + FIG.alPct + "% leased", ST.leasing, 2.1); pill(s, 10.6, 1.95, "Grandfathered OZ 1.0", "3F5567", 2.1);
stat(s, 8.4, 2.5, 2.1, "424", "residences", false, 22); stat(s, 10.6, 2.5, 2.1, "51,000 sf", "retail, Sprouts-anchored", false, 22);
T(s, [
  { text: "Two ten-story buildings on a 5.13-acre site with one-, two-, three- and four-bedroom apartments, townhome-style penthouses and six guest suites; more than 900 garage and surface parking spaces. Construction was substantially completed and leasing began in 2024.", options: { breakLine: true } },
  { text: " ", options: { breakLine: true, fontSize: 6 } },
  { text: "Purchase price: ", options: { bold: true, color: NAVY } }, { text: "$20.7 million for the site, $4.9 million for the adjacent parcel (August 2023) and $4.7 million for 1900 Fruitville Road, used for retail parking.", options: { breakLine: true } },
  { text: " ", options: { breakLine: true, fontSize: 6 } },
  { text: "Debt at " + FIG.debtAsOf + ": ", options: { bold: true, color: NAVY } }, { text: FIG.alDebtDrawn + " drawn on a " + FIG.alDebtFacility + " mortgage and mezzanine facility with SM Finance III LLC at one-month term SOFR (3.25% floor) plus 2.55% blended, interest-only, maturing October 2027 with two one-year extensions. Interest rate cap at 6.0% through October 15, 2027. Leasing percentage is a Company estimate as of " + FIG.leasingAsOf + "." },
], { x: 8.4, y: 3.85, w: 4.3, h: 2.5, fontSize: 10, color: N70 });

s = pres.addSlide({ masterName: "WHITE", sectionTitle: "Portfolio" });
label(s, "Completed · Leasing"); title(s, "Aster & Links: residences and amenities");
photo(s, "aster_amenity.jpg", 0.7, 1.95, 3.9, 3.5); photo(s, "aster_amenity2.jpg", 4.7, 1.95, 3.9, 3.5); photo(s, "aster_interior.jpg", 8.7, 1.95, 3.9, 3.5);
cap(s, "Resident amenity deck.", 0.7, 5.5, 3.9); cap(s, "Resident lounge.", 4.7, 5.5, 3.9); cap(s, "Residence interior.", 8.7, 5.5, 3.9);
T(s, "Grandfathered OZ 1.0 site in downtown Sarasota. In July 2026, CNN named Sarasota the number one town in America’s Best Towns to Visit, and Niche ranked Sarasota County the number one county for retirees in the nation for 2026. Sarasota is headquarters to Boar’s Head Provisions, CAE Healthcare, Helios Technologies and Voalte, and home to the University of South Florida, Florida State University’s College of Medicine campus, Ringling College and State College of Florida.", { x: 0.7, y: 5.85, w: 11.9, h: 0.7, fontSize: 10, color: N70 });

// ================= 8 VIV =================
s = pres.addSlide({ masterName: "WHITE", sectionTitle: "Portfolio" });
label(s, "Completed · Leasing"); title(s, "VIV, St. Petersburg, Florida");
photo(s, "viv_sunset.jpg", 0.7, 1.95, 7.3, 4.1);
cap(s, "VIV, 1000 1st Avenue North, EDGE District.", 0.7, 6.08, 7.3);
pill(s, 8.4, 1.95, "Leasing · >" + FIG.vivPct + "% leased", ST.leasing, 2.1); pill(s, 10.6, 1.95, "OZ 2.0 zone", FOREST, 2.1);
stat(s, 8.4, 2.5, 2.1, "269", "residences, two towers", false, 22); stat(s, 10.6, 2.5, 2.1, "15,500 sf", "ground-floor retail", false, 22);
T(s, [
  { text: "Two 11-story residential towers above a four-story parking structure on about 1.6 acres, one block from Tropicana Field and one mile from the downtown waterfront. Construction was substantially completed in 2025; leasing began October 2025 and first move-ins followed in November 2025.", options: { breakLine: true } },
  { text: " ", options: { breakLine: true, fontSize: 6 } },
  { text: "Purchase price: ", options: { bold: true, color: NAVY } }, { text: "$12.1 million for the site and $2.5 million for 900 First Avenue North, a two-tenant retail building whose development rights were transferred to VIV.", options: { breakLine: true } },
  { text: " ", options: { breakLine: true, fontSize: 6 } },
  { text: "Debt at " + FIG.debtAsOf + ": ", options: { bold: true, color: NAVY } }, { text: FIG.vivDebtDrawn + " drawn on a " + FIG.vivDebtFacility + " construction loan at one-month term SOFR plus 3.80% (7.55% floor), maturing June 28, 2027 with two one-year extensions. Interest rate cap at 6.25% on the full notional through July 1, 2027. Leasing percentage is a Company estimate as of " + FIG.leasingAsOf + "." },
], { x: 8.4, y: 3.85, w: 4.3, h: 2.5, fontSize: 10, color: N70 });

s = pres.addSlide({ masterName: "WHITE", sectionTitle: "Portfolio" });
label(s, "Completed · Leasing"); title(s, "VIV: pool deck, club and the downtown skyline");
photo(s, "viv_pool.jpg", 0.7, 1.95, 3.9, 3.5); photo(s, "viv_club.jpg", 4.7, 1.95, 3.9, 3.5); photo(s, "viv_wide.jpg", 8.7, 1.95, 3.9, 3.5);
cap(s, "Courtyard pool deck.", 0.7, 5.5, 3.9); cap(s, "Resident clubroom.", 4.7, 5.5, 3.9); cap(s, "Looking east to downtown and Tampa Bay.", 8.7, 5.5, 3.9);
T(s, "Amenities include a clubroom, fitness center, courtyard pool, shared working space and a leasing office. St. Petersburg is Florida’s fifth-largest city and placed 51st on Niche’s 2026 Best Cities to Live in America list. In January 2026, Site Selection Magazine ranked Tampa fifth among the top ten U.S. cities for corporate headquarters.", { x: 0.7, y: 5.85, w: 11.9, h: 0.6, fontSize: 10, color: N70 });

// ================= 9 Sarasota pipeline =================
s = pres.addSlide({ masterName: "CREAM", sectionTitle: "Portfolio" });
label(s, "Development pipeline · Grandfathered OZ 1.0 sites"); title(s, "Downtown Sarasota: the next phase");
const cards9 = [
  ["1700 Main Street", "Predevelopment", ST.pre, "Grandfathered OZ 1.0 site. 1.3-acre site planned as a 12-story community of about 130 apartment homes with roughly 6,000 square feet of retail and a 300-space garage. Purchase price $6.9 million. No debt."],
  ["1701 and 1710 Ringling Boulevard", "Development · Office renovation", ST.dev, "Grandfathered OZ 1.0 site. 1.6-acre site with a six-story office building to be renovated into about 80,000 square feet of modern office space, plus a 128-space lot. A 20-year lease covering about 42,000 square feet was signed with the existing tenant at acquisition. Purchase price $7.0 million. No debt."],
  ["1702 Ringling Boulevard", "Held for future development", ST.pre, "Grandfathered OZ 1.0 site. Fully leased 1,546 square foot single-tenant office building on 0.327 acres, held for future multifamily development. Purchase price $1.5 million. No debt."],
];
cards9.forEach((c, i) => {
  const x = 0.7 + i * 4.05;
  s.addShape(pres.ShapeType.rect, { x, y: 2.0, w: 3.85, h: 4.1, fill: { color: WHITE }, line: { color: LINE, width: 0.75 } });
  pill(s, x + 0.3, 2.3, c[1], c[2], 2.6);
  T(s, c[0], { x: x + 0.3, y: 2.8, w: 3.25, h: 0.7, fontSize: 15, bold: true, color: NAVY });
  T(s, c[3], { x: x + 0.3, y: 3.5, w: 3.25, h: 2.4, fontSize: 10, color: N70 });
});
src(s, "Source: " + D.report + ". Development plans are subject to change, financing, entitlements and market conditions. Zone designations reflect the Company’s assessment under current law and guidance (OZ 2.0 zones designated under the One Big Beautiful Bill Act of 2025; grandfathered OZ 1.0 sites designated in 2018) and are subject to final Treasury and IRS guidance.");

// ================= 10 Storrs =================
s = pres.addSlide({ masterName: "WHITE", sectionTitle: "Portfolio" });
label(s, "Development pipeline · Rural OZ 2.0 zone"); title(s, "Storrs, Connecticut: near UConn");
const cards10 = [
  ["497-501 Middle Turnpike and Cedar Swamp Road", "Predevelopment", ST.pre, "Qualified rural opportunity zone (OZ 2.0). About 60 acres (30 of former golf course, 30 of wetlands) plus an adjacent 1.1-acre parcel, less than a mile from UConn’s main campus. Planned as an approximately 261-apartment-home community with a single-family home, clubroom, fitness center and meadow. Purchase price: $3.8 million initial capital for the controlling interest, $1.6 million to redeem the remaining partner (March 2026), and $0.3 million for Cedar Swamp Road. No debt."],
  ["1750 Storrs Road", "Predevelopment", ST.pre, "Qualified rural opportunity zone (OZ 2.0). About 19 acres near UConn, planned as a multifamily mixed-use development with one-, two- and three-bedroom apartments and a clubhouse. Purchase price $5.5 million. No debt."],
  ["Storrs Road", "Predevelopment", ST.pre, "Qualified rural opportunity zone (OZ 2.0). 9.0-acre parcel near UConn intended for multifamily development. Purchase price $0.1 million. No debt."],
];
cards10.forEach((c, i) => {
  const x = 0.7 + i * 4.05;
  s.addShape(pres.ShapeType.rect, { x, y: 2.0, w: 3.85, h: 4.1, fill: { color: CREAM }, line: { color: CREAM } });
  pill(s, x + 0.3, 2.3, c[1], c[2], 2.6);
  T(s, c[0], { x: x + 0.3, y: 2.8, w: 3.25, h: 0.7, fontSize: 15, bold: true, color: NAVY });
  T(s, c[3], { x: x + 0.3, y: 3.5, w: 3.25, h: 2.5, fontSize: 10, color: N70 });
});
src(s, "Source: " + D.report + ". UConn ranked 32nd among Top Public Schools in the 2027 U.S. News rankings; over 21,000 undergraduates attended the Storrs campus in Fall 2025. Zone designations reflect the Company’s assessment under current law and guidance (OZ 2.0 zones designated under the One Big Beautiful Bill Act of 2025; grandfathered OZ 1.0 sites designated in 2018) and are subject to final Treasury and IRS guidance.");

// ================= 11 Nashville =================
s = pres.addSlide({ masterName: "CREAM", sectionTitle: "Portfolio" });
label(s, "Dispositions · OZ 2.0 zones"); title(s, "Nashville, Tennessee: land positioned for sale");
const cards11 = [
  ["900 8th Avenue South", "Under contract to be sold", ST.contract, "OZ 2.0 zone. 3.2-acre assemblage at the north end of the 8th Avenue South District, approved for up to 365 residential units and seven stories. Under contract with WP South Acquisitions, L.L.C. for $19.6 million; entitlements date November 30, 2026, inspection date 30 days later, closing up to 180 days after that. Purchase price $19.7 million. Debt: $8.5 million fixed-rate land loan at 9.50%, maturing July 2027."],
  ["690/1106, 1130 and 1400 Davidson Street", "Intended for sale", ST.sale, "OZ 2.0 zone. About 15.6 acres on the East Bank, rezoned from industrial to high-density multifamily and mixed commercial use. Purchase price $39.5 million ($21.0 million, $2.1 million and $16.4 million). No debt. A 47-acre former industrial site about one mile west sold in August 2025 for approximately $245 million. The Company intends to sell the Davidson properties to maximize value for unitholders."],
];
cards11.forEach((c, i) => {
  const x = 0.7 + i * 6.1;
  s.addShape(pres.ShapeType.rect, { x, y: 2.0, w: 5.8, h: 4.1, fill: { color: WHITE }, line: { color: LINE, width: 0.75 } });
  pill(s, x + 0.3, 2.3, c[1], c[2], 2.6);
  T(s, c[0], { x: x + 0.3, y: 2.8, w: 5.2, h: 0.5, fontSize: 16, bold: true, color: NAVY });
  T(s, c[3], { x: x + 0.3, y: 3.4, w: 5.2, h: 2.6, fontSize: 11, color: N70 });
});
src(s, "Source: " + D.report + ". The 900 8th Avenue South sale is subject to the buyer’s inspection and closing conditions and may not close. There is no assurance the Davidson properties will sell at any particular price. Zone designations reflect the Company’s assessment under current law and guidance (OZ 2.0 zones designated under the One Big Beautiful Bill Act of 2025; grandfathered OZ 1.0 sites designated in 2018) and are subject to final Treasury and IRS guidance.");

// ================= Plan =================
s = pres.addSlide({ masterName: "CREAM", sectionTitle: "Portfolio" });
label(s, "Objectives"); title(s, "The plan from here");
const steps2 = [
  ["Stabilize Aster & Links and VIV", "Complete lease-up of the two completed properties. Aster & Links was " + FIG.alQualifier + " " + FIG.alPct + "% leased and VIV " + FIG.vivQualifier + " " + FIG.vivPct + "% leased as of " + FIG.leasingAsOf + "."],
  ["Refinance the completed properties", "Both are financed with floating-rate loans maturing in 2027 with extension options. The objective is to refinance on stabilization, which depends on occupancy, net operating income, interest rates and lender appetite."],
  ["Monetize Nashville properties", "900 8th Avenue South is under contract to be sold. The Davidson Street properties could be sold, joint ventured or developed. Proceeds from any transaction could be used for distributions, unit repurchases, debt reduction or reinvestment, in each case subject to Board approval, loan covenants, cash needs and reserves."],
  ["Distributions when the Company is in a position to pay them", "Paying distributions to unitholders is one of the Company’s principal objectives once it is in a position to do so. Distributions are subject to many factors, including Board approval, the Company’s results and cash needs, compliance with loan covenants, and the maintenance of adequate reserves."],
];
steps2.forEach((st, i) => {
  const x = 0.7 + (i % 2) * 6.1, y = 1.95 + Math.floor(i / 2) * 2.15;
  s.addShape(pres.ShapeType.rect, { x, y, w: 5.9, h: 2.0, fill: { color: WHITE }, line: { color: LINE, width: 0.75 } });
  T(s, String(i + 1).padStart(2, "0"), { x: x + 0.3, y: y + 0.25, w: 0.6, h: 0.35, fontSize: 11, bold: true, color: SAGE });
  T(s, st[0], { x: x + 0.3, y: y + 0.55, w: 5.3, h: 0.4, fontSize: 13, bold: true, color: NAVY });
  T(s, st[1], { x: x + 0.3, y: y + 0.98, w: 5.3, h: 0.95, fontSize: 10, color: N70 });
});
src(s, "These objectives are forward-looking statements, not guidance or a commitment. There is no assurance that the properties will stabilize on any timetable, that refinancing will be available on acceptable terms, that the Nashville properties will sell, or that any distribution will be paid. See the risk factors in the Company’s Form 10-K and Form 10-Q and the Chairman and CEO letter to unitholders.");

// ================= 12 Tax benefits =================
pres.addSection({ title: "Opportunity Zones" });
s = pres.addSlide({ masterName: "WHITE", sectionTitle: "Opportunity Zones" });
label(s, "Opportunity Zone program"); title(s, "Defer, reduce, eliminate");
const tax = [
  ["Defer", "Reinvest an eligible capital gain in a Qualified Opportunity Fund within 180 days of the sale and defer the capital gain. Under OZ 1.0 the deferred gain is recognized on the earlier of December 31, 2026 or an inclusion event such as a sale of the investment. Under OZ 2.0, for investments made on or after January 1, 2027, the deferral runs five years from the date of investment."],
  ["Reduce", "The amount of deferred gain recognized is reduced by a basis step-up: 10% after five years under OZ 2.0, with 30% for qualified rural opportunity funds. Under OZ 1.0, the 10% step-up applied to investments that reached their five-year anniversary on or before December 31, 2026."],
  ["Eliminate", "Hold the Qualified Opportunity Fund investment for at least ten years and elect to step the basis up to fair market value on sale, eliminating federal capital gains tax on the appreciation of the investment. OZ 1.0 investments keep this benefit; under OZ 2.0 the election is available for up to 30 years."],
];
tax.forEach((t, i) => {
  const x = 0.7 + i * 4.05;
  s.addShape(pres.ShapeType.rect, { x, y: 2.0, w: 0.04, h: 3.3, fill: { color: FOREST }, line: { color: FOREST } });
  T(s, t[0], { x: x + 0.25, y: 2.0, w: 3.5, h: 0.55, fontSize: 24, bold: true, color: FOREST });
  T(s, t[1], { x: x + 0.25, y: 2.6, w: 3.5, h: 2.7, fontSize: 11, color: N70 });
});
T(s, [{ text: "Bridge from OZ 1.0 to OZ 2.0. ", options: { bold: true, color: NAVY } }, { text: "Exiting an OZ 1.0 investment any time before the December 31, 2026 recognition date triggers a new 180-day reinvestment period. Nothing in the OZ 1.0 or OZ 2.0 statutes, regulations or IRS guidance as it exists today prohibits overlapping that exit with the start of the OZ 2.0 program on or after January 1, 2027." }, { text: "2", options: { superscript: true } }, { text: " Review with a qualified tax professional. See note 2 on the final page." }], { x: 0.7, y: 5.45, w: 11.9, h: 0.85, fontSize: 10.5, color: N70 });
src(s, "Nothing here is tax advice. Benefits depend on each investor’s circumstances and on current law, which may change. Consult your own tax adviser.");

// ================= 13 How to invest =================
pres.addSection({ title: "Investing" });
s = pres.addSlide({ masterName: "NAVY", sectionTitle: "Investing" });
label(s, "Two ways in"); title(s, "How to invest");
const steps = [["01", "Realize a gain", "Sell stock, a business, real estate, crypto or another appreciated asset. Only the gain needs to be reinvested."], ["02", "Buy OZ within 180 days", [{ text: "Purchase Class A units through any brokerage account at the prevailing market price." }, { text: "1", options: { superscript: true } }, { text: " No minimum, no accreditation, no subscription documents." }]], ["03", "Elect and hold", "Make the deferral election on IRS Form 8949 and report the investment on Form 8997 with your tax return, then hold for the ten-year benefit."]];
steps.forEach((st, i) => {
  const x = 0.7 + i * 4.05;
  s.addShape(pres.ShapeType.rect, { x, y: 2.0, w: 0.04, h: 2.2, fill: { color: LGREEN }, line: { color: LGREEN } });
  T(s, st[0], { x: x + 0.25, y: 2.0, w: 1, h: 0.4, fontSize: 12, bold: true, color: LGREEN });
  T(s, st[1], { x: x + 0.25, y: 2.4, w: 3.5, h: 0.5, fontSize: 17, bold: true, color: WHITE });
  T(s, st[2], { x: x + 0.25, y: 2.95, w: 3.5, h: 1.3, fontSize: 11, color: "C9D1D8" });
});
T(s, "From time to time, in its sole discretion, the Company may offer Class A units directly under an effective registration statement. Availability and terms are subject to change.", { x: 0.7, y: 4.4, w: 11.9, h: 0.4, fontSize: 10, color: "8C98A4" });
T(s, "Investor Relations", { x: 0.7, y: 5.0, w: 6, h: 0.4, fontSize: 15, bold: true, color: WHITE });
T(s, [{ text: "203-883-1944  ·  IR@belpointeoz.com", options: { breakLine: true } }, { text: "255 Glenville Road, Greenwich, CT 06831", options: { breakLine: true } }, { text: "belpointeoz.com  ·  SEC filings at www.sec.gov" }], { x: 0.7, y: 5.45, w: 7, h: 1.0, fontSize: 12, color: "C9D1D8" });
T(s, "NYSE American: OZ", { x: 8.3, y: 5.0, w: 4.3, h: 0.5, fontSize: 22, bold: true, color: LGREEN, align: "right" });
T(s, "Class A units trade at prevailing market prices and may be above or below NAV. Brokerage commissions and expenses will reduce returns.", { x: 8.3, y: 5.55, w: 4.3, h: 0.8, fontSize: 9.5, color: "8C98A4", align: "right" });
T(s, [{ text: "1", options: { superscript: true } }, { text: " Based on current law, Treasury regulations and IRS guidance, and on Belpointe OZ’s structure as a Qualified Opportunity Fund, Class A units acquired with eligible capital gains within the 180-day period, whether purchased on NYSE American or directly from the Company, should be a qualifying investment. Opportunity Zone tax benefits depend on each investor’s circumstances. Consult your tax adviser.", options: { breakLine: true } }, { text: "2", options: { superscript: true } }, { text: " Based on current law, Treasury regulations and IRS guidance, each of which is subject to change. Whether gain recognized on exiting an existing Qualified Opportunity Fund investment is eligible for reinvestment, and when the 180-day period begins and ends, depend on each investor’s circumstances. This is general information, not tax advice. Consult your tax adviser." }], { x: 0.7, y: 6.24, w: 11.9, h: 0.6, fontSize: 7.5, color: "8C98A4", valign: "top" });

(async () => {
  const out = path.join(__dirname, "Belpointe OZ Investor Presentation.pptx");
  await pres.writeFile({ fileName: out });
  await applyTheme(out, THEME);
  console.log("wrote", out);
})();
