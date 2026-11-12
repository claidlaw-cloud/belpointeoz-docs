const fs = require("fs");
const path = require("path");
const { Document, Packer, Paragraph, TextRun, HeadingLevel, LevelFormat, AlignmentType, Table, TableRow, TableCell, WidthType, ImageRun, BorderStyle, ShadingType, PageNumber, Footer, Header, TabStopType, PageBreak } = require("docx");

const FIG = require(path.join(__dirname, "figures.json"));
const money2 = (n) => "$" + Number(n).toFixed(2);
const mill = (n, d) => "$" + (n / 1e6).toFixed(d) + " million";
const millShort = (n) => "$" + Math.round(n / 1e6) + "M";
const MONTHS_SHORT = { January: "Jan.", February: "Feb.", March: "Mar.", April: "Apr.", May: "May", June: "June", July: "July", August: "Aug.", September: "Sept.", October: "Oct.", November: "Nov.", December: "Dec." };
const shortDate = (long) => long.replace(/^([A-Z][a-z]+) /, (m, mo) => (MONTHS_SHORT[mo] || mo) + " ");
const D = {
  nav: money2(FIG.nav), price: money2(FIG.price),
  totalNavLong: mill(FIG.totalNav, 1), totalNavWords: "$" + Math.round(FIG.totalNav / 1e6) + " million", totalNavShort: millShort(FIG.totalNav),
  marketCapWords: "$" + Math.round(FIG.price * FIG.units / 1e6) + " million", marketCapShort: millShort(FIG.price * FIG.units),
  discount: Math.round((1 - FIG.price / FIG.nav) * 100) + "%", cents: Math.round((FIG.price / FIG.nav) * 100) + " cents",
  units: Number(FIG.units).toLocaleString("en-US"), priceDateShort: shortDate(FIG.priceDate), navAsOfShort: shortDate(FIG.navAsOf),
  al: FIG.alQualifier + " " + FIG.alPct + "% leased", viv: FIG.vivQualifier + " " + FIG.vivPct + "% leased",
};
const F = "Inter";
const NAVY = "1B2B3A", FOREST = "2E5E4E", N70 = "5F6A74", N50 = "8D949B", LINE = "D9DCDF", CREAM = "F5F3EF";
const logo = fs.readFileSync(path.join(__dirname, "img", "logo.png"));
const NB = { style: BorderStyle.NONE, size: 0, color: "FFFFFF" };
const noBorders = { top: NB, bottom: NB, left: NB, right: NB, insideHorizontal: NB, insideVertical: NB };

const r = (t, o = {}) => new TextRun({ text: t, font: F, size: 21, color: NAVY, ...o });
const b = (t, o = {}) => r(t, { bold: true, ...o });
const SUP = (n) => r(String(n), { size: 16, superScript: true });
const SUP1 = () => SUP(3);
const GF_NOTE = "¹ Based on current guidance and subject to change. OZ 1.0 zones were designated in 2018 and, based on current guidance, remain qualified opportunity zone property under the original program.";
const p = (x, o = {}) => new Paragraph({ spacing: { after: 160, line: 300 }, ...o, children: Array.isArray(x) ? x : [r(x)] });
const dp = (t) => new Paragraph({ spacing: { after: 120, line: 280 }, children: [r(t, { size: 19, color: N70 })] });
const small = (t, o = {}) => new Paragraph({ spacing: { after: 100, line: 260 }, ...o, children: [r(t, { size: 17, color: N50 })] });
const label = (t) => new Paragraph({ keepNext: true, spacing: { before: 120, after: 60 }, children: [r(t.toUpperCase(), { size: 16, bold: true, color: FOREST, characterSpacing: 40 })] });
const h1 = (t) => new Paragraph({ heading: HeadingLevel.HEADING_1, keepNext: true, spacing: { before: 360, after: 160 }, children: [r(t, { size: 34, bold: true })] });
const h2 = (t) => new Paragraph({ heading: HeadingLevel.HEADING_2, keepNext: true, spacing: { before: 240, after: 120 }, children: [r(t, { size: 26, bold: true })] });
const bullet = (x) => new Paragraph({ numbering: { reference: "bullets", level: 0 }, spacing: { after: 100, line: 300 }, children: Array.isArray(x) ? x : [r(x)] });
const numbered = (x, inst = 0) => new Paragraph({ numbering: { reference: "nums", level: 0, instance: inst }, spacing: { after: 100, line: 300 }, children: Array.isArray(x) ? x : [r(x)] });
const cell = (text, w, o = {}) => new TableCell({ width: { size: w, type: WidthType.DXA }, margins: { top: 80, bottom: 80, left: 120, right: 120 }, shading: o.shade ? { type: ShadingType.CLEAR, fill: o.shade, color: "auto" } : undefined, borders: { top: { style: BorderStyle.SINGLE, size: 4, color: LINE }, bottom: { style: BorderStyle.SINGLE, size: 4, color: LINE }, left: NB, right: NB }, verticalAlign: "center", children: [new Paragraph({ keepNext: true, keepLines: true, alignment: o.right ? AlignmentType.RIGHT : AlignmentType.LEFT, children: [r(text, { size: o.size || 19, bold: !!o.bold, color: o.color || NAVY })] })] });
const table = (rows, widths, opts = {}) => new Table({ width: { size: widths.reduce((a, c) => a + c, 0), type: WidthType.DXA }, columnWidths: widths, rows: rows.map((cells, i) => new TableRow({ tableHeader: i === 0, cantSplit: true, children: cells.map((c, j) => { const isNum = opts.numCols && opts.numCols.includes(j) && i > 0; const boldRow = opts.boldRows && opts.boldRows.includes(i); return i === 0 ? cell(c, widths[j], { shade: NAVY, color: "FFFFFF", bold: true, right: isNum || (opts.numCols && opts.numCols.includes(j)) }) : cell(c, widths[j], { right: isNum, bold: boldRow, size: opts.size }); }) })) });
const statTile = (value, caption, w, vs = 32, cs = 16) => new TableCell({ width: { size: w, type: WidthType.DXA }, margins: { top: vs < 30 ? 100 : 140, bottom: vs < 30 ? 100 : 140, left: 160, right: 120 }, shading: { type: ShadingType.CLEAR, fill: CREAM, color: "auto" }, borders: { top: NB, bottom: NB, left: { style: BorderStyle.SINGLE, size: 18, color: FOREST }, right: NB }, children: [new Paragraph({ spacing: { after: 40 }, children: [r(value, { size: vs, bold: true })] }), new Paragraph({ children: [r(caption, { size: cs, color: N70 })] })] });
const statRow = (tiles, gap = 160) => { const w = Math.floor((9360 - gap * (tiles.length - 1)) / tiles.length); const cells = []; tiles.forEach((t, i) => { cells.push(statTile(t[0], t[1], w)); if (i < tiles.length - 1) cells.push(new TableCell({ width: { size: gap, type: WidthType.DXA }, borders: noBorders, children: [new Paragraph("")] })); }); const widths = cells.map((c, i) => i % 2 === 0 ? w : gap); return new Table({ width: { size: 9360, type: WidthType.DXA }, columnWidths: widths, borders: noBorders, rows: [new TableRow({ children: cells })] }); };

const footer = (title) => new Footer({ children: [new Paragraph({ border: { top: { style: BorderStyle.SINGLE, size: 12, color: FOREST, space: 6 } }, spacing: { before: 0, after: 0 }, tabStops: [{ type: TabStopType.RIGHT, position: 9360 }], children: [r("Belpointe PREP, LLC  ·  NYSE American: OZ  ·  203-883-1944  ·  IR@belpointeoz.com  ·  belpointeoz.com", { size: 14, color: N50 }), r("\t", { size: 14 }), r("Page ", { size: 14, color: N50 }), new TextRun({ children: [PageNumber.CURRENT], font: F, size: 14, color: N50 })] })] });
const numbering = { config: [{ reference: "bullets", levels: [{ level: 0, format: LevelFormat.BULLET, text: "•", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 360, hanging: 240 } } } }] }, { reference: "nums", levels: [{ level: 0, format: LevelFormat.DECIMAL, text: "%1.", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 360, hanging: 300 } } } }] }] };
const styles = { default: { document: { run: { font: F, size: 21, color: NAVY } } }, paragraphStyles: [{ id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true, run: { font: F, size: 34, bold: true, color: NAVY } }, { id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true, run: { font: F, size: 26, bold: true, color: NAVY } }] };

const NAVTXT = "NAV per Class A unit is calculated quarterly by the Company’s Manager under the valuation policy described in the Company’s SEC filings. It is an estimate, not a guarantee of value, and the market price of Class A units may be above or below NAV and may never reach it.";

// ======================= WHITE PAPER =======================
function whitePaper() {
  const S = []; const push = (...x) => S.push(...x);
  push(new Paragraph({ spacing: { after: 600 }, children: [new ImageRun({ data: logo, transformation: { width: 180, height: 39 }, type: "png" })] }));
  push(label("White paper"));
  push(new Paragraph({ spacing: { after: 240 }, children: [r("Opportunity Zones in 2026 and beyond: the program, the January 2027 reset, and the case for a listed fund", { size: 44, bold: true })] }));
  push(p([r("How Qualified Opportunity Funds defer, reduce and eliminate capital gains tax, what changes when OZ 2.0 begins on January 1, 2027, and why Belpointe PREP, LLC (“Belpointe OZ”), the only Qualified Opportunity Fund listed on a national securities exchange, offers a different way in.", { size: 24, color: N70 }), SUP(1)]));
  push(p([r("NYSE American: OZ  ·  " + FIG.month, { bold: true, color: FOREST })]));
  push(new Paragraph({ spacing: { before: 2000 }, children: [] }));
  push(statRow([["180 days", "to reinvest an eligible gain in a Qualified Opportunity Fund"], ["Dec. 31, 2026", "OZ 1.0 recognition date for deferred gains"], ["Jan. 1, 2027", "OZ 2.0 investments begin, with a rolling five-year deferral"], ["10 years", "hold for the elimination benefit"]]));
  push(small("Numbered notes and important information appear on the last page.", { spacing: { before: 200 } }));
  push(new Paragraph({ children: [new PageBreak()] }));

  push(h1("Executive summary"));
  push(p("The Opportunity Zone program is the most generous capital gains incentive in the federal tax code: an investor who sells an appreciated asset and reinvests the gain in a Qualified Opportunity Fund within 180 days can defer that gain, reduce it through a basis step-up, and, after a ten-year hold, pay no federal capital gains tax on the appreciation of the fund investment itself."));
  push(p("Two dates now shape every Opportunity Zone decision. On December 31, 2026, gains deferred under the original program (“OZ 1.0”) are recognized. On January 1, 2027, the permanent program enacted in 2025 (“OZ 2.0”) opens with a rolling five-year deferral, a 30% step-up for rural funds and a new map of zones that is redrawn every ten years. From 2027 the program is no longer a closing window but a standing feature of the tax code, and the question for investors shifts from beating a deadline to choosing the fund."));
  push(p("Almost every Qualified Opportunity Fund is private: investors subscribe at net asset value, commit capital to properties that often do not exist yet, and have no market for their interests. Belpointe OZ is the exception. Its Class A units have traded on the NYSE American since October 2021, its two principal properties are built and leasing, and the units have recently traded at a material discount to the Company’s announced net asset value. This paper explains the program first, then the case for the listed fund."));

  push(h1("1. How the Opportunity Zone program works"));
  push(p("Congress created Opportunity Zones in the Tax Cuts and Jobs Act of 2017 to draw private capital into roughly 8,700 low-income census tracts nominated by governors and certified by the Treasury Department. The incentive runs through a Qualified Opportunity Fund (“QOF”): a corporation or partnership that elects QOF status and holds at least 90% of its assets in qualified opportunity zone property, tested twice a year."));
  push(h2("Who can use it"));
  push(p("Any taxpayer with an eligible capital gain: individuals, partnerships, corporations, trusts and estates. The gain can come from stock, a business, real estate, cryptocurrency, collectibles or any other capital asset sold to an unrelated party. Only the gain needs to be reinvested, not the full proceeds, and the investor keeps the rest."));
  push(h2("The 180-day window"));
  push(p("The investment must be made within 180 days of the sale. For gains allocated from a partnership, the window can begin at the partnership’s year end or the due date of its return, which for calendar-year partnerships pushes the deadline for 2026 gains well into 2027. Installment sales, Section 1231 gains and certain other situations have their own start dates, so the date should be confirmed with a tax adviser."));
  push(h2("What the fund must own"));
  push(p("Qualified opportunity zone property is stock or a partnership interest in a qualified opportunity zone business, or tangible property used in a trade or business inside a zone. The property must be acquired after 2017 from an unrelated party and either be put to its original use in the zone or be substantially improved, meaning the fund invests at least as much as it paid for the building within 30 months. In practice, that requirement pushes most funds toward ground-up development."));
  push(h2("The elections"));
  push(p("The investor makes the deferral election on IRS Form 8949 with the return for the year of the sale and reports the QOF investment each year on Form 8997. A fund that is a partnership, like Belpointe OZ, issues a Schedule K-1 to its investors."));

  push(h1("2. Defer, reduce, eliminate"));
  push(h2("Defer"));
  push(p("The reinvested capital gain is deferred, not recognized in the year of sale."));
  push(p([b("Under OZ 1.0 (investments through December 31, 2026): "), r("investors who hold through the December 31, 2026 recognition date, or who trigger an inclusion event before it such as a sale of the investment, recognize gain under Treasury Regulation § 1.1400Z2(b)-1(e)(4). Because Belpointe OZ is taxed as a partnership, the amount of gain included in gross income is equal to the lesser of (i) the product of (A) the percentage of the qualifying investment that gave rise to the inclusion event and (B) the remaining deferred gain, less any basis adjustments pursuant to section 1400Z-2(b)(2)(B)(iii) and (iv); or (ii) the gain that would be recognized on a fully taxable disposition at fair market value of the qualifying investment that gave rise to the inclusion event. The basis adjustments in (i)(B) include the 10% basis step-up for investments that reached their five-year anniversary on or before December 31, 2026.")]));
  push(p([b("Under OZ 2.0 (investments from January 1, 2027): "), r("the deferral runs five years from the date of each investment, on a rolling basis, or until an earlier inclusion event. There is no single recognition date for all investors.")]));
  push(h2("Reduce"));
  push(p("The amount of deferred gain that is ultimately recognized is reduced by a step-up in basis."));
  push(p([b("Under OZ 1.0: "), r("a 10% step-up applied only to investments that reached their five-year anniversary on or before December 31, 2026. Investments made after 2021 receive no step-up under OZ 1.0.")]));
  push(p([b("Under OZ 2.0: "), r("a 10% step-up after five years for any Qualified Opportunity Fund investment, and a 30% step-up after five years for investments in a qualified rural opportunity fund.")]));
  push(h2("Eliminate"));
  push(p("An investor who holds the Qualified Opportunity Fund investment for at least ten years may elect to step its basis up to fair market value on sale, eliminating federal capital gains tax on the appreciation of the fund investment itself."));
  push(p([b("Under OZ 1.0: "), r("the election remains available for sales through December 31, 2047.")]));
  push(p([b("Under OZ 2.0: "), r("the election is available for up to 30 years after the investment; at 30 years the basis steps up to fair market value automatically."), SUP(2)]));

  push(h1("3. OZ 1.0 vs. OZ 2.0: what changes on January 1, 2027"));
  push(p("The One Big Beautiful Bill Act, signed July 4, 2025, made the Opportunity Zone program permanent and rewrote several of its rules. The original program (“OZ 1.0”) governs investments made through December 31, 2026. The new program (“OZ 2.0”) governs investments made on or after January 1, 2027. The two overlap for existing investors, and the differences decide how much of the benefit an investor keeps."));
  push(table([["", "OZ 1.0 (investments through December 31, 2026)", "OZ 2.0 (investments from January 1, 2027)"],
    ["Program status", "Temporary; deferral ends December 31, 2026", "Permanent, with zones redesignated every ten years"],
    ["Deferral of the reinvested gain", "Until December 31, 2026, or an earlier inclusion event", "Five years from the date of each investment, on a rolling basis"],
    ["Basis step-up (reduction)", "10% at five years, only if the fifth anniversary fell on or before December 31, 2026", "10% at five years; 30% for investments in a qualified rural opportunity fund"],
    ["Ten-year exclusion (elimination)", "Election available for sales through December 31, 2047", "Election available for up to 30 years after the investment; basis steps up to fair market value automatically at 30 years"],
    ["Zone map", "Roughly 8,700 tracts designated in 2018; the current map stays in force for OZ 1.0 investments", "Governors nominate new tracts from July 1, 2026 with tighter income tests; the new map takes effect January 1, 2027 and is redrawn every ten years"],
    ["Rural incentive", "None", "Qualified rural opportunity funds get the 30% step-up and a lower 50% substantial-improvement threshold"],
    ["Reporting", "Form 8996 for the fund; Form 8997 for the investor", "Expanded fund and investor reporting, with penalties for funds that fail to report"],
    ["Belpointe OZ", "Aster & Links, 1700 Main Street and 1701-1710 Ringling Boulevard are in OZ 1.0 zones³", "VIV, 901-909 Central Avenue and the Nashville assets sit in OZ 2.0 zones; the Storrs parcels sit in a qualified rural opportunity zone"]], [2100, 3630, 3630], { size: 18 }));
  push(h2("What OZ 2.0 means for an investor in 2027 and after"));
  push(p("An investor with a gain in 2027 reinvests within 180 days, defers the gain for five years, recognizes 90% of it (70% for a rural fund) in year five, and pays no federal capital gains tax on the appreciation of the fund investment itself if the investment is held at least ten years. Because the deferral is rolling, there is no longer a single cliff date for everyone; each investment runs on its own clock. Because the program is permanent, a fund that already holds property in a designated OZ 2.0 zone can accept new capital for years rather than racing a sunset."));
  push(p("The practical question for a 2027 investor is the same one that applied under OZ 1.0: does the fund own real property in a qualifying zone, and is that property built or still a plan? Belpointe OZ already owns operating and development property in OZ 2.0 zones, including a rural zone, and its original Sarasota sites remain qualified under OZ 1.0."));
  push(h2("The bridge strategy"));
  push(p([r("Exiting an OZ 1.0 investment any time before the December 31, 2026 recognition date triggers a new 180-day reinvestment period. Nothing in the OZ 1.0 or OZ 2.0 statutes, regulations or IRS guidance as it exists today prohibits overlapping that exit with the start of the OZ 2.0 program on or after January 1, 2027. An investor who sells before year end and reinvests in January 2027 can therefore move from a program whose deferral is about to end into one with a fresh five-year deferral, a 10% reduction and the ten-year elimination. The Company strongly encourages investors to review this with a qualified tax professional."), SUP(4)]));

  push(h1("4. Why most Opportunity Zone funds are private, and what that costs the investor"));
  push(p([r("The substantial-improvement rule and the 90% asset test push funds toward buying land and building. Sponsors raise capital through private placements under Regulation D, which limits most offerings to accredited investors, and the capital is committed for a decade with no secondary market. The investor underwrites a business plan, not a building."), SUP(5)]));
  push(table([["Feature", "Private Qualified Opportunity Fund", "Belpointe OZ (NYSE American: OZ)"],
    ["Entry price", "Net asset value, set by the sponsor", "Market price, recently at a material discount to NAV"],
    ["Who can invest", "Usually accredited investors only", "Any investor with a brokerage account"],
    ["Minimum investment", "Typically $25,000 to $250,000 or more", "One Class A unit"],
    ["Paperwork", "Subscription agreement and investor questionnaire", "A brokerage order; the deferral election is made on the investor’s tax return"],
    ["Capital calls", "Common", "None"],
    ["What you are buying", "A development plan", "Completed, leasing properties plus a pipeline"],
    ["Liquidity", "Generally none until the fund sells its assets", "Units can be sold on the exchange at the prevailing price, subject to the holding periods needed for the tax benefits"],
    ["Reporting", "Sponsor reports to investors", "Annual and quarterly reports, proxy statements and current reports filed with the SEC; quarterly NAV announcements"],
    ["Governance", "Varies", "Board of directors with Audit, Compensation, Nominating and Corporate Governance, and Conflicts committees"]], [2200, 3480, 3680], { size: 18 }));
  push(p(""));
  push(p("Liquidity cuts both ways. Because Class A units trade on an exchange, the market price can sit well below or above the Company’s estimate of NAV. The current discount is the result of that mechanism."));

  push(new Paragraph({ heading: HeadingLevel.HEADING_1, keepNext: true, pageBreakBefore: true, spacing: { before: 0, after: 160 }, children: [r("5. The case for Belpointe OZ", { size: 34, bold: true })] }));
  push(p("Belpointe PREP, LLC is a Delaware limited liability company, the successor to Belpointe REIT, Inc., which was formed in 2019. It is externally managed by Belpointe PREP Manager, LLC (the “Manager”) and sponsored by Belpointe, LLC (the “Sponsor”), a Greenwich, Connecticut private investment firm whose predecessor was formed in 2001. Belpointe’s affiliates oversee more than $10 billion in assets across asset management, real estate development, construction and property operations. Class A units began trading on the NYSE American on October 18, 2021 under the symbol OZ."));
  push(h2("The contrarian entry point: a material discount to net asset value"));
  push(p([r("On " + FIG.navAnnounced + ", the Company announced that NAV as of " + FIG.navAsOf + " was " + D.nav + " per Class A unit, with total net asset value of approximately " + D.totalNavLong + ". On " + FIG.priceDate + ", Class A units closed at " + D.price + "."), SUP(6)], { keepNext: true }));
  push(statRow([[D.nav, "NAV per Class A unit, " + FIG.navAsOf], [D.price, "closing price, " + FIG.priceDate], [D.discount, "discount of " + D.priceDateShort + " closing price to " + D.navAsOfShort + " NAV"], [D.marketCapShort + " vs. " + D.totalNavShort, "market capitalization, " + D.priceDateShort + ", vs. net asset value, " + D.navAsOfShort]]));
  push(p("", { spacing: { after: 120 } }));
  push(p("At the " + FIG.priceDate + " closing price, a buyer on the exchange is paying about " + D.cents + " for each dollar of the Manager’s estimate of net asset value. Several factors can explain a persistent discount: a small float and limited trading volume; the lag between development spending and stabilized income; and interest expense on floating-rate property debt. None of these change the underlying real estate, but each may keep the price below NAV for an extended period."));
  push(h2("Completed, income-producing properties"));
  push(p([b("Aster & Links, Sarasota, Florida. "), r("OZ 1.0 zone."), SUP1(), r(" Two ten-story buildings on a 5.13-acre site in downtown Sarasota with 424 residences, including six guest suites, approximately 51,000 square feet of ground-floor retail anchored by Sprouts Farmers Market, and more than 900 parking spaces. Construction was substantially completed and leasing began in 2024. The property was " + D.al + " as of " + FIG.leasingAsOf + "."), SUP(7)]));
  push(p([b("VIV, St. Petersburg, Florida. "), r("OZ 2.0 zone. Two 11-story residential towers at 1000 1st Avenue North in the EDGE District with 269 residences and approximately 15,500 square feet of ground-floor retail. Leasing began in October 2025, and the property was " + D.viv + " as of " + FIG.leasingAsOf + "."), SUP(7)]));
  push(p([b("901-909 Central Avenue North, St. Petersburg. "), r("OZ 2.0 zone. A fully leased 5,328 square foot retail and office building.")]));
  push(h2("Construction complete at Aster & Links and VIV"));
  push(p("Construction risk on the two completed properties is behind the Company. Lease-up risk at each is substantially reduced, though not eliminated, until the property is stabilized. The development pipeline in Sarasota and Storrs carries its own construction risk. A new OZ 2.0 fund must buy land and build; here the two principal properties already exist and the investor can see what they are buying."));
  push(h2("A pipeline and land positioned for sale"));
  push(bullet([b("Downtown Sarasota (OZ 1.0 zones"), SUP1(), b("). "), r("1700 Main Street, planned as an approximately 130-residence community with ground-floor retail, and 1701-1710 Ringling Boulevard, an office renovation with a 20-year lease for approximately 42,000 square feet.")]));
  push(bullet([b("Storrs and Mansfield, Connecticut (qualified rural opportunity zone under OZ 2.0). "), r("497-501 Middle Turnpike, near the University of Connecticut, planned as an approximately 261-residence community, plus Storrs Road and 1750 Storrs Road held for future multifamily development.")]));
  push(bullet([b("Nashville, Tennessee, held for sale (OZ 2.0 zones). "), r("900 8th Avenue South, under contract for $19.6 million, and the Davidson Street properties, approximately 15.6 acres on the East Bank rezoned for high-density multifamily and commercial use.")]));

  push(h1("6. How to invest"));
  push(numbered([b("Realize a gain. "), r("Sell stock, a business, real estate, cryptocurrency or another appreciated asset. Only the capital gain needs to be reinvested.")], 1));
  push(numbered([b("Buy Class A units within 180 days. "), r("Place an order for OZ through any brokerage account at the prevailing market price. Qualified Opportunity Fund treatment applies to units purchased on the exchange with eligible gains within the window."), SUP(8)], 1));
  push(numbered([b("Elect and hold. "), r("Make the deferral election on IRS Form 8949 and report the investment on Form 8997 with your tax return. Belpointe OZ issues a Schedule K-1 to unitholders each year. Hold for at least ten years for the elimination benefit.")], 1));
  push(p("From time to time, in its sole discretion, the Company may offer Class A units directly under an effective registration statement. No such offering is currently open. Availability and terms are subject to change. Investors who want to discuss the Company can call Investor Relations at 203-883-1944 or email IR@belpointeoz.com."));

  push(new Paragraph({ children: [new PageBreak()] }));
  push(h1("Important information"));
  push(new Paragraph({ keepNext: true, spacing: { before: 160, after: 60 }, children: [r("Notes", { size: 22, bold: true })] }));
  const NOTES = [
    "Nothing in this paper is tax advice. Opportunity Zone benefits depend on each investor’s circumstances and on current law, which may change. Consult your own tax adviser.",
    "Most states conform to the federal statute; a few do not. Confirm the treatment in your state with your own tax professional.",
    "Zone designations reflect the Company’s assessment under current law and guidance. On November 5, 2026, the Treasury and IRS announced the census tracts designated as qualified opportunity zones under OZ 2.0, effective January 1, 2027. Zones designated under OZ 1.0 remain in effect through December 31, 2028, and under transitional guidance issued by the Treasury and IRS a qualified opportunity zone business operating in such a zone may continue to treat it as a qualified opportunity zone through December 31, 2047, subject to conditions, including conditions on new construction. New construction on sites in OZ 1.0 zones must be carried out under a working capital safe harbor plan adopted on or before December 31, 2026, and must meet funding and expenditure thresholds by that date. How those thresholds will be calculated is not yet clear, and new construction of property in OZ 1.0 zones after December 31, 2026 may not comply. The Company currently intends to manage its affairs so that it continues to meet the requirements for classification as a qualified opportunity fund under OZ 1.0, as well as the transitional guidance and the proposed regulations issued by the Treasury and IRS under OZ 2.0. These provisions remain untested, final OZ 2.0 regulations remain pending, and the Company’s ability to qualify and operate in conformity with future requirements is subject to uncertainty. A failure to continue to qualify as a qualified opportunity fund would constitute an inclusion event.",
    "Based on current law, Treasury regulations and IRS guidance, each of which is subject to change. Whether gain recognized on exiting an existing Qualified Opportunity Fund investment is eligible for reinvestment, and when the 180-day period begins and ends, depend on each investor’s circumstances. This is general information, not tax advice. Consult your tax adviser.",
    "Comparisons describe general features of private Qualified Opportunity Funds and are not a description of any specific fund.",
    NAVTXT,
    "Property information is from Belpointe PREP, LLC’s " + FIG.reportForm + " for " + FIG.reportPeriod + ". Leasing percentages are the Company’s estimates as of " + FIG.leasingAsOf + " and change over time.",    "Based on current law, Treasury regulations and IRS guidance, and on Belpointe OZ’s structure as a Qualified Opportunity Fund, Class A units acquired with eligible capital gains within the 180-day period, whether purchased on NYSE American or directly from the Company, should be a qualifying investment. Opportunity Zone tax benefits depend on each investor’s circumstances. Consult your tax adviser. Final OZ 2.0 regulations have not been issued, and transitional guidance does not have the force of regulations. There can be no assurance that Class A units acquired after December 31, 2026 will be eligible for the Opportunity Zone tax benefits. The law, regulations and guidance are subject to change.",
  ];
  const SUPS = ["¹","²","³","⁴","⁵","⁶","⁷","⁸"];
  NOTES.forEach((t, i) => push(dp(SUPS[i] + " " + t)));
  push(new Paragraph({ keepNext: true, spacing: { before: 160, after: 60 }, children: [r("Not an offer", { size: 22, bold: true })] }));
  push(dp("This white paper is for informational purposes only and does not constitute an offer to sell or the solicitation of an offer to buy any securities of Belpointe PREP, LLC. Class A units trade on the NYSE American under the symbol OZ. Investing in Class A units involves a high degree of risk, including the possible loss of the entire investment. Before investing, read the Company’s filings with the Securities and Exchange Commission, including the risk factors in its Annual Report on Form 10-K and Quarterly Reports on Form 10-Q, available at www.sec.gov."));
  push(new Paragraph({ keepNext: true, spacing: { before: 160, after: 60 }, children: [r("Forward-looking statements", { size: 22, bold: true })] }));
  push(dp("This paper contains forward-looking statements within the meaning of Section 27A of the Securities Act of 1933 and Section 21E of the Securities Exchange Act of 1934, including statements about leasing, stabilization, development plans, dispositions, net asset value and the Opportunity Zone program. Words such as “anticipate,” “estimate,” “expect,” “project,” “plan,” “intend,” “believe,” “may” and “will” identify forward-looking statements. They reflect current expectations and are subject to the risks and uncertainties described in the Company’s SEC filings, including delays in the lease-up and stabilization of its properties, fluctuations in occupancy and market rents, the availability and cost of financing, and changes in tax law. Actual results may differ materially, and the Company undertakes no obligation to update any forward-looking statement except as required by law."));
  push(new Paragraph({ keepNext: true, spacing: { before: 160, after: 60 }, children: [r("Net asset value and market data", { size: 22, bold: true })] }));
  push(dp("NAV per Class A unit is calculated quarterly by the Company’s Manager under the valuation policy described in the Company’s SEC filings. NAV is an estimate, not a guarantee of value. The market price of Class A units may be above or below NAV and may never reach it. Market price data is delayed and provided for information only. Figures in this document are updated automatically from public sources, including the Company’s filings with the SEC and third-party market data, and may contain errors, omissions or delays. They are provided for informational purposes only. The Company’s filings with the SEC, available at www.sec.gov, are the official source and control if there is any difference. Market capitalization is the " + D.price + " closing price on " + FIG.priceDate + " multiplied by " + D.units + " Class A units outstanding at " + FIG.unitsAsOf + ", as reported in the " + FIG.unitsSource + "."));
  push(new Paragraph({ keepNext: true, spacing: { before: 160, after: 60 }, children: [r("Tax matters", { size: 22, bold: true })] }));
  push(dp("Nothing in this paper is tax advice. Opportunity Zone benefits depend on each investor’s circumstances, holding periods and on current law and IRS guidance, which may change. Most states conform to the federal statute; a few do not. Investors should consult their own tax advisers."));
  push(new Paragraph({ keepNext: true, spacing: { before: 160, after: 60 }, children: [r("Sources", { size: 22, bold: true })] }));
  push(dp("Figures are from the Company’s " + FIG.reportForm + " for " + FIG.reportPeriod + " (accession " + FIG.reportAccession + ") and " + FIG.navSource + " unless otherwise stated. Leasing percentages are the Company’s estimates as of " + FIG.leasingAsOf + " and change over time. Belpointe OZ is managed by Belpointe PREP Manager, LLC and sponsored by Belpointe, LLC."));

  return new Document({ creator: "Cody Laidlaw", lastModifiedBy: "Cody Laidlaw", title: "Opportunity Zones in 2026 and Beyond", description: "Belpointe OZ white paper", numbering, styles, sections: [{ properties: { page: { margin: { top: 1440, bottom: 1440, left: 1440, right: 1440 } } }, footers: { default: footer("White Paper") }, children: S }] });
}

// ======================= FACT SHEET =======================
function factSheet() {
  const S = []; const push = (...x) => S.push(...x);
  const W = 9720;
  const hdr = new Table({ width: { size: W, type: WidthType.DXA }, columnWidths: [5220, 4500], borders: noBorders, rows: [new TableRow({ children: [
    new TableCell({ width: { size: 5220, type: WidthType.DXA }, borders: noBorders, children: [new Paragraph({ children: [new ImageRun({ data: logo, transformation: { width: 160, height: 35 }, type: "png" })] })] }),
    new TableCell({ width: { size: 4500, type: WidthType.DXA }, borders: noBorders, verticalAlign: "center", children: [new Paragraph({ alignment: AlignmentType.RIGHT, children: [r("NYSE American: OZ", { size: 22, bold: true, color: FOREST })] }), new Paragraph({ alignment: AlignmentType.RIGHT, children: [r("Fact sheet  ·  " + FIG.month, { size: 16, color: N50 })] })] }),
  ] })] });
  push(hdr);
  push(new Paragraph({ spacing: { before: 200, after: 90, line: 290 }, children: [r("The only Qualified Opportunity Fund listed on a national securities exchange", { size: 28, bold: true })] }));
  push(new Paragraph({ spacing: { after: 180, line: 272 }, children: [r("Belpointe PREP, LLC (“Belpointe OZ”) owns completed, income-producing multifamily and mixed-use properties in Florida Opportunity Zones, a development pipeline in Sarasota and Connecticut, and land held for sale in Nashville. Class A units trade on the NYSE American and have recently traded at a material discount to net asset value.", { size: 18, color: N70 })] }));

  // stats, full width
  const tiles = [[D.price, "Class A unit price, " + D.priceDateShort + " (delayed)"], [D.nav, "NAV per Class A unit, " + FIG.navAsOf + "¹"], [D.marketCapShort + " vs. " + D.totalNavShort, "market cap (" + D.priceDateShort + ") vs. NAV (" + D.navAsOfShort + ")"], [String(FIG.completedResidences), "completed residences"]];
  const gap = 140, tw = Math.floor((W - gap * 3) / 4);
  const cells = []; tiles.forEach((t, i) => { cells.push(statTile(t[0], t[1], tw, 22, 15)); if (i < 3) cells.push(new TableCell({ width: { size: gap, type: WidthType.DXA }, borders: noBorders, children: [new Paragraph("")] })); });
  push(new Table({ width: { size: W, type: WidthType.DXA }, columnWidths: cells.map((c, i) => i % 2 === 0 ? tw : gap), borders: noBorders, rows: [new TableRow({ children: cells })] }));
  push(new Paragraph({ spacing: { after: 60 }, children: [] }));

  const fp = (x, after = 90) => new Paragraph({ spacing: { after, line: 268 }, children: Array.isArray(x) ? x : [r(x, { size: 16 })] });
  const fb = (t) => r(t, { size: 16, bold: true });
  const fr = (t) => r(t, { size: 16 });
  const fbul = (x) => new Paragraph({ numbering: { reference: "bullets", level: 0 }, spacing: { after: 56, line: 252 }, children: Array.isArray(x) ? x : [fr(x)] });
  const fl = (t) => new Paragraph({ keepNext: true, spacing: { before: 150, after: 80 }, border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: FOREST, space: 5 } }, children: [r(t, { size: 22, bold: true })] });
  const col = (kids, w) => new TableCell({ width: { size: w, type: WidthType.DXA }, borders: noBorders, margins: { top: 0, bottom: 0, left: 0, right: 0 }, verticalAlign: "top", children: kids });
  const two = (leftKids, rightKids) => new Table({ width: { size: W, type: WidthType.DXA }, columnWidths: [4680, 360, 4680], borders: noBorders, rows: [new TableRow({ children: [col(leftKids, 4680), col([new Paragraph("")], 360), col(rightKids, 4680)] })] });

  push(two([
    fl("The contrarian entry point"),
    fp([fr("Class A units have recently traded at a material discount to the Company’s most recently announced net asset value."), r("1", { size: 16, superScript: true }), fr(" At the " + D.priceDateShort + " close of " + D.price + " against " + D.navAsOfShort + " NAV of " + D.nav + ", a buyer on the exchange acquires the underlying real estate at roughly " + D.cents + " on the dollar of the Manager’s estimate of value. Company-wide, the market values all Class A units at about " + D.marketCapWords + " against net asset value of " + D.totalNavWords + ".")], 100),
    fp("The entry price is set by the market.", 40),
  ], [
    fl("Why Belpointe OZ"),
    fbul([fb("Completed, income-producing properties. "), fr("Aster & Links and VIV are built and leasing: " + FIG.alQualifier + " " + FIG.alPct + "% and " + FIG.vivPct + "% leased as of " + FIG.leasingAsOf + ".")]),
    fbul([fb("Construction complete at Aster & Links and VIV. "), fr("Lease-up risk remains until each is stabilized. The development pipeline carries its own construction risk.")]),
    fbul([fb("Listed on NYSE American. "), fr("Any brokerage account, no accreditation, no subscription documents, no capital calls, no lock-up. Liquidity depends on market conditions.")]),
    fbul([fb("Public company transparency. "), fr("SEC reports, proxy materials and quarterly NAV announcements.")]),
  ]));
  push(fl("Portfolio"));
  const three = (a, bb, c) => { const w = 3080, g = 240; return new Table({ width: { size: W, type: WidthType.DXA }, columnWidths: [w, g, w, g, w], borders: noBorders, rows: [new TableRow({ children: [col(a, w), col([new Paragraph("")], g), col(bb, w), col([new Paragraph("")], g), col(c, w)] })] }); };
  push(three([
    fbul([fb("Aster & Links, Sarasota, FL. "), fr("OZ 1.0 zone. 424 residences,"), r("2", { size: 16, superScript: true }), fr(" ~51,000 sf retail anchored by Sprouts, 900+ parking spaces. Completed.")]),
    fbul([fb("VIV, St. Petersburg, FL. "), fr("OZ 2.0 zone. 269 residences in two 11-story towers, ~15,500 sf retail, EDGE District. Completed.")]),
  ], [
    fbul([fb("901-909 Central Avenue North, St. Petersburg. "), fr("OZ 2.0 zone. Fully leased retail/office building.")]),
    fbul([fb("Sarasota pipeline. "), fr("OZ 1.0 zone. 1700 Main Street (~130 residences, predevelopment) and 1701-1710 Ringling Boulevard (office renovation, 20-year lease signed).")]),
  ], [
    fbul([fb("Storrs, CT pipeline. "), fr("Rural OZ 2.0 zone. 497-501 Middle Turnpike (~261 residences, predevelopment), 1750 Storrs Road and Storrs Road, near UConn.")]),
    fbul([fb("Nashville, held for sale. "), fr("OZ 2.0 zones. 900 8th Avenue South under contract to be sold; Davidson Street, ~15.6 acres rezoned for high-density use.")]),
  ]));
  push(fl("Tax benefits: defer, reduce, eliminate"));
  push(two([
    fbul([fb("Defer "), fr("the capital gain by reinvesting in a Qualified Opportunity Fund within 180 days. Under OZ 1.0 the deferred gain is recognized on the earlier of December 31, 2026 or an inclusion event; under OZ 2.0, for investments from January 1, 2027, the deferral runs five years.")]),
    fbul([fb("Reduce "), fr("the gain recognized with a 10% basis step-up after five years (30% for qualified rural opportunity funds under OZ 2.0).")]),
  ], [
    fbul([fb("Eliminate "), fr("federal capital gains tax on the appreciation of the Qualified Opportunity Fund investment after a ten-year hold.")]),
    fbul([fb("Bridge. "), fr("Exiting an OZ 1.0 investment before December 31, 2026 starts a new 180-day window that can overlap the start of OZ 2.0."), r("3", { size: 16, superScript: true }), fr(" Review with a qualified tax professional.")]),
  ]));
  push(fl("How to invest"));
  push(fp([fb("1. "), fr("Realize a capital gain.   "), fb("2. "), fr("Buy OZ through any brokerage account within 180 days."), r("4", { size: 16, superScript: true }), fr("   "), fb("3. "), fr("Make the deferral election on IRS Form 8949 and Form 8997 with your return, then hold.")], 60));
  push(new Paragraph({ spacing: { before: 40 }, children: [r("¹ ² ³ ⁴ See Important information on page 2. This fact sheet is not an offer to sell or a solicitation of an offer to buy any securities. Investing in Class A units involves a high degree of risk, including loss of the entire investment.", { size: 13, color: N50 })] }));

  // ---- Page 2: Important information ----
  const dh = (t) => new Paragraph({ keepNext: true, spacing: { before: 120, after: 50 }, children: [r(t, { size: 20, bold: true })] });
  const dq = (t) => new Paragraph({ spacing: { after: 40, line: 236 }, children: [r(t, { size: 14, color: N70 })] });
  push(new Paragraph({ pageBreakBefore: true, spacing: { before: 0, after: 100 }, children: [r("Important information", { size: 30, bold: true })] }));
  push(dh("Not an offer"));
  push(dq("This fact sheet is for informational purposes only and does not constitute an offer to sell or the solicitation of an offer to buy any securities of Belpointe PREP, LLC. Class A units trade on the NYSE American under the symbol OZ. Investing in Class A units involves a high degree of risk, including the possible loss of the entire investment. Before investing, read the Company’s filings with the Securities and Exchange Commission, including the risk factors in its Annual Report on Form 10-K and Quarterly Reports on Form 10-Q, available at www.sec.gov. From time to time, in its sole discretion, the Company may offer Class A units directly under an effective registration statement. No such offering is currently open. Availability and terms are subject to change."));
  push(dh("Forward-looking statements"));
  push(dq("This fact sheet contains forward-looking statements within the meaning of Section 27A of the Securities Act of 1933 and Section 21E of the Securities Exchange Act of 1934, including statements about leasing, stabilization, development plans, dispositions, net asset value and the Opportunity Zone program. Words such as “anticipate,” “estimate,” “expect,” “project,” “plan,” “intend,” “believe,” “may” and “will” identify forward-looking statements. They reflect current expectations and are subject to the risks and uncertainties described in the Company’s SEC filings, including delays in the lease-up and stabilization of its properties, fluctuations in occupancy and market rents, the availability and cost of financing, and changes in tax law. Actual results may differ materially, and the Company undertakes no obligation to update any forward-looking statement except as required by law."));
  push(dh("Properties and Opportunity Zone designations"));
  push(dq("Property information is from the Company’s " + FIG.reportForm + " for " + FIG.reportPeriod + " (accession " + FIG.reportAccession + "). Leasing percentages are the Company’s estimates as of " + FIG.leasingAsOf + " and change over time; a property is not stabilized until it reaches and sustains its target occupancy. Completed properties remain subject to lease-up, operating and refinancing risk. Development and predevelopment properties, and properties held for sale, are subject to entitlement, construction, financing, market and disposition risk, and plans for them may change. Zone designations reflect the Company’s assessment under current law and guidance. On November 5, 2026, the Treasury and IRS announced the census tracts designated as qualified opportunity zones under OZ 2.0, effective January 1, 2027. Zones designated under OZ 1.0 remain in effect through December 31, 2028, and under transitional guidance issued by the Treasury and IRS a qualified opportunity zone business operating in such a zone may continue to treat it as a qualified opportunity zone through December 31, 2047, subject to conditions, including conditions on new construction."));
  push(dh("Tax matters"));
  push(dq("Nothing in this fact sheet is tax advice. Opportunity Zone benefits depend on each investor’s circumstances, holding periods and on current law and IRS guidance, which may change. Most states conform to the federal statute; a few do not. The deferral election is made by the investor on IRS Form 8949 and Form 8997. Investors should consult their own tax advisers before investing. The Company currently intends to manage its affairs so that it continues to meet the requirements for classification as a qualified opportunity fund under OZ 1.0, as well as the transitional guidance and the proposed regulations issued by the Treasury and IRS under OZ 2.0. These provisions remain untested, final OZ 2.0 regulations remain pending, and the Company’s ability to qualify and operate in conformity with future requirements is subject to uncertainty. A failure to continue to qualify as a qualified opportunity fund would constitute an inclusion event."));
  push(dh("¹ Net asset value and market data"));
  push(dq("NAV per Class A unit is calculated quarterly by the Company’s Manager under the valuation policy described in the Company’s SEC filings. NAV is an estimate, not a guarantee of value. The market price of Class A units may be above or below NAV and may never reach it. Market price data is delayed and provided for information only. Figures in this document are updated automatically from public sources, including the Company’s filings with the SEC and third-party market data, and may contain errors, omissions or delays. They are provided for informational purposes only. The Company’s filings with the SEC, available at www.sec.gov, are the official source and control if there is any difference. Market capitalization is the " + D.price + " closing price on " + FIG.priceDate + " multiplied by " + D.units + " Class A units outstanding at " + FIG.unitsAsOf + ", as reported in the " + FIG.unitsSource + ". Any discount to NAV reflects market conditions and is not a prediction of future prices."));
  push(dh("² Residences"));
  push(dq("The 424 residences at Aster & Links include six guest suites."));
  push(dh("³ Bridge from OZ 1.0 to OZ 2.0"));
  push(dq("Based on current law, Treasury regulations and IRS guidance, each of which is subject to change. Whether gain recognized on exiting an existing Qualified Opportunity Fund investment is eligible for reinvestment, and when the 180-day period begins and ends, depend on each investor’s circumstances. This is general information, not tax advice. Consult your tax adviser."));
  push(dh("⁴ Qualifying investment"));
  push(dq("Based on current law, Treasury regulations and IRS guidance, and on Belpointe OZ’s structure as a Qualified Opportunity Fund, Class A units acquired with eligible capital gains within the 180-day period, whether purchased on NYSE American or directly from the Company, should be a qualifying investment. Opportunity Zone tax benefits depend on each investor’s circumstances. Consult your tax adviser. Final OZ 2.0 regulations have not been issued, and transitional guidance does not have the force of regulations. There can be no assurance that Class A units acquired after December 31, 2026 will be eligible for the Opportunity Zone tax benefits. The law, regulations and guidance are subject to change."));
  push(dh("Sources and management"));
  push(dq("Figures are from the Company’s " + FIG.reportForm + " for " + FIG.reportPeriod + " and " + FIG.navSource + " unless otherwise stated. Belpointe OZ is managed by Belpointe PREP Manager, LLC and sponsored by Belpointe, LLC. Belpointe PREP, LLC, 255 Glenville Road, Greenwich, CT 06831. Investor Relations: 203-883-1944, IR@belpointeoz.com, belpointeoz.com."));
  return new Document({ creator: "Cody Laidlaw", lastModifiedBy: "Cody Laidlaw", title: "Belpointe OZ Fact Sheet", numbering, styles, sections: [{ properties: { page: { margin: { top: 560, bottom: 600, left: 1000, right: 1000, footer: 300 } } }, footers: { default: new Footer({ children: [new Paragraph({ border: { top: { style: BorderStyle.SINGLE, size: 12, color: FOREST, space: 6 } }, tabStops: [{ type: TabStopType.RIGHT, position: W }], children: [r("Investor Relations  ·  203-883-1944  ·  IR@belpointeoz.com  ·  255 Glenville Road, Greenwich, CT 06831  ·  belpointeoz.com  ·  NYSE American: OZ", { size: 14, color: N50 }), r("\t", { size: 14 }), r("Page ", { size: 14, color: N50 }), new TextRun({ children: [PageNumber.CURRENT], font: F, size: 14, color: N50 })] })] }) }, children: S }] });
}

(async () => {
  fs.writeFileSync(path.join(__dirname, "Belpointe OZ White Paper.docx"), await Packer.toBuffer(whitePaper()));
  fs.writeFileSync(path.join(__dirname, "Belpointe OZ Fact Sheet.docx"), await Packer.toBuffer(factSheet()));
  console.log("ok");
})();
