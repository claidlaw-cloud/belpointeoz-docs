// Merges the relay's facts response into figures.json.
// Usage: node update-figures.js facts.json   (facts.json = JSON from <relay>?type=facts, which includes facts, nav and quote)
// Exits 0 and prints "changed" when any figure changed, "unchanged" otherwise. Exits 2 if the input is unusable.
const fs = require("fs");
const path = require("path");
const figPath = path.join(__dirname, "figures.json");
const fig = JSON.parse(fs.readFileSync(figPath, "utf8"));
const input = JSON.parse(fs.readFileSync(process.argv[2], "utf8"));
const before = JSON.stringify(fig);

const longDate = (d) => new Date(d + "T12:00:00Z").toLocaleDateString("en-US", { year: "numeric", month: "long", day: "numeric", timeZone: "America/New_York" });
const cap = (s) => s.charAt(0).toUpperCase() + s.slice(1);

const nav = input.nav, facts = input.facts, quote = input.quote;
if (!nav || !nav.nav || !nav.asOf) { console.error("No NAV in input"); process.exit(2); }

fig.nav = nav.nav;
fig.navAsOf = nav.asOf;
fig.navAnnounced = nav.announced || longDate(nav.filingDate);
fig.navSource = (nav.form && nav.form.indexOf("424B") === 0 ? "prospectus supplement" : "Form " + (nav.form || "8-K")) + " dated " + fig.navAnnounced;
if (nav.totalNav) fig.totalNav = nav.totalNav;
if (nav.unitsOutstanding) fig.unitsOutstanding = nav.unitsOutstanding, fig.units = nav.unitsOutstanding;


if (facts) {
  if (facts.report && facts.report.form && facts.report.periodLabel) {
    fig.reportForm = "Form " + facts.report.form;
    fig.reportPeriod = facts.report.periodLabel;
    if (facts.report.url) { const m = facts.report.url.match(/\/(\d{10})(\d{2})(\d{6})\//); if (m) fig.reportAccession = m[1] + "-" + m[2] + "-" + m[3]; }
  }
  const qual = (q) => (q === "over" || q === "more than" ? "more than" : q);
  if (facts.al && facts.al.pct) { fig.alPct = facts.al.pct; fig.alQualifier = qual(facts.al.qualifier || "approximately"); fig.leasingAsOf = facts.al.asOf || fig.leasingAsOf; }
  if (facts.viv && facts.viv.pct) { fig.vivPct = facts.viv.pct; fig.vivQualifier = qual(facts.viv.qualifier || "approximately"); if (facts.viv.asOf && (!facts.al || facts.viv.asOf >= (facts.al.asOf || ""))) fig.leasingAsOf = facts.viv.asOf; }
}

// The unit price is refreshed only when something else changed (a new NAV, report or leasing figure), so the documents
// are not regenerated every week for a price move alone. Set FORCE_PRICE=1 to refresh the price regardless.
const substantive = JSON.stringify(fig) !== before;
if ((substantive || process.env.FORCE_PRICE) && quote && quote.price) {
  fig.price = Math.round(quote.price * 100) / 100;
  fig.priceDate = longDate(new Date().toISOString().slice(0, 10));
}
if (substantive || process.env.FORCE_PRICE) fig.month = new Date().toLocaleDateString("en-US", { year: "numeric", month: "long", timeZone: "America/New_York" });

fs.writeFileSync(figPath, JSON.stringify(fig, null, 2) + "\n");
console.log(JSON.stringify(fig) === before ? "unchanged" : "changed");
