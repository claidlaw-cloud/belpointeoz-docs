// Merges the relay's facts response into figures.json.
// Usage: node update-figures.js facts.json   (facts.json = JSON from <relay>?type=facts, which includes facts, nav and quote,
//        plus "closes": the last ten or more daily closes [{d:"YYYY-MM-DD",c:number}] from <relay>?type=history)
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
// Units outstanding come from the cover page of the latest 10-Q or 10-K (kept in figures.json with their date and source).
// A NAV announcement replaces them only when its as-of date is later than the unit count already on file.
if (nav.unitsOutstanding && (!fig.unitsAsOf || new Date(nav.asOf) > new Date(fig.unitsAsOf))) { fig.units = nav.unitsOutstanding; fig.unitsAsOf = nav.asOf; fig.unitsSource = fig.navSource; }


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

// The unit price is the last completed weekly close: the latest daily close dated on or before the most recent Friday
// (New York time). It therefore changes once a week, and always carries the date of the close it came from.
const closes = Array.isArray(input.closes) ? input.closes.filter((p) => p && /^\d{4}-\d{2}-\d{2}$/.test(p.d) && p.c > 0) : [];
if (closes.length) {
  const todayNY = new Date().toLocaleDateString("en-CA", { timeZone: "America/New_York" });
  const t = new Date(todayNY + "T12:00:00Z");
  const back = (t.getUTCDay() + 2) % 7; // days since the most recent Friday (0 on a Friday)
  t.setUTCDate(t.getUTCDate() - back);
  const friday = t.toISOString().slice(0, 10);
  const pick = closes.filter((p) => p.d <= friday).sort((a, b) => (a.d < b.d ? -1 : 1)).pop();
  if (pick) { fig.price = Math.round(pick.c * 100) / 100; fig.priceDate = longDate(pick.d); }
} else if (process.env.FORCE_PRICE && quote && quote.price) {
  fig.price = Math.round(quote.price * 100) / 100;
  fig.priceDate = longDate(new Date().toLocaleDateString("en-CA", { timeZone: "America/New_York" }));
}
if (JSON.stringify(fig) !== before) fig.month = new Date().toLocaleDateString("en-US", { year: "numeric", month: "long", timeZone: "America/New_York" });

fs.writeFileSync(figPath, JSON.stringify(fig, null, 2) + "\n");
console.log(JSON.stringify(fig) === before ? "unchanged" : "changed");
