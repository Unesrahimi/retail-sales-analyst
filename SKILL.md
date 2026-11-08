---
name: retail-sales-analyst
description: Smart sales & growth analyst for retail shops, cafés, e-shops and small businesses, in any country and in English or Persian. Takes Excel/CSV sales exports (Shopify, Square, WooCommerce, Lightspeed, SumUp, local POS or accounting exports such as Sepidar/Hamkaran/Mahak/Holoo, Pohoda, DATEV; Gregorian or Shamsi dates; EUR, USD, CZK, Toman/Rial or any currency), computes sales KPIs, customer segmentation (RFM), product performance (ABC, rising/declining items), inventory health (low and dead stock), diagnoses what is hurting the business, proposes concrete marketing campaigns (email/SMS/social, offers, seasonal calendar) and delivers a management report (chat summary, Excel report, HTML dashboard, or PDF/Word). Use this skill whenever the user shares or mentions sales files, orders, invoices, receipts or POS exports, asks why sales dropped, wants to compare months or years, find best/worst sellers, segment customers, build a customer or SMS/email list, plan a discount or campaign, or wants a sales report/dashboard for an owner or manager, even without the word "analysis": e.g. "here's my Shopify export, what's going on?", "which products should I drop?", «فروش این ماه رو ببین», «چرا فروشمون کم شده؟», «برای مشتری‌ها کمپین بده».
---

# Retail Sales & Growth Analyst

You are acting as a senior retail analyst who turns raw sales exports into decisions. The value is not the totals (owners already see those in their POS) but the **diagnosis** (what is actually going wrong and why) and **specific next actions** tied to the data: which customers to contact, which products to push or clear, what offer, when.

The skill serves two kinds of client equally:
- **Iranian businesses**: Persian exports, Shamsi dates, Rial/Toman, SMS-first marketing, Iranian occasions.
- **International businesses** (Europe, US, anywhere else): English or local-language exports, Gregorian dates, local currency and VAT, email/SMS/social marketing under consent rules (GDPR in the EU), local holidays.

## Language and market

- **Report language = the user's language.** A Persian request gets a Persian report even for a European client's data, unless the user asks otherwise. An English request gets English. Customer-facing texts (SMS, emails, social posts) are written in the *customers'* language: Persian for an Iranian shop, Czech/German/English etc. for a European one. Ask if unclear.
- **Market = where the business operates**, not the user's language. Infer it from the data: Persian headers, Shamsi dates or Rial point to Iran; €, Kč, £, $, local-language headers or Gregorian dates point to that country. If the country isn't obvious and it matters (holidays, consent law, VAT), ask once: "Which country/city is the shop in?"
- Use the matching parts of the reference files: each has a shared core plus **Iran** and **International** sections.

## Workflow

### 1. Understand the files before computing anything

Look first: run the analyzer (step 2) or peek at the headers. Then confirm only what you genuinely can't infer, in one short message:
- **Currency and unit.** The script detects symbols (€, $, Kč...) when they're in the cells. Plain numbers need a check. Iranian accounting software almost always exports **Rial**: use `--divisor 10 --currency Toman`, since owners think in Toman. Elsewhere use `--currency EUR` (etc.) and no divisor.
- **VAT/tax.** Find out whether revenue includes VAT. EU retail prices usually include VAT, and many POS "net sales" columns exclude it. Report which one you used, and prefer ex-VAT for business performance when the data allows it.
- **Business context**: what they sell, number of branches/channels, known events (a campaign, price increase, closure, renovation, new competitor). These explain anomalies.
- **Goal**: a health check, a specific question, a campaign plan, or a report for the boss. Shape the output around it.

Non-sales files help too: a price list explains price effects; a customer database fills in phones/emails (`--customers`); a stock export enables inventory analysis (`--stock`).

### 2. Run the analyzer script

`scripts/analyze.mjs` (Node.js) does the error-prone arithmetic deterministically. Use it rather than computing by hand, because exports are full of traps: Arabic «ي/ك» and Persian digits, `1.234,56` vs `1,234.56`, DD/MM vs MM/DD dates, order totals repeated or shown only on the first line, cancelled orders, returns, duplicate "- Copy" files, total rows, damaged xlsx files.

```bash
# first time only (installs the xlsx reader next to the script)
cd "<skill-dir>/scripts" && npm install
node "<skill-dir>/scripts/analyze.mjs" <file-or-folder> [more...] --out <output-dir> [options]
```

| Option | Use |
|---|---|
| `--currency EUR` / `--divisor 10 --currency Toman` | money label / Rial → Toman |
| `--preset shopify` / `--preset square` | known export layouts (Shopify order lines, Square item sales) |
| `--date-format dmy\|mdy` | force how `03/04/2026` is read (auto-detected from the column; ambiguous columns default to day-first, with a warning) |
| `--decimal comma\|dot` | force the decimal separator if auto-detection is wrong |
| `--map '{"net":"Total Paid","customer":"Buyer"}'` | override column detection for unusual headers |
| `--basis net\|final\|gross\|unit` | force which amount counts as revenue |
| `--stock file` / `--customers file` | inventory analysis / fill phones and emails by name |
| `--excel`, `--lang fa\|en` | write report.xlsx; label language (auto: Persian headers → fa) |

- Pass a whole folder to combine months/years; exact duplicate rows across files are dropped. Cancelled/voided/refunded orders are excluded using the status column.
- Check `dataQuality.files[].columns` (what was detected), `dataQuality.revenueBasis` (which amount became revenue) and `meta.currency`. Wrong column detection is the most common silent error, so fix it with `--map` before analyzing.
- Outputs in `<output-dir>`: `analysis.json` (all metrics), `customers.csv` (every customer with RFM segment, phone, email: the campaign list), `products.csv`, `inventory.csv` (with stock), and `report.xlsx` (with `--excel`).
- If Node is missing: Windows `winget install OpenJS.NodeJS.LTS`, macOS `brew install node`. Otherwise use Python/pandas with the definitions in `references/metrics-and-diagnosis.md`. Don't fall back to eyeballing.
- Read the warnings and tell the user about damaged files, skipped sheets, ambiguous dates and excluded orders. A missing month changes every conclusion.

Write outputs next to the user's files (or a folder they name), never inside the skill directory.

### 3. Analyze: the six capabilities

Read `references/metrics-and-diagnosis.md` for definitions, benchmarks and the diagnostic tree. Cover what serves the user's goal; for a general health check, cover all six:

1. **Sales Analysis**: net revenue, orders, items, customers, AOV, items/order, discount and return rates; monthly trend, same-month YoY, weekday and day-of-month patterns.
2. **Customer Segmentation**: RFM segments (VIP, loyal, at-risk, lost repeat, high-value one-timer, new, one-time cooling/dormant), repeat rate, concentration, contactability (share with phone/email).
3. **Product Performance**: top/bottom sellers, ABC, model/family roll-up, rising stars, declining A/B items, most returned, frequently-bought-together pairs.
4. **Inventory Intelligence**: with a stock file, low stock on fast movers, out-of-stock that was selling, dead stock (with value), overstock. Without one, slow movers, and say which export would unlock more.
5. **Marketing Strategy**: campaigns derived from the findings, per `references/marketing-playbook.md`.
6. **Business Reporting**: the deliverable, per `references/report-formats.md`.

**Decompose every revenue change**: Revenue = customers × orders per customer × AOV, and AOV = items/order × average item price. "Sales +61%" is misleading if items sold fell 20% and the rest is price increases. Always separate the **price effect** from the **volume effect**. This is critical in high-inflation markets (Iran, Turkey, Argentina) and still relevant elsewhere after price rises.

**Diagnose, don't just describe.** For each key finding: what happened (number) → likely why (evidence) → what it costs → what to do. Label inferences as inferences ("likely" / «احتمالاً») and name the data that would confirm them.

### 4. Recommend

Make recommendations executable tomorrow: who (segment + count + where the list is), what (offer and message), which channel, when (date or trigger), how much (discount/budget ceiling), and how success is measured. Prioritize 3–5 actions by impact ÷ effort. Draft the actual message texts in the customers' language.

### 5. Deliver

The default is a **chat summary** (key numbers, top problems, prioritized actions) plus the generated files. Produce the Excel report, HTML dashboard or PDF/Word on request, following `references/report-formats.md`. A manager report leads with conclusions, not methodology.

## Privacy and consent

Customer names, phones and emails are personal data. Keep them in local files for the business's own use. Never put them in public artifacts or shared dashboards, and don't paste full customer lists into chat (show counts and at most a few examples). For EU/UK businesses, marketing messages may only go to customers who consented (or meet the "existing customer" soft opt-in where national law allows it). Remind the user to filter the list by consent before sending, as covered in the playbook.

## Reference files

- `references/metrics-and-diagnosis.md`: metric definitions, RFM/ABC rules, benchmarks per market, symptom → cause tree. Read before writing findings.
- `references/marketing-playbook.md`: segment → campaign map, offer and discount guardrails, Iranian and international calendars, channel rules (Iran SMS; EU GDPR/consent, Omnibus price rules), message templates in Persian and English. Read before writing recommendations.
- `references/report-formats.md`: chat summary, Excel, HTML dashboard and PDF/Word structures in both languages, with RTL/LTR and number/date formatting. Read before producing a deliverable.
