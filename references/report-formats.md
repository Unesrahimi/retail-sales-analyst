# Report Formats

Whatever the format or language, a management report answers in this order: **What's the situation? What's wrong? What do we do?** Methodology and data caveats go last and stay short.

## Contents
1. Locale conventions
2. Chat summary (default)
3. Excel report
4. HTML dashboard
5. PDF / Word report

## 1. Locale conventions

| | Iranian client | International client |
|---|---|---|
| Direction | RTL (`dir="rtl"`, RTL sheets) | LTR |
| Dates | Shamsi, Persian month names (فروردین ۱۴۰۵) | Gregorian. ISO (2026-03-21) in tables; local style in text (21 March 2026 / 21. 3. 2026 / March 21, 2026) |
| Money | Toman, thousand separators; «میلیون/میلیارد تومان» for big totals | Local currency and format via `Intl.NumberFormat` (e.g. 1 234,50 Kč · 1.234,50 € · €1,234.50 · $1,234.50); state whether amounts include VAT |
| Week | starts Saturday | starts Monday (Sunday in the US) |
| Fonts | Vazirmatn (fallback Tahoma) | Inter / system UI |

The report language follows the user; customer-facing text follows the customers' language. A Persian report about a Czech shop uses Gregorian dates and CZK, with a Persian narrative.

## 2. Chat summary (default)

English skeleton (mirror it in Persian: «خلاصه وضعیت / مهم‌ترین مشکلات / فرصت‌ها / اقدامات پیشنهادی / فایل‌ها / نکات داده»):

```
## Snapshot ([period])
- 3–5 key numbers, each with its comparison (vs previous month / same month last year)

## Biggest problems
1. [problem] — [number] — [likely cause] — [cost/impact]
(3–5, most important first)

## Opportunities
- …

## Recommended actions (by priority)
| Action | Target (count) | When | Measurable goal |

## Files
- paths to report.xlsx / customers.csv / …
## Data notes
- damaged/missing files, currency and VAT assumption, ambiguous dates, limitations
```

Keep it scannable, ideally one screen, and offer the deeper formats.

## 3. Excel report

`--excel` produces a plain `report.xlsx` (Summary, Monthly, YoY, Weekday, Products, Segments, Customers, Inventory, Data quality), with sheet names and direction following `--lang`. For a polished manager workbook, build one with the xlsx skill if available (or openpyxl / exceljs) on top of `analysis.json`:
- First sheet "Dashboard" / «داشبورد»: KPI tiles (revenue, orders, customers, AOV, repeat rate) with the change vs the previous period, and the top 3 problems and top 3 actions as text.
- Then: the monthly trend with a chart (revenue and items as two series, so price vs volume is visible); products with ABC color-coding (A green, C grey); segments; a campaign list (customer, phone/email, segment, message, send date, sent?, and consent? for EU); inventory with status colors.
- Freeze the header row, use locale number formats and fit the column widths. Persian workbooks get RTL sheet direction.
- Never overwrite the user's source files. Write a new file such as `Sales report [period].xlsx` / `گزارش فروش [بازه].xlsx`.

## 4. HTML dashboard

A single self-contained page (inline CSS/JS; a chart library from a CDN is fine) built from `analysis.json`:
- Direction, fonts, dates and money per the locale table above.
- Layout: a header (business, period, currency, VAT basis) → a KPI tile row with colored deltas → monthly revenue and items chart → YoY comparison → segments (customers and revenue share) → top products with ABC badges → weekday pattern → inventory alerts → "Problems & actions" cards → a data-notes footer.
- Make it readable on a phone, since owners will open it on mobile.
- **Privacy:** aggregated customer data only. No names, phones or emails in a dashboard meant to be shared.
- If an artifact/publishing tool is available and the user wants to share it, publish it as a private page. Otherwise save `dashboard.html` next to the other outputs.

## 5. PDF / Word report

Use the docx or pdf skill if available. Structure:
1. Cover: title, business, period, date.
2. Executive summary (half a page): the situation, the 3 problems and the 3 actions.
3. Sales performance: a KPI table, the monthly chart (as an image) and the price/volume decomposition.
4. Customers: segments, repeat rate, concentration, contactability.
5. Products and inventory: top/bottom products, ABC, stock alerts.
6. Marketing plan: a campaign table (target, offer, channel, timing, budget cap, KPI) and the message texts.
7. Appendix: data sources, currency/VAT assumption, data-quality notes.

Persian documents need an RTL layout and a Persian-capable font (Vazirmatn, B Nazanin, Tahoma); check that the charts render Persian labels with correct shaping before delivering.
