# Metrics & Diagnosis

## Contents
1. Core metric definitions (as computed by analyze.mjs)
2. Revenue decomposition (price vs volume)
3. RFM segmentation rules
4. ABC and product signals
5. Inventory statuses
6. Rough benchmarks for small retail
7. Diagnostic tree: symptom → likely causes → what to check

## 1. Core metrics

| Metric | Definition | Note |
|---|---|---|
| Net revenue | Per line, using the first amount column that is filled on ≥90% of lines: net → final − returns → gross − discount + tax − returns → unit price × qty − discount − returns | Order-level totals shown only on an order's first line are thereby skipped. Whether VAT is included depends on the column; say which one was used |
| Invoices / orders | Distinct invoice/order numbers with net > 0; cancelled/voided/refunded orders (status column) are excluded entirely | Without an invoice column: date + customer |
| Items sold | Σ(qty − returned qty) | |
| Customers | Distinct identities: phone, email, customer ID or cleaned name (honorifics such as خانم/آقای/Mr/Mrs removed, Arabic letters normalized); identifiers seen on the same order are merged | Generic names (مشتری عمومی, walk-in, guest, anonymous...) excluded |
| AOV | Net revenue of sale invoices ÷ invoices | میانگین مبلغ فاکتور |
| Items / invoice | Items ÷ invoices | basket size |
| Avg item price | Net revenue ÷ items | the inflation lens |
| Discount rate | Σ discount ÷ Σ gross | also % of invoices with any discount |
| Return rate | Σ return amount ÷ Σ gross | |
| Repeat rate | Customers with ≥2 purchase days ÷ identified customers | same-day multiple invoices count once |
| YoY | Same calendar month, consecutive years | `yearComparison` compares only months present in the latest year |

Customer identity is name-based when there is no phone/ID column. Two different people with the same name merge, and spelling variants split one person. Mention this when repeat rate matters, and prefer `--customers` with a phone database.

## 2. Revenue decomposition

```
Revenue = Customers × (Invoices / Customer) × AOV
AOV     = (Items / Invoice) × Avg item price
```

Report changes in each factor side by side. Interpretation:
- Revenue ↑ but items ↓ and avg price ↑ strongly → **price-driven growth; real volume is shrinking.** Common in high-inflation markets: Iran (prices often +30–60% a year), Turkey, Argentina. In the eurozone/US, compare against inflation of roughly 2–5% and the shop's own price changes. Compare revenue growth with the price increase (from price lists, or the avg item price trend) to estimate *real* growth.
- Customers ↓ with stable frequency → acquisition / footfall problem.
- Customers stable, frequency ↓ → retention problem.
- AOV ↓ with items/invoice stable → shift to cheaper products, or heavier discounts.

## 3. RFM segments (rule-based, works when most customers buy once)

The reference date is the last date in the data. Thresholds scale with the data span: recent = max(30, 15% of span), active = max(90, 35%), lapsed = max(180, 60%). The values used are in `customers.stats.segmentThresholdsDays`. M score = monetary quintile 1–5.

| Segment | Rule | Typical action |
|---|---|---|
| vip | ≥2 purchases, recency ≤ active, and (≥3 purchases or M ≥4) | recognition, early access, gift, no blanket discounts |
| loyal | ≥2 purchases, recency ≤ active | cross-sell, referral |
| atRisk | ≥2 purchases, active < recency ≤ lapsed | personal win-back before they're lost |
| lostRepeat | ≥2 purchases, recency > lapsed | strong win-back offer, or ask why |
| highValueOnce | 1 purchase, top-20% spend | convert to a 2nd purchase: the biggest lever in low-repeat shops |
| newRecent | 1 purchase, recency ≤ recent | thank-you + 2nd-purchase incentive with a deadline |
| oneTimeCooling | 1 purchase, recent < recency ≤ active | reminder + new arrivals |
| oneTimeLost | 1 purchase, recency > active | only cheap channels (bulk SMS at big occasions) |

The second purchase is the critical moment: a customer who buys twice is several times more likely to buy again. In shops with a repeat rate under 15%, converting one-timers usually beats any acquisition campaign.

## 4. Product signals

- **ABC**: A = products making the first 80% of revenue, B = next 15%, C = last 5%. A items must never be out of stock. C items with stock are candidates for bundles or clearance.
- **Concentration**: if fewer than 20% of products make 80% of revenue, the range is long-tailed; if more than 40% do, there's no clear hero product.
- **Rising stars**: 2nd-half revenue ≥ 1.5× 1st-half. Push them and restock.
- **Declining**: A/B items with 2nd-half revenue ≤ 60% of 1st-half. Check price, stock-outs, competition, season.
- **Models/families**: `models` groups by the first word of the product name (e.g. a model in several colors). Useful for "which line is growing"; verify it makes sense for the user's naming.
- **Basket pairs**: frequent pairs suggest bundles and cross-sell. If multi-item invoices are under 10%, basket-building is itself an opportunity (second item at a discount).
- **Returns**: products with returns above ~5% of qty point to quality, sizing or expectation problems.

## 5. Inventory statuses (with --stock)

Velocity = qty sold in the last min(90, span) days ÷ days. Days of cover = stock ÷ velocity.

| Status | Rule | Action |
|---|---|---|
| out_of_stock_selling | stock ≤ 0 and sold recently | reorder now; quantify lost sales ≈ velocity × days out |
| low_stock | cover < lead days (default 21) | reorder; A items first |
| healthy | lead ≤ cover ≤ 180 | — |
| overstock | cover > 180 | slow the reorders; feature in marketing |
| dead_no_recent_sales | stock > 0, no sales in the window | bundle, clearance, return to supplier |
| dead_never_sold | stock item that never appears in sales | check the listing/display; possibly a naming mismatch |

Without a stock file, `products.slowMovers` lists products not sold in ≥60 days. Ask for a stock export (موجودی کالا) to go further.

## 6. Rough benchmarks (use as orientation, not truth; they vary by sector)

- Repeat rate (share of customers with ≥2 purchases over a year): fashion/accessories 15–30%; cafés, groceries and beauty 40%+; e-commerce 20–30%. Under 10% is weak everywhere. It's only measurable where customers are identified (name/phone/email/loyalty card); with many anonymous walk-ins, say so.
- Top 10% of customers ≈ 30–50% of revenue is typical.
- Discount rate: above 15–20% of gross is heavy; check whether discounted months grow *volume* or just shift purchases.
- Multi-item orders: 10–25% typical for accessories; cafés and groceries are much higher.
- Return rate: physical retail 2–8%; online fashion 20–30% is normal.
- Weekday peaks: Iran usually Thursday/Friday (the week starts Saturday); Europe/US physical retail usually Saturday (the week starts Monday, and Sunday trading is restricted in some countries such as Germany). Online shops usually peak Sunday evening and Monday. Paydays (start/end of month) lift sales in most markets. Check the weekday and day-of-month results instead of assuming.
- Seasonality: always compare against the same month last year before calling a drop a problem. December (Christmas) or Esfand (pre-Nowruz) can be 2–3× a normal month.

## 7. Diagnostic tree

**Revenue down**
1. Is volume (items) down, or price/mix? → decomposition.
2. Fewer customers? → footfall, marketing activity, competitors, location, seasonality (compare the same month last year).
3. Fewer repeat buyers? → after-sales contact, loyalty program, product freshness.
4. Lower AOV? → cheaper mix, fewer multi-item baskets, discounts.
5. Missing data? → a damaged/missing file or a month only partly exported (check `meta.from/to` and the rows per month).
6. Channel shift? → sales may have moved online/to marketplaces, or to another branch, rather than been lost. Ask whether other channels exist.

**Revenue up but the business feels worse** → price-driven growth with falling units; margin squeeze from discounts; growth concentrated in a few customers or products (fragile).

**Peaks only in discount months** → customers are trained to wait for sales. Test non-price triggers (new collection, gift with purchase, VIP early access) and shorter, targeted discounts instead of store-wide ones.

**Many one-time customers** → there's no second-purchase mechanism: no phone capture, no follow-up, no reason to return. Fix capture first (phone completeness), then automate a post-purchase message.

**High discount rate with low growth** → discounts are subsidizing purchases that would have happened anyway. Compare AOV and items/invoice for discounted vs full-price invoices (`discounts`).

**Stock problems** → A items out of stock = lost revenue; dead stock = frozen cash. Quantify both in money.
