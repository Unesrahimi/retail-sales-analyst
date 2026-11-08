# Retail Sales Analyst: a Claude skill

**English** · [فارسی](#فارسی)

A [Claude](https://claude.ai) skill that turns a shop's sales export (Excel or CSV) into a diagnosis and a marketing plan. It works for Iranian businesses (Persian, Shamsi dates, Toman/Rial, SMS) and international ones (English, Gregorian dates, EUR/USD/CZK, email/SMS under GDPR).

## What it does

| Capability | Output |
|---|---|
| Sales analysis | revenue, orders, average basket, discounts, returns; monthly, same month last year, weekday; price effect vs real volume |
| Customer segmentation | RFM groups (VIP, loyal, at risk, high-value one-timers, new, dormant) and contactability |
| Product performance | best/worst sellers, ABC classes, rising and declining items, products bought together |
| Inventory intelligence | low stock on fast movers, dead stock and its value, overstock (with a stock export) |
| Marketing strategy | prioritised campaigns with target list, offer, channel, timing, budget cap, KPI and message texts |
| Business reporting | chat summary, Excel report, HTML dashboard, or PDF/Word |

It reads Shopify, Square, WooCommerce, Lightspeed, Iranian accounting exports (Sepidar, Hamkaran, Mahak, Holoo), and German/Czech POS and accounting exports. Cancelled orders, duplicate files, damaged xlsx files, `1.234,56` vs `1,234.56` and DD/MM vs MM/DD are handled automatically.

## Install

**Claude Code** (needs [Node.js](https://nodejs.org)):

```bash
npx skills add <github-user>/retail-sales-analyst -g
cd ~/.claude/skills/retail-sales-analyst/scripts && npm install
```

or copy this folder to `~/.claude/skills/retail-sales-analyst/` and run `npm install` in `scripts/`.

**Claude.ai**: Settings → Capabilities → Skills → Upload skill, then choose a zip of this folder.

## Use

Give Claude a sales file and ask in plain words:

> Here's our Shopify export for 2026. Why is revenue up but the shop feels quieter?
>
> فروش این شش ماه رو تحلیل کن و برای یلدا کمپین پیامکی پیشنهاد بده.

The analyzer can also be run on its own:

```bash
node scripts/analyze.mjs ./exports --out ./report --currency EUR --excel
node scripts/analyze.mjs ./فروش --out ./گزارش --divisor 10 --currency Toman --excel --customers customers.xlsx
```

Options: `--preset shopify|square`, `--stock`, `--customers`, `--map`, `--date-format dmy|mdy`, `--decimal comma|dot`, `--basis`, `--lang fa|en`. See the header of `scripts/analyze.mjs`.

## Structure

```
SKILL.md                              workflow and rules Claude follows
references/metrics-and-diagnosis.md   definitions, benchmarks, symptom → cause tree
references/marketing-playbook.md      segments → campaigns, Iranian & international calendars, consent rules, templates
references/report-formats.md          chat / Excel / dashboard / PDF layouts, RTL and LTR
scripts/analyze.mjs                   deterministic analyzer (Node.js + xlsx)
```

## Privacy

Everything runs locally. Customer names, phones and emails stay in the output files on your computer. Shared dashboards show totals only. For EU/UK shops, only message customers who gave consent.

---

## فارسی

اسکیلی برای Claude که خروجی فروش فروشگاه (Excel یا CSV) را تحلیل می‌کند، مشکلات کسب‌وکار را تشخیص می‌دهد و کمپین بازاریابی پیشنهاد می‌کند. برای کسب‌وکارهای ایرانی (تاریخ شمسی، تومان/ریال، پیامک) و خارجی (میلادی، یورو/دلار، ایمیل و GDPR).

**شش قابلیت:** تحلیل فروش · دسته‌بندی مشتریان (RFM) · عملکرد محصولات (ABC) · هوش موجودی · استراتژی بازاریابی · گزارش مدیریتی (چت، Excel، داشبورد، PDF).

خروجی سپیدار، همکاران، محک، هلو و همچنین Shopify، Square و WooCommerce را مستقیم می‌خواند. «ي/ك» عربی، اعداد فارسی، مبالغ ریالی، فایل‌های تکراری و خراب و فاکتورهای لغوشده را خودش درست می‌کند.

**نصب:** در Claude Code دستور بالا را بزنید. در Claude.ai از Settings ← Capabilities ← Skills فایل زیپ را آپلود کنید.

**استفاده:** فایل فروش را بدهید و ساده بپرسید، مثلاً «چرا فروشمون کم شده؟» یا «برای مشتری‌های VIP کمپین بده».

اطلاعات مشتری‌ها روی کامپیوتر خودتان می‌ماند و جایی ارسال نمی‌شود.
