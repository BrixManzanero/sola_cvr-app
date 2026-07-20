# SOLA CVR — Traffic Conversion Extractor

A web app that extracts **Conversion Rate (CVR)** and related traffic metrics from
Shopee, TikTok, and Lazada Excel exports, and produces clean, plain Excel reports.
Sales/GMV/Revenue data is automatically excluded.

---

## How to use (for anyone — no setup needed)

1. Open the app link in any browser (phone or computer).
2. Pick the platform tab: **Shopee**, **TikTok**, or **Lazada**.
3. Drag in your export file(s) for that platform, or click **Browse files**.
4. Click **Process**.
5. Review the on-screen table (per-month CVR), then click **Download Excel**.

### Continuing from last month (so you don't re-upload everything)
- Open the **📁 Files** section at the top of a tab.
- Upload last month's downloaded Excel as the baseline.
- Add only the new month's file below, then Process.
- The app merges them and keeps all your history.

---

## What each platform extracts

All three platforms now use ONE unified, comparable formula: **CVR = Orders ÷ Visitors**.

| Platform | Output columns | CVR formula |
|---|---|---|
| Shopee | Month, Visitors, Paid Orders, CVR | Paid Orders ÷ Visitors |
| TikTok | Month, Page Views, Visitors, Product Clicks, Orders, CVR | Orders ÷ Visitors |
| Lazada | Month, Pageviews, Visitors, Buyers, Orders, CVR | Orders ÷ Visitors |
| Shopify | Month, Visitors, Orders, CVR | Orders ÷ Visitors |

**Shopify is manual entry.** There's no clean Shopify export with unique Visitors +
Orders together, so the Shopify tab lets you TYPE the two numbers per month
(from Shopify Analytics → Online store **visitors** (not Sessions) + Orders).
You can also load a previous Shopify output to keep your history and just add the new month.

**Shopee needs TWO files per month** because the numbers live in different exports:
the **Overview** export supplies **Visitors**, and the **Traffic** export supplies
**Orders** (placed / confirmed / paid). Upload both for the same month and the app
merges them. The Traffic file's month is read from its filename, so keep it named
like `Shopee Traffic_January_2026.xlsx`.

Each output shows **per-month** CVR only — no aggregate TOTAL row, because the
visitor de-dup basis differs across platforms (Lazada/Shopify are true monthly-unique;
TikTok/Shopee are summed-daily), so a single overall CVR would not be apples-to-apples.

### Smart file handling (automatic)
- **Overlapping months** (e.g. May 1–15 and May 1–31): keeps the most complete one.
- **Complementary halves** (e.g. May 1–15 and May 16–31): combines into one full month.
- **Exact duplicates**: blocked, with a message to remove one file.

---

## For developers / maintainers

**Run locally:**
```bash
pip install -r requirements.txt
streamlit run app.py
```

**Project structure:**
```
app.py                 entry point
backend/               pure Python (no Streamlit) — extraction + logic
  utils.py             formatters, safe readers, plain-Excel writer
  shopee.py            Shopee extractor + Excel builder
  tiktok.py            TikTok extractor + Excel builder
  lazada.py            Lazada extractor + Excel builder
  processor.py         pipeline + smart overlap resolver
  auditor.py           post-extraction CVR verifier
  reader.py            reads previous output Excel as baseline
frontend/              Streamlit UI only
  config.py            platform settings
  styles.py            CSS
  components.py        UI rendering
```

**Deployment:** hosted on Streamlit Community Cloud (share.streamlit.io),
connected to this GitHub repository. Pushing changes to the main branch
auto-redeploys the live app.
