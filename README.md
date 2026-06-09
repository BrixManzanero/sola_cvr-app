# SOLA CVR — Traffic Conversion Extractor

A web app that extracts **Conversion Rate (CVR)** and related traffic metrics from
Shopee, TikTok, and Lazada Excel exports, and produces clean, plain Excel reports.
Sales / GMV / Revenue data is automatically excluded.

## How to use (no setup needed)
1. Open the app link in any browser.
2. Pick the platform tab: Shopee, TikTok, or Lazada.
3. Drag in your export file(s), or click Browse files.
4. Click Process.
5. Review the table (with TOTAL row), then click Download Excel.

## Continuing from last month
- Open the **📁 Files** section at the top of a tab.
- Upload last month's downloaded Excel as the baseline.
- Add only the new month's file below, then Process.

## What each platform extracts
| Platform | Output columns | CVR formula |
|---|---|---|
| Shopee | Month, Paid Orders, Product Clicks, CVR | read directly from source |
| TikTok | Month, Page Views, Visitors, Product Clicks, Orders, CVR | Customers ÷ Visitors |
| Lazada | Month, Pageviews, Visitors, Buyers, Orders, CVR | Buyers ÷ Visitors |

## Run locally
```
pip install -r requirements.txt
streamlit run app.py
```
