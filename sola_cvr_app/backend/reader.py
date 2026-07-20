"""
backend/reader.py
Reads a previously generated SOLA CVR output Excel back into the pipeline.
This lets users treat their last output as a baseline and just upload new files on top.

Shopee output  → 3 tabs: Placed Order | Confirmed Order | Paid Order
TikTok output  → 1 tab: Traffic Conversion
Lazada output  → 1 tab: Traffic Conversion
"""

import io
import calendar as cal
from datetime import datetime

import pandas as pd


# ─────────────────────────────────────────────────────────────────────────────
# DATE HELPERS
# ─────────────────────────────────────────────────────────────────────────────

def _parse_label(label: str):
    """'January 2025' → datetime(2025,1,1)   'Jan 2025' → same."""
    for fmt in ("%B %Y", "%b %Y"):
        try:
            return datetime.strptime(str(label).strip(), fmt)
        except ValueError:
            pass
    return None


def _month_last(dt: datetime) -> datetime:
    last = cal.monthrange(dt.year, dt.month)[1]
    return datetime(dt.year, dt.month, last)


def _full_month_meta(dt: datetime) -> dict:
    """Return _day_count / _period_start / _period_end for a complete month."""
    end = _month_last(dt)
    return {
        "_day_count":    (end - dt).days + 1,
        "_period_start": dt,
        "_period_end":   end,
    }


# ─────────────────────────────────────────────────────────────────────────────
# SHEET READER — extracts monthly-breakdown rows from any tab
# ─────────────────────────────────────────────────────────────────────────────

def _extract_monthly_rows(sheet_df: pd.DataFrame) -> list[dict]:
    """
    Scan a sheet for the "Monthly Breakdown" section and return rows as dicts.
    Columns: [Month/Period, col1, col2, col3, ...]
    Stops at the first row that can't be parsed as a month label.
    """
    rows = []
    in_section = False
    headers    = []

    for _, row in sheet_df.iterrows():
        vals  = [v for v in row.tolist()]
        first = str(vals[0]).strip() if pd.notna(vals[0]) else ""

        if "Monthly Breakdown" in first:
            in_section = True
            continue

        if in_section:
            # Header row (Month / Period, Clicks, …)
            if not headers and "Month" in first:
                headers = [str(v).strip() for v in vals]
                continue

            # Data row — first cell must be a parseable month label
            dt = _parse_label(first)
            if dt is None:
                break          # end of section

            if not headers:
                headers = [f"col{i}" for i in range(len(vals))]

            row_dict = dict(zip(headers, vals))
            row_dict["_dt"] = dt
            rows.append(row_dict)

    return rows


# ─────────────────────────────────────────────────────────────────────────────
# SHOPEE READER
# ─────────────────────────────────────────────────────────────────────────────

def read_shopee(uploaded_file) -> tuple:
    """
    Read a previously generated Shopee CVR Excel (single 'Paid Order' tab) back
    into a fake extract dict. Layout: Month | Paid Orders | Product Clicks | CVR.
    Returns (extract_dict, error_string).  error_string is None on success.
    """
    uploaded_file.seek(0)
    data = uploaded_file.read()

    # New unified output has a single 'Traffic Conversion' tab:
    #   Month | Visitors | Paid Orders | CVR
    df = None
    for tab in ("Traffic Conversion", "Paid Order"):
        try:
            df = pd.read_excel(io.BytesIO(data), sheet_name=tab, header=None)
            break
        except Exception:
            continue
    if df is None:
        return None, "No 'Traffic Conversion' tab found in previous output"

    rows_raw = _extract_monthly_rows_simple(df)
    if not rows_raw:
        return None, "No monthly data found in previous Shopee output"

    def _cell(vals, idx):
        try:
            v = vals[idx]
            return v if pd.notna(v) else None
        except Exception:
            return None

    rows = []
    for r in rows_raw:
        dt   = r["_dt"]
        vals = r["_vals"]
        meta = _full_month_meta(dt)
        rows.append({
            "Date":             dt,
            "Month Label":      dt.strftime("%B %Y"),
            "Visitors":         _cell(vals, 1),
            "Placed Orders":    None,
            "Confirmed Orders": None,
            "Paid Orders":      _cell(vals, 2),
            "_role":            "both",
            "_from_output":     True,
            **meta,
        })

    monthly_df = pd.DataFrame(rows)
    first = rows[0]["Date"]
    summary = pd.Series({
        "Date": str(first), "Visitors": "—",
        "Placed CVR": "—", "Confirmed CVR": "—", "Paid CVR": "—",
    })

    return {
        "summary":       summary,
        "monthly":       monthly_df,
        "start_date":    first,
        "name":          f"[Previous output] {uploaded_file.name}",
        "_from_output":  True,
        "_kind":         "output",
    }, None


def _extract_monthly_rows_simple(sheet_df: pd.DataFrame) -> list[dict]:
    """
    Extract data rows from a single-tab output (Month | … | … | CVR).
    Skips the title and header row, stops at TOTAL or unparseable month.
    Accepts both 'Jan-25' and 'January 2025' month formats.
    """
    rows = []
    for _, row in sheet_df.iterrows():
        vals  = list(row.tolist())
        first = str(vals[0]).strip() if pd.notna(vals[0]) else ""
        if not first or first.upper() == "TOTAL" or "Month" in first or "Traffic Conversion" in first:
            continue
        dt = _parse_label_any(first)
        if dt is None:
            continue
        rows.append({"_dt": dt, "_vals": vals})
    return rows


def _parse_label_any(label: str):
    """Parse 'Jan-25', 'January 2025', 'Jan 2025' → datetime."""
    s = str(label).strip()
    for fmt in ("%b-%y", "%B %Y", "%b %Y", "%b-%Y"):
        try:
            return datetime.strptime(s, fmt)
        except ValueError:
            pass
    return None


# ─────────────────────────────────────────────────────────────────────────────
# TIKTOK READER
# ─────────────────────────────────────────────────────────────────────────────

def read_tiktok(uploaded_file) -> tuple:
    uploaded_file.seek(0)
    data = uploaded_file.read()

    try:
        df = pd.read_excel(io.BytesIO(data), sheet_name="Traffic Conversion", header=None)
    except Exception as e:
        return None, f"Cannot read Traffic Conversion tab: {e}"

    monthly_rows_raw = _extract_monthly_rows(df)
    if not monthly_rows_raw:
        return None, "No monthly data found in previous TikTok output"

    # Detect if a Customers column exists (header row holds the names)
    header_names = []
    for _, row in df.iterrows():
        first = str(row.tolist()[0]).strip() if pd.notna(row.tolist()[0]) else ""
        if "Month" in first:
            header_names = [str(v).strip() for v in row.tolist()]
            break
    cust_idx = next((i for i, h in enumerate(header_names) if h == "Customers"), None)

    rows = []
    for r in monthly_rows_raw:
        dt   = r["_dt"]
        meta = _full_month_meta(dt)
        vals = list(r.values())

        def safe(idx, default="0"):
            try:
                v = vals[idx]
                return str(v) if pd.notna(v) else default
            except Exception:
                return default

        row_out = {
            "Date":            dt,
            "Month Label":     dt.strftime("%B %Y"),
            "Page Views":      safe(1),
            "Visitors":        safe(2),
            "Product Clicks":  safe(3),
            "Orders":          safe(4),
            "Conversion Rate": safe(len(header_names) - 1, "N/A") if header_names else safe(5, "N/A"),
            "_from_output":    True,
            **meta,
        }
        if cust_idx is not None:
            row_out["Customers"] = safe(cust_idx)
        rows.append(row_out)

    monthly_df = pd.DataFrame(rows)
    first = rows[0]["Date"]
    summary = {
        "Period":          f"{first.strftime('%d/%m/%Y')}–present",
        "Page Views":      "—", "Visitors": "—",
        "Product Clicks":  "—", "Orders":   "—",
        "Conversion Rate": "—",
    }

    return {
        "summary":      summary,
        "monthly":      monthly_df,
        "start_date":   first,
        "name":         f"[Previous output] {uploaded_file.name}",
        "_from_output": True,
    }, None


# ─────────────────────────────────────────────────────────────────────────────
# LAZADA READER
# ─────────────────────────────────────────────────────────────────────────────

def read_lazada(uploaded_file) -> tuple:
    uploaded_file.seek(0)
    data = uploaded_file.read()

    try:
        df = pd.read_excel(io.BytesIO(data), sheet_name="Traffic Conversion", header=None)
    except Exception as e:
        return None, f"Cannot read Traffic Conversion tab: {e}"

    monthly_rows_raw = _extract_monthly_rows(df)
    if not monthly_rows_raw:
        return None, "No monthly data found in previous Lazada output"

    rows = []
    for r in monthly_rows_raw:
        dt   = r["_dt"]
        meta = _full_month_meta(dt)
        vals = list(r.values())

        def safe(idx, default="0"):
            try:
                v = vals[idx]
                return str(v) if pd.notna(v) else default
            except Exception:
                return default

        rows.append({
            "Date":            dt,
            "Month Label":     dt.strftime("%B %Y"),
            "Pageviews":       safe(1),
            "Visitors":        safe(2),
            "Buyers":          safe(3),
            "Orders":          safe(4),
            "Conversion Rate": safe(5, "N/A"),
            "_from_output":    True,
            **meta,
        })

    monthly_df = pd.DataFrame(rows)
    first = rows[0]["Date"]
    row   = rows[0]
    summary = {k: v for k, v in row.items() if not k.startswith("_")}

    return {
        "row":          row,
        "summary":      summary,
        "period":       first.strftime("%Y-%m-%d") + "~" + _month_last(first).strftime("%Y-%m-%d"),
        "start_date":   first,
        "name":         f"[Previous output] {uploaded_file.name}",
        "_from_output": True,
    }, None


# ─────────────────────────────────────────────────────────────────────────────
# DISPATCHER
# ─────────────────────────────────────────────────────────────────────────────

READERS = {
    "shopee": read_shopee,
    "tiktok": read_tiktok,
    "lazada": read_lazada,
}


def read_previous_output(uploaded_file, platform: str) -> tuple:
    """Read a previously generated output Excel for any platform."""
    reader = READERS.get(platform)
    if reader is None:
        return None, f"Unknown platform: {platform}"
    return reader(uploaded_file)
