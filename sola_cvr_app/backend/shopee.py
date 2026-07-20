"""
backend/shopee.py
Shopee CVR extractor — UNIFIED FORMULA: CVR = Orders ÷ Visitors.

Shopee splits the two numbers we need across TWO different exports:
  • Overview export  → 'overview' sheet, daily rows, has 'Product Visitors (Visit)'
                       → this is where VISITORS come from.
  • Traffic export   → '(placed)/(confirmed)/(paid) Product Traffic' sheets,
                       per-product rows, has 'Orders' + 'Product Clicks'
                       → this is where ORDERS come from (per order status).

Because the Traffic export has NO date inside it, the month for a Traffic file
is read from its FILENAME (e.g. 'Shopee Traffic_January_2026.xlsx' → Jan 2026).
The Overview file's month(s) are read from its own Date column.

The processor merges the Visitors row (Overview) and the Orders row (Traffic)
for the same month, then computes CVR = Orders ÷ Visitors for each order status.
"""

import re
import calendar
from datetime import datetime

import pandas as pd
from openpyxl import Workbook

from backend.utils import (
    fmt_num, safe_read, safe_xl, to_xlsx_bytes, write_plain_sheet,
)


# ─────────────────────────────────────────────────────────────────────────────
# CONSTANTS
# ─────────────────────────────────────────────────────────────────────────────

OVERVIEW_VISITORS_COL = "Product Visitors (Visit)"
TRAFFIC_ORDERS_COL    = "Orders"
STATUS_KEYS           = ["placed", "confirmed", "paid"]
ORDER_COLS            = ["Placed Orders", "Confirmed Orders", "Paid Orders"]

_MONTHS = {m.lower(): i for i, m in enumerate(calendar.month_name) if m}
_MONTHS.update({m.lower(): i for i, m in enumerate(calendar.month_abbr) if m})


# ─────────────────────────────────────────────────────────────────────────────
# SHARED DATE HELPERS (kept for processor imports)
# ─────────────────────────────────────────────────────────────────────────────

def parse_date(val):
    """Parse a date. Accepts a datetime or 'dd/mm/yyyy' string."""
    if isinstance(val, datetime):
        return val
    try:
        return datetime.strptime(str(val).strip(), "%d/%m/%Y")
    except Exception:
        return None


def parse_summary_date(val) -> datetime:
    dt = parse_date(val)
    return dt if dt else datetime.min


def to_month_label(val) -> str:
    dt = parse_date(val)
    return dt.strftime("%B %Y") if dt else str(val)


def _month_from_filename(name: str):
    """'...January_2026...' or '...Jan-2026...' → datetime(2026, 1, 1)."""
    s = str(name)
    ym = re.search(r"(20\d{2})", s)
    year = int(ym.group(1)) if ym else None
    for token in re.split(r"[\s_\-.]+", s):
        key = token.strip().lower()
        if key in _MONTHS and year:
            return datetime(year, _MONTHS[key], 1)
    return None


def _month_end(dt: datetime) -> datetime:
    last = calendar.monthrange(dt.year, dt.month)[1]
    return datetime(dt.year, dt.month, last)


def _to_float(v) -> float:
    try:
        s = str(v).replace(",", "").replace("%", "").strip()
        if s in ("", "-", "nan", "None"):
            return 0.0
        return float(s)
    except Exception:
        return 0.0


# ─────────────────────────────────────────────────────────────────────────────
# FILE-KIND DETECTION
# ─────────────────────────────────────────────────────────────────────────────

def _detect_kind(uploaded_file):
    """Return 'overview' | 'traffic' | None by inspecting sheets/columns."""
    try:
        xl = safe_xl(uploaded_file)
    except Exception:
        return None
    names = xl.sheet_names

    if any("product traffic" in str(n).lower() for n in names):
        return "traffic"

    for n in names:
        if str(n).strip().lower() == "overview":
            return "overview"

    for n in names:
        try:
            df = safe_read(uploaded_file, sheet_name=n, header=0, nrows=1)
            cols = {str(c).strip() for c in df.columns}
        except Exception:
            continue
        if OVERVIEW_VISITORS_COL in cols:
            return "overview"
        if TRAFFIC_ORDERS_COL in cols and "Product Clicks" in cols:
            return "traffic"
    return None


# ─────────────────────────────────────────────────────────────────────────────
# VALIDATOR
# ─────────────────────────────────────────────────────────────────────────────

def validate(uploaded_file) -> tuple:
    """(True, kind) for a valid Shopee Overview or Traffic export, else (False, reason)."""
    if not uploaded_file.name.lower().endswith(".xlsx"):
        return False, "Expected a .xlsx file"
    kind = _detect_kind(uploaded_file)
    if kind == "overview":
        return True, "overview"
    if kind == "traffic":
        if _month_from_filename(uploaded_file.name) is None:
            return False, ("Traffic file needs the month in its filename, e.g. "
                           "'Shopee Traffic_January_2026.xlsx'")
        return True, "traffic"
    return False, ("Not a recognised Shopee export — expected an Overview sheet "
                   "(with 'Product Visitors (Visit)') or Product Traffic sheets (with 'Orders').")


# ─────────────────────────────────────────────────────────────────────────────
# EXTRACTORS
# ─────────────────────────────────────────────────────────────────────────────

def _extract_overview(uploaded_file) -> dict:
    """Monthly Visitors from an Overview export (sum of daily Product Visitors)."""
    xl = safe_xl(uploaded_file)
    sheet = next((n for n in xl.sheet_names if str(n).strip().lower() == "overview"),
                 xl.sheet_names[0])
    df = safe_read(uploaded_file, sheet_name=sheet, header=0)
    df.columns = [str(c).strip() for c in df.columns]

    df["_dt"] = pd.to_datetime(df["Date"], dayfirst=True, errors="coerce")
    df = df.dropna(subset=["_dt"])
    df["_vis"] = df[OVERVIEW_VISITORS_COL].apply(_to_float)
    df["_month"] = df["_dt"].dt.to_period("M")

    rows = []
    for month, g in df.groupby("_month"):
        mstart = month.to_timestamp().to_pydatetime()
        rows.append({
            "Date":             mstart,
            "Month Label":      mstart.strftime("%B %Y"),
            "Visitors":         int(g["_vis"].sum()),
            "Placed Orders":    None,
            "Confirmed Orders": None,
            "Paid Orders":      None,
            "_role":            "visitors",
            "_day_count":       int(g["_dt"].dt.day.nunique()),
            "_period_start":    g["_dt"].min().to_pydatetime(),
            "_period_end":      g["_dt"].max().to_pydatetime(),
        })
    monthly = pd.DataFrame(rows)
    start = monthly["Date"].min() if not monthly.empty else datetime.min
    summary = pd.Series({
        "Date": str(start), "Visitors": "—",
        "Placed CVR": "—", "Confirmed CVR": "—", "Paid CVR": "—",
    })
    return {"summary": summary, "monthly": monthly, "start_date": start,
            "name": uploaded_file.name, "_kind": "overview"}


def _extract_traffic(uploaded_file) -> dict:
    """Monthly Orders (per status) from a Traffic export. Month from filename."""
    xl = safe_xl(uploaded_file)
    mstart = _month_from_filename(uploaded_file.name)
    orders = {"placed": None, "confirmed": None, "paid": None}

    for sheet in xl.sheet_names:
        low = str(sheet).lower()
        status = next((k for k in STATUS_KEYS if k in low), None)
        if status is None:
            continue
        df = safe_read(uploaded_file, sheet_name=sheet, header=0)
        df.columns = [str(c).strip() for c in df.columns]
        if TRAFFIC_ORDERS_COL not in df.columns:
            continue
        orders[status] = float(df[TRAFFIC_ORDERS_COL].apply(_to_float).sum())

    row = {
        "Date":             mstart,
        "Month Label":      mstart.strftime("%B %Y"),
        "Visitors":         None,
        "Placed Orders":    orders["placed"],
        "Confirmed Orders": orders["confirmed"],
        "Paid Orders":      orders["paid"],
        "_role":            "orders",
        "_day_count":       calendar.monthrange(mstart.year, mstart.month)[1],
        "_period_start":    mstart,
        "_period_end":      _month_end(mstart),
    }
    monthly = pd.DataFrame([row])
    summary = pd.Series({
        "Date": str(mstart), "Visitors": "—",
        "Placed CVR": "—", "Confirmed CVR": "—", "Paid CVR": "—",
    })
    return {"summary": summary, "monthly": monthly, "start_date": mstart,
            "name": uploaded_file.name, "_kind": "traffic"}


def extract(uploaded_file) -> dict:
    """Dispatch to the Overview or Traffic extractor based on file kind."""
    kind = _detect_kind(uploaded_file)
    if kind == "overview":
        return _extract_overview(uploaded_file)
    if kind == "traffic":
        return _extract_traffic(uploaded_file)
    raise ValueError("Unrecognised Shopee file (neither Overview nor Traffic).")


# ─────────────────────────────────────────────────────────────────────────────
# CVR HELPER (shared with processor / excel / auditor)
# ─────────────────────────────────────────────────────────────────────────────

def cvr_str(orders, visitors) -> str:
    """Orders ÷ Visitors → '6.73%' string, or 'N/A'."""
    if orders is None:
        return "N/A"
    o = _to_float(orders)
    v = _to_float(visitors)
    if not v:
        return "N/A"
    return f"{o / v * 100:.2f}%"


# ─────────────────────────────────────────────────────────────────────────────
# EXCEL BUILDER
# ─────────────────────────────────────────────────────────────────────────────

def build_excel(monthly_df: pd.DataFrame, summaries: list) -> bytes:
    """
    Plain Shopee CVR Excel — single 'Traffic Conversion' tab:
        Month | Visitors | Paid Orders | CVR   + TOTAL row
    CVR = Paid Orders ÷ Visitors. TOTAL = total Paid Orders ÷ total Visitors.
    """
    wb = Workbook()
    wb.remove(wb.active)
    ws = wb.create_sheet("Traffic Conversion")
    ws.sheet_view.showGridLines = True

    rows = []
    t_vis = 0.0
    t_paid = 0.0

    for _, row in monthly_df.iterrows():
        vis = _to_float(row.get("Visitors", 0))
        t_vis += vis
        opaid = row.get("Paid Orders")
        t_paid += _to_float(opaid)

        dt = parse_date(row.get("Date"))
        month_lbl = dt.strftime("%b-%y") if dt else str(row.get("Date"))
        rows.append([
            month_lbl,
            f"{int(vis):,}" if vis else "0",
            f"{int(round(_to_float(opaid))):,}" if opaid is not None else "N/A",
            cvr_str(opaid, vis),
        ])

    total_row = [
        "TOTAL",
        f"{int(t_vis):,}",
        f"{int(round(t_paid)):,}",
        cvr_str(t_paid, t_vis),
    ]

    write_plain_sheet(
        ws,
        headers=["Month", "Visitors", "Paid Orders", "CVR"],
        rows=rows,
        total_row=total_row,
        col_widths=[14, 14, 14, 12],
    )
    return to_xlsx_bytes(wb)
