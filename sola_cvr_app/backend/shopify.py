"""
backend/shopify.py
Shopify CVR — unified formula CVR = Orders ÷ Visitors (Online store visitors,
NOT sessions).

Shopify keeps traffic (Visitors, in the `sessions` dataset) and Orders (in the
`sales` dataset) in DIFFERENT datasets, so ShopifyQL can't return both in one
query. This module therefore accepts the two ShopifyQL exports (a Visitors file
and an Orders file — or a single file that already has both columns) plus an
optional manual-entry table, and MERGES everything by month.

Output tab: 'Traffic Conversion' — Month | Visitors | Orders | CVR + TOTAL.
"""

import io
import calendar
from datetime import datetime

import pandas as pd
from openpyxl import Workbook

from backend.utils import to_xlsx_bytes, write_plain_sheet


_MONTH_FMTS = ("%B %Y", "%b %Y", "%b-%y", "%b-%Y", "%B-%Y",
               "%Y-%m", "%Y-%m-%d", "%m/%Y", "%d/%m/%Y", "%Y/%m/%d")


def _to_float(v) -> float:
    try:
        s = str(v).replace(",", "").replace("%", "").replace("₱", "").strip()
        if s in ("", "-", "nan", "None"):
            return 0.0
        return float(s)
    except Exception:
        return 0.0


def parse_month(val):
    """Accept many month/date formats → datetime(first of month), else None."""
    if val is None:
        return None
    if isinstance(val, datetime):
        return datetime(val.year, val.month, 1)
    try:
        if isinstance(val, pd.Timestamp):
            return datetime(val.year, val.month, 1)
    except Exception:
        pass
    s = str(val).strip()
    if not s or s.lower() in ("nan", "none", "total"):
        return None
    for fmt in _MONTH_FMTS:
        try:
            dt = datetime.strptime(s, fmt)
            return datetime(dt.year, dt.month, 1)
        except ValueError:
            pass
    # last resort: let pandas try
    try:
        dt = pd.to_datetime(s, errors="coerce")
        if pd.notna(dt):
            return datetime(dt.year, dt.month, 1)
    except Exception:
        pass
    return None


def _cvr(orders, visitors) -> str:
    v = _to_float(visitors)
    if not v or orders is None:
        return "N/A"
    return f"{_to_float(orders) / v * 100:.2f}%"


# ─────────────────────────────────────────────────────────────────────────────
# COLUMN DETECTION for ShopifyQL / exploration exports
# ─────────────────────────────────────────────────────────────────────────────

def _find_col(cols, want):
    """Find a column by fuzzy name. want ∈ {'month','visitors','orders'}."""
    low = {c: str(c).strip().lower() for c in cols}
    if want == "month":
        for c, l in low.items():
            if l in ("month", "date", "day", "week", "period"):
                return c
        for c, l in low.items():
            if "month" in l or "date" in l:
                return c
    if want == "visitors":
        for c, l in low.items():
            if "visitor" in l:
                return c
    if want == "orders":
        bad = ("reorder", "per order", "per_order", "average", "value",
               "rate", "returned", "cancel")
        for c, l in low.items():
            if "order" in l and not any(b in l for b in bad):
                return c
    return None


def read_file(uploaded_file):
    """
    Read a ShopifyQL / exploration export (.csv or .xlsx).
    Returns (source_rows, error). Each source row is a dict:
        {"month": datetime, "visitors": float|None, "orders": float|None}
    """
    name = uploaded_file.name
    uploaded_file.seek(0)
    data = uploaded_file.read()
    try:
        if name.lower().endswith(".csv"):
            df = pd.read_csv(io.BytesIO(data))
        else:
            df = pd.read_excel(io.BytesIO(data))
    except Exception as exc:
        return [], f"Cannot read '{name}': {exc}"

    if df.empty:
        return [], f"'{name}' has no rows"

    m_col = _find_col(df.columns, "month")
    v_col = _find_col(df.columns, "visitors")
    o_col = _find_col(df.columns, "orders")

    if m_col is None:
        return [], (f"'{name}': no Month/Date column found. Export with "
                    f"'GROUP BY month' (columns like Month + Visitors/Orders).")
    if v_col is None and o_col is None:
        return [], (f"'{name}': no Visitors or Orders column found "
                    f"(columns: {list(df.columns)})")

    rows = []
    for _, r in df.iterrows():
        dt = parse_month(r.get(m_col))
        if dt is None:
            continue
        row = {"month": dt, "visitors": None, "orders": None}
        if v_col is not None:
            row["visitors"] = _to_float(r.get(v_col))
        if o_col is not None:
            row["orders"] = _to_float(r.get(o_col))
        rows.append(row)
    if not rows:
        return [], f"'{name}': no parseable month rows"
    kind = "both" if (v_col is not None and o_col is not None) else \
           ("visitors" if v_col is not None else "orders")
    return rows, None


def rows_from_manual(records: list):
    """Editor records [{Month, Visitors, Orders}] → source rows."""
    out = []
    for r in records or []:
        dt = parse_month(r.get("Month"))
        if dt is None:
            continue
        v = _to_float(r.get("Visitors"))
        o = _to_float(r.get("Orders"))
        if v == 0 and o == 0:
            continue
        out.append({"month": dt, "visitors": v if v else None,
                    "orders": o if o else None})
    return out


def rows_from_previous(uploaded_file):
    """Previous Shopify output (Month|Visitors|Orders|CVR) → source rows."""
    uploaded_file.seek(0)
    data = uploaded_file.read()
    try:
        raw = pd.read_excel(io.BytesIO(data), sheet_name="Traffic Conversion", header=0)
    except Exception:
        return []
    out = []
    for _, r in raw.iterrows():
        dt = parse_month(r.get("Month"))
        if dt is None:
            continue
        out.append({"month": dt, "visitors": _to_float(r.get("Visitors")),
                    "orders": _to_float(r.get("Orders"))})
    return out


# ─────────────────────────────────────────────────────────────────────────────
# MERGE  — combine Visitors + Orders across all sources, by month
# ─────────────────────────────────────────────────────────────────────────────

def merge_sources(sources: list) -> pd.DataFrame:
    """
    sources: list of {"month": dt, "visitors": float|None, "orders": float|None}
    Merge by month (Visitors from whichever source has it, Orders likewise),
    compute CVR = Orders ÷ Visitors. Returns sorted DataFrame:
        Date | Month | Visitors | Orders | CVR
    """
    by_month = {}
    for s in sources:
        dt = s.get("month")
        if dt is None:
            continue
        cur = by_month.setdefault(dt, {"visitors": None, "orders": None})
        if cur["visitors"] is None and s.get("visitors") not in (None, 0, 0.0):
            cur["visitors"] = s["visitors"]
        if cur["orders"] is None and s.get("orders") not in (None, 0, 0.0):
            cur["orders"] = s["orders"]

    rows = []
    for dt in sorted(by_month):
        vis = by_month[dt]["visitors"]
        ordr = by_month[dt]["orders"]
        rows.append({
            "Date":     dt,
            "Month":    dt.strftime("%B %Y"),
            "Visitors": int(vis) if vis is not None else 0,
            "Orders":   int(ordr) if ordr is not None else None,
            "CVR":      _cvr(ordr, vis),
        })
    return pd.DataFrame(rows)


# ─────────────────────────────────────────────────────────────────────────────
# EXCEL + AUDIT
# ─────────────────────────────────────────────────────────────────────────────

def build_excel(df: pd.DataFrame) -> bytes:
    """Plain Shopify Excel — Month | Visitors | Orders | CVR + TOTAL row.
    Months without Orders show N/A and are excluded from the TOTAL."""
    wb = Workbook()
    wb.remove(wb.active)
    ws = wb.create_sheet("Traffic Conversion")
    ws.sheet_view.showGridLines = True

    rows = []
    t_vis = t_ord = 0.0
    for _, r in df.iterrows():
        vis = _to_float(r["Visitors"])
        ordr = r["Orders"]
        has_ord = ordr is not None and not pd.isna(ordr)
        if has_ord:
            t_vis += vis
            t_ord += _to_float(ordr)
        dt = r["Date"] if isinstance(r["Date"], datetime) else parse_month(r["Month"])
        month_lbl = dt.strftime("%B %Y") if dt else str(r["Month"])
        rows.append([
            month_lbl,
            f"{int(vis):,}" if vis else "0",
            f"{int(_to_float(ordr)):,}" if has_ord else "N/A",
            _cvr(ordr if has_ord else None, vis),
        ])

    total_row = ["TOTAL", f"{int(t_vis):,}", f"{int(t_ord):,}", _cvr(t_ord, t_vis)]
    write_plain_sheet(
        ws,
        headers=["Month", "Visitors", "Orders", "CVR"],
        rows=rows,
        total_row=total_row,
        col_widths=[14, 14, 12, 12],
    )
    return to_xlsx_bytes(wb)


def verify(df: pd.DataFrame) -> list:
    checks = []
    if df.empty:
        return [{"status": "warn", "message": "No data yet"}]
    for _, r in df.iterrows():
        v = _to_float(r["Visitors"])
        o = r["Orders"]
        if o is None or pd.isna(o):
            checks.append({"status": "warn",
                           "message": f"{r['Month']}: no Orders — upload the Orders export "
                                      f"for this month (shown as N/A, excluded from TOTAL)"})
        elif v <= 0:
            checks.append({"status": "warn",
                           "message": f"{r['Month']}: no Visitors — upload the Visitors export"})
        elif _to_float(o) > v:
            checks.append({"status": "warn",
                           "message": f"{r['Month']}: Orders ({int(_to_float(o)):,}) > Visitors "
                                      f"({int(v):,})? Baka Sessions ang nagamit, hindi Visitors."})
        else:
            checks.append({"status": "ok",
                           "message": f"{r['Month']}: Orders({int(_to_float(o)):,}) ÷ "
                                      f"Visitors({int(v):,}) = {_cvr(o, v)} ✓"})
    checks.append({"status": "ok",
                   "message": "Formula: Orders ÷ Visitors (unified across platforms)"})
    return checks
