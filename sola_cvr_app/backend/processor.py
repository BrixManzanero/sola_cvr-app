"""
backend/processor.py
Central processing engine — validates, extracts, resolves overlaps,
merges, sorts, audits, and builds output.

Smart overlap resolution (replaces simple duplicate blocking):
  For each calendar month covered by multiple files:
    • More _day_count  → winner  (partial auto-replaced, shown as info)
    • Equal _day_count → true duplicate → blocked with error
No Streamlit imports — pure Python only.
"""

from collections import defaultdict
from datetime import datetime

import pandas as pd

from backend import shopee, tiktok, lazada
from backend import auditor
from backend.shopee import parse_date as sh_parse_date, to_month_label as sh_month_label


_VALIDATORS  = {"shopee": shopee.validate,  "tiktok": tiktok.validate,  "lazada": lazada.validate}
_EXTRACTORS  = {"shopee": shopee.extract,   "tiktok": tiktok.extract,   "lazada": lazada.extract}
_XL_BUILDERS = {"shopee": shopee.build_excel, "tiktok": tiktok.build_excel, "lazada": lazada.build_excel}


# ─────────────────────────────────────────────────────────────────────────────
# HELPERS
# ─────────────────────────────────────────────────────────────────────────────

def _get_month_label(platform: str, row: pd.Series) -> str:
    if platform == "shopee":
        return sh_month_label(row["Date"])
    return str(row.get("Month Label", row.get("Date", "")))


def _get_rows(data: dict, platform: str):
    """Yield (month_label, row_series) for every monthly row in an extract."""
    if platform == "lazada":
        row = pd.Series(data["row"])
        yield _get_month_label(platform, row), row
    else:
        for _, row in data["monthly"].iterrows():
            yield _get_month_label(platform, row), row


# ─────────────────────────────────────────────────────────────────────────────
# SMART OVERLAP RESOLVER
# ─────────────────────────────────────────────────────────────────────────────

def _to_dt(val):
    """Convert any date-like value to datetime."""
    if isinstance(val, datetime):
        return val
    try:
        return pd.Timestamp(val).to_pydatetime()
    except Exception:
        return datetime.min


def _ranges_overlap(s_a, e_a, s_b, e_b) -> bool:
    """Return True if date ranges [s_a, e_a] and [s_b, e_b] share any days."""
    try:
        return _to_dt(s_a) <= _to_dt(e_b) and _to_dt(s_b) <= _to_dt(e_a)
    except Exception:
        return True   # assume overlap if dates can't be compared


def _parse_num(val) -> float:
    """Parse a formatted number string like '110,259' to float."""
    try:
        return float(str(val).replace(",", "").strip())
    except Exception:
        return 0.0


def _parse_cvr(val) -> float:
    """Parse a CVR string like '2.16%' to float (e.g. 2.16)."""
    try:
        v = float(str(val).replace("%", "").strip())
        return v * 100 if v < 1 else v
    except Exception:
        return 0.0


def _combine_lazada(entries: list) -> pd.Series:
    """Combine non-overlapping Lazada rows using summed counts and recomputed CVR."""
    total_pageviews = sum(_parse_num(r.get("Pageviews", 0)) for _, r in entries)
    total_visitors  = sum(_parse_num(r.get("Visitors",  0)) for _, r in entries)
    total_buyers    = sum(_parse_num(r.get("Buyers",    0)) for _, r in entries)
    total_orders    = sum(_parse_num(r.get("Orders",    0)) for _, r in entries)
    # Unified CVR = Orders ÷ Visitors
    cvr = f"{total_orders / total_visitors * 100:.2f}%" if total_visitors > 0 else "N/A"
    combined = dict(entries[0][1])
    combined["Pageviews"]       = f"{int(total_pageviews):,}"
    combined["Visitors"]        = f"{int(total_visitors):,}"
    combined["Buyers"]          = f"{int(total_buyers):,}"
    combined["Orders"]          = f"{int(total_orders):,}"
    combined["Conversion Rate"] = cvr
    combined["_day_count"]      = sum(int(r.get("_day_count", 0)) for _, r in entries)
    combined["_period_start"]   = min(_to_dt(r.get("_period_start", datetime.min)) for _, r in entries)
    combined["_period_end"]     = max(_to_dt(r.get("_period_end",   datetime.min)) for _, r in entries)
    combined["Date"]            = combined["_period_start"]
    return pd.Series(combined)


def _merge_shopee(extracts: list) -> tuple:
    """
    Shopee-specific merge: Visitors come from Overview files, Orders come from
    Traffic files. For each month we take Visitors from whichever file has it and
    the per-status Orders from whichever file has them, then compute
    CVR = Orders ÷ Visitors for placed / confirmed / paid.

    Returns (resolved_rows, info_msgs, warn_msgs, winning_names).
    """
    def _has_val(v):
        return v is not None and str(v).strip() not in ("", "None", "nan", "—", "N/A")

    by_month = defaultdict(list)   # label -> [(src, row), ...]
    for data in extracts:
        src = data["name"]
        mdf = data.get("monthly")
        if mdf is None:
            continue
        for _, row in mdf.iterrows():
            label = str(row.get("Month Label") or sh_month_label(row.get("Date")))
            by_month[label].append((src, row))

    resolved, info_msgs, warn_msgs, winners = [], [], [], set()

    for label, entries in by_month.items():
        visitors, vis_src, date_val = None, None, None
        orders  = {"Placed Orders": None, "Confirmed Orders": None, "Paid Orders": None}
        ord_src = None

        for src, row in entries:
            if date_val is None and _has_val(row.get("Date")):
                date_val = row.get("Date")
            if visitors is None and _has_val(row.get("Visitors")):
                visitors, vis_src = row.get("Visitors"), src
            for oc in orders:
                if orders[oc] is None and _has_val(row.get(oc)):
                    orders[oc], ord_src = row.get(oc), src

        merged = {"Date": date_val, "Month Label": label, "Visitors": visitors, **orders}
        merged["Placed CVR"]    = shopee.cvr_str(orders["Placed Orders"],    visitors)
        merged["Confirmed CVR"] = shopee.cvr_str(orders["Confirmed Orders"], visitors)
        merged["Paid CVR"]      = shopee.cvr_str(orders["Paid Orders"],      visitors)
        resolved.append(pd.Series(merged))

        if vis_src:
            winners.add(vis_src)
        if ord_src:
            winners.add(ord_src)

        if visitors is None:
            warn_msgs.append({"month": label, "kind": "missing_visitors"})
        elif all(orders[o] is None for o in orders):
            warn_msgs.append({"month": label, "kind": "missing_orders"})
        elif vis_src and ord_src and vis_src != ord_src:
            info_msgs.append({"month": label, "vis_src": vis_src, "ord_src": ord_src})

    def _key(s):
        dt = sh_parse_date(s.get("Date"))
        return dt if dt else datetime.min

    resolved.sort(key=_key)
    return resolved, info_msgs, warn_msgs, winners


def resolve_conflicts(extracts: list, platform: str) -> tuple:
    """
    For every calendar month covered by multiple files, apply smart resolution:

    OVERLAPPING ranges (e.g. May 1–25 vs May 1–31):
      → Keep the more complete one (more days). Discard the partial (blue info).
      → If equal days: true duplicate → blocked (red error).

    NON-OVERLAPPING / COMPLEMENTARY (e.g. May 1–7 vs May 8–31):
      → COMBINE into one row using visitor-weighted CVR.
      → Both files contribute; result is accurate for the full month.

    Returns
    -------
    resolved_rows  : list of pd.Series — one row per month, sorted
    info_messages  : list of dicts  type="partial_replaced" | "combined"
    error_messages : list of dicts  type="duplicate"
    winning_names  : set of filenames that contributed ≥ 1 winning row
    """
    all_entries = []
    for data in extracts:
        src = data["name"]
        for label, row in _get_rows(data, platform):
            all_entries.append((label, src, row))

    by_month = defaultdict(list)
    for label, src, row in all_entries:
        by_month[label].append((src, row))

    resolved_rows  = []
    info_messages  = []
    error_messages = []
    winning_names  = set()

    for label, entries in by_month.items():
        if len(entries) == 1:
            src, row = entries[0]
            resolved_rows.append(row)
            winning_names.add(src)
            continue

        # ── Classify relationships between all pairs ──────────────────────
        has_overlap       = False
        all_non_overlapping = True

        for i, (_, r_a) in enumerate(entries):
            for _, r_b in entries[i+1:]:
                s_a = _to_dt(r_a.get("_period_start", datetime.min))
                e_a = _to_dt(r_a.get("_period_end",   datetime.min))
                s_b = _to_dt(r_b.get("_period_start", datetime.min))
                e_b = _to_dt(r_b.get("_period_end",   datetime.min))
                if _ranges_overlap(s_a, e_a, s_b, e_b):
                    has_overlap = True
                    all_non_overlapping = False

        # ── Case A: ALL non-overlapping → COMBINE ────────────────────────
        # NOTE: Shopee never reaches resolve_conflicts — it uses _merge_shopee()
        # in process(). Only TikTok / Lazada flow through here.
        if all_non_overlapping:
            if platform == "lazada":
                combined_row = _combine_lazada(entries)
            else:
                # TikTok: non-overlapping monthly groups merged at daily level
                # (handled upstream, shouldn't occur here)
                combined_row = entries[0][1]   # fallback: use first

            total_days = sum(int(r.get("_day_count", 0)) for _, r in entries)
            resolved_rows.append(combined_row)
            for src, _ in entries:
                winning_names.add(src)
            info_messages.append({
                "type":   "combined",
                "month":  label,
                "files":  [src for src, _ in entries],
                "total_days": total_days,
            })
            continue

        # ── Case B: OVERLAPPING → pick by priority ───────────────────────
        # Priority: (1) new source file > previous output  (2) more days > fewer days
        def _pri(r):
            is_out  = bool(r.get("_from_output", False))
            day_cnt = int(r.get("_day_count", 0))
            return (0 if is_out else 1, day_cnt)

        entries_sorted = sorted(entries, key=lambda x: _pri(x[1]), reverse=True)
        winner_src, winner_row = entries_sorted[0]
        best_pri = _pri(winner_row)

        for src, row in entries_sorted[1:]:
            pri     = _pri(row)
            is_out  = bool(row.get("_from_output", False))
            day_cnt = int(row.get("_day_count", 0))
            w_cnt   = int(winner_row.get("_day_count", 0))

            if pri == best_pri and not is_out:
                # Same type, same days → true duplicate
                error_messages.append({
                    "type":   "duplicate",
                    "month":  label,
                    "winner": winner_src,
                    "loser":  src,
                })
            else:
                # Either previous output displaced by new source, or smaller partial replaced
                info_messages.append({
                    "type":          "partial_replaced",
                    "month":         label,
                    "partial_file":  src,
                    "partial_days":  day_cnt,
                    "complete_file": winner_src,
                    "complete_days": w_cnt,
                })

        resolved_rows.append(winner_row)
        winning_names.add(winner_src)

    # Sort chronologically
    def _sort_key(row):
        if platform == "shopee":
            dt = sh_parse_date(row.get("Date"))
            return dt if dt else datetime.min
        try:
            d = row.get("Date")
            return pd.to_datetime(d) if not isinstance(d, datetime) else d
        except Exception:
            return datetime.min

    resolved_rows.sort(key=_sort_key)
    return resolved_rows, info_messages, error_messages, winning_names


# ─────────────────────────────────────────────────────────────────────────────
# MAIN PIPELINE
# ─────────────────────────────────────────────────────────────────────────────

def process(files: list, platform: str, prev_extracts: list = None) -> tuple:
    """
    Full pipeline:
      0. Restore    — include previous output data as baseline extracts
      1. Validate   — reject wrong-platform / malformed files
      2. Extract    — pull CVR metrics using platform-specific logic
      3. Audit      — verify computed CVR against source numbers
      4. Resolve    — smart overlap: partial auto-replaced, true dupes blocked
      5. Merge+sort — combine all winning rows chronologically
      6. Build Excel— format and return as bytes

    Returns
    -------
    valid         : list of accepted UploadedFile objects
    errors        : list of {name, reason, type}
                    types: validation | extraction | partial_replaced | duplicate
    preview       : pd.DataFrame for display  (None on failure)
    n_months      : int  (None on failure)
    excel         : bytes  (None on failure)
    audit_results : list of {file, checks}
    """
    valid         = []
    errors        = []
    extracts      = list(prev_extracts) if prev_extracts else []
    audit_results = []

    # ── Step 1: Validate ──────────────────────────────────────────────────
    for f in files:
        ok, result = _VALIDATORS[platform](f)
        if ok:    valid.append(f)
        else:     errors.append({"name": f.name, "reason": result, "type": "validation"})

    if not valid and not extracts:
        return [], errors, None, None, None, []

    # ── Step 2: Extract ───────────────────────────────────────────────────
    for f in valid:
        try:
            data = _EXTRACTORS[platform](f)
            extracts.append(data)

            # ── Step 3: Audit ─────────────────────────────────────────────
            if platform == "shopee":
                checks = auditor.verify_shopee(data)
            elif platform == "tiktok":
                checks = auditor.verify_tiktok(data, data.get("source_cvr_raw", 0))
                for w in data.get("cvr_warnings", []):
                    checks.insert(0, auditor._warn(w))
            else:
                checks = auditor.verify_lazada(data)

            audit_results.append({"file": f.name, "checks": checks})

        except Exception as exc:
            errors.append({"name": f.name, "reason": str(exc), "type": "extraction"})

    # Skip-out only if nothing at all to work with
    if not extracts and not valid:
        return valid, errors, None, None, None, audit_results

    # ── Step 4: Smart overlap resolution ─────────────────────────────────
    if platform == "shopee":
        # Shopee merges Visitors (Overview) + Orders (Traffic) per month.
        resolved_rows, sh_info, sh_warn, winning_names = _merge_shopee(extracts)
        dup_errors = []
        for m in sh_info:
            errors.append({
                "name":   f"{m['month']} — merged",
                "reason": (f"Visitors from '{m['vis_src']}' + Orders from "
                           f"'{m['ord_src']}' combined into one month"),
                "type":   "partial_replaced",   # blue info
            })
        for w in sh_warn:
            if w["kind"] == "missing_visitors":
                errors.append({
                    "name":   w["month"],
                    "reason": ("No Visitors found — upload the Shopee OVERVIEW file "
                               "for this month so CVR can be computed"),
                    "type":   "validation",
                })
            else:
                errors.append({
                    "name":   w["month"],
                    "reason": ("No Orders found — upload the Shopee TRAFFIC file "
                               "for this month so CVR can be computed"),
                    "type":   "validation",
                })
        if not resolved_rows:
            return valid, errors, None, None, None, audit_results

        combined = pd.DataFrame(resolved_rows).reset_index(drop=True)
        summaries = sorted(
            [{"start_date": d["start_date"], "summary": d["summary"], "name": d["name"]}
             for d in extracts if d["name"] in winning_names],
            key=lambda x: x["start_date"],
        )
        excel = _XL_BUILDERS[platform](combined, summaries)

        preview = combined.copy()
        preview["Month"]  = preview["Date"].apply(sh_month_label)
        preview["CVR"]    = preview["Paid CVR"]
        # 'Orders' column (value = Paid Orders) — matches the downloaded Excel header
        preview["Orders"] = preview["Paid Orders"]
        preview = preview[["Month", "Visitors", "Orders", "CVR"]].copy()
        preview["Orders"] = preview["Orders"].apply(
            lambda v: f"{int(round(float(str(v).replace(',', '')))):,}"
            if v is not None and str(v).strip() not in ("", "None", "nan", "N/A") else "N/A"
        )
        preview["Visitors"] = preview["Visitors"].apply(
            lambda v: f"{int(float(str(v).replace(',', ''))):,}"
            if v is not None and str(v).strip() not in ("", "None", "nan") else "0"
        )
        return valid, errors, preview, len(combined), excel, audit_results

    resolved_rows, info_msgs, dup_errors, winning_names = resolve_conflicts(extracts, platform)

    # Surface info messages
    for m in info_msgs:
        if m["type"] == "partial_replaced":
            errors.append({
                "name":   m["partial_file"],
                "reason": (f"Partial data ({m['partial_days']} days) auto-replaced by "
                           f"'{m['complete_file']}' ({m['complete_days']} days) "
                           f"for {m['month']}"),
                "type": "partial_replaced",
            })
        elif m["type"] == "combined":
            files_str = " + ".join(f"'{f}'" for f in m["files"])
            errors.append({
                "name":   f"{m['month']} — combined",
                "reason": (f"Non-overlapping files merged into one complete month "
                           f"({m['total_days']} days total): {files_str}"),
                "type": "partial_replaced",   # show as blue info
            })

    # Surface true duplicate errors (red — block)
    for d in dup_errors:
        errors.append({
            "name":   "Duplicate months",
            "reason": (f"{d['month']} appears in both '{d['winner']}' and "
                       f"'{d['loser']}' with identical coverage — remove one file"),
            "type": "duplicate",
        })

    if dup_errors:
        return valid, errors, None, None, None, audit_results

    if not resolved_rows:
        return valid, errors, None, None, None, audit_results

    # ── Step 5: Build combined DataFrame from winning rows ────────────────
    combined = pd.DataFrame(resolved_rows).reset_index(drop=True)

    # Summaries: only files that contributed at least one winning row
    summaries = sorted(
        [{"start_date": d["start_date"], "summary": d["summary"], "name": d["name"]}
         for d in extracts if d["name"] in winning_names],
        key=lambda x: x["start_date"],
    )

    # ── Step 6: Build Excel ───────────────────────────────────────────────
    excel = _XL_BUILDERS[platform](combined, summaries)

    # ── Build preview DataFrame ───────────────────────────────────────────
    # (Shopee is handled and returned earlier — see the shopee branch above.)
    if platform == "tiktok":
        has_cust = "Customers" in combined.columns
        cols_sel = ["Month Label", "Page Views", "Visitors", "Product Clicks", "Orders"]
        new_names = ["Month", "Page Views", "Visitors", "Clicks", "Orders"]
        if has_cust:
            cols_sel.append("Customers")
            new_names.append("Customers")
        cols_sel.append("Conversion Rate")
        new_names.append("CVR")
        preview = combined[cols_sel].copy()
        preview.columns = new_names

    else:
        preview = combined[[
            "Month Label", "Pageviews", "Visitors",
            "Buyers", "Orders", "Conversion Rate",
        ]].copy()
        preview.columns = ["Month", "Pageviews", "Visitors", "Buyers", "Orders", "CVR"]

    return valid, errors, preview, len(combined), excel, audit_results
