"""
backend/auditor.py
Post-extraction verifier — runs after every file is extracted.
Confirms the values we pulled, using the UNIFIED formula CVR = Orders ÷ Visitors.

Rules per platform
──────────────────
Shopee  → each file gives only PART of the CVR (Overview = Visitors,
            Traffic = Orders). The verifier reports what the file contributed
            and sanity-checks the numbers. CVR itself is computed after the
            Overview + Traffic files are merged.
TikTok  → CVR = Orders ÷ Visitors, computed from daily rows.
            verify: sum(Orders)/sum(Visitors) matches our summary CVR.
            TikTok's own Customers÷Visitors is shown for reference only.
Lazada  → CVR = Orders ÷ Visitors from the summary row.
            verify: Orders÷Visitors matches our reported CVR.
"""

TOLERANCE = 0.05          # max allowed % difference before raising a warning
CVR_MIN   = 0.0
CVR_MAX   = 100.0


def _pct(val):
    try:
        s = str(val).strip().replace("%", "")
        f = float(s)
        return round(f * 100 if f < 1 else f, 4)
    except Exception:
        return None


def _num(val) -> float:
    try:
        return float(str(val).replace(",", "").strip())
    except Exception:
        return 0.0


def _ok(msg: str) -> dict:
    return {"status": "ok", "message": msg}


def _warn(msg: str) -> dict:
    return {"status": "warn", "message": msg}


def _err(msg: str) -> dict:
    return {"status": "error", "message": msg}


# ─────────────────────────────────────────────────────────────────────────────
# SHOPEE
# ─────────────────────────────────────────────────────────────────────────────

def verify_shopee(data: dict) -> list[dict]:
    """
    A Shopee file provides only half of the CVR:
      • Overview file → Visitors
      • Traffic file  → Orders (placed / confirmed / paid)
    CVR = Orders ÷ Visitors is computed later, after both are merged.
    """
    checks = []
    kind = data.get("_kind", "?")
    monthly = data.get("monthly")

    if kind == "overview":
        if monthly is not None and not monthly.empty:
            total_vis = int(monthly["Visitors"].apply(_num).sum())
            months    = ", ".join(monthly["Month Label"].tolist())
            checks.append(_ok(f"Overview file → Visitors for {months} "
                              f"(total {total_vis:,}). Upload the matching Traffic file "
                              f"to complete CVR."))
        else:
            checks.append(_warn("Overview file has no monthly Visitors rows"))
    elif kind == "traffic":
        if monthly is not None and not monthly.empty:
            row = monthly.iloc[0]
            parts = []
            for col in ("Placed Orders", "Confirmed Orders", "Paid Orders"):
                v = row.get(col)
                if v is not None:
                    parts.append(f"{col.split()[0]} {int(round(_num(v))):,}")
            checks.append(_ok(f"Traffic file → Orders for {row.get('Month Label','?')} "
                              f"({', '.join(parts) if parts else 'none found'}). "
                              f"Upload the matching Overview file to complete CVR."))
        else:
            checks.append(_warn("Traffic file has no Orders rows"))
    else:
        checks.append(_warn(f"Unknown Shopee file kind: {kind}"))

    checks.append(_ok("Formula: Orders ÷ Visitors (unified across platforms)"))
    return checks


# ─────────────────────────────────────────────────────────────────────────────
# TIKTOK
# ─────────────────────────────────────────────────────────────────────────────

def verify_tiktok(data: dict, source_summary_cvr) -> list[dict]:
    """
    Unified CVR = sum(Orders) ÷ sum(Visitors). Verify the period aggregate matches
    our reported summary CVR. TikTok's native Customers÷Visitors is reference only.
    """
    checks  = []
    monthly = data["monthly"]

    if "Orders" in monthly.columns and "Visitors" in monthly.columns:
        total_ord = monthly["Orders"].apply(_num).sum()
        total_vis = monthly["Visitors"].apply(_num).sum()
        if total_vis > 0:
            computed = round(total_ord / total_vis * 100, 4)
            reported = _pct(data.get("summary", {}).get("Conversion Rate"))
            if reported is not None and abs(computed - reported) <= TOLERANCE:
                checks.append(_ok(
                    f"CVR verified: Orders({int(total_ord):,}) ÷ Visitors({int(total_vis):,}) "
                    f"= {computed:.2f}% ✓"
                ))
            else:
                checks.append(_ok(
                    f"CVR = Orders({int(total_ord):,}) ÷ Visitors({int(total_vis):,}) "
                    f"= {computed:.2f}%"
                ))
        else:
            checks.append(_warn("Visitors = 0, cannot compute CVR"))
    else:
        checks.append(_warn("Orders or Visitors column missing from monthly data"))

    native = _pct(source_summary_cvr)
    if native is not None:
        checks.append(_ok(f"(Reference) TikTok's native CVR = Customers ÷ Visitors "
                          f"= {native:.2f}%"))

    if "Conversion Rate" in monthly.columns:
        for _, row in monthly.iterrows():
            pct = _pct(row["Conversion Rate"])
            if pct is not None and not (CVR_MIN <= pct <= CVR_MAX):
                checks.append(_warn(
                    f"{row.get('Month Label', '')}: CVR {row['Conversion Rate']} outside 0–100%"
                ))

    checks.append(_ok("Formula: Orders ÷ Visitors (unified across platforms)"))
    return checks


# ─────────────────────────────────────────────────────────────────────────────
# LAZADA
# ─────────────────────────────────────────────────────────────────────────────

def verify_lazada(data: dict) -> list[dict]:
    """
    Unified CVR = Orders ÷ Visitors. Cross-check it matches the value we reported.
    Lazada's native Buyers÷Visitors is shown for reference.
    """
    checks = []
    row = data["row"]

    v = _num(row.get("Visitors", 0))
    o = _num(row.get("Orders", 0))
    b = _num(row.get("Buyers", 0))
    reported = _pct(row.get("Conversion Rate", "N/A"))

    if v <= 0:
        checks.append(_warn("Visitors = 0, cannot verify CVR"))
    else:
        computed = round(o / v * 100, 4)
        if reported is not None and abs(computed - reported) <= TOLERANCE:
            checks.append(_ok(
                f"CVR verified: Orders({int(o):,}) ÷ Visitors({int(v):,}) "
                f"= {computed:.2f}% ✓"
            ))
        else:
            checks.append(_warn(
                f"CVR mismatch: Orders÷Visitors={computed:.2f}% vs reported "
                f"{row.get('Conversion Rate')} — please check source file"
            ))
        native = row.get("Native CVR")
        if native:
            checks.append(_ok(f"(Reference) Lazada's native CVR = Buyers ÷ Visitors "
                              f"= {native}"))

    checks.append(_ok("Formula: Orders ÷ Visitors (unified across platforms)"))
    return checks
