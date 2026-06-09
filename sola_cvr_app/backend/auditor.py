"""
backend/auditor.py
Post-extraction verifier — runs after every file is extracted.
Compares our computed values against the platform's own source numbers
and returns structured check results shown in the UI.

Rules per platform
──────────────────
Shopee  → CVR read directly as a string ('2.66%') — no computation → always exact
TikTok  → CVR computed as Customers÷Visitors from daily rows
            verify: sum(Customers)/sum(Visitors) for the period
            must match TikTok's own summary CVR within TOLERANCE
Lazada  → CVR read directly as a string ('3.36%') — no computation → always exact
           cross-check: Buyers÷Visitors must equal CVR string within TOLERANCE
"""

TOLERANCE = 0.05          # max allowed % difference before raising a warning
CVR_MIN   = 0.0           # sanity: CVR must be ≥ 0%
CVR_MAX   = 100.0         # sanity: CVR must be ≤ 100%


def _pct(val) -> float | None:
    """Parse any CVR representation to a plain float percentage."""
    try:
        s = str(val).strip().replace("%", "")
        f = float(s)
        return round(f * 100 if f < 1 else f, 4)
    except Exception:
        return None


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
    Shopee CVR is read directly from the source string — never computed.
    Checks: CVR strings are parseable, within sane range, no N/A values.
    """
    checks = []
    summary = data["summary"]

    for col in ("Placed Order CVR", "Confirmed Order CVR", "Paid Order CVR"):
        val = summary.get(col, "N/A")
        if val == "N/A":
            checks.append(_warn(f"{col} is N/A — may be missing in source file"))
            continue
        pct = _pct(val)
        if pct is None:
            checks.append(_err(f"{col}: cannot parse '{val}' as a percentage"))
        elif not (CVR_MIN <= pct <= CVR_MAX):
            checks.append(_warn(f"{col}: value {val} is outside expected range 0–100%"))
        else:
            checks.append(_ok(f"{col}: {val} — read directly from source ✓"))

    # Verify ordering: Placed ≥ Confirmed ≥ Paid (Shopee typical pattern)
    p = _pct(summary.get("Placed Order CVR", "N/A"))
    c = _pct(summary.get("Confirmed Order CVR", "N/A"))
    d = _pct(summary.get("Paid Order CVR", "N/A"))
    if all(v is not None for v in (p, c, d)):
        if not (p >= c >= d):
            checks.append(_warn(
                f"CVR ordering unexpected: Placed({p:.2f}%) Confirmed({c:.2f}%) Paid({d:.2f}%). "
                f"Typically Placed ≥ Confirmed ≥ Paid — please cross-check with Shopee."
            ))

    return checks


# ─────────────────────────────────────────────────────────────────────────────
# TIKTOK
# ─────────────────────────────────────────────────────────────────────────────

def verify_tiktok(data: dict, source_summary_cvr) -> list[dict]:
    """
    TikTok CVR is computed as sum(Customers)/sum(Visitors) per month.
    Verify: the period-level aggregate matches TikTok's own summary CVR.
    """
    checks = []
    monthly = data["monthly"]

    # Parse TikTok's official summary CVR
    official = _pct(source_summary_cvr)
    if official is None:
        checks.append(_warn(f"Cannot parse TikTok summary CVR: '{source_summary_cvr}'"))
        return checks

    # Compute overall CVR from monthly data
    import pandas as pd

    def to_num(series):
        return (
            series.astype(str)
            .str.replace(",", "", regex=False)
            .apply(pd.to_numeric, errors="coerce")
            .fillna(0)
        )

    # Monthly DataFrame has Customers and Visitors embedded in the extraction
    # We verify via the stored monthly CVR values weighted by Visitors
    if "Customers" in monthly.columns and "Visitors" in monthly.columns:
        total_cust = to_num(monthly["Customers"]).sum()
        total_vis  = to_num(monthly["Visitors"]).sum()
        if total_vis > 0:
            computed = round(total_cust / total_vis * 100, 4)
            diff     = abs(computed - official)
            if diff <= TOLERANCE:
                checks.append(_ok(
                    f"CVR verified: computed {computed:.2f}% vs TikTok official {official:.2f}% "
                    f"(diff {diff:.3f}% ≤ tolerance {TOLERANCE}%)"
                ))
            else:
                checks.append(_warn(
                    f"CVR diff {diff:.3f}% > tolerance {TOLERANCE}%: "
                    f"computed {computed:.2f}% vs TikTok official {official:.2f}%. "
                    f"This may indicate a format change — please verify."
                ))
    else:
        checks.append(_warn(
            "Cannot cross-check CVR: 'Customers' or 'Visitors' column missing from monthly data"
        ))

    # Sanity check individual month CVRs
    if "Conversion Rate" in monthly.columns:
        for _, row in monthly.iterrows():
            pct = _pct(row["Conversion Rate"])
            if pct is not None and not (CVR_MIN <= pct <= CVR_MAX):
                checks.append(_warn(
                    f"{row.get('Month Label', '')}: CVR {row['Conversion Rate']} is outside 0–100%"
                ))

    checks.append(_ok("Formula: Customers ÷ Visitors (TikTok's official definition)"))
    return checks


# ─────────────────────────────────────────────────────────────────────────────
# LAZADA
# ─────────────────────────────────────────────────────────────────────────────

def verify_lazada(data: dict) -> list[dict]:
    """
    Lazada CVR is read directly from the source string — never computed.
    Cross-check: Buyers ÷ Visitors should equal the CVR string within tolerance.
    """
    checks = []
    row = data["row"]

    cvr_str  = str(row.get("Conversion Rate", "N/A"))
    visitors = str(row.get("Visitors", "0")).replace(",", "")
    buyers   = str(row.get("Buyers", "0")).replace(",", "")

    try:
        v = int(float(visitors))
        b = int(float(buyers))
        cvr_source   = _pct(cvr_str)
        cvr_computed = round(b / v * 100, 4) if v > 0 else None

        if cvr_source is None:
            checks.append(_err(f"Cannot parse CVR string: '{cvr_str}'"))
        elif cvr_computed is None:
            checks.append(_warn("Visitors = 0, cannot verify CVR formula"))
        else:
            diff = abs(cvr_computed - cvr_source)
            if diff <= TOLERANCE:
                checks.append(_ok(
                    f"CVR verified: Buyers({b:,}) ÷ Visitors({v:,}) = {cvr_computed:.2f}% "
                    f"matches source {cvr_str} ✓"
                ))
            else:
                checks.append(_warn(
                    f"CVR mismatch: Buyers÷Visitors={cvr_computed:.2f}% "
                    f"vs source={cvr_str} (diff={diff:.3f}%) — please check source file"
                ))
    except Exception as exc:
        checks.append(_warn(f"Lazada CVR verification error: {exc}"))

    checks.append(_ok("Formula: Buyers ÷ Visitors (Lazada's official definition)"))
    return checks
