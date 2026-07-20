"""
frontend/components.py
All reusable Streamlit UI components for the SOLA CVR Dashboard.
Imports Streamlit but never touches pandas/openpyxl directly —
data comes pre-processed from backend/processor.py.
"""

import pandas as pd
import streamlit as st

from frontend.config import PLATFORMS
from backend.processor import process


# ─────────────────────────────────────────────────────────────────────────────
# HEADER
# ─────────────────────────────────────────────────────────────────────────────

def render_header():
    """Render the top app bar with logo dot, title, and subtitle."""
    st.markdown(
        """
        <div class="app-header">
          <div class="header-dot" style="background:#1E5799"></div>
          <div>
            <p class="header-title">SOLA CVR Dashboard</p>
            <p class="header-sub">Shopee · TikTok · Lazada · Shopify — Traffic Conversion Extractor</p>
          </div>
        </div>
        """,
        unsafe_allow_html=True,
    )


# ─────────────────────────────────────────────────────────────────────────────
# FILE STATUS LIST
# ─────────────────────────────────────────────────────────────────────────────

def render_file_status(valid_files: list, errors: list, all_files: list):
    """
    Show a coloured row for every uploaded file:
      green ✓  — accepted and processed
      red   ✗  — skipped (validation failed)
    """
    skipped_names = {e["name"] for e in errors if e["type"] == "validation"}
    html = '<div class="file-status-wrap">'

    for f in all_files:
        if f.name in skipped_names:
            reason = next((e["reason"] for e in errors if e["name"] == f.name), "")
            html += (
                f'<div class="file-status-row file-skip">'
                f'<div class="file-dot" style="background:#dc2626"></div>'
                f'<span class="file-name-text" title="{reason}">{f.name}</span>'
                f'<span style="font-size:.7rem;opacity:.7">skipped</span>'
                f'</div>'
            )
        else:
            html += (
                f'<div class="file-status-row file-ok">'
                f'<div class="file-dot" style="background:#16a34a"></div>'
                f'<span class="file-name-text">{f.name}</span>'
                f'</div>'
            )

    html += "</div>"
    st.markdown(html, unsafe_allow_html=True)


# ─────────────────────────────────────────────────────────────────────────────
# STATUS BANNERS
# ─────────────────────────────────────────────────────────────────────────────

def render_errors(errors: list):
    """
    Render status banners for all processing messages.
    partial_replaced → blue info  (auto-resolved, non-blocking)
    duplicate        → red error  (blocked, must remove file)
    validation       → yellow warning (file skipped)
    extraction       → red error
    """
    for e in errors:
        if e["type"] == "partial_replaced":
            st.markdown(
                f'<div class="banner banner-info">'
                f'ℹ️ &nbsp;<b>{e["name"]}</b> — {e["reason"]}</div>',
                unsafe_allow_html=True,
            )
        elif e["type"] == "duplicate":
            st.markdown(
                f'<div class="banner banner-error">'
                f'⛔ &nbsp;<b>{e["name"]}</b> — {e["reason"]}</div>',
                unsafe_allow_html=True,
            )
        elif e["type"] == "validation":
            st.markdown(
                f'<div class="banner banner-warning">'
                f'⚠️ &nbsp;<b>{e["name"]}</b> skipped — {e["reason"]}</div>',
                unsafe_allow_html=True,
            )
        elif e["type"] == "extraction":
            st.markdown(
                f'<div class="banner banner-error">'
                f'❌ &nbsp;<b>{e["name"]}</b> extraction failed — {e["reason"]}</div>',
                unsafe_allow_html=True,
            )


def render_success_banner(n_valid: int, n_months: int, n_skipped: int):
    """Render the green success banner after a successful run."""
    skip_txt = (
        f" &nbsp;·&nbsp; {n_skipped} file{'s' if n_skipped > 1 else ''} skipped"
        if n_skipped else ""
    )
    st.markdown(
        f'<div class="banner banner-success">'
        f'✅ &nbsp;<b>{n_valid} file{"s" if n_valid > 1 else ""}</b> processed'
        f' &nbsp;·&nbsp; <b>{n_months} month{"s" if n_months > 1 else ""}</b>'
        f' &nbsp;·&nbsp; sorted oldest → newest{skip_txt}</div>',
        unsafe_allow_html=True,
    )


# ─────────────────────────────────────────────────────────────────────────────
# METRIC CARDS
# ─────────────────────────────────────────────────────────────────────────────

def _parse_num(series: pd.Series) -> pd.Series:
    """Strip commas/% and parse to float."""
    return (
        series
        .astype(str)
        .str.replace(",", "", regex=False)
        .str.replace("%", "", regex=False)
        .apply(pd.to_numeric, errors="coerce")
    )


def _weighted_cvr(preview: pd.DataFrame, cvr_col: str,
                  weight_col: str = "Visitors") -> str:
    """
    Weighted average CVR using Visitors as weights.
    Matches how Shopee/Lazada compute overall CVR (orders ÷ visitors).
    Falls back to simple mean when weight column is unavailable.
    """
    cvr     = _parse_num(preview[cvr_col])
    weights = _parse_num(preview[weight_col]) if weight_col in preview.columns else None

    if weights is not None:
        valid = cvr.notna() & weights.notna() & (weights > 0)
        if valid.sum() > 0:
            result = (cvr[valid] * weights[valid]).sum() / weights[valid].sum()
            return f"{result:.2f}%"

    # Fallback: simple mean
    vals = cvr.dropna()
    return f"{vals.mean():.2f}%" if len(vals) else "—"


def _col_sum(preview: pd.DataFrame, col: str) -> str:
    total = _parse_num(preview[col]).sum()
    return f"{int(total):,}"


def _date_range(preview: pd.DataFrame) -> str:
    """Return 'Jan 2025 – May 2026' from the Month column."""
    try:
        from datetime import datetime
        dates = preview["Month"].apply(
            lambda x: datetime.strptime(str(x).strip(), "%B %Y")
        ).dropna().sort_values()
        if len(dates) == 0:
            return "—"
        first = dates.iloc[0].strftime("%b %Y")
        last  = dates.iloc[-1].strftime("%b %Y")
        return first if first == last else f"{first} – {last}"
    except Exception:
        return "—"


def render_metrics(platform: str, preview: pd.DataFrame, n_months: int):
    """
    Render summary metric cards — all CVR-related values pulled from the data.
    CVR is computed as a visitor-weighted average across months
    (same method Shopee/TikTok/Lazada use internally).
    """
    color    = PLATFORMS[platform]["color"]
    date_rng = _date_range(preview)

    if platform == "shopee":
        # Unified CVR = Paid Orders ÷ Visitors.
        cards = [
            ("Months",         str(n_months),                            ""),
            ("Period",         date_rng,                                 ""),
            ("Total Visitors", _col_sum(preview, "Visitors"),            ""),
            ("Paid Orders",    _col_sum(preview, "Paid Orders"),         ""),
            ("Paid CVR",       _weighted_cvr(preview, "CVR", "Visitors"), "#2E5E0D"),
        ]

    elif platform == "tiktok":
        cards = [
            ("Months",         str(n_months),                              ""),
            ("Period",         date_rng,                                   ""),
            ("Conversion CVR", _weighted_cvr(preview, "CVR"),              color),
            ("Total Orders",   _col_sum(preview, "Orders"),                ""),
            ("Total Visitors", _col_sum(preview, "Visitors"),              ""),
        ]

    elif platform == "shopify":
        cards = [
            ("Months",         str(n_months),                             ""),
            ("Period",         date_rng,                                  ""),
            ("Total Visitors", _col_sum(preview, "Visitors"),             ""),
            ("Total Orders",   _col_sum(preview, "Orders"),               ""),
            ("CVR",            _weighted_cvr(preview, "CVR", "Visitors"),  color),
        ]

    else:  # lazada
        cards = [
            ("Months",         str(n_months),                              ""),
            ("Period",         date_rng,                                   ""),
            ("Conversion CVR", _weighted_cvr(preview, "CVR", "Visitors"),  color),
            ("Total Buyers",   _col_sum(preview, "Buyers"),                ""),
            ("Total Orders",   _col_sum(preview, "Orders"),                ""),
        ]

    html = '<div class="metrics-wrap">'
    for label, val, clr in cards:
        vc = f'style="color:{clr}"' if clr else ""
        html += (
            f'<div class="metric-card">'
            f'<div class="metric-label">{label}</div>'
            f'<div class="metric-value" {vc}>{val}</div>'
            f'</div>'
        )
    html += "</div>"
    st.markdown(html, unsafe_allow_html=True)


# ─────────────────────────────────────────────────────────────────────────────
# PREVIEW TABLE
# ─────────────────────────────────────────────────────────────────────────────

def render_table(platform: str, preview: pd.DataFrame):
    """Render the CVR preview table with a bold TOTAL row at the bottom."""
    color = PLATFORMS[platform]["color"]
    cols  = list(preview.columns)

    def _num(v) -> float:
        try:
            return float(str(v).replace(",", "").replace("%", "").strip())
        except Exception:
            return 0.0

    # Build header cells
    heads = "".join(
        f'<th>{c}</th>' if c == cols[0] else f'<th class="r">{c}</th>'
        for c in cols
    )

    def make_td(val: str, col: str) -> str:
        if col == cols[0]:
            return f"<td>{val}</td>"
        if "CVR" in col or col == "CVR":
            return f'<td class="cvr" style="color:{color}">{val}</td>'
        return f'<td class="r">{val}</td>'

    rows_html = "".join(
        "<tr>" + "".join(make_td(str(row[c]), c) for c in cols) + "</tr>"
        for _, row in preview.iterrows()
    )

    # ── Build TOTAL row ────────────────────────────────────────────────────
    # Sum every numeric column; CVR computed from the right ratio per platform
    totals = {cols[0]: "TOTAL"}
    for c in cols[1:]:
        if "CVR" in c or c == "CVR":
            totals[c] = ""   # filled below
        else:
            totals[c] = f"{int(sum(_num(v) for v in preview[c])):,}"

    # CVR totals — unified formula = ΣOrders ÷ ΣVisitors, computed as a
    # visitor-weighted average of each month's CVR (works for every CVR column,
    # including Shopee's Placed / Confirmed / Paid CVR).
    for cc in [c for c in cols if ("CVR" in c or c == "CVR")]:
        totals[cc] = _weighted_cvr(preview, cc, "Visitors")

    def make_total_td(val: str, col: str) -> str:
        if col == cols[0]:
            return f'<td style="font-weight:700">{val}</td>'
        if "CVR" in col or col == "CVR":
            return f'<td class="r" style="font-weight:700;color:{color}">{val}</td>'
        return f'<td class="r" style="font-weight:700">{val}</td>'

    total_html = (
        '<tr style="border-top:2px solid #cbd5e1;background:#f8fafc">'
        + "".join(make_total_td(str(totals[c]), c) for c in cols)
        + "</tr>"
    )

    st.markdown(
        f'<div class="table-wrap">'
        f'<table class="preview-table">'
        f'<thead style="background:{color}"><tr>{heads}</tr></thead>'
        f'<tbody>{rows_html}{total_html}</tbody>'
        f'</table></div>',
        unsafe_allow_html=True,
    )


# ─────────────────────────────────────────────────────────────────────────────
# FULL TAB RENDERER
# ─────────────────────────────────────────────────────────────────────────────

def render_audit(audit_results: list):
    """
    Render the verification panel — shows per-file CVR check results.
    Green = verified ✓   Yellow = warning   Red = error
    """
    if not audit_results:
        return

    st.markdown("**Data Verification**")
    for result in audit_results:
        fname  = result["file"]
        checks = result["checks"]

        has_err  = any(c["status"] == "error"  for c in checks)
        has_warn = any(c["status"] == "warn"   for c in checks)
        icon     = "❌" if has_err else ("⚠️" if has_warn else "✅")

        with st.expander(f"{icon}  {fname}", expanded=has_err or has_warn):
            for c in checks:
                if c["status"] == "ok":
                    st.markdown(
                        f'<div style="font-size:.82rem;color:var(--color-text-success);'
                        f'padding:3px 0">✅ {c["message"]}</div>',
                        unsafe_allow_html=True,
                    )
                elif c["status"] == "warn":
                    st.markdown(
                        f'<div style="font-size:.82rem;color:#92400e;background:#fffaeb;'
                        f'padding:5px 8px;border-radius:6px;margin:3px 0">'
                        f'⚠️ {c["message"]}</div>',
                        unsafe_allow_html=True,
                    )
                else:
                    st.markdown(
                        f'<div style="font-size:.82rem;color:#991b1b;background:#fdf0ef;'
                        f'padding:5px 8px;border-radius:6px;margin:3px 0">'
                        f'❌ {c["message"]}</div>',
                        unsafe_allow_html=True,
                    )


def render_shopify_tab(platform: str, cfg: dict):
    """
    Shopify flow: merge the two ShopifyQL exports (Visitors file + Orders file,
    or one file with both) BY MONTH → CVR = Orders ÷ Visitors. A manual-entry
    table and a previous-output loader are also merged in.
    """
    from backend import shopify

    st.caption(
        "Upload your **ShopifyQL exports** — the **Visitors** file and the **Orders** "
        "file (or a single file that has both columns). The app merges them by month "
        "into **CVR = Orders ÷ Visitors**. Walang file? Gamitin ang manual entry sa ibaba."
    )

    files = st.file_uploader(
        "Shopify exports (Visitors + Orders)",
        type=["xlsx", "csv"],
        accept_multiple_files=True,
        help=cfg["file_hint"],
        key=f"ul_{platform}",
        label_visibility="collapsed",
    )

    with st.expander("✏️  Or type manually (Online store visitors + Orders)", expanded=not files):
        base = pd.DataFrame({"Month": ["January 2026"], "Visitors": [0], "Orders": [0]})
        edited = st.data_editor(
            base, num_rows="dynamic", use_container_width=True, key=f"editor_{platform}",
            column_config={
                "Month":    st.column_config.TextColumn("Month", help="e.g. January 2026"),
                "Visitors": st.column_config.NumberColumn("Visitors", min_value=0, step=1),
                "Orders":   st.column_config.NumberColumn("Orders", min_value=0, step=1),
            },
        )

    with st.expander("📁  Load previous Shopify output (optional)", expanded=False):
        prev = st.file_uploader(
            "Previous Shopify output", type=["xlsx"],
            key=f"prev_{platform}", label_visibility="collapsed",
        )

    run = st.button("⚡  Compute CVR", key=f"btn_{platform}",
                    type="primary", use_container_width=True)
    if not run:
        return

    sources, file_errs = [], []
    for f in files or []:
        rows, err = shopify.read_file(f)
        if err:
            file_errs.append(err)
        else:
            sources.extend(rows)
    sources.extend(shopify.rows_from_manual(edited.to_dict("records")))
    if prev is not None:
        sources.extend(shopify.rows_from_previous(prev))

    for err in file_errs:
        st.markdown(f'<div class="banner banner-warning">⚠️ &nbsp;{err}</div>',
                    unsafe_allow_html=True)

    df = shopify.merge_sources(sources)
    if df.empty:
        st.markdown(
            '<div class="banner banner-error">❌ &nbsp;No data — upload a Visitors and/or '
            'Orders export, or type at least one month manually.</div>',
            unsafe_allow_html=True,
        )
        return

    excel   = shopify.build_excel(df)
    preview = df[["Month", "Visitors", "Orders", "CVR"]].copy()
    preview["Visitors"] = preview["Visitors"].apply(lambda v: f"{int(_to_num(v)):,}")
    preview["Orders"]   = preview["Orders"].apply(
        lambda v: f"{int(_to_num(v)):,}" if v is not None and not pd.isna(v) else "N/A"
    )

    render_success_banner(len(df), len(df), 0)
    st.markdown("---")
    render_metrics(platform, preview, len(df))
    render_table(platform, preview)
    st.markdown("---")
    render_audit([{"file": "Shopify (merged)", "checks": shopify.verify(df)}])
    st.markdown("---")
    st.download_button(
        label="📥  Download Excel",
        data=excel,
        file_name=cfg["dl_name"],
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        key=f"dl_{platform}",
        use_container_width=True,
    )
    st.caption("💡 Save this Excel and upload it next time as 'previous output' "
               "to keep your history — then just add the new month.")


def _to_num(v) -> float:
    try:
        return float(str(v).replace(",", "").strip())
    except Exception:
        return 0.0


def render_tab(tab, platform: str):
    cfg = PLATFORMS[platform]

    with tab:
        st.markdown(
            f'<div class="platform-badge"'
            f' style="background:{cfg["light"]};color:{cfg["dark_text"]}">'
            f'{cfg["tab_icon"]}  {cfg["label"]} — Traffic Conversion'
            f'</div>',
            unsafe_allow_html=True,
        )

        # ── Custom-input platforms (e.g. Shopify) take a different path ──────
        if cfg.get("input_mode"):
            render_shopify_tab(platform, cfg)
            return

        # ── Previous output loader ─────────────────────────────────────────
        with st.expander("📁  Files", expanded=False):
            st.caption(
                "Upload the Excel file this app generated last time. "
                "Its months are restored as a baseline, then merged with new files below. "
                "This way you never need to re-upload historical source files."
            )
            prev_file = st.file_uploader(
                "Previous output Excel",
                type=["xlsx"],
                key=f"prev_{platform}",
                label_visibility="collapsed",
            )

        # ── New files uploader ─────────────────────────────────────────────
        files = st.file_uploader(
            "Upload new export files",
            type=cfg["file_types"],
            accept_multiple_files=True,
            help=cfg["file_hint"],
            key=f"ul_{platform}",
            label_visibility="collapsed",
        )

        has_prev  = prev_file is not None
        has_files = bool(files)

        if not has_prev and not has_files:
            types_str = ", ".join("." + t for t in cfg["file_types"])
            st.caption(
                f"↑  Drag & drop your {cfg['label']} export files here, "
                f"or click to browse.  Accepted: {types_str}"
            )
            return

        n = len(files) if files else 0
        src_parts = []
        if has_prev:  src_parts.append("previous output")
        if n > 0:     src_parts.append(f"{n} new file{'s' if n > 1 else ''}")

        col_btn, col_info = st.columns([2, 3])
        with col_btn:
            run = st.button(
                f"⚡  Process {' + '.join(src_parts)}",
                key=f"btn_{platform}",
                type="primary",
                use_container_width=True,
            )
        with col_info:
            st.caption(" + ".join(src_parts))

        if not run:
            return

        # ── Load previous output ───────────────────────────────────────────
        prev_extracts = []
        if has_prev:
            from backend.reader import read_previous_output
            prev_data, err = read_previous_output(prev_file, platform)
            if err:
                st.markdown(
                    f'<div class="banner banner-error">'
                    f'❌ &nbsp;Could not read previous output: {err}</div>',
                    unsafe_allow_html=True,
                )
            elif prev_data:
                prev_extracts = [prev_data]
                mon_count = len(prev_data.get("monthly", pd.DataFrame()))
                st.markdown(
                    f'<div class="banner banner-info">'
                    f'📂 &nbsp;<b>Previous output loaded</b> — '
                    f'{mon_count} months restored as baseline</div>',
                    unsafe_allow_html=True,
                )

        # ── Run pipeline ───────────────────────────────────────────────────
        with st.spinner(f"Processing {' + '.join(src_parts)}…"):
            valid, errors, preview, n_months, excel, audit_results = process(
                files or [], platform, prev_extracts=prev_extracts
            )

        render_errors(errors)

        if preview is None:
            has_dup = any(e["type"] == "duplicate" for e in errors)
            if not has_dup:
                st.markdown(
                    '<div class="banner banner-error">'
                    '❌ &nbsp;No valid files could be processed. '
                    'Check the warnings above.</div>',
                    unsafe_allow_html=True,
                )
            return

        render_file_status(valid, errors, files or [])
        render_success_banner(len(valid), n_months, len(files or []) - len(valid))

        st.markdown("---")
        render_metrics(platform, preview, n_months)
        render_table(platform, preview)

        st.markdown("---")
        render_audit(audit_results)

        st.markdown("---")
        st.download_button(
            label="📥  Download Excel",
            data=excel,
            file_name=cfg["dl_name"],
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key=f"dl_{platform}",
            use_container_width=True,
        )

        st.caption(
            "💡 Save this Excel and upload it next time as 'previous output' "
            "— then you only need to add new monthly files."
        )
