"""
frontend/styles.py
All CSS for the SOLA CVR Dashboard.
Import APP_CSS and inject it once at the top of app.py.
"""

APP_CSS: str = """
<style>

/* ── Base layout ─────────────────────────────────────────────────────────── */
.block-container {
    padding: 1.5rem 1.25rem 3rem;
    max-width: 1200px;
}

/* ── App header ───────────────────────────────────────────────────────────── */
.app-header {
    display: flex;
    align-items: center;
    gap: 12px;
    padding-bottom: 1.25rem;
    border-bottom: 1px solid rgba(128,128,128,.15);
    margin-bottom: 1.25rem;
}
.header-dot {
    width: 10px;
    height: 10px;
    border-radius: 50%;
    flex-shrink: 0;
}
.header-title {
    font-size: 1.25rem;
    font-weight: 600;
    margin: 0;
    line-height: 1.2;
}
.header-sub {
    font-size: 0.8rem;
    opacity: .55;
    margin: 0;
    margin-top: 1px;
}

/* ── Platform badge ───────────────────────────────────────────────────────── */
.platform-badge {
    display: inline-flex;
    align-items: center;
    gap: 6px;
    padding: 4px 12px;
    border-radius: 999px;
    font-size: 0.75rem;
    font-weight: 600;
    letter-spacing: .03em;
    margin-bottom: 1rem;
}

/* ── Metric cards ─────────────────────────────────────────────────────────── */
.metrics-wrap {
    display: grid;
    grid-template-columns: repeat(auto-fit, minmax(130px, 1fr));
    gap: 10px;
    margin-bottom: 1.25rem;
}
.metric-card {
    background: rgba(128,128,128,.06);
    border-radius: 10px;
    padding: 12px 14px;
}
.metric-label {
    font-size: 0.7rem;
    font-weight: 500;
    letter-spacing: .04em;
    text-transform: uppercase;
    opacity: .5;
    margin-bottom: 4px;
}
.metric-value {
    font-size: 1.5rem;
    font-weight: 600;
    line-height: 1;
}

/* ── File status list ─────────────────────────────────────────────────────── */
.file-status-wrap {
    display: flex;
    flex-direction: column;
    gap: 6px;
    margin-bottom: 1rem;
}
.file-status-row {
    display: flex;
    align-items: center;
    gap: 8px;
    padding: 7px 10px;
    border-radius: 8px;
    font-size: 0.8rem;
    font-weight: 500;
    border: 1px solid transparent;
}
.file-ok   { background: #edf7ee; border-color: #b7deba; color: #1c5e22; }
.file-skip { background: #fdf0ef; border-color: #f5bcba; color: #991b1b; }
.file-dot  { width: 7px; height: 7px; border-radius: 50%; flex-shrink: 0; }
.file-name-text {
    flex: 1;
    overflow: hidden;
    text-overflow: ellipsis;
    white-space: nowrap;
}

/* ── Status banners ───────────────────────────────────────────────────────── */
.banner {
    display: flex;
    align-items: flex-start;
    gap: 10px;
    padding: 10px 14px;
    border-radius: 10px;
    font-size: 0.82rem;
    font-weight: 500;
    margin-bottom: 1rem;
    line-height: 1.5;
    border: 1px solid transparent;
}
.banner-success { background: #edf7ee; border-color: #b7deba; color: #1c5e22; }
.banner-error   { background: #fdf0ef; border-color: #f5bcba; color: #991b1b; }
.banner-warning { background: #fffaeb; border-color: #fde68a; color: #92400e; }
.banner-info    { background: #eff6ff; border-color: #bfdbfe; color: #1e40af; }

/* ── Preview table ────────────────────────────────────────────────────────── */
.table-wrap {
    overflow-x: auto;
    border-radius: 10px;
    border: 1px solid rgba(128,128,128,.15);
    margin-bottom: 1rem;
}
.preview-table {
    width: 100%;
    border-collapse: collapse;
    font-size: 0.8rem;
    min-width: 520px;
}
.preview-table thead th {
    padding: 9px 12px;
    text-align: left;
    font-weight: 600;
    font-size: 0.7rem;
    letter-spacing: .03em;
    white-space: nowrap;
    color: #fff;
}
.preview-table thead th.r { text-align: right; }
.preview-table tbody tr:nth-child(even) { background: rgba(128,128,128,.04); }
.preview-table tbody tr:hover           { background: rgba(128,128,128,.08); }
.preview-table tbody td {
    padding: 7px 12px;
    border-top: 1px solid rgba(128,128,128,.09);
    font-size: 0.8rem;
    white-space: nowrap;
}
.preview-table tbody td.r   { text-align: right; opacity: .7; font-variant-numeric: tabular-nums; }
.preview-table tbody td.cvr { text-align: right; font-weight: 600; font-variant-numeric: tabular-nums; }

/* ── Download button ──────────────────────────────────────────────────────── */
div[data-testid="stDownloadButton"] button {
    width: 100%;
    border-radius: 10px !important;
    font-weight: 600 !important;
    font-size: 0.88rem !important;
    padding: 0.6rem 1rem !important;
}

/* ── File uploader ────────────────────────────────────────────────────────── */
div[data-testid="stFileUploader"] section {
    border-radius: 10px !important;
    border-width: 2px !important;
    border-style: dashed !important;
}

/* ── Tabs ─────────────────────────────────────────────────────────────────── */
.stTabs [data-baseweb="tab-list"] {
    gap: 4px;
    border-bottom: 1px solid rgba(128,128,128,.15);
}
.stTabs [data-baseweb="tab"] {
    border-radius: 8px 8px 0 0 !important;
    padding: 8px 18px !important;
    font-weight: 500 !important;
}

/* ── Divider ──────────────────────────────────────────────────────────────── */
hr { margin: 1rem 0; opacity: .12; }

/* ── Responsive: mobile ───────────────────────────────────────────────────── */
@media (max-width: 640px) {
    .block-container { padding: 1rem 0.75rem 2rem; }
    .header-title    { font-size: 1.05rem; }
    .metric-value    { font-size: 1.25rem; }
    .metrics-wrap    { grid-template-columns: repeat(2, 1fr); }
    .preview-table thead th,
    .preview-table tbody td { padding: 6px 8px; }
}

/* ── Responsive: tablet ───────────────────────────────────────────────────── */
@media (min-width: 641px) and (max-width: 900px) {
    .metrics-wrap { grid-template-columns: repeat(3, 1fr); }
}

/* ── Dark mode overrides ──────────────────────────────────────────────────── */
@media (prefers-color-scheme: dark) {
    .file-ok      { background: #0d2e11; border-color: #1a5c20; color: #6ee27a; }
    .file-skip    { background: #2d1212; border-color: #6b2020; color: #f98080; }
    .banner-success { background: #0d2e11; border-color: #1a5c20; color: #6ee27a; }
    .banner-error   { background: #2d1212; border-color: #6b2020; color: #f98080; }
    .banner-warning { background: #2d2008; border-color: #7a5c00; color: #fcd34d; }
}

</style>
"""
