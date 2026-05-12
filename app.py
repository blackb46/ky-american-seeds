"""Kentucky American Seeds — Transaction Manager.

Streamlit app entry point. Three tabs:
    1. Upload & Process — extract invoices from PDFs, review, edit, approve.
    2. Dashboard — KPIs, filters, charts, grower/invoice/product summaries.
    3. Grower Map — geocoded markers with click-through invoice history.

The workbook is the source of truth, hosted on Google Drive and shared with
the service account configured in secrets. Sheet1 stays in its existing format;
finance details (loan #, ACH, batch, etc.) live on a second sheet.
"""
from __future__ import annotations
import io
import json
from datetime import datetime
from pathlib import Path

import pandas as pd
import streamlit as st

from lib import auth, drive, extract, geocode, normalize as norm, pdf_render
from lib import theme, validate, verify, workbook as wb_mod


# ---------------------------------------------------------------------------
# Page config + theme
# ---------------------------------------------------------------------------
LOGO_PATH = Path(__file__).resolve().parent / "logo.png"

st.set_page_config(
    page_title="KAS — Transaction Manager",
    page_icon=str(LOGO_PATH) if LOGO_PATH.exists() else "🌾",
    layout="wide",
    initial_sidebar_state="expanded",
)
theme.inject_css()


# ---------------------------------------------------------------------------
# Workbook cache (Drive download)
# ---------------------------------------------------------------------------
def _is_test_mode() -> bool:
    return bool(st.secrets.get("TEST_MODE", False))


@st.cache_data(ttl=120, show_spinner="Loading workbook from Google Drive...")
def fetch_workbook_bytes(file_id: str, _cache_key: str) -> bytes:
    """Download the .xlsx. _cache_key forces a refresh when changed."""
    return drive.download_xlsx(file_id)


@st.cache_data(ttl=10, show_spinner=False)
def _cached_file_metadata(file_id: str) -> dict:
    """Cache Drive file_metadata for 10 seconds. Without this, every Streamlit
    rerun (including st_folium chatter, cookie-iframe events, etc.) hits the
    Drive API. The user's diag log showed 20+ metadata calls in a 9-minute
    session — every one is an SSL handshake and a chance for transient
    failures. 10s is short enough that concurrent-edit detection still works."""
    return drive.file_metadata(file_id)


def reload_workbook() -> tuple[bytes, str, dict]:
    """Return (workbook bytes, modifiedTime, meta dict). modifiedTime acts as
    the optimistic-lock baseline for the next save. The meta dict is returned
    so callers don't need to make a second file_metadata round trip."""
    file_id = st.secrets["GOOGLE_DRIVE_FILE_ID"]
    with _diag_timing("drive.file_metadata (cached 10s)"):
        meta = _cached_file_metadata(file_id)
    mtime = meta["modifiedTime"]
    with _diag_timing("fetch_workbook_bytes (cached if mtime unchanged)"):
        data = fetch_workbook_bytes(file_id, _cache_key=mtime)
    st.session_state["wb_modified_time"] = mtime
    return data, mtime, meta


import sys as _sys
import time as _time
import traceback as _tb
from contextlib import contextmanager


def _diag(msg: str, *, level: str = "INFO") -> None:
    """Append a timestamped message to the diagnostic log AND emit to stderr
    so it also shows in the Streamlit Cloud server log (Manage app → Logs).
    The session-state log shows in a collapsible sidebar panel.

    Levels: INFO (default), WARN, ERROR, TIMING, CLICK, RENDER, SAVE, DRIVE.
    """
    ts = datetime.now().strftime("%H:%M:%S.%f")[:-3]
    line = f"{ts}  [{level:<6}] {msg}"
    if "_diag_log" not in st.session_state:
        st.session_state["_diag_log"] = []
    log = st.session_state["_diag_log"]
    log.append(line)
    # Cap to most-recent 500 entries (per session).
    if len(log) > 500:
        del log[: len(log) - 500]
    # Mirror to stderr so it survives session crashes and shows up in
    # Streamlit Cloud's server log. Print is line-buffered → flushes promptly.
    print(f"[KAS] {line}", file=_sys.stderr, flush=True)


@contextmanager
def _diag_timing(name: str):
    """Time an operation and log the duration. Usage:
        with _diag_timing("map render"):
            ...build map...
    """
    start = _time.perf_counter()
    try:
        yield
    finally:
        elapsed_ms = (_time.perf_counter() - start) * 1000
        _diag(f"{name}: {elapsed_ms:.0f}ms", level="TIMING")


def _diag_error(msg: str, exc: BaseException | None = None) -> None:
    """Log an error with optional traceback to both session log and stderr."""
    if exc is not None:
        tb_str = "".join(_tb.format_exception(type(exc), exc, exc.__traceback__))
        _diag(f"{msg}\n{tb_str}", level="ERROR")
    else:
        _diag(msg, level="ERROR")


def _invalidate_workbook_caches() -> None:
    """Targeted cache invalidation after a successful save. Prefer this over
    st.cache_data.clear() — the global clear wipes filter option lists,
    grower index, etc. that don't actually depend on the workbook contents.

    Looks names up via globals() because some callers (the sidebar Backfill
    button) run BEFORE the cache-decorated functions are defined later in
    the module. A NameError there would break the user-visible action.
    """
    for name in ("_dataframes", "_existing_invoice_numbers_cached",
                  "_existing_keys_cached", "_build_grower_index",
                  "_filter_options", "_prepare_map_dataframe"):
        fn = globals().get(name)
        if fn is None:
            continue
        try:
            fn.clear()
        except Exception as e:
            # Don't swallow silently — a Streamlit version mismatch in .clear()
            # would leave callers reading stale data and surface as cryptic
            # downstream errors. Print so it shows in Streamlit Cloud logs.
            print(f"Cache clear failed for {name}: {e}")


def write_workbook(content: bytes) -> None:
    if _is_test_mode():
        st.toast("TEST MODE — write skipped", icon="🧪")
        return
    file_id = st.secrets["GOOGLE_DRIVE_FILE_ID"]
    expected = st.session_state.get("wb_modified_time", "")
    try:
        meta = drive.upload_xlsx_if_unchanged(file_id, content, expected)
    except drive.ConcurrentEditError as e:
        st.error(str(e))
        st.stop()
    st.session_state["wb_modified_time"] = meta["modifiedTime"]
    fetch_workbook_bytes.clear()
    try:
        _cached_file_metadata.clear()
    except Exception:
        pass


# ---------------------------------------------------------------------------
# Top header
# ---------------------------------------------------------------------------
# ---------------------------------------------------------------------------
# Auth (rendered before the header so the login is at the top of the page)
# ---------------------------------------------------------------------------
# Session-lifecycle markers. Streamlit reruns the script on every
# interaction, so logging a "rerun started" line lets us count and time
# reruns. The first rerun in a fresh session also stamps a SESSION START.
if "_session_started" not in st.session_state:
    st.session_state["_session_started"] = True
    st.session_state["_rerun_count"] = 0
    _diag("=" * 50, level="INFO")
    _diag("SESSION START (cold first render)", level="INFO")
st.session_state["_rerun_count"] = st.session_state.get("_rerun_count", 0) + 1
_diag(f"Rerun #{st.session_state['_rerun_count']}", level="RENDER")
# One memory snapshot per rerun so we can correlate growth with crashes.
_log_memory_snapshot(f"rerun #{st.session_state['_rerun_count']}")

if not auth.require_login():
    st.stop()

theme.header(LOGO_PATH)
if _is_test_mode():
    st.markdown(
        '<div class="test-mode-banner">🧪 TEST MODE — '
        'extracted data is shown but NOT written to the spreadsheet.</div>',
        unsafe_allow_html=True,
    )


# ---------------------------------------------------------------------------
# Sidebar — workbook status + actions
# ---------------------------------------------------------------------------
with st.sidebar:
    st.markdown("### 📊 Workbook")
    # Validate required secrets — missing keys here used to crash the app
    # on first load with no clear error message.
    _required_secrets = ("GOOGLE_DRIVE_FILE_ID",)
    _missing = [k for k in _required_secrets if k not in st.secrets]
    if _missing:
        st.error(f"Missing secret(s): {', '.join(_missing)}. "
                 "Add them in Streamlit Cloud secrets and reload.")
        st.stop()
    try:
        _bytes, _mtime, meta = reload_workbook()
        st.caption(f"**{meta.get('name', '?')}**")
        st.caption(f"Updated: {meta.get('modifiedTime', '?')[:19].replace('T', ' ')}")
        st.caption(f"Size: {int(meta.get('size', 0)) / 1024:.1f} KB")
    except Exception as e:
        st.error(f"Drive load failed: {e}")
        if st.button("🔁 Retry Drive load", use_container_width=True):
            fetch_workbook_bytes.clear()
            st.rerun()
        st.stop()

    file_id = st.secrets["GOOGLE_DRIVE_FILE_ID"]
    st.link_button(
        "📂 Open in Google Sheets",
        f"https://docs.google.com/spreadsheets/d/{file_id}/edit",
        use_container_width=True,
    )

    _gcs_bucket = st.secrets.get("GCS_PDF_BUCKET") or ""
    st.caption(f"PDF storage: {'☁️ GCS: ' + _gcs_bucket if _gcs_bucket else '⚠️ GCS not configured'}")

    if st.button("🔄 Reload from Drive", use_container_width=True):
        fetch_workbook_bytes.clear()
        st.rerun()

    if st.button("🔍 Verify data accuracy", use_container_width=True,
                  help="Compares the app's view to the raw .xlsx cell-by-cell."):
        with st.spinner("Verifying..."):
            try:
                report = verify.verify_workbook(_bytes)
                if report.passed:
                    st.success(
                        f"✅ All checks passed — {report.app_rows:,} rows, "
                        f"${report.app_total:,.2f} total revenue, "
                        f"matches Excel exactly."
                    )
                else:
                    st.error(
                        f"❌ {len(report.failures)} check(s) failed."
                    )
                with st.expander("Detailed results", expanded=not report.passed):
                    for c in report.checks:
                        icon = "✅" if c.passed else "❌"
                        st.markdown(
                            f"{icon} **{c.name}** — "
                            f"app: `{c.app_value}` · raw: `{c.raw_value}`"
                        )
            except Exception as e:
                st.error(f"Verification failed: {e}")

    st.divider()
    with st.expander("🔬 Diagnostics", expanded=False):
        st.caption("Recent events (most recent at bottom). Server-side logs "
                    "with the same content are available in Streamlit Cloud "
                    "→ Manage app → Logs.")
        log = st.session_state.get("_diag_log", [])
        st.code("\n".join(log[-60:]) if log else "(no events yet)",
                language=None)
        c_dl, c_clr = st.columns(2)
        with c_dl:
            st.download_button(
                "💾 Download full log",
                data="\n".join(log),
                file_name=f"kas_diag_{datetime.now():%Y%m%d_%H%M%S}.txt",
                mime="text/plain",
                use_container_width=True,
                disabled=not log,
                key="_diag_download",
            )
        with c_clr:
            if st.button("Clear log", use_container_width=True, key="_diag_clear"):
                st.session_state["_diag_log"] = []
                st.rerun()
    if st.button("🚪 Sign out", use_container_width=True):
        auth.logout()


# ---------------------------------------------------------------------------
# Workbook → DataFrames (cached)
# ---------------------------------------------------------------------------
@st.cache_data(ttl=120)
def _dataframes(_xlsx: bytes, cache_key: str) -> tuple[pd.DataFrame, pd.DataFrame]:
    """Cache dataframes keyed on the workbook's modifiedTime. ``_xlsx`` is
    underscore-prefixed so it doesn't participate in the cache key (it's
    huge bytes); ``cache_key`` (mtime) is what actually drives invalidation."""
    wb = wb_mod.load(_xlsx)
    return wb_mod.read_sheet1_dataframe(wb), wb_mod.read_finance_dataframe(wb)


df_sheet1, df_finance = _dataframes(_bytes, _mtime)


def _get_pdf_folder_id() -> str:
    """Return the Google Drive folder ID where processed PDFs are stored.

    Priority:
      1. PROCESSED_PDFS_DRIVE_FOLDER_ID secret — a folder the user created
         manually in their own Drive and shared with the service account.
         This is the recommended approach because the service account cannot
         create folders visible to the user without explicit share access.
      2. (Legacy fallback) find_or_create in the workbook's parent folder —
         this silently creates the folder in the service account's own Drive
         if the service account lacks write access to the user's folder,
         making the PDFs invisible.

    Set PROCESSED_PDFS_DRIVE_FOLDER_ID in Streamlit Cloud secrets to fix
    the "folder not visible in my Drive" problem.
    """
    folder_id = st.secrets.get("PROCESSED_PDFS_DRIVE_FOLDER_ID")
    if folder_id:
        return folder_id
    # Fallback: try to find/create alongside the workbook (may end up in
    # the service account's Drive if the account lacks folder write access).
    folder_name = st.secrets.get("PROCESSED_PDFS_DRIVE_FOLDER_NAME", "kas_processed_pdfs")
    workbook_meta = drive.file_metadata(st.secrets["GOOGLE_DRIVE_FILE_ID"])
    parents = workbook_meta.get("parents", [])
    parent_id = parents[0] if parents else None
    return drive.find_or_create_folder(folder_name, parent_id=parent_id)


def _upload_pdf(filename: str, content: bytes) -> str | None:
    """Upload a PDF and return a storable reference (GCS URL or Drive file ID).

    When a single PDF contains multiple invoices, the user clicks "Approve &
    Save" for each one. Without dedup, that would upload the same PDF to GCS
    multiple times under the same filename (overwrite) — wasteful, but more
    importantly the OLD code with Drive could create duplicate files. We cache
    the upload result by content hash in session state so all invoices from
    one PDF share a single uploaded copy.

    Prefers GCS (service accounts have no Drive storage quota on personal
    accounts). Falls back to Drive folder upload if GCS_PDF_BUCKET is not set.
    """
    import hashlib
    content_hash = hashlib.sha256(content).hexdigest()
    cache = st.session_state.setdefault("_pdf_upload_cache", {})
    if content_hash in cache:
        return cache[content_hash]

    bucket = st.secrets.get("GCS_PDF_BUCKET")
    if bucket:
        ref = drive.upload_pdf_to_gcs(bucket, filename, content)
    else:
        folder_id = _get_pdf_folder_id()
        result = drive.upload_pdf(folder_id, filename, content)
        ref = result.get("id")
    cache[content_hash] = ref
    return ref


@st.cache_data(ttl=120)
def _existing_invoice_numbers_cached(_xlsx: bytes, cache_key: str) -> set[str]:
    """Cached invoice-number lookup. Avoids re-parsing the whole workbook
    on every rerun while the user is typing in a review form."""
    return wb_mod.existing_invoice_numbers(wb_mod.load(_xlsx))


@st.cache_data(ttl=120)
def _existing_keys_cached(_xlsx: bytes, cache_key: str) -> set[tuple]:
    """Cached set of (invoice_no, item_description, quantity) keys already
    in Sheet1. Used for per-line-item duplicate detection in the review
    form so the user can pick which items to add when an invoice is partially
    in the spreadsheet."""
    return wb_mod.existing_keys(wb_mod.load(_xlsx))


def _line_item_dup_key(invoice_no_str: str, item: dict) -> tuple:
    """Build the same dedup key the workbook uses, so per-row checkbox
    defaults match what append_invoice would actually do.

    Keyed on (invoice, qty, ext_amount) — numeric fields are exact and
    don't suffer from description formatting drift between the sheet
    and Claude's extraction."""
    qty = wb_mod._round_or_none(item.get("quantity"))
    total = wb_mod._round_or_none(item.get("ext_amount"))
    return (invoice_no_str, qty, total)


@st.cache_data(ttl=300, show_spinner=False, max_entries=8)
def _render_pdf_pages_cached(_pdf_bytes: bytes, content_hash: str, dpi: int = 110) -> list:
    """Cache rendered PDF page images so the 'View full PDF' expander doesn't
    re-rasterize the entire document on every keystroke / rerun. Keyed on the
    content hash; max_entries=8 caps memory."""
    return pdf_render.render_pdf_pages(_pdf_bytes, dpi=dpi)


@st.cache_data(ttl=300, show_spinner=False, max_entries=20)
def _render_pdf_thumbnail_cached(_pdf_bytes: bytes, content_hash: str, dpi: int = 120) -> bytes:
    """Cache the per-PDF thumbnail used at the top of the review form.

    Previously this called PyMuPDF on every rerun for every PDF in the form
    — hundreds of native renders per session. Caching by content hash drops
    that to one render per PDF and is the leading-suspect fix for the
    intermittent segfaults seen on post-save reruns."""
    return pdf_render.render_pdf_thumbnail(_pdf_bytes, page=0, dpi=dpi)


def _log_memory_snapshot(label: str) -> None:
    """Log current process RSS (resident set size) to the diagnostic log.
    Streamlit Cloud kills processes ~1 GB; this lets us see growth trends
    and correlate with segfaults."""
    try:
        import resource
        rss_kb = resource.getrusage(resource.RUSAGE_SELF).ru_maxrss
        # On Linux ru_maxrss is in KB; on macOS it's in bytes. Streamlit
        # Cloud is Linux so KB. Display as MB.
        _diag(f"memory peak: {rss_kb / 1024:.0f} MB  ({label})", level="MEM")
    except Exception:
        pass  # resource module unavailable on Windows; silent fail


@st.cache_data(ttl=300, show_spinner=False)
def _prepare_map_dataframe(_xlsx: bytes, cache_key: str) -> pd.DataFrame:
    """Heavy dataframe prep for the map. df.apply over thousands of rows
    was the dominant cost on cold map render (~24s observed). Cache it on
    workbook mtime so subsequent renders pull a precomputed dataframe with
    __raw_label / __grower / __addr1 / __city / __state / __zip already
    populated."""
    wb = wb_mod.load(_xlsx)
    df = wb_mod.read_sheet1_dataframe(wb)
    df["__raw_label"] = df.apply(_grower_label, axis=1)
    canonical = _canonical_names_by_id(df)
    df["__grower"] = df.apply(
        lambda r: canonical.get(r["Grower ID"], r["__raw_label"]),
        axis=1,
    )

    def _norm(s):
        return str(s).strip().upper() if pd.notnull(s) and str(s).strip() else ""
    df["__addr1"] = df["Grower Address1"].apply(_norm)
    df["__city"] = df["Grower City"].apply(_norm)
    df["__state"] = df["Grower State"].apply(_norm)
    df["__zip"] = df["Grower ZIP CODE"].apply(
        lambda z: str(int(z)) if pd.notnull(z) and str(z).strip() else ""
    )
    df = df[df["__addr1"] != ""]
    return df


@st.cache_data(ttl=120, show_spinner=False)
def _filter_options(_df: pd.DataFrame, cache_key: str) -> dict:
    """Pre-compute distinct option lists for the dashboard sidebar filters.

    Rebuilding these on every keystroke (any filter widget change triggers a
    rerun) was a major source of dashboard sluggishness with thousands of rows.
    Cached on the workbook mtime — invalidated only when the data actually
    changes.
    """
    df_dates = _df["Invoice Date"].dropna()
    years = sorted({d.year for d in df_dates})
    retailers = sorted({norm.normalize_retailer(r) or r
                         for r in _df["Retailer Name"].dropna().unique()})
    finance_cos = sorted({norm.normalize_finance_company(f) or f
                           for f in _df["Finance Company"].dropna().unique()})
    manufacturers = sorted(_df["Manufacturer Name"].dropna().unique().tolist())
    growers = sorted({_grower_label(r) for _, r in _df.iterrows() if _grower_label(r)})
    products = sorted(_df["Item Description/Brand"].dropna().unique().tolist())
    if len(df_dates):
        min_d, max_d = df_dates.min().date(), df_dates.max().date()
    else:
        min_d = max_d = datetime.now().date()
    return {
        "years": years,
        "retailers": retailers,
        "finance_cos": finance_cos,
        "manufacturers": manufacturers,
        "growers": growers,
        "products": products,
        "min_d": min_d,
        "max_d": max_d,
    }


# ---------------------------------------------------------------------------
# Navigation (st.tabs resets on every rerun, dropping the user back to the
# Upload tab whenever any widget triggers a refresh. A radio keeps state.)
# ---------------------------------------------------------------------------
# Feature flags — flip to True to re-enable. The Dashboard and Map pages
# are kept in code but hidden from the navigation; their function bodies
# never execute when the flags are False, so no extra work / network
# traffic / chart building happens at runtime.
ENABLE_DASHBOARD = False
ENABLE_MAP = False

PAGES = ["📄 Upload & Process"]
if ENABLE_DASHBOARD:
    PAGES.append("📊 Dashboard")
if ENABLE_MAP:
    PAGES.append("🗺️ Grower Map")

# Only show the page selector if there's more than one page enabled.
if len(PAGES) > 1:
    active_page = st.radio(
        "Section", PAGES, horizontal=True, label_visibility="collapsed",
        key="active_page",
    )
    _prev_page = st.session_state.get("_last_active_page")
    if _prev_page != active_page:
        _diag(f"Page changed: {_prev_page} → {active_page}", level="CLICK")
        st.session_state["_last_active_page"] = active_page
    st.divider()
else:
    active_page = PAGES[0]



# ===========================================================================
# Tab 1 — Upload & Process
# ===========================================================================
def _conf(level: str | None) -> str:
    return theme.confidence_pill(level)


def _render_review_form(idx: int, pdf_name: str, pdf_bytes: bytes,
                        result: dict, existing_invoice_nums: set[str],
                        existing_keys: set[tuple]):
    """Render an editable review form for one extracted PDF."""
    final = result["final"]
    pass1 = result["pass1"]
    invoices = final.get("invoices", [])

    st.markdown(f"### 📄 {pdf_name}")
    cols_top = st.columns([2, 5])
    with cols_top[0]:
        with st.expander("PDF preview", expanded=True):
            try:
                import hashlib as _hl
                _pdf_h = _hl.sha256(pdf_bytes).hexdigest()
                doc_pages = _render_pdf_thumbnail_cached(pdf_bytes, _pdf_h, dpi=120)
                st.image(doc_pages, use_container_width=True)
                st.caption("Page 1")
            except Exception as e:
                st.warning(f"Couldn't render preview: {e}")

    with cols_top[1]:
        st.markdown(
            f"**Verifier confidence:** {result.get('pass2', {}).get('overall_confidence', '?')}  •  "
            f"**Invoices found:** {len(invoices)}  •  "
            f"**API tokens:** {result['usage']['input_tokens']:,} in / {result['usage']['output_tokens']:,} out"
        )
        agree = result.get("agreement", {})
        if agree:
            for inv_no, info in agree.items():
                diffs = info.get("discrepancies", [])
                if diffs:
                    st.warning(f"Inv {inv_no}: pass-1 vs pass-2 differ — {'; '.join(diffs)}")

    if not invoices:
        st.error("No invoices were extracted from this PDF.")
        return

    for inv_idx, inv in enumerate(invoices):
        # Normalize invoice number to plain integer string so "1094912.0"
        # and "1094912" both compare equal against whatever's in the workbook.
        _raw_no = inv.get("invoice_number")
        try:
            invoice_no = str(int(float(str(_raw_no).strip()))) if _raw_no else "?"
        except (ValueError, TypeError):
            invoice_no = str(_raw_no or "?").strip()
        is_dup = invoice_no in existing_invoice_nums
        state_key = f"review_{idx}_{inv_idx}"

        with st.container(border=True):
            top = st.columns([4, 2, 2])
            top[0].markdown(
                f"#### Invoice **#{invoice_no}** {_conf(inv.get('invoice_number_confidence'))}",
                unsafe_allow_html=True,
            )
            top[1].markdown(
                f"**Total:** ${(inv.get('invoice_total') or 0):,.2f}  "
                f"{_conf(inv.get('invoice_total_confidence'))}",
                unsafe_allow_html=True,
            )
            if is_dup:
                top[2].error("⚠️ Already in spreadsheet")

            # Duplicate handling: invoice number already in spreadsheet.
            # We no longer block Approve & Save — the Include checkboxes
            # in the line items table let the user pick exactly which rows
            # to add. Skip and Attach PDF are still here for convenience.
            if is_dup:
                already_attached = st.session_state.get(f"{state_key}_done") == "attached"
                if already_attached:
                    st.success(
                        f"📎 PDF already attached to invoice {invoice_no}."
                    )
                d1, d2, d3 = st.columns([1, 1, 3])
                if d1.button("⏭️ Skip this invoice",
                              key=f"{state_key}_dup_skip"):
                    st.session_state[f"{state_key}_done"] = "skipped"
                    st.rerun()
                if d2.button("📎 Attach PDF only",
                              key=f"{state_key}_dup_attach",
                              help="Uploads the PDF and links it to existing rows of this invoice. Use this if you only want the PDF link, not new line items."):
                    if _attach_pdf_to_existing(invoice_no, pdf_name, pdf_bytes, inv):
                        st.session_state[f"{state_key}_done"] = "attached"
                        st.rerun()

            # Validate — wrapped so a crash inside validate doesn't take
            # the whole review form down. Logs the offending inv dict to
            # the diagnostic log so we can see what shape Claude returned.
            try:
                vr = validate.validate_invoice(inv)
            except Exception as _vexc:
                import json as _json
                try:
                    _inv_dump = _json.dumps(inv, default=str, indent=2)
                except Exception:
                    _inv_dump = repr(inv)
                _diag_error(
                    f"validate.validate_invoice CRASHED for {pdf_name} "
                    f"invoice_no={invoice_no!r}. Extracted JSON below:\n{_inv_dump}",
                    _vexc,
                )
                st.error(
                    f"⚠️ This invoice (`{invoice_no}` in `{pdf_name}`) hit an "
                    "extraction-format error during validation. The full "
                    "details have been written to 🔬 Diagnostics in the "
                    "sidebar — please download the log so we can fix the "
                    "underlying bug. Skip this invoice for now."
                )
                continue  # to the next invoice in this PDF
            if vr.issues:
                with st.expander(
                    f"🔍 Validation: {vr.summary()}",
                    expanded=bool(vr.errors),
                ):
                    for issue in vr.issues:
                        cls = f"issue-{issue.severity}"
                        st.markdown(
                            f'<div class="{cls}"><b>{issue.field}</b> — {issue.message}</div>',
                            unsafe_allow_html=True,
                        )

            # ---- Editable fields ----
            st.markdown("##### Invoice header")
            h1, h2, h3, h4 = st.columns(4)
            inv["invoice_number"] = h1.text_input(
                "Invoice #", value=invoice_no, key=f"{state_key}_invno"
            )
            inv["patron_number"] = h2.text_input(
                "Patron #", value=str(inv.get("patron_number") or ""), key=f"{state_key}_patron"
            )
            inv["sold_date"] = h3.text_input(
                "Sold Date (YYYY-MM-DD)", value=str(inv.get("sold_date") or ""),
                key=f"{state_key}_date",
            )
            inv["retailer_name"] = h4.text_input(
                "Retailer", value=str(inv.get("retailer_name") or ""),
                key=f"{state_key}_retailer",
            )

            st.markdown("##### Grower")
            g = inv.setdefault("grower", {})
            g1, g2, g3 = st.columns(3)
            g["first_name"] = g1.text_input("First", value=g.get("first_name") or "",
                                             key=f"{state_key}_gfn")
            g["last_name"] = g2.text_input("Last", value=g.get("last_name") or "",
                                            key=f"{state_key}_gln")
            g["company_name"] = g3.text_input("Company", value=g.get("company_name") or "",
                                               key=f"{state_key}_gco")
            ga1, ga2, ga3, ga4 = st.columns([3, 2, 1, 1])
            g["address1"] = ga1.text_input("Address", value=g.get("address1") or "",
                                            key=f"{state_key}_addr")
            g["city"] = ga2.text_input("City", value=g.get("city") or "",
                                        key=f"{state_key}_city")
            g["state"] = ga3.text_input("State", value=g.get("state") or "KY",
                                         key=f"{state_key}_state", max_chars=2)
            g["zip"] = ga4.text_input("ZIP", value=str(g.get("zip") or ""),
                                       key=f"{state_key}_zip", max_chars=5)

            st.markdown("##### Line items")
            li_df = pd.DataFrame(inv.get("line_items") or [])
            for col in ("item_number", "description", "unit", "manufacturer"):
                if col not in li_df.columns:
                    li_df[col] = None
            for col in ("quantity", "unit_price", "ext_amount"):
                if col not in li_df.columns:
                    li_df[col] = None

            # Per-row duplicate detection. Mark each line item as already-
            # in-sheet using the same (invoice_no, desc, qty) key the
            # workbook uses. Add "Already in sheet?" status + "Include"
            # checkbox columns. Default: ALL unchecked — the user must
            # opt-in to add rows to the sheet (per requirement).
            already_flags = []
            for r in li_df.to_dict(orient="records"):
                already_flags.append(_line_item_dup_key(invoice_no, r) in existing_keys)
            li_df["__already"] = ["✓ already in sheet" if a else "➕ new" for a in already_flags]
            li_df["Include"] = False  # default unchecked

            n_already = sum(already_flags)
            n_new = len(already_flags) - n_already
            if n_already > 0 and n_new > 0:
                st.info(f"📋 {n_already} of {len(already_flags)} line item(s) "
                          f"already in the spreadsheet. {n_new} are new. "
                          "Use the **Include** checkbox to choose which to add.")
            elif n_already == len(already_flags) and n_already > 0:
                st.info(f"📋 All {n_already} line item(s) already in the spreadsheet. "
                          "Use **📎 Attach PDF** below to (re)link the PDF, or check "
                          "Include if you want to override.")
            else:
                st.info(f"📋 {n_new} new line item(s). "
                          "Tick **Include** on the rows you want to add.")

            qa, qb, _ = st.columns([1, 1, 4])
            with qa:
                if st.button("✔ Select all new", key=f"{state_key}_sel_new",
                              help="Check Include for every line item not already in the sheet."):
                    li_df["Include"] = [not a for a in already_flags]
                    st.session_state[f"{state_key}_li_override"] = li_df
                    st.rerun()
            with qb:
                if st.button("✗ Clear all", key=f"{state_key}_sel_none",
                              help="Uncheck Include on every row."):
                    li_df["Include"] = False
                    st.session_state[f"{state_key}_li_override"] = li_df
                    st.rerun()

            # Apply any quick-button override (consumed once).
            override = st.session_state.pop(f"{state_key}_li_override", None)
            if override is not None:
                li_df = override

            li_df = li_df[["Include", "__already", "item_number", "description",
                            "unit", "quantity", "unit_price", "ext_amount", "manufacturer"]]
            edited_li = st.data_editor(
                li_df,
                num_rows="fixed",  # disabling row add/delete to avoid breaking the dup-status column
                use_container_width=True,
                key=f"{state_key}_li",
                column_config={
                    "Include": st.column_config.CheckboxColumn(
                        "Include",
                        help="Check to add this line item to the spreadsheet.",
                        default=False,
                    ),
                    "__already": st.column_config.TextColumn(
                        "Status", disabled=True,
                        help="✓ already in sheet means this exact line item "
                             "(same invoice + description + quantity) is already saved.",
                    ),
                    "quantity": st.column_config.NumberColumn(format="%.2f"),
                    "unit_price": st.column_config.NumberColumn(format="$%.2f"),
                    "ext_amount": st.column_config.NumberColumn(format="$%.2f"),
                },
            )
            edited_records = edited_li.to_dict(orient="records")
            # Strip helper columns before passing to save logic.
            for r in edited_records:
                r.pop("__already", None)
            inv["line_items"] = edited_records

            # Live recompute math summary
            try:
                recomputed_sum = sum(float(r.get("ext_amount") or 0) for r in inv["line_items"])
            except (TypeError, ValueError):
                recomputed_sum = 0.0
            inv_total_input = st.number_input(
                "Invoice total",
                value=float(inv.get("invoice_total") or recomputed_sum),
                key=f"{state_key}_total",
                format="%.2f",
            )
            inv["invoice_total"] = inv_total_input
            mc1, mc2, mc3 = st.columns(3)
            mc1.metric("Sum of line items", f"${recomputed_sum:,.2f}")
            mc2.metric("Invoice total", f"${inv_total_input:,.2f}")
            mc3.metric("Difference", f"${recomputed_sum - inv_total_input:,.2f}",
                       delta_color="inverse")

            st.markdown("##### Prepaid split")
            ps1, ps2 = st.columns(2)
            inv["account_charge_amount"] = ps1.number_input(
                "Charged to account", value=float(inv.get("account_charge_amount") or 0),
                key=f"{state_key}_chg", format="%.2f",
            )
            inv["prepaid_amount"] = ps2.number_input(
                "Applied to prepaid", value=float(inv.get("prepaid_amount") or 0),
                key=f"{state_key}_prep", format="%.2f",
            )

            st.markdown("##### Finance details")
            f = inv.setdefault("finance", {})
            f1, f2, f3, f4 = st.columns(4)
            f["finance_company"] = f1.text_input("Finance Co",
                                                   value=f.get("finance_company") or "",
                                                   key=f"{state_key}_fco")
            f["loan_number"] = f2.text_input("Loan #", value=f.get("loan_number") or "",
                                              key=f"{state_key}_loan")
            f["loan_year"] = f3.text_input("Loan Year", value=str(f.get("loan_year") or ""),
                                            key=f"{state_key}_lyr")
            f["batch_number"] = f4.text_input("Batch #",
                                               value=str(f.get("batch_number") or ""),
                                               key=f"{state_key}_batch")
            f5, f6, f7, f8 = st.columns(4)
            f["product_rate"] = f5.text_input("Product/Rate",
                                                value=f.get("product_rate") or "",
                                                key=f"{state_key}_prate")
            f["manufacturer_from_notes"] = f6.text_input("Manufacturer",
                                                          value=f.get("manufacturer_from_notes") or "",
                                                          key=f"{state_key}_mfg")
            f["amount_to_retailer"] = f7.number_input(
                "Amount to Retailer", value=float(f.get("amount_to_retailer") or 0),
                key=f"{state_key}_a2r", format="%.2f",
            )
            f["ach_date"] = f8.text_input("ACH Date (YYYY-MM-DD)",
                                            value=str(f.get("ach_date") or ""),
                                            key=f"{state_key}_ach")

            inv["extraction_notes"] = st.text_area(
                "Notes / extraction comments",
                value=inv.get("extraction_notes") or "",
                key=f"{state_key}_notes",
            )

            # Action buttons. The save button is disabled when this invoice
            # is already in the spreadsheet — duplicate-handling buttons
            # (Skip / Attach PDF only) are shown above instead.
            b1, b2, b3, b4 = st.columns([1, 1, 1, 3])
            # Count how many rows are checked Include=True. Disable Save
            # if zero — saves a confused click when nothing is selected.
            n_to_include = sum(1 for r in inv["line_items"] if r.get("Include"))
            approve = b1.button(
                f"✅ Approve & Save ({n_to_include} item{'s' if n_to_include != 1 else ''})",
                type="primary",
                disabled=not vr.passes or n_to_include == 0,
                help=("Tick Include on at least one line item to enable Save."
                      if n_to_include == 0 else
                      f"Saves the {n_to_include} ticked line item(s) and "
                      "uploads the PDF."),
                key=f"{state_key}_approve",
            )
            skip = b2.button("⏭️ Skip", key=f"{state_key}_skip")
            # "View full PDF" is a TOGGLE backed by session state so the
            # expander stays open across reruns instead of disappearing on the
            # next keystroke (which used to make the page jump dramatically as
            # full-page images vanished).
            _view_key = f"{state_key}_view_open"
            if b3.button("👀 View full PDF", key=f"{state_key}_view_btn"):
                st.session_state[_view_key] = not st.session_state.get(_view_key, False)
            if st.session_state.get(_view_key):
                with st.expander("Full PDF pages", expanded=True):
                    import hashlib, base64
                    pdf_hash = hashlib.sha256(pdf_bytes).hexdigest()
                    blocks = _render_pdf_pages_cached(pdf_bytes, pdf_hash, dpi=110)
                    for i, blk in enumerate(blocks):
                        st.image(base64.b64decode(blk["source"]["data"]),
                                 caption=f"Page {i+1}", use_container_width=True)

            if approve:
                status, pdf_failed = _save_invoice(inv, pdf_name, pdf_bytes)
                if status == "saved":
                    st.session_state[f"{state_key}_done"] = "saved"
                    st.session_state[f"{state_key}_pdf_failed"] = bool(pdf_failed)
                    st.rerun()
                # status == "duplicate" or "error": leave the form open so the
                # user can see the warning above and choose Skip / Attach.
            if skip:
                st.session_state[f"{state_key}_done"] = "skipped"

        if st.session_state.get(f"{state_key}_done") == "saved":
            st.success(f"✅ Invoice {invoice_no} saved.")
            if st.session_state.get(f"{state_key}_pdf_failed"):
                st.warning(
                    "⚠️ The PDF couldn't be uploaded — invoice data is saved but "
                    "won't have a viewable link from the map. Click below to retry."
                )
                if st.button("📎 Attach PDF now", key=f"{state_key}_attach_retry"):
                    if _attach_pdf_to_existing(invoice_no, pdf_name, pdf_bytes, inv):
                        st.session_state[f"{state_key}_pdf_failed"] = False
                        st.rerun()
        elif st.session_state.get(f"{state_key}_done") == "skipped":
            st.info(f"⏭️ Invoice {invoice_no} skipped.")
        # "attached" state is rendered inline above (next to the buttons) so
        # the Attach PDF button stays available for re-clicking.


def _attach_pdf_to_existing(invoice_no: str, pdf_name: str,
                              pdf_bytes: bytes, inv: dict) -> bool:
    """Upload PDF to Drive/GCS and write/update a Finance Details row that
    links it to an existing invoice. Also writes the PDF Link cell into
    Sheet1 column 23 for every line of this invoice.

    Returns True on success, False on any failure. Catches all exceptions so
    the caller's UI state machine never sees an implicit None.
    """
    try:
        return _attach_pdf_to_existing_inner(invoice_no, pdf_name, pdf_bytes, inv)
    except Exception as e:
        st.error(f"PDF attach failed: {e}")
        return False


def _attach_pdf_to_existing_inner(invoice_no: str, pdf_name: str,
                                    pdf_bytes: bytes, inv: dict) -> bool:
    _diag(f"_attach_pdf_to_existing start: invoice={invoice_no} pdf={pdf_name}",
          level="SAVE")
    if _is_test_mode():
        st.toast("TEST MODE — PDF not actually uploaded", icon="🧪")
        return True  # Treat as success in test mode so the UI advances.

    drive_id = None
    try:
        drive_id = _upload_pdf(pdf_name, pdf_bytes)
    except Exception as e:
        st.error(f"PDF upload failed: {e}")
        return False

    # Write a Finance Details row (or update if one already exists) with the
    # PDF Drive ID so the map can link to it.
    try:
        xlsx_bytes, _mtime, _ = reload_workbook()
        wb = wb_mod.load(xlsx_bytes)
        wb_mod._ensure_finance_sheet(wb)
        fws = wb[wb_mod.FINANCE_SHEET_NAME]

        # Look for an existing row for this invoice in Finance Details
        existing_row = None
        for r in range(2, fws.max_row + 1):
            v = fws.cell(row=r, column=1).value
            if v is not None and str(v).strip() == invoice_no:
                existing_row = r
                break

        finance = inv.get("finance") or {}

        if existing_row is None:
            target = fws.max_row + 1
        else:
            target = existing_row

        fields = {
            1: invoice_no,
            2: inv.get("patron_number"),
            3: finance.get("loan_number"),
            4: finance.get("loan_year"),
            5: finance.get("finance_company"),
            6: finance.get("product_rate"),
            7: finance.get("batch_number"),
            8: finance.get("ach_date"),
            9: inv.get("invoice_total"),
            10: finance.get("amount_to_retailer"),
            11: inv.get("prepaid_amount"),
            12: inv.get("account_charge_amount"),
            13: inv.get("merchandised_by"),
            14: pdf_name,
            15: drive_id,
            16: datetime.now(),
            17: False,
            18: "Attached to existing invoice (line items not re-imported).",
        }
        for col_idx, val in fields.items():
            if existing_row and col_idx == 15 and val:
                # always overwrite PDF Drive ID
                fws.cell(row=target, column=col_idx, value=val)
            elif existing_row and val in (None, ""):
                continue  # don't clobber existing values with blanks
            else:
                fws.cell(row=target, column=col_idx, value=val)

        # Also write the PDF Link into Sheet1 column 23 for every line item
        # of this invoice, so the .xlsx itself has a clickable hyperlink per
        # row (not just the Finance Details cross-reference).
        s1 = wb[wb_mod.SHEET1_NAME]
        wb_mod._ensure_sheet1_extra_headers(s1)
        url = wb_mod._pdf_url_from_ref(drive_id)
        rows_linked = 0
        if url:
            for r in range(2, s1.max_row + 1):
                inv_cell = s1.cell(row=r, column=17)
                if inv_cell.value is None:
                    continue
                if wb_mod._norm_invoice_no(inv_cell.value) == invoice_no:
                    link_cell = s1.cell(row=r, column=23,
                                        value=f'=HYPERLINK("{url}","Open PDF")')
                    link_cell.hyperlink = url
                    rows_linked += 1

        # Auto-backfill: also fill in column W for any other historical rows
        # whose invoices have a PDF on file but aren't linked yet. Free piggyback
        # on the existing write — keeps the sheet self-healing without a button.
        try:
            wb_mod.backfill_pdf_links(wb)
        except Exception as e:
            print(f"Auto-backfill failed (non-fatal): {e}")

        out = wb_mod.save_to_bytes(wb)
        write_workbook(out)
        _invalidate_workbook_caches()
        st.toast(f"PDF attached — {rows_linked} row{'s' if rows_linked != 1 else ''} linked.",
                  icon="📎")
        return True
    except Exception as e:
        st.error(f"Couldn't update Finance Details sheet: {e}")
        return False


def _save_invoice(inv: dict, pdf_name: str, pdf_bytes: bytes) -> tuple[str, bool]:
    """Append invoice to workbook + upload PDF to processed folder.

    Returns (status, pdf_failed):
      status: "saved" | "duplicate" | "error"
      pdf_failed: True if the PDF upload failed (only meaningful when saved)
    """
    _diag(f"_save_invoice start: invoice={inv.get('invoice_number')!r} pdf={pdf_name}",
          level="SAVE")
    # Filter to ONLY the line items the user checked Include. The review
    # form puts an "Include" boolean on each row; other rows are kept out
    # of the save. This is what powers per-line-item dedup.
    inv_filtered = dict(inv)
    inv_filtered["line_items"] = [
        {k: v for k, v in r.items() if k != "Include"}  # drop UI flag
        for r in (inv.get("line_items") or [])
        if r.get("Include")
    ]
    if not inv_filtered["line_items"]:
        st.error("No line items checked. Tick the Include checkbox on the "
                 "rows you want to save.")
        return ("error", False)
    _diag(f"  filtered line_items: {len(inv_filtered['line_items'])} of {len(inv.get('line_items') or [])}",
          level="SAVE")

    bundles = extract.to_invoice_bundles({"invoices": [inv_filtered]}, pdf_filename=pdf_name)
    if not bundles:
        st.error("Nothing to save (no line items).")
        return ("error", False)
    bundle = bundles[0]

    # Optionally upload PDF first to capture Drive ID into the finance row
    drive_id = None
    pdf_upload_failed = False
    if not _is_test_mode():
        try:
            drive_id = _upload_pdf(pdf_name, pdf_bytes)
        except Exception as e:
            st.warning(f"PDF upload failed (continuing without): {e}")
            pdf_upload_failed = True
    if bundle.finance is not None and drive_id:
        bundle.finance.pdf_drive_id = drive_id
        bundle.finance.pdf_source_file = pdf_name
    elif drive_id:
        # No finance data extracted — create a minimal Finance Details row
        # so the PDF Drive ID is still stored and linkable from the map.
        bundle.finance = wb_mod.FinanceDetail(
            invoice_number=bundle.invoice_number,
            pdf_source_file=pdf_name,
            pdf_drive_id=drive_id,
        )

    # Reload latest workbook bytes (concurrency-safe baseline)
    xlsx_bytes, _mtime, _ = reload_workbook()
    wb = wb_mod.load(xlsx_bytes)
    counts = wb_mod.append_invoice(wb, bundle)

    # If the workbook backstop refused the write because the invoice number
    # already exists, don't write anything — surface a clear message instead.
    # invoice_already_exists path is dead now that the workbook backstop
    # was removed (per-line-item dedup is enough). Defensive check kept
    # in case the workbook layer ever re-introduces it.
    if counts.get("invoice_already_exists"):
        return ("duplicate", pdf_upload_failed)

    # Apply the new PDF link to EVERY existing row of this invoice in
    # Sheet1 column W — including rows that were already in the sheet
    # before this upload. This way, when the user partially adds new
    # items from a multi-item invoice, all line items (old + new) end
    # up pointing to the same latest PDF instead of having a mix of
    # linked / unlinked / stale-linked rows.
    rows_synced = 0
    if drive_id:
        url = wb_mod._pdf_url_from_ref(drive_id)
        inv_no_str = wb_mod._norm_invoice_no(bundle.invoice_number)
        if url and inv_no_str:
            s1 = wb[wb_mod.SHEET1_NAME]
            wb_mod._ensure_sheet1_extra_headers(s1)
            for r in range(2, s1.max_row + 1):
                inv_cell = s1.cell(row=r, column=17)
                if inv_cell.value is None:
                    continue
                if wb_mod._norm_invoice_no(inv_cell.value) == inv_no_str:
                    link_cell = s1.cell(row=r, column=23,
                                        value=f'=HYPERLINK("{url}","Open PDF")')
                    link_cell.hyperlink = url
                    rows_synced += 1
            _diag(f"  PDF link synced to {rows_synced} row(s) of invoice {inv_no_str}",
                  level="SAVE")

    # Auto-backfill PDF links: catches any historical rows on OTHER
    # invoices whose Finance Details has a PDF Drive ID but column W is
    # still empty. Silent — no toast.
    try:
        wb_mod.backfill_pdf_links(wb)
    except Exception as e:
        print(f"Auto-backfill failed (non-fatal): {e}")

    out = wb_mod.save_to_bytes(wb)
    write_workbook(out)
    _invalidate_workbook_caches()

    # Auto-verify data consistency after every save. Combine save + verify
    # results into ONE toast so the user doesn't see a stack of three
    # animations flickering in and out (which felt like errors flashing).
    save_msg = (
        f"Saved {counts['line_items_added']} line items"
        + (f" ({counts['duplicates_skipped']} dupes skipped)" if counts["duplicates_skipped"] else "")
        + (f" · PDF linked to {rows_synced} row(s) for this invoice" if rows_synced else "")
    )
    try:
        report = verify.verify_workbook(out)
        if report.passed:
            st.toast(f"✅ {save_msg} · data check passed ({report.app_rows} rows).",
                      icon="✅")
        else:
            st.toast(f"✅ {save_msg}", icon="✅")
            st.error(
                "⚠️ Data consistency check FAILED after save. "
                f"{len(report.failures)} check(s) failed:\n"
                + "\n".join(
                    f"- **{c.name}**: app={c.app_value} vs raw={c.raw_value}"
                    for c in report.failures
                )
            )
    except Exception as e:
        st.toast(f"✅ {save_msg}", icon="✅")
        st.warning(f"Data verification skipped: {e}")

    return ("saved", pdf_upload_failed)


if active_page == PAGES[0]:
    st.markdown("Upload one or more invoice PDFs. The app will extract every line item, "
                "verify with a second pass, and let you review/edit before saving.")

    uploaded = st.file_uploader(
        "Drop PDFs here", type=["pdf"], accept_multiple_files=True,
        key="pdf_uploader",
    )

    if uploaded:
        # 100% accuracy on the dedup check: force a fresh spreadsheet read
        # before deciding which line items are already in. Bypasses the
        # 10s metadata cache and the 120s existing_keys cache so the
        # "✓ already in sheet" / "➕ new" badges reflect the spreadsheet's
        # CURRENT state, not a stale snapshot.
        with st.spinner("Checking the latest spreadsheet…"):
            try:
                _cached_file_metadata.clear()
                _existing_invoice_numbers_cached.clear()
                _existing_keys_cached.clear()
                _bytes, _mtime, _meta_fresh = reload_workbook()
                _diag(f"forced fresh dedup check: workbook mtime={_mtime}",
                      level="DRIVE")
            except Exception as e:
                _diag(f"Fresh workbook fetch failed; using last-known cache: {e}",
                      level="WARN")
        existing_inv_nums = _existing_invoice_numbers_cached(_bytes, _mtime)
        existing_keys = _existing_keys_cached(_bytes, _mtime)
        st.caption(f"📊 Checked against spreadsheet last modified "
                    f"`{_mtime[:19].replace('T', ' ')} UTC`. "
                    f"{len(existing_inv_nums):,} invoices, "
                    f"{len(existing_keys):,} line items already on file.")
        api_key = st.secrets["ANTHROPIC_API_KEY"]

        # Run extraction (cached in session_state by file content hash).
        # Cap the cache at 10 entries — extraction results can be ~MB each and
        # session state gets serialized between runs; unbounded growth was
        # making tab switches feel slow after processing many PDFs in a row.
        _EXTRACT_CACHE_MAX = 10
        for i, uf in enumerate(uploaded):
            file_bytes = uf.getvalue()
            cache_key = f"extract_{uf.name}_{len(file_bytes)}"
            if cache_key not in st.session_state:
                with st.spinner(f"Extracting {uf.name} (dual-pass verification)..."):
                    try:
                        result = extract.extract_pdf(file_bytes, api_key=api_key, verify=True)
                        st.session_state[cache_key] = result
                    except Exception as e:
                        st.error(f"Extraction failed for {uf.name}: {e}")
                        continue
                # Evict oldest extraction-cache entries beyond the cap.
                extract_keys = [k for k in st.session_state.keys() if k.startswith("extract_")]
                if len(extract_keys) > _EXTRACT_CACHE_MAX:
                    for old in extract_keys[:-_EXTRACT_CACHE_MAX]:
                        try:
                            del st.session_state[old]
                        except KeyError:
                            pass
            result = st.session_state[cache_key]
            _render_review_form(i, uf.name, file_bytes, result, existing_inv_nums, existing_keys)
    else:
        st.info("👆 Upload PDFs to get started.")


# ===========================================================================
# Tab 2 — Dashboard
# ===========================================================================
def _render_dashboard():
    _diag("_render_dashboard() called", level="RENDER")
    _dash_start = _time.perf_counter()
    if df_sheet1.empty:
        st.info("No data yet. Upload some PDFs to populate the dashboard.")
        return

    # Dashboard filters live in the main pane (NOT the sidebar) to avoid the
    # ~250px sidebar reflow that happens when switching tabs. Collapsed by
    # default — most users glance at the dashboard without filtering.
    opts = _filter_options(df_sheet1, _mtime)
    min_d, max_d = opts["min_d"], opts["max_d"]
    # Clamp stale date range from session state (e.g. after a fresh save
    # changed the data bounds) before we render the widget.
    prev = st.session_state.get("dash_dates")
    if isinstance(prev, tuple) and len(prev) == 2:
        lo, hi = prev
        if lo < min_d or hi > max_d or lo > hi:
            st.session_state["dash_dates"] = (min_d, max_d)

    with st.expander("🔎 Filters", expanded=False):
        fc1, fc2, fc3 = st.columns(3)
        with fc1:
            date_range = st.date_input(
                "Date range", value=(min_d, max_d),
                min_value=min_d, max_value=max_d, key="dash_dates",
            )
            sel_years = st.multiselect("Year", opts["years"], default=[],
                                         placeholder="All years", key="dash_years")
        with fc2:
            sel_retailers = st.multiselect("Retailer", opts["retailers"], default=[],
                                              placeholder="All retailers", key="dash_ret")
            sel_finance = st.multiselect("Finance Company", opts["finance_cos"], default=[],
                                            placeholder="All finance companies",
                                            key="dash_fin")
        with fc3:
            sel_mfg = st.multiselect("Manufacturer", opts["manufacturers"], default=[],
                                        placeholder="All manufacturers", key="dash_mfg")
            sel_growers = st.multiselect("Grower", opts["growers"], default=[],
                                            placeholder="All growers (type to search)",
                                            key="dash_grow")
        sel_products = st.multiselect(
            "Product", opts["products"], default=[],
            placeholder=f"All products ({len(opts['products'])} available — type to search)",
            key="dash_prod",
        )

    # Apply filters
    df = df_sheet1.copy()
    df["__retailer"] = df["Retailer Name"].apply(lambda r: norm.normalize_retailer(r) or r)
    df["__finance"] = df["Finance Company"].apply(lambda f: norm.normalize_finance_company(f) or f)
    df["__grower"] = df.apply(_grower_label, axis=1).replace("", pd.NA)

    if isinstance(date_range, tuple) and len(date_range) == 2:
        df = df[df["Invoice Date"].between(pd.Timestamp(date_range[0]),
                                            pd.Timestamp(date_range[1]) + pd.Timedelta(days=1))]
    if sel_years:
        df = df[df["Invoice Date"].dt.year.isin(sel_years)]
    if sel_retailers:
        df = df[df["__retailer"].isin(sel_retailers)]
    if sel_finance:
        df = df[df["__finance"].isin(sel_finance)]
    if sel_mfg:
        df = df[df["Manufacturer Name"].isin(sel_mfg)]
    if sel_growers:
        df = df[df["__grower"].isin(sel_growers)]
    if sel_products:
        df = df[df["Item Description/Brand"].isin(sel_products)]

    # KPI strip
    st.markdown("#### Overview")
    k1, k2, k3, k4, k5, k6 = st.columns(6)
    total_sales = float(df["Sum Total Price"].sum())
    n_invoices = df["Invoice Number"].nunique()
    n_growers = df["__grower"].nunique()
    avg_invoice = total_sales / n_invoices if n_invoices else 0
    n_lines = len(df)
    advances = float(df_finance["Amount To Retailer"].sum() if not df_finance.empty else 0)
    theme.kpi(k1, "Total Sales", f"${total_sales:,.0f}")
    theme.kpi(k2, "Invoices", f"{n_invoices:,}")
    theme.kpi(k3, "Line Items", f"{n_lines:,}")
    theme.kpi(k4, "Unique Growers", f"{n_growers:,}", gold=True)
    theme.kpi(k5, "Avg Invoice", f"${avg_invoice:,.0f}")
    theme.kpi(k6, "Advances to Retailer", f"${advances:,.0f}", gold=True)

    # Charts
    import plotly.express as px
    px.defaults.color_discrete_sequence = theme.PLOTLY_COLORWAY

    st.markdown("#### Trends")
    c1, c2 = st.columns(2)
    monthly = (df.dropna(subset=["Invoice Date"])
               .assign(month=lambda d: d["Invoice Date"].dt.to_period("M").dt.to_timestamp())
               .groupby("month", as_index=False)["Sum Total Price"].sum())
    # Plotly tickformat ",.0f" gives "$1,234,567" instead of "1.23M". Apply
    # to every $-axis chart so the dashboard reads consistently.
    _CURRENCY_AXIS = dict(tickprefix="$", tickformat=",.0f")
    _CURRENCY_HOVER = "$%{y:,.2f}"
    _CURRENCY_HOVER_X = "$%{x:,.2f}"

    if not monthly.empty:
        fig = px.line(monthly, x="month", y="Sum Total Price",
                       markers=True, title="Monthly Revenue")
        fig.update_traces(line_color=theme.KAS_GREEN, marker_color=theme.KAS_GOLD,
                          hovertemplate="%{x|%b %Y}<br>" + _CURRENCY_HOVER + "<extra></extra>")
        fig.update_yaxes(**_CURRENCY_AXIS)
        c1.plotly_chart(fig, use_container_width=True)

    by_ret = (df.groupby("__retailer", as_index=False)["Sum Total Price"].sum()
              .sort_values("Sum Total Price", ascending=False))
    if not by_ret.empty:
        fig = px.bar(by_ret, x="__retailer", y="Sum Total Price",
                     title="Revenue by Retailer Location",
                     color="__retailer")
        fig.update_traces(hovertemplate="%{x}<br>" + _CURRENCY_HOVER + "<extra></extra>")
        fig.update_yaxes(**_CURRENCY_AXIS)
        fig.update_layout(showlegend=False)
        c2.plotly_chart(fig, use_container_width=True)

    c3, c4 = st.columns(2)
    by_mfg = (df.groupby("Manufacturer Name", as_index=False)["Sum Total Price"].sum()
              .sort_values("Sum Total Price", ascending=False).head(15))
    if not by_mfg.empty:
        fig = px.pie(by_mfg, names="Manufacturer Name", values="Sum Total Price",
                      title="Revenue by Manufacturer", hole=0.4)
        fig.update_traces(hovertemplate="%{label}<br>$%{value:,.2f} (%{percent})<extra></extra>")
        c3.plotly_chart(fig, use_container_width=True)

    by_fin = (df.groupby("__finance", as_index=False)["Sum Total Price"].sum())
    if not by_fin.empty:
        fig = px.pie(by_fin, names="__finance", values="Sum Total Price",
                      title="Finance Company Mix", hole=0.4)
        fig.update_traces(hovertemplate="%{label}<br>$%{value:,.2f} (%{percent})<extra></extra>")
        c4.plotly_chart(fig, use_container_width=True)

    c5, c6 = st.columns(2)
    top_growers = (df.groupby("__grower", as_index=False)["Sum Total Price"].sum()
                   .sort_values("Sum Total Price", ascending=True).tail(15))
    if not top_growers.empty:
        fig = px.bar(top_growers, x="Sum Total Price", y="__grower", orientation="h",
                     title="Top 15 Growers by Spend")
        fig.update_traces(marker_color=theme.KAS_GREEN,
                          hovertemplate="%{y}<br>" + _CURRENCY_HOVER_X + "<extra></extra>")
        fig.update_xaxes(**_CURRENCY_AXIS)
        c5.plotly_chart(fig, use_container_width=True)

    top_products = (df.groupby("Item Description/Brand", as_index=False)["Sum Total Price"].sum()
                    .sort_values("Sum Total Price", ascending=True).tail(20))
    if not top_products.empty:
        fig = px.bar(top_products, x="Sum Total Price", y="Item Description/Brand", orientation="h",
                     title="Top 20 Products by Revenue")
        fig.update_traces(marker_color=theme.KAS_GOLD,
                          hovertemplate="%{y}<br>" + _CURRENCY_HOVER_X + "<extra></extra>")
        fig.update_xaxes(**_CURRENCY_AXIS)
        c6.plotly_chart(fig, use_container_width=True)

    # Detail tables
    st.markdown("#### Grower summary")
    grower_summary = (df.groupby("__grower")
                      .agg(total_spend=("Sum Total Price", "sum"),
                           n_line_items=("Invoice Number", "size"),
                           n_invoices=("Invoice Number", "nunique"),
                           last_purchase=("Invoice Date", "max"))
                      .sort_values("total_spend", ascending=False)
                      .reset_index()
                      .rename(columns={"__grower": "Grower"}))
    # Use pandas Styler.format() for thousand-separator commas. Streamlit
    # Cloud's NumberColumn JS printf doesn't accept the %, flag, but Styler
    # uses Python's format spec which does. Numeric sorting still works
    # because Styler keeps the underlying values numeric.
    grower_styled = grower_summary.style.format({
        "total_spend": "${:,.2f}",
        "n_line_items": "{:,}",
        "n_invoices": "{:,}",
        "last_purchase": lambda v: v.strftime("%Y-%m-%d") if pd.notnull(v) else "",
    })
    st.dataframe(
        grower_styled,
        use_container_width=True, hide_index=True,
        column_config={
            "total_spend": st.column_config.Column("Total Spend"),
            "n_line_items": st.column_config.Column("Line Items"),
            "n_invoices": st.column_config.Column("Invoices"),
            "last_purchase": st.column_config.Column("Last Invoice"),
        },
    )

    st.markdown("#### Filtered line items")
    line_items_df = (df[["Invoice Date", "Invoice Number", "__grower", "Item Description/Brand",
                          "Manufacturer Name", "Standard Unit Of Measure", "Quantity", "Sum Total Price",
                          "__retailer", "__finance"]]
                      .rename(columns={"__grower": "Grower", "__retailer": "Retailer",
                                        "__finance": "Finance"})
                      .sort_values("Invoice Date", ascending=False))
    line_items_styled = line_items_df.style.format({
        "Sum Total Price": "${:,.2f}",
        "Quantity": "{:,.2f}",
        "Invoice Number": "{:.0f}",
        "Invoice Date": lambda v: v.strftime("%Y-%m-%d") if pd.notnull(v) else "",
    })
    st.dataframe(
        line_items_styled,
        use_container_width=True, hide_index=True, height=380,
    )

    # Export
    csv = df.drop(columns=["__retailer", "__finance", "__grower"]).to_csv(index=False).encode()
    st.download_button("⬇️ Export filtered to CSV", csv,
                        file_name="kas_filtered.csv", mime="text/csv")


def _canonical_names_by_id(df: pd.DataFrame) -> dict:
    """One DISPLAY label per Grower ID.

    Picks the most common name spelling per ID. If two different IDs share
    the same most-common spelling, both get suffixed with their ID so the
    UI can tell them apart (e.g. "TRIPLE H FARMS (#101137)" vs
    "TRIPLE H FARMS (#102137)"). This guarantees one display name per
    real-world entity.
    """
    base: dict = {}
    for gid, sub in df.dropna(subset=["Grower ID"]).groupby("Grower ID"):
        labels = [s for s in sub.get("__raw_label", []) if pd.notnull(s) and s]
        if labels:
            base[gid] = max(set(labels), key=labels.count)
    name_to_ids: dict = {}
    for gid, name in base.items():
        name_to_ids.setdefault(name, []).append(gid)
    out: dict = {}
    for gid, name in base.items():
        if len(name_to_ids[name]) > 1:
            try:
                out[gid] = f"{name} (#{int(gid)})"
            except (TypeError, ValueError):
                out[gid] = f"{name} (#{gid})"
        else:
            out[gid] = name
    return out


@st.cache_data(ttl=120, show_spinner=False)
def _build_grower_index(_xlsx_bytes: bytes, cache_key: str = "") -> tuple:
    """Pre-compute per-grower line-item subsets, summaries, and PDF lookup.

    Uses CANONICAL names (one canonical spelling per Grower ID) so this
    matches the map's view exactly: same totals, same name strings.
    """
    wb = wb_mod.load(_xlsx_bytes)
    df = wb_mod.read_sheet1_dataframe(wb)
    fdf = wb_mod.read_finance_dataframe(wb)

    df["__raw_label"] = df.apply(_grower_label, axis=1).replace("", pd.NA)
    canonical = _canonical_names_by_id(df)
    df["__grower"] = df.apply(
        lambda r: canonical.get(r["Grower ID"], r["__raw_label"]),
        axis=1,
    )
    df = df.dropna(subset=["__grower"])

    # PDF lookup keyed by invoice number string.
    pdf_lookup: dict[str, str] = {}
    if not fdf.empty and "PDF Drive ID" in fdf.columns:
        for _, r in fdf.iterrows():
            inv = r.get("Invoice Number")
            pid = r.get("PDF Drive ID")
            if pd.notnull(inv) and pid:
                pdf_lookup[str(int(inv))] = str(pid)

    subs: dict = {}
    summaries: dict = {}
    for grower, sub in df.groupby("__grower"):
        sub = sub.sort_values("Invoice Date", ascending=False)
        subs[grower] = sub
        s = (sub.groupby(["Invoice Date", "Invoice Number"])
              .agg(line_items=("Item Description/Brand", "size"),
                   total=("Sum Total Price", "sum"),
                   retailer=("Retailer Name", "first"),
                   finance=("Finance Company", "first"))
              .reset_index()
              .sort_values("Invoice Date", ascending=False))
        def _pdf_url(n):
            if not pd.notnull(n):
                return None
            val = pdf_lookup.get(str(int(n)))
            if not val:
                return None
            # GCS URLs are stored as full https:// URLs; legacy entries are Drive file IDs.
            return val if val.startswith("https://") else f"https://drive.google.com/file/d/{val}/view"
        s["PDF"] = s["Invoice Number"].apply(_pdf_url)
        summaries[grower] = s
    return subs, summaries


def _grower_detail(grower_name: str) -> tuple:
    """Cheap lookup using the precomputed index."""
    subs, summaries = _build_grower_index(_bytes, cache_key=_mtime)
    return subs.get(grower_name), summaries.get(grower_name)


def _grower_label(r) -> str:
    def _clean(v):
        # Treat NaN/None/empty/whitespace as missing.
        if v is None or (isinstance(v, float) and pd.isna(v)):
            return ""
        s = str(v).strip()
        if not s or s.lower() == "nan":
            return ""
        return s.upper()

    company = _clean(r.get("Grower Company Name"))
    if company:
        return company
    fn = _clean(r.get("Grower First Name"))
    ln = _clean(r.get("Grower Last Name"))
    if ln and fn:
        return f"{ln}, {fn}"
    return ln or fn or ""


if ENABLE_DASHBOARD and active_page == "📊 Dashboard":
    try:
        _render_dashboard()
    except Exception as _e:
        _diag_error("Dashboard render crashed", _e)
        st.error(
            "Dashboard rendering hit an unexpected error. The app is still "
            "running — open the 🔬 Diagnostics panel in the sidebar and "
            "send the log so we can fix it."
        )


# ===========================================================================
# Tab 3 — Grower Map
# ===========================================================================
def _render_map():
    import folium
    from streamlit_folium import st_folium

    _diag("_render_map() called", level="RENDER")
    _map_render_start = _time.perf_counter()
    if df_sheet1.empty:
        st.info("No data yet.")
        return

    # Consume any pending grower selection from a prior marker click. We seed
    # the widget state HERE (before the dropdown renders) instead of mutating
    # widget keys after the widget has rendered — the latter pattern caused
    # intermittent "click did nothing" bugs because Streamlit's widget state
    # machine doesn't reliably handle mid-run del + rerun.
    pending = st.session_state.pop("_map_pending_grower", None)
    if pending is not None:
        _diag(f"consuming pending grower selection: {pending!r}")
        st.session_state["map_search_grower"] = pending
        st.session_state["map_selected_grower"] = pending
        # Set the dropdown's stored value DIRECTLY (don't pop). Popping the
        # key forced the widget to re-initialize, which in some cases
        # appears to trigger the on_change callback with a stale value
        # and reset map_selected_grower mid-flight. Direct assignment
        # before the widget renders is safe.
        st.session_state["map_search_select"] = pending

    # Log current selection state so we can trace what's setting/clearing it.
    _diag(f"render-start: map_selected_grower={st.session_state.get('map_selected_grower')!r} "
          f"map_search_select={st.session_state.get('map_search_select')!r}")

    # Step 1: pre-computed dataframe (canonical grower labels + normalized
    # address columns). Cached on workbook mtime so this only runs when
    # data actually changes.
    with _diag_timing("_prepare_map_dataframe"):
        df = _prepare_map_dataframe(_bytes, _mtime).copy()

    # ---- Step 2: geocode every distinct address string FIRST ----
    addr_keys = df[["__addr1", "__city", "__state", "__zip"]].drop_duplicates()
    coords_for: dict[tuple, tuple[float, float] | None] = {}
    needs_geocode = []
    for _, ar in addr_keys.iterrows():
        key = (ar["__addr1"], ar["__city"], ar["__state"], ar["__zip"])
        coords_for[key] = None  # placeholder
        needs_geocode.append((key, ar))

    if st.button("🌐 Geocode addresses", type="primary",
                  help="Run once after adding new growers."):
        bar = st.progress(0.0, text="Geocoding...")
        for i, (key, ar) in enumerate(needs_geocode):
            coords_for[key] = geocode.geocode_address(
                ar["__addr1"], ar["__city"],
                ar["__state"] or "KY", ar["__zip"],
            )
            bar.progress((i + 1) / len(needs_geocode),
                          text=f"Geocoded {i + 1}/{len(needs_geocode)}")
        bar.empty()
        st.session_state.pop("_coords_cache", None)
        st.rerun()

    # Pull cached coordinates (no API calls if already cached).
    # Memoize the result in session state keyed on the address-set hash so we
    # don't re-iterate every keystroke. Only invalidates when the underlying
    # address set changes.
    addr_hash = hash(tuple(sorted(coords_for.keys())))
    cached = st.session_state.get("_coords_cache")
    if cached and cached.get("hash") == addr_hash:
        coords_for.update(cached["coords"])
    else:
        for key, ar in needs_geocode:
            coords_for[key] = geocode.geocode_address(
                ar["__addr1"], ar["__city"],
                ar["__state"] or "KY", ar["__zip"],
            )
        st.session_state["_coords_cache"] = {"hash": addr_hash, "coords": dict(coords_for)}

    # ---- Step 3: group rows by their geocoded coordinate (rounded ~10m) ----
    def _coord_key(r):
        c = coords_for.get((r["__addr1"], r["__city"], r["__state"], r["__zip"]))
        if not c:
            return None
        return (round(c[0], 4), round(c[1], 4))

    df["__coord_key"] = df.apply(_coord_key, axis=1)
    geocoded = df[df["__coord_key"].notnull()]
    missing = df[df["__coord_key"].isnull()]

    # ONE marker per coordinate. Multi-grower locations get a multi-card
    # popup with ‹/› arrows; clicking the marker loads the highest-spender's
    # table below.
    locations = (
        geocoded.groupby("__coord_key")
        .agg(total_spend=("Sum Total Price", "sum"),
              n_invoices=("Invoice Number", "nunique"),
              last_invoice=("Invoice Date", "max"),
              grower_ids=("Grower ID", lambda s: sorted({x for x in s if pd.notnull(x)})),
              addr1=("Grower Address1", "first"),
              city=("Grower City", "first"),
              state=("Grower State", "first"),
              zip_code=("Grower ZIP CODE", "first"))
        .reset_index()
    )

    # The search box is the single source of truth for which grower's details
    # appear below the map. The map itself just shows the dots — clicks open a
    # popup with summary info, but the table load happens here.
    search_options = sorted({n for n in df["__grower"].dropna().unique() if n})
    # Initialize the selectbox's stored value once (so we don't fight other
    # widgets that update map_search_grower programmatically).
    options_with_blank = ["— none selected —"] + search_options
    if "map_search_select" not in st.session_state:
        cur = st.session_state.get("map_search_grower")
        st.session_state["map_search_select"] = (
            cur if cur in search_options else "— none selected —"
        )

    def _on_search_change():
        v = st.session_state["map_search_select"]
        _diag(f"dropdown on_change fired: new value={v!r}", level="CLICK")
        if v == "— none selected —":
            st.session_state["map_search_grower"] = None
            st.session_state["map_selected_grower"] = None
        else:
            st.session_state["map_search_grower"] = v
            st.session_state["map_selected_grower"] = v

    st.selectbox(
        "🔎 Pick a grower to view their invoice history",
        options_with_blank,
        key="map_search_select",
        on_change=_on_search_change,
        help="Type any part of the name to filter. Selection drives the table below the map.",
    )

    # Pre-compute "growers at each location" for the multi-grower radio.
    growers_by_coord = {}
    for _ck, _grp in geocoded.dropna(subset=["__grower"]).groupby("__coord_key"):
        growers_by_coord[_ck] = _grp["__grower"].unique().tolist()

    # Grower-level totals (used by detail panel and map markers).
    grower_totals = (
        df.dropna(subset=["__grower"])
        .groupby("__grower")
        .agg(spend=("Sum Total Price", "sum"),
              invoices=("Invoice Number", "nunique"),
              last=("Invoice Date", "max"))
    )

    # ---- Detail panel (above the map so user doesn't have to scroll) ----
    selected = st.session_state.get("map_selected_grower")

    # Multi-grower radio: shown when the last marker click had >1 grower.
    last_sig = st.session_state.get("_last_marker_click_sig")
    growers_at_last_click: list[str] = []
    if last_sig and last_sig in growers_by_coord:
        names_at = growers_by_coord[last_sig]
        if len(names_at) > 1:
            growers_at_last_click = (
                grower_totals.loc[grower_totals.index.intersection(names_at)]
                .sort_values("spend", ascending=False)
                .index.tolist()
            )

    if growers_at_last_click:
        st.markdown(f"#### {len(growers_at_last_click)} growers at this address")
        radio_key = f"multi_grower_radio_{last_sig}"
        if radio_key not in st.session_state:
            st.session_state[radio_key] = (
                selected if selected in growers_at_last_click
                else growers_at_last_click[0]
            )

        def _on_radio_change():
            new_val = st.session_state[radio_key]
            st.session_state["map_selected_grower"] = new_val
            st.session_state["map_search_grower"] = new_val
            st.session_state["map_search_select"] = new_val

        st.radio(
            "Pick which grower to view",
            growers_at_last_click,
            horizontal=True,
            label_visibility="collapsed",
            key=radio_key,
            on_change=_on_radio_change,
        )
        selected = st.session_state.get("map_selected_grower")

    if selected:
        st.divider()
        st.markdown(f"### 📍 {selected}")
        try:
            sub, inv_summary = _grower_detail(selected)
        except Exception as _e:
            st.error(f"Error loading grower detail: {_e}")
            sub, inv_summary = None, None
        if sub is None or sub.empty:
            st.info("No invoices found. Try 🔄 Reload from Drive in the sidebar.")
        else:
            cols = st.columns(4)
            cols[0].metric("Total spend", f"${sub['Sum Total Price'].sum():,.2f}")
            cols[1].metric("Invoices", f"{sub['Invoice Number'].nunique()}")
            cols[2].metric("Line items", f"{len(sub)}")
            last = sub["Invoice Date"].max()
            cols[3].metric("Last purchase",
                            last.strftime("%Y-%m-%d") if pd.notnull(last) else "—")

            st.markdown("##### Invoice history")
            inv_summary_styled = inv_summary.style.format({
                "total": "${:,.2f}",
                "Invoice Number": "{:.0f}",
                "Invoice Date": lambda v: v.strftime("%Y-%m-%d") if pd.notnull(v) else "",
            })
            st.dataframe(
                inv_summary_styled,
                use_container_width=True, hide_index=True,
                column_config={
                    "PDF": st.column_config.LinkColumn(
                        "PDF",
                        display_text="📄 Open",
                        help="Open the source PDF in Google Drive.",
                    ),
                },
            )

            st.markdown("##### All line items")
            line_items_sub = sub[["Invoice Date", "Invoice Number", "Item Description/Brand",
                                   "Manufacturer Name", "Standard Unit Of Measure", "Quantity",
                                   "Sum Total Price", "Retailer Name", "Finance Company"]]
            line_items_sub_styled = line_items_sub.style.format({
                "Sum Total Price": "${:,.2f}",
                "Quantity": "{:,.2f}",
                "Invoice Number": "{:.0f}",
                "Invoice Date": lambda v: v.strftime("%Y-%m-%d") if pd.notnull(v) else "",
            })
            st.dataframe(
                line_items_sub_styled,
                use_container_width=True, hide_index=True, height=300,
            )
        st.divider()

    st.markdown("#### 🗺️ Map — click a marker to load that grower above")
    # Map view: remember whatever the user last panned/zoomed to so reruns
    # don't reset the view. Only on first visit (no saved state) we use the
    # western-Kentucky default.
    map_center = st.session_state.get("map_center") or [37.3, -87.5]
    map_zoom = st.session_state.get("map_zoom") or 8

    # IMPORTANT: do NOT cache the folium Map across reruns. Re-using the same
    # Python Map object with st_folium can suppress click-event propagation
    # (the iframe thinks nothing changed). We accept a fresh build per render
    # in exchange for reliable click handling.
    SPEND_TIERS = [
        ("#4A148C", "$100,000+",       100_000),  # deep purple
        ("#1565C0", "$50,000 – $100k",  50_000),  # blue
        ("#2E7D32", "$20,000 – $50k",   20_000),  # green
        ("#EF6C00", "$5,000 – $20k",     5_000),  # orange
        ("#C62828", "Under $5,000",          0),  # red
    ]

    if False:  # cache disabled — see comment above
        m = None
    else:
        # Map without a default tile layer — we add all four below so we control
        # both the initial view (`show=True`) and the labels in the layer control.
        m = folium.Map(
            location=map_center, zoom_start=map_zoom,
            tiles=None,
            control_scale=True,
        )
        folium.TileLayer(
            tiles="https://mt1.google.com/vt/lyrs=m&x={x}&y={y}&z={z}",
            attr="Google", name="Map", overlay=False, control=True, show=True,
        ).add_to(m)
        folium.TileLayer(
            tiles="https://mt1.google.com/vt/lyrs=s&x={x}&y={y}&z={z}",
            attr="Google", name="Satellite", overlay=False, control=True, show=False,
        ).add_to(m)
        folium.TileLayer(
            tiles="https://mt1.google.com/vt/lyrs=p&x={x}&y={y}&z={z}",
            attr="Google", name="Terrain", overlay=False, control=True, show=False,
        ).add_to(m)
        folium.TileLayer(
            tiles="https://mt1.google.com/vt/lyrs=y&x={x}&y={y}&z={z}",
            attr="Google", name="Hybrid", overlay=False, control=True, show=False,
        ).add_to(m)

        def _tier_color(spend: float) -> str:
            for color, _label, threshold in SPEND_TIERS:
                if spend >= threshold:
                    return color
            return SPEND_TIERS[-1][0]

        plotted = 0
        MARKER_RADIUS = 9

        for _, row in locations.iterrows():
            lat, lon = row["__coord_key"]
            spend = float(row["total_spend"])
            marker_color = _tier_color(spend)
            last = pd.to_datetime(row["last_invoice"]).strftime("%Y-%m-%d") if pd.notnull(row["last_invoice"]) else "—"

            # Distinct growers at this location, ordered by their TOTAL spend
            # (across all their addresses). Stats shown in the popup match what
            # the table below the map will show when this grower is selected.
            loc_growers = (
                geocoded[geocoded["__coord_key"] == row["__coord_key"]]
                .dropna(subset=["__grower"])["__grower"].unique().tolist()
            )
            per_grower = (
                grower_totals.loc[grower_totals.index.intersection(loc_growers)]
                .sort_values("spend", ascending=False)
                .reset_index()
            )
            if per_grower.empty:
                per_grower = pd.DataFrame([{
                    "__grower": "(unknown grower)",
                    "spend": spend,
                    "invoices": row["n_invoices"],
                    "last": row["last_invoice"],
                }])

            names = per_grower["__grower"].tolist()
            primary = names[0]
            tooltip_label = primary if len(names) == 1 else f"{len(names)} growers"

            zip_str = row["zip_code"] if pd.notnull(row["zip_code"]) else ""
            zip_str = str(int(zip_str)) if zip_str and str(zip_str).replace(".0", "").isdigit() else (zip_str or "")
            addr_block = (
                f"<i style='color:#555'>{row['addr1']}<br>"
                f"{row['city']}, {row['state']} {zip_str}</i>"
            )

            # Attribute-escape grower name for the data-grower attribute (HTML
            # quotes / ampersands).
            import html as _html
            cards_html = ""
            for idx, g in per_grower.iterrows():
                g_last = pd.to_datetime(g["last"]).strftime("%Y-%m-%d") if pd.notnull(g["last"]) else "—"
                display = "block" if idx == 0 else "none"
                grower_attr = _html.escape(str(g["__grower"]), quote=True)
                cards_html += (
                    f"<div class='kas-grower-card' data-grower=\"{grower_attr}\" "
                    f"style='display:{display}'>"
                    f"<div style='font-weight:700;font-size:14px;color:#1B5E20;"
                    f"margin-bottom:4px'>{g['__grower']}</div>"
                    f"<div style='font-size:12px;line-height:1.6'>"
                    f"<b>Total Spend:</b> ${float(g['spend']):,.2f}<br>"
                    f"<b>Invoices:</b> {int(g['invoices'])}<br>"
                    f"<b>Last Purchase:</b> {g_last}"
                    f"</div></div>"
                )

            # Popup: show ALL growers at this location at once (stacked) + a hint
            # that they can switch the table view via the radio below the map for
            # multi-grower spots. No popup arrows, no JS bridging.
            if len(per_grower) == 1:
                popup_html = (
                    f"<div style='min-width:240px;font-family:sans-serif'>"
                    f"{cards_html}<div style='margin-top:10px'>{addr_block}</div>"
                    f"</div>"
                )
            else:
                # Show all cards stacked with a separator between each.
                stacked = ""
                for idx, g in per_grower.iterrows():
                    g_last = pd.to_datetime(g["last"]).strftime("%Y-%m-%d") if pd.notnull(g["last"]) else "—"
                    stacked += (
                        f"<div style='padding:6px 0;"
                        f"{'border-bottom:1px solid #eee;' if idx < len(per_grower)-1 else ''}'>"
                        f"<div style='font-weight:700;font-size:13px;color:#1B5E20'>"
                        f"{g['__grower']}</div>"
                        f"<div style='font-size:11px;color:#444;line-height:1.5'>"
                        f"<b>Total:</b> ${float(g['spend']):,.2f} · "
                        f"<b>Invoices:</b> {int(g['invoices'])} · "
                        f"<b>Last:</b> {g_last}"
                        f"</div></div>"
                    )
                popup_html = (
                    f"<div style='min-width:280px;font-family:sans-serif'>"
                    f"<div style='font-size:11px;color:#6D4C00;background:#fff8e1;"
                    f"padding:5px 8px;border-radius:4px;margin-bottom:8px;"
                    f"text-align:center'>"
                    f"📍 <b>{len(per_grower)} growers</b> at this address"
                    f"</div>"
                    f"{stacked}"
                    f"<div style='margin-top:10px;padding-top:8px;"
                    f"border-top:1px solid #eee'>{addr_block}</div>"
                    f"</div>"
                )

            # Outer dark halo: keeps dots visible on bright satellite imagery.
            folium.CircleMarker(
                location=[lat, lon], radius=MARKER_RADIUS + 2,
                color="#222", fill=False, weight=1, opacity=0.7,
            ).add_to(m)
            folium.CircleMarker(
                location=[lat, lon], radius=MARKER_RADIUS,
                color="white", fill=True, fill_color=marker_color,
                fill_opacity=1.0, weight=2.5,
                popup=folium.Popup(popup_html, max_width=320),
                tooltip=f"{tooltip_label} — ${spend:,.0f}",
            ).add_to(m)
            plotted += 1

        # Legend (HTML overlay, bottom-right).
        legend_rows = "".join(
            f'<div style="display:flex;align-items:center;margin:3px 0;">'
            f'<span style="display:inline-block;width:14px;height:14px;'
            f'border-radius:50%;background:{c};border:2px solid white;'
            f'box-shadow:0 0 0 1px #999;margin-right:8px;"></span>'
            f'<span style="font-size:12px;color:#222;">{label}</span>'
            f'</div>'
            for c, label, _ in SPEND_TIERS
        )
        legend_html = (
            '<div style="position: fixed; bottom: 30px; right: 18px; z-index: 9999; '
            'background: white; padding: 10px 14px; border: 1px solid #ccc; '
            'border-radius: 8px; box-shadow: 0 2px 8px rgba(0,0,0,0.15); '
            'font-family: sans-serif;">'
            '<div style="font-size:11px;font-weight:700;color:#666;'
            'text-transform:uppercase;letter-spacing:0.5px;margin-bottom:6px;">'
            'Total Spend</div>'
            f'{legend_rows}</div>'
        )
        m.get_root().html.add_child(folium.Element(legend_html))

        folium.LayerControl(position="topright", collapsed=False).add_to(m)

        # Cache disabled (see top of map-build block) — clicks need fresh map.
        pass


    # Only capture marker clicks. NOT zoom/center — including those would
    # fire a rerun on every pan/zoom, making the map painfully laggy.
    map_state = st_folium(
        m, width=None, height=600,
        returned_objects=["last_object_clicked"],
        key="kas_map", use_container_width=True,
    )

    # Marker click handling. We track BOTH the last-clicked sig (so the
    # multi-grower radio knows which location to show alternates for) AND
    # whether the resulting selection would change (so same-marker re-clicks
    # are no-ops without consuming a dedup slot).
    clicked = (map_state or {}).get("last_object_clicked")
    if clicked and isinstance(clicked, dict) and "lat" in clicked:
        sig = (round(clicked["lat"], 4), round(clicked["lng"], 4))
        growers_here = (
            geocoded[geocoded["__coord_key"] == sig]
            .dropna(subset=["__grower"])["__grower"].unique().tolist()
        )
        ordered = (
            grower_totals.loc[grower_totals.index.intersection(growers_here)]
            .sort_values("spend", ascending=False)
        )
        _diag(f"marker click sig={sig} growers={len(growers_here)} ordered={len(ordered)}")
        if len(ordered):
            primary = ordered.index[0]
            sel_changed = st.session_state.get("map_selected_grower") != primary
            sig_changed = st.session_state.get("_last_marker_click_sig") != sig
            if sel_changed or sig_changed:
                st.session_state["_last_marker_click_sig"] = sig
                st.session_state["_map_pending_grower"] = primary
                st.session_state["_scroll_to_details"] = True
                _diag(f"  → loading grower={primary!r} (sel_changed={sel_changed}, sig_changed={sig_changed})")
                st.rerun()
            else:
                _diag(f"  → no-op (already selected, same sig)")

    if len(missing) > 0:
        st.caption(f"{plotted} locations on map. "
                    f"{missing['Grower ID'].nunique()} grower(s) couldn't be placed "
                    f"(usually a typo'd address).")
    else:
        st.caption(f"{plotted} locations on map.")

    _diag(f"_render_map total: {(_time.perf_counter() - _map_render_start) * 1000:.0f}ms",
          level="TIMING")


if ENABLE_MAP and active_page == "🗺️ Grower Map":
    try:
        _render_map()
    except Exception as _e:
        _diag_error("Map render crashed", _e)
        st.error(
            "Map rendering hit an unexpected error. The app is still "
            "running — open the 🔬 Diagnostics panel in the sidebar and "
            "send the log so we can fix it."
        )
