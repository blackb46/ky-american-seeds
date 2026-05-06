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


def reload_workbook() -> tuple[bytes, str]:
    """Return (workbook bytes, modifiedTime). modifiedTime acts as the
    optimistic-lock baseline for the next save."""
    file_id = st.secrets["GOOGLE_DRIVE_FILE_ID"]
    meta = drive.file_metadata(file_id)
    mtime = meta["modifiedTime"]
    data = fetch_workbook_bytes(file_id, _cache_key=mtime)
    st.session_state["wb_modified_time"] = mtime
    return data, mtime


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


# ---------------------------------------------------------------------------
# Top header
# ---------------------------------------------------------------------------
# ---------------------------------------------------------------------------
# Auth (rendered before the header so the login is at the top of the page)
# ---------------------------------------------------------------------------
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
    try:
        _bytes, _mtime = reload_workbook()
        meta = drive.file_metadata(st.secrets["GOOGLE_DRIVE_FILE_ID"])
        st.caption(f"**{meta.get('name', '?')}**")
        st.caption(f"Updated: {meta.get('modifiedTime', '?')[:19].replace('T', ' ')}")
        st.caption(f"Size: {int(meta.get('size', 0)) / 1024:.1f} KB")
    except Exception as e:
        st.error(f"Drive load failed: {e}")
        st.stop()

    file_id = st.secrets["GOOGLE_DRIVE_FILE_ID"]
    st.link_button(
        "📂 Open in Google Sheets",
        f"https://docs.google.com/spreadsheets/d/{file_id}/edit",
        use_container_width=True,
    )

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
    if st.button("🚪 Sign out", use_container_width=True):
        auth.logout()


# ---------------------------------------------------------------------------
# Workbook → DataFrames (cached)
# ---------------------------------------------------------------------------
@st.cache_data(ttl=120)
def _dataframes(_xlsx: bytes) -> tuple[pd.DataFrame, pd.DataFrame]:
    wb = wb_mod.load(_xlsx)
    return wb_mod.read_sheet1_dataframe(wb), wb_mod.read_finance_dataframe(wb)


df_sheet1, df_finance = _dataframes(_bytes)


# ---------------------------------------------------------------------------
# Navigation (st.tabs resets on every rerun, dropping the user back to the
# Upload tab whenever any widget triggers a refresh. A radio keeps state.)
# ---------------------------------------------------------------------------
PAGES = ["📄 Upload & Process", "📊 Dashboard", "🗺️ Grower Map"]
active_page = st.radio(
    "Section", PAGES, horizontal=True, label_visibility="collapsed",
    key="active_page",
)
st.divider()



# ===========================================================================
# Tab 1 — Upload & Process
# ===========================================================================
def _conf(level: str | None) -> str:
    return theme.confidence_pill(level)


def _render_review_form(idx: int, pdf_name: str, pdf_bytes: bytes,
                        result: dict, existing_invoice_nums: set[str]):
    """Render an editable review form for one extracted PDF."""
    final = result["final"]
    pass1 = result["pass1"]
    invoices = final.get("invoices", [])

    st.markdown(f"### 📄 {pdf_name}")
    cols_top = st.columns([2, 5])
    with cols_top[0]:
        with st.expander("PDF preview", expanded=True):
            try:
                doc_pages = pdf_render.render_pdf_thumbnail(pdf_bytes, page=0, dpi=120)
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
        invoice_no = str(inv.get("invoice_number") or "?").strip()
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

            # Duplicate handling: offer to skip, or attach the PDF to the
            # existing invoice (no row writes, just a Drive link added to the
            # Finance Details sheet so the map can show the source PDF).
            if is_dup and st.session_state.get(f"{state_key}_done") not in ("saved", "skipped", "attached"):
                st.warning(
                    f"**Invoice {invoice_no} is already in the spreadsheet.** "
                    "Choose an action below — uploading line items again would "
                    "create duplicate rows. You can either skip, or attach this "
                    "PDF to the existing invoice so it's clickable from the map."
                )
                d1, d2, d3 = st.columns([1, 1, 3])
                if d1.button("⏭️ Skip this invoice",
                              key=f"{state_key}_dup_skip"):
                    st.session_state[f"{state_key}_done"] = "skipped"
                    st.rerun()
                if d2.button("📎 Attach PDF only",
                              type="primary",
                              key=f"{state_key}_dup_attach",
                              help="Uploads this PDF to Drive and links it to the existing invoice. No spreadsheet changes."):
                    _attach_pdf_to_existing(invoice_no, pdf_name, pdf_bytes, inv)
                    st.session_state[f"{state_key}_done"] = "attached"
                    st.rerun()
                # Still allow review-and-save in case user wants to override,
                # but make the duplicate state visible — they can scroll past
                # the warning to see the form below.

            # Validate
            vr = validate.validate_invoice(inv)
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

            st.markdown("##### Line items (edit, add rows, or delete)")
            li_df = pd.DataFrame(inv.get("line_items") or [])
            for col in ("item_number", "description", "unit", "manufacturer"):
                if col not in li_df.columns:
                    li_df[col] = None
            for col in ("quantity", "unit_price", "ext_amount"):
                if col not in li_df.columns:
                    li_df[col] = None
            li_df = li_df[["item_number", "description", "unit", "quantity",
                           "unit_price", "ext_amount", "manufacturer"]]
            edited_li = st.data_editor(
                li_df,
                num_rows="dynamic",
                use_container_width=True,
                key=f"{state_key}_li",
                column_config={
                    "quantity": st.column_config.NumberColumn(format="%.2f"),
                    "unit_price": st.column_config.NumberColumn(format="$%.2f"),
                    "ext_amount": st.column_config.NumberColumn(format="$%.2f"),
                },
            )
            inv["line_items"] = edited_li.to_dict(orient="records")

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

            # Action buttons
            b1, b2, b3, b4 = st.columns([1, 1, 1, 3])
            approve = b1.button("✅ Approve & Save",
                                 type="primary",
                                 disabled=not vr.passes,
                                 key=f"{state_key}_approve")
            skip = b2.button("⏭️ Skip", key=f"{state_key}_skip")
            view_pdf = b3.button("👀 View full PDF", key=f"{state_key}_view")
            if view_pdf:
                with st.expander("Full PDF pages", expanded=True):
                    blocks = pdf_render.render_pdf_pages(pdf_bytes, dpi=110)
                    for i, blk in enumerate(blocks):
                        import base64
                        st.image(base64.b64decode(blk["source"]["data"]),
                                 caption=f"Page {i+1}", use_container_width=True)

            if approve:
                _save_invoice(inv, pdf_name, pdf_bytes)
                st.session_state[f"{state_key}_done"] = "saved"
                st.rerun()
            if skip:
                st.session_state[f"{state_key}_done"] = "skipped"

        if st.session_state.get(f"{state_key}_done") == "saved":
            st.success(f"✅ Invoice {invoice_no} saved.")
        elif st.session_state.get(f"{state_key}_done") == "skipped":
            st.info(f"⏭️ Invoice {invoice_no} skipped.")
        elif st.session_state.get(f"{state_key}_done") == "attached":
            st.success(f"📎 PDF attached to existing invoice {invoice_no}. "
                        "It's now clickable from the map.")


def _attach_pdf_to_existing(invoice_no: str, pdf_name: str,
                              pdf_bytes: bytes, inv: dict) -> None:
    """Upload PDF to Drive and write/update a Finance Details row that links
    it to an existing invoice. Does NOT touch Sheet1 — the line items are
    already there.
    """
    if _is_test_mode():
        st.toast("TEST MODE — PDF not actually uploaded", icon="🧪")
        return

    drive_id = None
    try:
        folder_name = st.secrets.get("PROCESSED_PDFS_DRIVE_FOLDER_NAME", "kas_processed_pdfs")
        workbook_meta = drive.file_metadata(st.secrets["GOOGLE_DRIVE_FILE_ID"])
        parents = workbook_meta.get("parents", [])
        parent_id = parents[0] if parents else None
        folder_id = drive.find_or_create_folder(folder_name, parent_id=parent_id)
        up = drive.upload_pdf(folder_id, pdf_name, pdf_bytes)
        drive_id = up.get("id")
    except Exception as e:
        st.error(f"PDF upload to Drive failed: {e}")
        return

    # Write a Finance Details row (or update if one already exists) with the
    # PDF Drive ID so the map can link to it.
    try:
        xlsx_bytes, _mtime = reload_workbook()
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
        cells = [
            wb_mod._to_int_or(invoice_no),  # type: ignore[attr-defined]
        ] if False else None  # placeholder; we set cells inline below.

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

        out = wb_mod.save_to_bytes(wb)
        write_workbook(out)
        st.cache_data.clear()
        st.toast("PDF attached and Drive ID saved.", icon="📎")
    except Exception as e:
        st.error(f"Couldn't update Finance Details sheet: {e}")


def _save_invoice(inv: dict, pdf_name: str, pdf_bytes: bytes) -> None:
    """Append invoice to workbook + upload PDF to processed folder."""
    bundles = extract.to_invoice_bundles({"invoices": [inv]}, pdf_filename=pdf_name)
    if not bundles:
        st.error("Nothing to save (no line items).")
        return
    bundle = bundles[0]

    # Optionally upload PDF first to capture Drive ID into the finance row
    drive_id = None
    if not _is_test_mode():
        try:
            folder_name = st.secrets.get("PROCESSED_PDFS_DRIVE_FOLDER_NAME", "kas_processed_pdfs")
            workbook_meta = drive.file_metadata(st.secrets["GOOGLE_DRIVE_FILE_ID"])
            parents = workbook_meta.get("parents", [])
            parent_id = parents[0] if parents else None
            folder_id = drive.find_or_create_folder(folder_name, parent_id=parent_id)
            up = drive.upload_pdf(folder_id, pdf_name, pdf_bytes)
            drive_id = up.get("id")
        except Exception as e:
            st.warning(f"PDF upload failed (continuing without): {e}")
    if bundle.finance is not None and drive_id:
        bundle.finance.pdf_drive_id = drive_id

    # Reload latest workbook bytes (concurrency-safe baseline)
    xlsx_bytes, _mtime = reload_workbook()
    wb = wb_mod.load(xlsx_bytes)
    counts = wb_mod.append_invoice(wb, bundle)
    out = wb_mod.save_to_bytes(wb)
    write_workbook(out)
    st.cache_data.clear()
    st.toast(
        f"Saved {counts['line_items_added']} line items"
        + (f" ({counts['duplicates_skipped']} dupes skipped)" if counts["duplicates_skipped"] else ""),
        icon="✅",
    )

    # Auto-verify data consistency after every save. If anything is off, the
    # user gets an immediate alert instead of discovering it later.
    try:
        report = verify.verify_workbook(out)
        if report.passed:
            st.toast(
                f"Data check passed ({report.app_rows} rows, "
                f"${report.app_total:,.2f} total).",
                icon="🔍",
            )
        else:
            st.error(
                "⚠️ Data consistency check FAILED after save. "
                f"{len(report.failures)} check(s) failed:\n"
                + "\n".join(
                    f"- **{c.name}**: app={c.app_value} vs raw={c.raw_value}"
                    for c in report.failures
                )
            )
    except Exception as e:
        st.warning(f"Data verification skipped: {e}")


if active_page == PAGES[0]:
    st.markdown("Upload one or more invoice PDFs. The app will extract every line item, "
                "verify with a second pass, and let you review/edit before saving.")

    uploaded = st.file_uploader(
        "Drop PDFs here", type=["pdf"], accept_multiple_files=True,
        key="pdf_uploader",
    )

    if uploaded:
        existing_inv_nums = wb_mod.existing_invoice_numbers(wb_mod.load(_bytes))
        api_key = st.secrets["ANTHROPIC_API_KEY"]

        # Run extraction (cached in session_state by file content hash)
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
            result = st.session_state[cache_key]
            _render_review_form(i, uf.name, file_bytes, result, existing_inv_nums)
    else:
        st.info("👆 Upload PDFs to get started.")


# ===========================================================================
# Tab 2 — Dashboard
# ===========================================================================
def _render_dashboard():
    if df_sheet1.empty:
        st.info("No data yet. Upload some PDFs to populate the dashboard.")
        return

    # Sidebar filters live in the main sidebar
    with st.sidebar:
        st.markdown("### 🔎 Dashboard filters")
        df_dates = df_sheet1["Invoice Date"].dropna()
        if len(df_dates):
            min_d, max_d = df_dates.min().date(), df_dates.max().date()
        else:
            min_d = max_d = datetime.now().date()
        date_range = st.date_input(
            "Date range", value=(min_d, max_d), min_value=min_d, max_value=max_d,
            key="dash_dates",
        )
        years = sorted({d.year for d in df_dates})
        sel_years = st.multiselect("Year", years, default=[],
                                     placeholder="All years", key="dash_years")

        retailers = sorted({norm.normalize_retailer(r) or r
                             for r in df_sheet1["Retailer Name"].dropna().unique()})
        sel_retailers = st.multiselect("Retailer", retailers, default=[],
                                          placeholder="All retailers", key="dash_ret")

        finance_cos = sorted({norm.normalize_finance_company(f) or f
                               for f in df_sheet1["Finance Company"].dropna().unique()})
        sel_finance = st.multiselect("Finance Company", finance_cos, default=[],
                                        placeholder="All finance companies",
                                        key="dash_fin")

        manufacturers = sorted(df_sheet1["Manufacturer Name"].dropna().unique().tolist())
        sel_mfg = st.multiselect("Manufacturer", manufacturers, default=[],
                                    placeholder="All manufacturers", key="dash_mfg")

        growers = sorted({_grower_label(r) for _, r in df_sheet1.iterrows() if _grower_label(r)})
        sel_growers = st.multiselect("Grower", growers, default=[],
                                        placeholder="All growers (type to search)",
                                        key="dash_grow")

        products = sorted(df_sheet1["Item Description/Brand"].dropna().unique().tolist())
        sel_products = st.multiselect(
            "Product", products, default=[],
            placeholder=f"All products ({len(products)} available — type to search)",
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
    if not monthly.empty:
        fig = px.line(monthly, x="month", y="Sum Total Price",
                       markers=True, title="Monthly Revenue")
        fig.update_traces(line_color=theme.KAS_GREEN, marker_color=theme.KAS_GOLD)
        c1.plotly_chart(fig, use_container_width=True)

    by_ret = (df.groupby("__retailer", as_index=False)["Sum Total Price"].sum()
              .sort_values("Sum Total Price", ascending=False))
    if not by_ret.empty:
        fig = px.bar(by_ret, x="__retailer", y="Sum Total Price",
                     title="Revenue by Retailer Location",
                     color="__retailer")
        fig.update_layout(showlegend=False)
        c2.plotly_chart(fig, use_container_width=True)

    c3, c4 = st.columns(2)
    by_mfg = (df.groupby("Manufacturer Name", as_index=False)["Sum Total Price"].sum()
              .sort_values("Sum Total Price", ascending=False).head(15))
    if not by_mfg.empty:
        fig = px.pie(by_mfg, names="Manufacturer Name", values="Sum Total Price",
                      title="Revenue by Manufacturer", hole=0.4)
        c3.plotly_chart(fig, use_container_width=True)

    by_fin = (df.groupby("__finance", as_index=False)["Sum Total Price"].sum())
    if not by_fin.empty:
        fig = px.pie(by_fin, names="__finance", values="Sum Total Price",
                      title="Finance Company Mix", hole=0.4)
        c4.plotly_chart(fig, use_container_width=True)

    c5, c6 = st.columns(2)
    top_growers = (df.groupby("__grower", as_index=False)["Sum Total Price"].sum()
                   .sort_values("Sum Total Price", ascending=True).tail(15))
    if not top_growers.empty:
        fig = px.bar(top_growers, x="Sum Total Price", y="__grower", orientation="h",
                     title="Top 15 Growers by Spend")
        fig.update_traces(marker_color=theme.KAS_GREEN)
        c5.plotly_chart(fig, use_container_width=True)

    top_products = (df.groupby("Item Description/Brand", as_index=False)["Sum Total Price"].sum()
                    .sort_values("Sum Total Price", ascending=True).tail(20))
    if not top_products.empty:
        fig = px.bar(top_products, x="Sum Total Price", y="Item Description/Brand", orientation="h",
                     title="Top 20 Products by Revenue")
        fig.update_traces(marker_color=theme.KAS_GOLD)
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
    st.dataframe(
        grower_summary,
        use_container_width=True, hide_index=True,
        column_config={
            "total_spend": st.column_config.NumberColumn("Total Spend", format="$%.2f"),
            "n_line_items": "Line Items",
            "n_invoices": "Invoices",
            "last_purchase": "Last Invoice",
        },
    )

    st.markdown("#### Filtered line items")
    st.dataframe(
        df[["Invoice Date", "Invoice Number", "__grower", "Item Description/Brand",
            "Manufacturer Name", "Standard Unit Of Measure", "Quantity", "Sum Total Price",
            "__retailer", "__finance"]]
        .rename(columns={"__grower": "Grower", "__retailer": "Retailer",
                          "__finance": "Finance"})
        .sort_values("Invoice Date", ascending=False),
        use_container_width=True, hide_index=True, height=380,
        column_config={
            "Sum Total Price": st.column_config.NumberColumn(format="$%.2f"),
            "Quantity": st.column_config.NumberColumn(format="%.2f"),
            "Invoice Number": st.column_config.NumberColumn(format="%d"),
            "Invoice Date": st.column_config.DateColumn(format="YYYY-MM-DD"),
        },
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
def _build_grower_index(_xlsx_bytes: bytes) -> tuple:
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
        s["PDF"] = s["Invoice Number"].apply(
            lambda n: f"https://drive.google.com/file/d/{pdf_lookup[str(int(n))]}/view"
            if pd.notnull(n) and str(int(n)) in pdf_lookup else None
        )
        summaries[grower] = s
    return subs, summaries


def _grower_detail(grower_name: str) -> tuple:
    """Cheap lookup using the precomputed index."""
    subs, summaries = _build_grower_index(_bytes)
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


if active_page == PAGES[1]:
    _render_dashboard()


# ===========================================================================
# Tab 3 — Grower Map
# ===========================================================================
def _render_map():
    import folium
    from streamlit_folium import st_folium

    if df_sheet1.empty:
        st.info("No data yet.")
        return

    df = df_sheet1.copy()

    # ---- Step 1: assign one DISPLAY label per Grower ID ----
    # Grower ID is the unique entity key. Two different Grower IDs are
    # always different entities, even if they share a name. We pick the
    # most common spelling per ID; if names collide between IDs, we
    # disambiguate by appending the ID in parens.
    df["__raw_label"] = df.apply(_grower_label, axis=1)
    canonical_name = _canonical_names_by_id(df)
    df["__grower"] = df.apply(
        lambda r: canonical_name.get(r["Grower ID"], r["__raw_label"]),
        axis=1,
    )

    # Skip rows with no street address.
    def _norm(s):
        return str(s).strip().upper() if pd.notnull(s) and str(s).strip() else ""
    df["__addr1"] = df["Grower Address1"].apply(_norm)
    df["__city"] = df["Grower City"].apply(_norm)
    df["__state"] = df["Grower State"].apply(_norm)
    df["__zip"] = df["Grower ZIP CODE"].apply(
        lambda z: str(int(z)) if pd.notnull(z) and str(z).strip() else ""
    )
    df = df[df["__addr1"] != ""]

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
        st.rerun()

    # Pull cached coordinates (no API calls if already cached).
    for key, ar in needs_geocode:
        coords_for[key] = geocode.geocode_address(
            ar["__addr1"], ar["__city"],
            ar["__state"] or "KY", ar["__zip"],
        )

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

    # Pre-compute "growers at each location" so the detail section can show
    # a radio for multi-grower spots after a marker click.
    growers_by_coord = {}
    for _ck, _grp in geocoded.dropna(subset=["__grower"]).groupby("__coord_key"):
        growers_by_coord[_ck] = _grp["__grower"].unique().tolist()

    # Map view: remember whatever the user last panned/zoomed to so reruns
    # don't reset the view. Only on first visit (no saved state) we use the
    # western-Kentucky default.
    map_center = st.session_state.get("map_center") or [37.3, -87.5]
    map_zoom = st.session_state.get("map_zoom") or 8

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

    # Spend tiers (color, label, threshold). Order matters: highest first.
    # Colors chosen for max perceptual distinction across both light (Map) and
    # dark (Satellite) basemaps. Each is bold + saturated; the white outline
    # plus shadow on every marker keeps them visible regardless of background.
    SPEND_TIERS = [
        ("#4A148C", "$100,000+",       100_000),  # deep purple
        ("#1565C0", "$50,000 – $100k",  50_000),  # blue
        ("#2E7D32", "$20,000 – $50k",   20_000),  # green
        ("#EF6C00", "$5,000 – $20k",     5_000),  # orange
        ("#C62828", "Under $5,000",          0),  # red
    ]

    def _tier_color(spend: float) -> str:
        for color, _label, threshold in SPEND_TIERS:
            if spend >= threshold:
                return color
        return SPEND_TIERS[-1][0]

    plotted = 0
    MARKER_RADIUS = 9
    # Pre-compute grower-level totals across ALL locations (matches table).
    grower_totals = (
        df.dropna(subset=["__grower"])
        .groupby("__grower")
        .agg(spend=("Sum Total Price", "sum"),
              invoices=("Invoice Number", "nunique"),
              last=("Invoice Date", "max"))
    )

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


    # Only capture marker clicks. NOT zoom/center — including those would
    # fire a rerun on every pan/zoom, making the map painfully laggy.
    map_state = st_folium(
        m, width=None, height=600,
        returned_objects=["last_object_clicked"],
        key="kas_map", use_container_width=True,
    )

    # Marker click → load the FIRST grower shown in the popup. The popup
    # orders by total grower-level spend (so the "$X" in the popup matches
    # the table). We use the same ordering here so the marker-click result
    # always matches the popup's top card.
    clicked = (map_state or {}).get("last_object_clicked")
    if clicked and isinstance(clicked, dict) and "lat" in clicked:
        sig = (round(clicked["lat"], 4), round(clicked["lng"], 4))
        if sig != st.session_state.get("_last_marker_click_sig"):
            st.session_state["_last_marker_click_sig"] = sig
            growers_here = (
                geocoded[geocoded["__coord_key"] == sig]
                .dropna(subset=["__grower"])["__grower"].unique().tolist()
            )
            ordered = (
                grower_totals.loc[grower_totals.index.intersection(growers_here)]
                .sort_values("spend", ascending=False)
            )
            if len(ordered):
                primary = ordered.index[0]
                st.session_state["map_selected_grower"] = primary
                st.session_state["map_search_grower"] = primary
                st.session_state["_scroll_to_details"] = True
                # Force the dropdown to re-init with the new value next run
                # (we can't write to its key here — it already rendered).
                if "map_search_select" in st.session_state:
                    del st.session_state["map_search_select"]
                # Also reset the radio for this location so it re-seeds
                # to the primary grower on the next render.
                _radio_key = f"multi_grower_radio_{sig}"
                if _radio_key in st.session_state:
                    del st.session_state[_radio_key]
                st.rerun()

    if len(missing) > 0:
        st.caption(f"{plotted} locations on map. "
                    f"{missing['Grower ID'].nunique()} grower(s) couldn't be placed "
                    f"(usually a typo'd address).")
    else:
        st.caption(f"{plotted} locations on map.")

    # The selected grower comes from either the search box above the map, or
    # the marker click handler above (which sets `map_selected_grower` and
    # triggers a single rerun).
    selected = st.session_state.get("map_selected_grower")

    # If the last-clicked location has MULTIPLE growers, show a horizontal
    # radio so the user can switch the table view between them. Streamlit's
    # native widget — fully reliable, no JS bridging needed.
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
        st.divider()
        st.markdown(f"#### {len(growers_at_last_click)} growers at this address")

        radio_key = f"multi_grower_radio_{last_sig}"

        # Seed the radio's stored value once (only when this widget hasn't
        # been rendered yet for the current location). Streamlit uses the
        # stored value on subsequent renders; setting `index=` each time
        # would race with user clicks, causing a feedback loop.
        if radio_key not in st.session_state:
            st.session_state[radio_key] = (
                selected if selected in growers_at_last_click
                else growers_at_last_click[0]
            )

        def _on_radio_change():
            new_val = st.session_state[radio_key]
            st.session_state["map_selected_grower"] = new_val
            st.session_state["map_search_grower"] = new_val
            # Keep the dropdown display in sync with the radio choice.
            st.session_state["map_search_select"] = new_val

        st.radio(
            "Pick which grower to view",
            growers_at_last_click,
            horizontal=True,
            label_visibility="collapsed",
            key=radio_key,
            on_change=_on_radio_change,
        )
        # Re-read selected so the table below uses the latest choice.
        selected = st.session_state.get("map_selected_grower")

    if selected:
        # Only add a divider if we didn't already render one for the
        # multi-grower radio just above.
        if not growers_at_last_click:
            st.divider()
        st.markdown('<div id="kas-details"></div>', unsafe_allow_html=True)
        if st.session_state.pop("_scroll_to_details", False):
            import streamlit.components.v1 as components
            components.html(
                """
                <script>
                  setTimeout(function(){
                    var el = window.parent.document.getElementById('kas-details');
                    if (el) el.scrollIntoView({behavior:'smooth', block:'start'});
                  }, 50);
                </script>
                """,
                height=0,
            )
        st.markdown(f"### 📍 {selected}")
        sub, inv_summary = _grower_detail(selected)
        if sub is None or sub.empty:
            st.info("No invoices found for this grower.")
        else:
            cols = st.columns(4)
            cols[0].metric("Total spend", f"${sub['Sum Total Price'].sum():,.2f}")
            cols[1].metric("Invoices", f"{sub['Invoice Number'].nunique()}")
            cols[2].metric("Line items", f"{len(sub)}")
            last = sub["Invoice Date"].max()
            cols[3].metric("Last purchase",
                            last.strftime("%Y-%m-%d") if pd.notnull(last) else "—")

            st.markdown("##### Invoice history")
            st.dataframe(
                inv_summary,
                use_container_width=True, hide_index=True,
                column_config={
                    "total": st.column_config.NumberColumn(format="$%.2f"),
                    "Invoice Number": st.column_config.NumberColumn(format="%d"),
                    "Invoice Date": st.column_config.DateColumn(format="YYYY-MM-DD"),
                    "PDF": st.column_config.LinkColumn(
                        "PDF",
                        display_text="📄 Open",
                        help="Open the source PDF (only available for invoices processed through this app).",
                    ),
                },
            )

            st.markdown("##### All line items")
            st.dataframe(
                sub[["Invoice Date", "Invoice Number", "Item Description/Brand",
                     "Manufacturer Name", "Standard Unit Of Measure", "Quantity",
                     "Sum Total Price", "Retailer Name", "Finance Company"]],
                use_container_width=True, hide_index=True, height=400,
                column_config={
                    "Sum Total Price": st.column_config.NumberColumn(format="$%.2f"),
                    "Quantity": st.column_config.NumberColumn(format="%.2f"),
                    "Invoice Number": st.column_config.NumberColumn(format="%d"),
                    "Invoice Date": st.column_config.DateColumn(format="YYYY-MM-DD"),
                },
            )


if active_page == PAGES[2]:
    _render_map()
