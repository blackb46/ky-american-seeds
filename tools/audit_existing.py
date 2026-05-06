"""Audit existing spreadsheet rows against their source PDFs.

For each PDF in ./invoices/ whose invoice number is already in Sheet1, run the
extractor (single pass, no verifier — saves ~50% tokens) and compare each
line item to what's in the spreadsheet. Report mismatches.
"""
from __future__ import annotations
import json
import sys
try:
    sys.stdout.reconfigure(encoding="utf-8")
except Exception:
    pass
import time
import unicodedata
from pathlib import Path
from collections import defaultdict

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

import pandas as pd
from lib import extract, workbook as wb_mod, normalize as norm

# Load API key from secrets
secrets_text = (Path(__file__).resolve().parent.parent / ".streamlit" / "secrets.toml").read_text()
API_KEY = secrets_text.split('ANTHROPIC_API_KEY = "')[1].split('"')[0]

PROJECT = Path(__file__).resolve().parent.parent
INVOICES_DIR = PROJECT / "invoices"
_xlsx_candidates = list(PROJECT.glob("*.xlsx"))
XLSX = _xlsx_candidates[0] if _xlsx_candidates else PROJECT / "transactions.xlsx"


def _norm_desc(s: str | None) -> str:
    if not s:
        return ""
    s = unicodedata.normalize("NFKC", str(s))
    s = " ".join(s.split()).upper()
    return s


def main():
    wb = wb_mod.load(XLSX)
    df = wb_mod.read_sheet1_dataframe(wb)
    df["__inv_str"] = df["Invoice Number"].apply(lambda x: str(int(x)) if pd.notnull(x) else "")
    by_inv: dict[str, list] = defaultdict(list)
    for _, r in df.iterrows():
        if r["__inv_str"]:
            by_inv[r["__inv_str"]].append(r)
    print(f"Spreadsheet has {len(by_inv)} unique invoice numbers, "
          f"{len(df)} line items.\n")

    pdfs = sorted([p for p in INVOICES_DIR.glob("*.pdf")])
    audit_results = []

    for pdf in pdfs:
        print(f"=== {pdf.name} ===")
        try:
            with open(pdf, "rb") as f:
                pdf_bytes = f.read()
            t0 = time.time()
            result = extract.extract_pdf(pdf_bytes, api_key=API_KEY, verify=False)
            elapsed = time.time() - t0
            print(f"  extract: {elapsed:.1f}s, "
                  f"{result['usage']['input_tokens']:,}+{result['usage']['output_tokens']:,} tok")
        except Exception as e:
            print(f"  EXTRACTION FAILED: {e}")
            continue

        for inv in result["final"].get("invoices", []):
            inv_no = str(inv.get("invoice_number") or "").strip()
            if not inv_no or inv_no not in by_inv:
                if inv_no:
                    print(f"  invoice {inv_no}: NOT in spreadsheet (skipped)")
                continue
            sheet_rows = by_inv[inv_no]
            mismatches = compare_invoice(inv, sheet_rows)
            audit_results.append({
                "pdf": pdf.name,
                "invoice": inv_no,
                "sheet_line_count": len(sheet_rows),
                "pdf_line_count": len(inv.get("line_items", [])),
                "mismatches": mismatches,
            })
            if mismatches:
                print(f"  ❌ Invoice {inv_no}: {len(mismatches)} issue(s)")
                for m in mismatches:
                    print(f"     - {m}")
            else:
                print(f"  ✅ Invoice {inv_no}: matches "
                      f"({len(sheet_rows)} line items)")

    # Final summary
    print("\n" + "=" * 70)
    print("AUDIT SUMMARY")
    print("=" * 70)
    n_clean = sum(1 for r in audit_results if not r["mismatches"])
    n_dirty = sum(1 for r in audit_results if r["mismatches"])
    print(f"Clean: {n_clean} invoices, with issues: {n_dirty} invoices")
    if n_dirty:
        print("\nInvoices needing your review:")
        for r in audit_results:
            if r["mismatches"]:
                print(f"  - Invoice {r['invoice']} (from {r['pdf']}): "
                      f"{len(r['mismatches'])} issue(s)")
                for m in r["mismatches"][:5]:
                    print(f"      {m}")
                if len(r["mismatches"]) > 5:
                    print(f"      ... and {len(r['mismatches']) - 5} more")

    out_path = PROJECT / "audit_report.json"
    with open(out_path, "w") as f:
        json.dump(audit_results, f, indent=2, default=str)
    print(f"\nFull report: {out_path}")


def compare_invoice(inv: dict, sheet_rows: list) -> list[str]:
    """Compare an extracted invoice's line items vs the spreadsheet rows."""
    issues: list[str] = []
    pdf_items = inv.get("line_items") or []
    inv_no = inv.get("invoice_number")

    # Build map: (description_normalized) -> sheet row(s)
    sheet_map: dict[str, list] = defaultdict(list)
    for row in sheet_rows:
        sheet_map[_norm_desc(row.get("Item Description/Brand"))].append(row)

    pdf_descs = [_norm_desc(li.get("description")) for li in pdf_items]
    sheet_descs = [_norm_desc(r.get("Item Description/Brand")) for r in sheet_rows]

    if len(pdf_items) != len(sheet_rows):
        issues.append(
            f"line item count: PDF has {len(pdf_items)}, sheet has {len(sheet_rows)}"
        )

    # Match each PDF line item to a sheet row by description
    matched_sheet_idx: set[int] = set()
    for li in pdf_items:
        pdf_desc = _norm_desc(li.get("description"))
        best_idx = None
        for i, sd in enumerate(sheet_descs):
            if i in matched_sheet_idx:
                continue
            if pdf_desc == sd or (pdf_desc and (pdf_desc in sd or sd in pdf_desc)):
                best_idx = i
                break
        if best_idx is None:
            issues.append(f"PDF item '{li.get('description')}' has no matching sheet row")
            continue
        matched_sheet_idx.add(best_idx)
        sr = sheet_rows[best_idx]
        # Compare quantity
        pdf_qty = li.get("quantity")
        sh_qty = sr.get("Quantity")
        if pdf_qty is not None and sh_qty is not None:
            try:
                if abs(float(pdf_qty) - float(sh_qty)) > 0.01:
                    issues.append(
                        f"qty mismatch on '{li.get('description')}': "
                        f"PDF {pdf_qty} vs sheet {sh_qty}"
                    )
            except (TypeError, ValueError):
                pass
        # Compare ext amount
        pdf_ext = li.get("ext_amount")
        sh_ext = sr.get("Sum Total Price")
        if pdf_ext is not None and sh_ext is not None:
            try:
                if abs(float(pdf_ext) - float(sh_ext)) > 0.05:
                    issues.append(
                        f"price mismatch on '{li.get('description')}': "
                        f"PDF ${float(pdf_ext):.2f} vs sheet ${float(sh_ext):.2f}"
                    )
            except (TypeError, ValueError):
                pass
        # Compare unit
        pdf_unit = norm.normalize_unit(li.get("unit"))
        sh_unit = norm.normalize_unit(sr.get("Standard Unit Of Measure"))
        if pdf_unit and sh_unit and pdf_unit != sh_unit:
            issues.append(
                f"unit mismatch on '{li.get('description')}': "
                f"PDF {pdf_unit} vs sheet {sh_unit}"
            )

    # Sheet rows with no matching PDF item
    unmatched = [i for i in range(len(sheet_rows)) if i not in matched_sheet_idx]
    for i in unmatched:
        sr = sheet_rows[i]
        issues.append(
            f"sheet row '{sr.get('Item Description/Brand')}' "
            f"(qty {sr.get('Quantity')}, ${sr.get('Sum Total Price')}) has no PDF match"
        )

    # Compare invoice total
    pdf_total = inv.get("invoice_total")
    sheet_total = sum(float(r.get("Sum Total Price") or 0) for r in sheet_rows)
    if pdf_total is not None and abs(float(pdf_total) - sheet_total) > 0.05:
        issues.append(
            f"invoice total: PDF ${float(pdf_total):.2f} vs sum of sheet rows ${sheet_total:.2f}"
        )

    return issues


if __name__ == "__main__":
    import pandas as pd
    main()
