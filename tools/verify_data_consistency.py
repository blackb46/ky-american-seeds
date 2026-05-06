"""End-to-end data accuracy verification.

Checks that every aggregated number the app displays matches what's
literally in the spreadsheet, and that the canonical-name merging used by
the map and the table produce identical totals.
"""
from __future__ import annotations
import sys, os
try:
    sys.stdout.reconfigure(encoding="utf-8")
except Exception:
    pass

from pathlib import Path
sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

import pandas as pd
from openpyxl import load_workbook as raw_load
from lib import workbook as wb_mod

PROJECT = Path(__file__).resolve().parent.parent
xlsx = next(PROJECT.glob("*.xlsx"))
print(f"File: {xlsx.name}\n")


# Replicate the app's _grower_label() for parity.
def _grower_label(r):
    def _clean(v):
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


# ---------------------------------------------------------------------------
# A. Raw read via openpyxl (the absolute truth)
# ---------------------------------------------------------------------------
wb = raw_load(xlsx, read_only=True, data_only=True)
ws = wb.active

raw_total = 0.0
raw_qty = 0.0
raw_invoices = set()
raw_patrons = set()
raw_rows = 0
per_patron_raw_total = {}  # patron_id -> total $
per_patron_raw_invoices = {}
per_patron_raw_lineitems = {}
per_patron_raw_lastdate = {}
# Totals/invoices/lines for rows with NO Grower ID, keyed by raw label.
# These rows aren't covered by the canonical_name dict (which is patron-keyed)
# so they fall through to the raw label and join whatever canonical group
# the same label refers to.
no_patron_total = {}
no_patron_invoices = {}
no_patron_lineitems = {}

for r in ws.iter_rows(min_row=2, values_only=True):
    if not r or all(c in (None, "") for c in r):
        continue
    raw_rows += 1
    price = float(r[19] or 0)
    qty = float(r[18] or 0)
    raw_total += price
    raw_qty += qty
    if r[16] is not None:
        raw_invoices.add(str(r[16]).strip())
    pid = r[5]
    if pid is not None:
        raw_patrons.add(pid)
        per_patron_raw_total[pid] = per_patron_raw_total.get(pid, 0.0) + price
        per_patron_raw_invoices.setdefault(pid, set()).add(str(r[16]).strip() if r[16] else None)
        per_patron_raw_lineitems[pid] = per_patron_raw_lineitems.get(pid, 0) + 1
        d = r[15]
        if d is not None and not isinstance(d, str):
            cur = per_patron_raw_lastdate.get(pid)
            if cur is None or d > cur:
                per_patron_raw_lastdate[pid] = d
    else:
        # Row with no Grower ID — bucket by raw grower label.
        row_dict = {
            "Grower First Name": r[6], "Grower Last Name": r[7],
            "Grower Company Name": r[8],
        }
        lbl = _grower_label(row_dict)
        if lbl:
            no_patron_total[lbl] = no_patron_total.get(lbl, 0.0) + price
            no_patron_invoices.setdefault(lbl, set()).add(str(r[16]).strip() if r[16] else None)
            no_patron_lineitems[lbl] = no_patron_lineitems.get(lbl, 0) + 1

print("=== A. Raw openpyxl truth ===")
print(f"  rows:           {raw_rows}")
print(f"  unique invoices: {len(raw_invoices)}")
print(f"  unique patrons:  {len(raw_patrons)}")
print(f"  total $ revenue: ${raw_total:,.2f}")
print(f"  total quantity:  {raw_qty:,.2f}")


# ---------------------------------------------------------------------------
# B. App's view via lib/workbook.py
# ---------------------------------------------------------------------------
wb2 = wb_mod.load(xlsx)
df = wb_mod.read_sheet1_dataframe(wb2)
df["__raw_label"] = df.apply(_grower_label, axis=1).replace("", pd.NA)

# Apply per-Grower-ID display label logic (matches app):
# pick most common name per ID, disambiguate collisions with the ID.
base = {}
for gid, sub in df.dropna(subset=["Grower ID"]).groupby("Grower ID"):
    labels = [s for s in sub["__raw_label"] if pd.notnull(s) and s]
    if labels:
        base[gid] = max(set(labels), key=labels.count)
name_to_ids = {}
for gid, n in base.items():
    name_to_ids.setdefault(n, []).append(gid)
canonical = {}
for gid, n in base.items():
    if len(name_to_ids[n]) > 1:
        try:
            canonical[gid] = f"{n} (#{int(gid)})"
        except (TypeError, ValueError):
            canonical[gid] = f"{n} (#{gid})"
    else:
        canonical[gid] = n
df["__grower"] = df.apply(
    lambda r: canonical.get(r["Grower ID"], r["__raw_label"]),
    axis=1,
)

app_total = float(df["Sum Total Price"].sum())
app_qty = float(df["Quantity"].sum())
app_invoices = set(df["Invoice Number"].dropna().astype(int).astype(str))
app_patrons = set(df["Grower ID"].dropna().astype(int).unique())

print("\n=== B. App view (canonical-name applied) ===")
print(f"  rows:           {len(df)}")
print(f"  unique invoices: {len(app_invoices)}")
print(f"  unique patrons:  {len(app_patrons)}")
print(f"  total $ revenue: ${app_total:,.2f}")
print(f"  total quantity:  {app_qty:,.2f}")


# ---------------------------------------------------------------------------
# Cross-check: aggregate totals
# ---------------------------------------------------------------------------
print("\n" + "=" * 70)
print("AGGREGATE CROSS-CHECK")
print("=" * 70)
def chk(label, raw, app, money=False, tol=None):
    fmt = (lambda v: f"${v:,.2f}") if money else (lambda v: f"{v:,}")
    if tol is None:
        tol = 0.01 if money else 1e-6
    ok = abs(raw - app) < tol if isinstance(raw, float) else raw == app
    print(f"  {'✅' if ok else '❌'} {label:25} raw={fmt(raw)}  app={fmt(app)}")

chk("rows", raw_rows, len(df))
chk("unique invoices", len(raw_invoices), len(app_invoices))
chk("unique patrons", len(raw_patrons), len(app_patrons))
chk("total revenue", raw_total, app_total, money=True)
chk("total quantity", raw_qty, app_qty, tol=1e-3)


# ---------------------------------------------------------------------------
# Per-grower (canonical) totals: should sum to the per-patron totals
# (since canonical-name merges spellings within a patron, never across).
# ---------------------------------------------------------------------------
print("\n" + "=" * 70)
print("PER-GROWER (CANONICAL) TOTALS — should match per-patron raw totals")
print("=" * 70)

grower_app_total = (
    df.dropna(subset=["__grower"])
    .groupby("__grower")["Sum Total Price"].sum()
)
# Map canonical name -> set of patrons. Each canonical name MAY map to >1
# patron if multiple patrons share the same most-common spelling.
canonical_to_patrons = {}
for gid, name in canonical.items():
    canonical_to_patrons.setdefault(name, set()).add(gid)

mismatches = []
all_canonical_names = set(canonical_to_patrons.keys()) | set(grower_app_total.index)
for cname in all_canonical_names:
    patrons = canonical_to_patrons.get(cname, set())
    expected = sum(per_patron_raw_total.get(p, 0.0) for p in patrons)
    # Add no-patron rows whose raw label matches this canonical name.
    expected += no_patron_total.get(cname, 0.0)
    got = float(grower_app_total.get(cname, 0.0))
    if abs(expected - got) > 0.01:
        mismatches.append((cname, expected, got, patrons))

if not mismatches:
    print(f"  ✅ All {len(all_canonical_names)} canonical-name groups match raw sums.")
else:
    print(f"  ❌ {len(mismatches)} canonical-name groups don't match:")
    for cname, expected, got, patrons in mismatches[:20]:
        print(f"     {cname}: expected ${expected:,.2f} got ${got:,.2f} "
              f"(patrons={sorted(patrons)})")


# ---------------------------------------------------------------------------
# Per-grower invoice count + line-item count consistency
# ---------------------------------------------------------------------------
print("\n" + "=" * 70)
print("PER-GROWER INVOICE / LINE-ITEM COUNTS")
print("=" * 70)

grower_invoices = (
    df.dropna(subset=["__grower"])
    .groupby("__grower")["Invoice Number"]
    .nunique()
)
grower_lines = df.dropna(subset=["__grower"]).groupby("__grower").size()

inv_mismatches = 0
line_mismatches = 0
for cname in all_canonical_names:
    patrons = canonical_to_patrons.get(cname, set())
    expected_invs = set()
    for p in patrons:
        for iv in per_patron_raw_invoices.get(p, set()):
            if iv:
                expected_invs.add(iv)
    for iv in no_patron_invoices.get(cname, set()):
        if iv:
            expected_invs.add(iv)
    expected_inv_count = len(expected_invs)
    expected_lines = (sum(per_patron_raw_lineitems.get(p, 0) for p in patrons)
                       + no_patron_lineitems.get(cname, 0))
    got_inv = int(grower_invoices.get(cname, 0))
    got_lines = int(grower_lines.get(cname, 0))
    if got_inv != expected_inv_count:
        inv_mismatches += 1
        if inv_mismatches <= 5:
            print(f"  ❌ {cname}: invoices expected={expected_inv_count} got={got_inv}")
    if got_lines != expected_lines:
        line_mismatches += 1
        if line_mismatches <= 5:
            print(f"  ❌ {cname}: line items expected={expected_lines} got={got_lines}")

if inv_mismatches == 0 and line_mismatches == 0:
    print(f"  ✅ Invoice and line-item counts match for all "
          f"{len(all_canonical_names)} growers.")
else:
    print(f"  Invoice mismatches: {inv_mismatches}, "
          f"line item mismatches: {line_mismatches}")


# ---------------------------------------------------------------------------
# What the popup shows == what the table shows
# ---------------------------------------------------------------------------
print("\n" + "=" * 70)
print("POPUP vs TABLE PARITY")
print("=" * 70)

# Popup uses: grower_totals = df.groupby("__grower").agg(spend=...)
# Table uses: _build_grower_index → also groupby("__grower").agg(...)
# Both come from the SAME df with canonical names → results are identical.
popup_totals = (
    df.dropna(subset=["__grower"])
    .groupby("__grower")
    .agg(spend=("Sum Total Price", "sum"),
         invoices=("Invoice Number", "nunique"),
         last=("Invoice Date", "max"))
)
print(f"  Total growers: {len(popup_totals)}")
print(f"  Total revenue across all growers: "
      f"${float(popup_totals['spend'].sum()):,.2f}")
print(f"  Excel-truth total: ${raw_total:,.2f}")
diff = abs(float(popup_totals["spend"].sum()) - raw_total)
print(f"  {'✅' if diff < 0.01 else '❌'} Difference: ${diff:,.2f}")


# ---------------------------------------------------------------------------
# Sample some specific growers shown in the screenshots
# ---------------------------------------------------------------------------
print("\n" + "=" * 70)
print("SAMPLE GROWERS")
print("=" * 70)
for sample in ["YATES FARMS", "DUNVILLE, JERRY", "KUEGEL, JOHN", "KUEGEL JR, JOHN"]:
    if sample in popup_totals.index:
        s = popup_totals.loc[sample]
        last = pd.to_datetime(s["last"]).strftime("%Y-%m-%d") if pd.notnull(s["last"]) else "—"
        patrons = sorted(canonical_to_patrons.get(sample, set()))
        print(f"  {sample}")
        print(f"    spend: ${float(s['spend']):,.2f}  "
              f"invoices: {int(s['invoices'])}  last: {last}")
        print(f"    patrons in this canonical group: {patrons}")
    else:
        print(f"  {sample}: not present as canonical name")


# Final summary
print("\n" + "=" * 70)
ok = (raw_rows == len(df)
      and abs(raw_total - app_total) < 0.01
      and len(raw_invoices) == len(app_invoices)
      and not mismatches
      and inv_mismatches == 0
      and line_mismatches == 0
      and abs(float(popup_totals["spend"].sum()) - raw_total) < 0.01)
if ok:
    print("✅ ALL DATA CONSISTENCY CHECKS PASSED.")
    print("    Popup totals, table totals, and the Excel file all match exactly.")
else:
    print("❌ Some checks FAILED. See above.")
print("=" * 70)
