"""Data integrity verification for the workbook.

Used by the app to:
  1. Run a quick consistency check after every approve+save
  2. Provide a sidebar "Verify data consistency" button for on-demand audits

The check confirms that totals computed by the app's canonical-name logic
match the raw spreadsheet cell totals to the cent.
"""
from __future__ import annotations
from dataclasses import dataclass
from io import BytesIO
from typing import Any

import pandas as pd
from openpyxl import load_workbook as _raw_load


@dataclass
class CheckResult:
    name: str
    passed: bool
    raw_value: Any
    app_value: Any
    detail: str = ""


@dataclass
class VerifyReport:
    checks: list[CheckResult]
    raw_total: float
    app_total: float
    raw_rows: int
    app_rows: int

    @property
    def passed(self) -> bool:
        return all(c.passed for c in self.checks)

    @property
    def failures(self) -> list[CheckResult]:
        return [c for c in self.checks if not c.passed]


def _grower_label(r: dict) -> str:
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


def verify_workbook(xlsx_bytes: bytes) -> VerifyReport:
    """Verify that the app's view of the workbook matches the raw .xlsx."""
    # ---- Raw read ----
    wb = _raw_load(BytesIO(xlsx_bytes), read_only=True, data_only=True)
    ws = wb.active

    raw_total = 0.0
    raw_rows = 0
    raw_invoices: set[str] = set()
    per_patron_total: dict = {}
    per_patron_invs: dict = {}
    per_patron_lines: dict = {}
    no_patron_total: dict = {}
    no_patron_invs: dict = {}
    no_patron_lines: dict = {}

    for r in ws.iter_rows(min_row=2, values_only=True):
        if not r or all(c in (None, "") for c in r):
            continue
        raw_rows += 1
        price = float(r[19] or 0)
        raw_total += price
        if r[16] is not None:
            raw_invoices.add(str(r[16]).strip())
        pid = r[5]
        inv = str(r[16]).strip() if r[16] is not None else None
        if pid is not None:
            per_patron_total[pid] = per_patron_total.get(pid, 0.0) + price
            if inv:
                per_patron_invs.setdefault(pid, set()).add(inv)
            per_patron_lines[pid] = per_patron_lines.get(pid, 0) + 1
        else:
            lbl = _grower_label({
                "Grower First Name": r[6],
                "Grower Last Name": r[7],
                "Grower Company Name": r[8],
            })
            if lbl:
                no_patron_total[lbl] = no_patron_total.get(lbl, 0.0) + price
                if inv:
                    no_patron_invs.setdefault(lbl, set()).add(inv)
                no_patron_lines[lbl] = no_patron_lines.get(lbl, 0) + 1

    # ---- App view (same logic as _build_grower_index / _render_map) ----
    from . import workbook as wb_mod
    wb2 = wb_mod.load(xlsx_bytes)
    df = wb_mod.read_sheet1_dataframe(wb2)
    df["__raw_label"] = df.apply(_grower_label, axis=1).replace("", pd.NA)

    # Pick most-common label per Grower ID; if two IDs collide on name,
    # disambiguate with the ID. Same logic as the app's _canonical_names_by_id.
    base: dict = {}
    for gid, sub in df.dropna(subset=["Grower ID"]).groupby("Grower ID"):
        labels = [s for s in sub["__raw_label"] if pd.notnull(s) and s]
        if labels:
            base[gid] = max(set(labels), key=labels.count)
    name_to_ids: dict = {}
    for gid, n in base.items():
        name_to_ids.setdefault(n, []).append(gid)
    canonical: dict = {}
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
    app_rows = len(df)
    app_invoices = set(df["Invoice Number"].dropna().astype(int).astype(str))
    grower_total = (df.dropna(subset=["__grower"]).groupby("__grower")
                    ["Sum Total Price"].sum())
    grower_invs = (df.dropna(subset=["__grower"]).groupby("__grower")
                    ["Invoice Number"].nunique())
    grower_lines = df.dropna(subset=["__grower"]).groupby("__grower").size()

    # Each canonical display label now corresponds to AT MOST ONE Grower ID
    # (we disambiguate above). No-Grower-ID rows roll up under their raw label.
    canonical_to_patrons: dict = {}
    for gid, name in canonical.items():
        canonical_to_patrons.setdefault(name, set()).add(gid)

    # ---- Build checks ----
    checks: list[CheckResult] = []

    checks.append(CheckResult(
        "Row count", raw_rows == app_rows, raw_rows, app_rows,
    ))
    checks.append(CheckResult(
        "Unique invoices",
        len(raw_invoices) == len(app_invoices),
        len(raw_invoices), len(app_invoices),
    ))
    checks.append(CheckResult(
        "Total revenue",
        abs(raw_total - app_total) < 0.01,
        f"${raw_total:,.2f}", f"${app_total:,.2f}",
    ))

    # Per-grower drill-down
    all_names = set(canonical_to_patrons.keys()) | set(grower_total.index)
    bad_totals = 0
    bad_invs = 0
    bad_lines = 0
    for cname in all_names:
        patrons = canonical_to_patrons.get(cname, set())
        expected_total = (sum(per_patron_total.get(p, 0.0) for p in patrons)
                           + no_patron_total.get(cname, 0.0))
        got_total = float(grower_total.get(cname, 0.0))
        if abs(expected_total - got_total) > 0.01:
            bad_totals += 1
        expected_invs: set = set()
        for p in patrons:
            for iv in per_patron_invs.get(p, set()):
                if iv:
                    expected_invs.add(iv)
        for iv in no_patron_invs.get(cname, set()):
            if iv:
                expected_invs.add(iv)
        if int(grower_invs.get(cname, 0)) != len(expected_invs):
            bad_invs += 1
        expected_lines = (sum(per_patron_lines.get(p, 0) for p in patrons)
                           + no_patron_lines.get(cname, 0))
        if int(grower_lines.get(cname, 0)) != expected_lines:
            bad_lines += 1

    checks.append(CheckResult(
        "Per-grower spend (popup = table)",
        bad_totals == 0,
        f"{len(all_names)} growers", f"{bad_totals} mismatched",
    ))
    checks.append(CheckResult(
        "Per-grower invoice counts",
        bad_invs == 0,
        f"{len(all_names)} growers", f"{bad_invs} mismatched",
    ))
    checks.append(CheckResult(
        "Per-grower line-item counts",
        bad_lines == 0,
        f"{len(all_names)} growers", f"{bad_lines} mismatched",
    ))

    return VerifyReport(
        checks=checks,
        raw_total=raw_total, app_total=app_total,
        raw_rows=raw_rows, app_rows=app_rows,
    )
