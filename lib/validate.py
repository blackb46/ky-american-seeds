"""Validation checks run on extracted invoice data before write.

All checks return ValidationResult with severity: 'error' (blocks save),
'warning' (visible but not blocking), or 'info'.
"""
from __future__ import annotations
from dataclasses import dataclass
from datetime import datetime, timedelta
from typing import Literal


Severity = Literal["error", "warning", "info"]


@dataclass
class Issue:
    severity: Severity
    field: str
    message: str
    expected: float | str | None = None
    actual: float | str | None = None


@dataclass
class ValidationResult:
    issues: list[Issue]

    @property
    def errors(self) -> list[Issue]:
        return [i for i in self.issues if i.severity == "error"]

    @property
    def warnings(self) -> list[Issue]:
        return [i for i in self.issues if i.severity == "warning"]

    @property
    def passes(self) -> bool:
        return not self.errors

    def summary(self) -> str:
        if not self.issues:
            return "All checks passed."
        return f"{len(self.errors)} error(s), {len(self.warnings)} warning(s)."


def _money_close(a: float | None, b: float | None, tol: float = 0.02) -> bool:
    if a is None or b is None:
        return False
    return abs(a - b) <= tol


def _safe_dict(v) -> dict:
    """Return v if it's a dict, otherwise {}. Use this anywhere we read
    fields with .get() from data that came back from Claude's JSON —
    Claude sometimes returns lists or strings where the schema expects
    a dict, and `value or {}` doesn't help because non-empty lists/
    strings are truthy."""
    return v if isinstance(v, dict) else {}


def _safe_list(v) -> list:
    """Return v if it's a list, otherwise []. Same rationale as _safe_dict."""
    return v if isinstance(v, list) else []


def validate_invoice(inv: dict) -> ValidationResult:
    """Validate a single extracted invoice dict (the schema from extract.py)."""
    issues: list[Issue] = []
    if not isinstance(inv, dict):
        issues.append(Issue("error", "_root",
                            f"Top-level invoice is not a dict (got {type(inv).__name__})."))
        return ValidationResult(issues=issues)

    # Required fields
    if not inv.get("invoice_number"):
        issues.append(Issue("error", "invoice_number", "Invoice number is required."))

    grower = _safe_dict(inv.get("grower"))
    has_name = grower.get("first_name") or grower.get("last_name") or grower.get("company_name")
    if not has_name:
        issues.append(Issue("error", "grower", "Grower name (first/last or company) is required."))

    line_items = _safe_list(inv.get("line_items"))
    if not line_items:
        issues.append(Issue("error", "line_items", "At least one line item is required."))

    # Line-item math: qty × unit_price ≈ ext_amount
    line_sum = 0.0
    for i, li_raw in enumerate(line_items):
        li = _safe_dict(li_raw)
        if li_raw is not None and not isinstance(li_raw, dict):
            issues.append(Issue("warning", f"line_items[{i}]",
                                f"Line item is not a dict (got {type(li_raw).__name__}); skipped."))
            continue
        qty = li.get("quantity")
        up = li.get("unit_price")
        ext = li.get("ext_amount")
        if not li.get("description"):
            issues.append(Issue("error", f"line_items[{i}].description",
                                "Line item description is required."))
        if qty is None or ext is None:
            issues.append(Issue("error", f"line_items[{i}]",
                                "Quantity and ext_amount are required."))
        elif up is not None and qty is not None:
            try:
                expected = round(float(qty) * float(up), 2)
                if not _money_close(expected, ext, tol=0.05):
                    issues.append(Issue("warning", f"line_items[{i}]",
                                        f"qty × unit_price ({expected:.2f}) ≠ ext_amount ({ext:.2f})",
                                        expected=expected, actual=ext))
            except (TypeError, ValueError):
                pass  # non-numeric qty/up — caught elsewhere
        if ext is not None:
            try:
                line_sum += float(ext)
            except (TypeError, ValueError):
                pass

    # Invoice total = sum of line items
    inv_total = inv.get("invoice_total")
    if inv_total is not None and line_items:
        if not _money_close(round(line_sum, 2), inv_total, tol=0.02):
            issues.append(Issue("warning", "invoice_total",
                                f"Sum of line items ({line_sum:.2f}) ≠ invoice_total ({inv_total:.2f})",
                                expected=round(line_sum, 2), actual=inv_total))

    # Prepaid + account charge = invoice total
    prep = inv.get("prepaid_amount")
    chg = inv.get("account_charge_amount")
    if inv_total is not None and (prep is not None or chg is not None):
        s = (prep or 0) + (chg or 0)
        if s > 0 and not _money_close(s, inv_total, tol=0.02):
            issues.append(Issue("warning", "prepaid_split",
                                f"prepaid ({prep or 0}) + account charge ({chg or 0}) = {s:.2f} ≠ invoice_total ({inv_total:.2f})",
                                expected=inv_total, actual=s))

    # CHS amount cross-check — finance can come back from Claude as a list
    # or string instead of a dict on atypical invoice formats (CFA / cash
    # farm advance / partial receipts), so use the isinstance guard.
    finance = _safe_dict(inv.get("finance"))
    a2r = finance.get("amount_to_retailer")
    if a2r is not None and inv_total is not None:
        if not _money_close(a2r, inv_total, tol=0.50):
            issues.append(Issue("info", "amount_to_retailer",
                                f"CHS amount_to_retailer ({a2r}) differs from invoice_total ({inv_total})",
                                expected=inv_total, actual=a2r))

    # Date sanity
    sold = inv.get("sold_date")
    if sold:
        try:
            d = datetime.fromisoformat(str(sold).split("T")[0])
            now = datetime.now()
            if d > now + timedelta(days=2):
                issues.append(Issue("warning", "sold_date",
                                    f"Sold date {sold} is in the future."))
            if d < now - timedelta(days=365 * 5):
                issues.append(Issue("warning", "sold_date",
                                    f"Sold date {sold} is more than 5 years old."))
        except (ValueError, TypeError):
            issues.append(Issue("warning", "sold_date", f"Unparseable date: {sold!r}"))

    # ZIP/state sanity
    zp = grower.get("zip")
    if zp and not str(zp).strip().isdigit():
        issues.append(Issue("warning", "grower.zip", f"ZIP not numeric: {zp!r}"))
    elif zp and not (4 <= len(str(zp).strip()) <= 5):
        issues.append(Issue("warning", "grower.zip", f"ZIP wrong length: {zp!r}"))

    state = grower.get("state")
    if state and len(str(state).strip()) != 2:
        issues.append(Issue("warning", "grower.state", f"State should be 2 letters: {state!r}"))

    # Low-confidence flags
    if inv.get("invoice_number_confidence") == "low":
        issues.append(Issue("warning", "invoice_number", "Low confidence on invoice number."))
    if inv.get("invoice_total_confidence") == "low":
        issues.append(Issue("warning", "invoice_total", "Low confidence on invoice total."))

    return ValidationResult(issues=issues)
