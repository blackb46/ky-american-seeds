"""Dual-pass extraction of KAS invoices + CHS Capital confirmations from PDFs.

Pass 1 (Extract): Claude vision reads all PDF pages and emits a strict JSON
                  payload with per-field confidence ratings.
Pass 2 (Verify):  Claude is shown the same PDF pages PLUS pass-1 JSON and asked
                  to identify errors, missing fields, or values that don't match
                  the document. Returns a corrected JSON.

The UI surfaces both versions and a per-field agreement indicator so the user
can audit before approving.
"""
from __future__ import annotations
import json
import re
from datetime import datetime
from typing import Any

from anthropic import Anthropic

from . import pdf_render
from . import normalize as norm


MODEL = "claude-sonnet-4-6"

EXTRACT_SYSTEM = """You are an expert at extracting structured data from agricultural invoices for Kentucky American Seeds (KAS).

Each PDF contains EITHER:
- (A) A KAS Sales Invoice (typed, with header "Sales Invoice" and a 7-digit invoice number),
      OPTIONALLY followed by a CHS Capital advance confirmation email page that lists loan number,
      batch number, ACH date, "To Retailer" amount, and product/rate.
- (B) A handwritten KAS delivery ticket (header says "KENTUCKY AMERICAN SEEDS, LLC", 4-digit
      ticket number in red top-right, line items in pencil/pen), with no CHS confirmation.
- (C) Multiple invoices in one PDF — extract ALL of them.

ACCURACY RULES (CRITICAL):
1. Read every digit carefully. Quantities and prices are NEVER rounded.
2. For each line item, verify mentally that quantity × unit_price ≈ ext_amount. If they
   don't match, prefer ext_amount as the source of truth and lower confidence on the others.
3. invoice_total must equal the sum of line ext_amounts. If they don't match, mark
   invoice_total confidence as "low" and report the discrepancy in extraction_notes.
4. Manufacturer is NOT on the KAS invoice. It comes from the CHS confirmation's "Notes"
   or "Product/Rate" column (e.g. "Accolade CP Syngenta 7.25%" → manufacturer="SYNGENTA").
   If no CHS page, leave manufacturer null.
5. "Charged to account" + "applied to prepaid" must equal invoice_total. Capture both.
6. Dates: format as ISO YYYY-MM-DD. The invoice date format is typically MM/DD/YY HH:MM.
7. Numbers: no commas, no dollar signs. Use null (not 0 or "") for unknown.
8. Per-field _confidence: "high" if clearly legible and consistent, "medium" if guessed
   from context, "low" if smudged/handwriting/uncertain.

OUTPUT FORMAT:
Return ONLY a JSON object (no markdown, no prose) with this schema:

{
  "invoices": [
    {
      "invoice_number": "1093943",
      "invoice_number_confidence": "high",
      "patron_number": "100992",
      "patron_number_confidence": "high",
      "sold_date": "2026-04-10",
      "sold_date_confidence": "high",
      "retailer_name": "KAS TX",
      "retailer_city": "FREDONIA",
      "retailer_state": "KY",
      "merchandised_by": null,
      "grower": {
        "first_name": "DAVID",
        "last_name": "DENNIE",
        "company_name": "BROWN'S RAMSEY CREEK FARM",
        "address1": "6381 ST RT 270 E",
        "address2": null,
        "city": "CLAY", "state": "KY", "zip": "42404",
        "_confidence": "high"
      },
      "line_items": [
        {
          "item_number": "10711",
          "description": "Boundary 6.5EC BULK",
          "epa_info": "100-1162",
          "unit": "GAL", "quantity": 537.0, "unit_price": 40.5, "ext_amount": 21748.50,
          "manufacturer": null,
          "_confidence": "high"
        }
      ],
      "invoice_total": 21748.50,
      "invoice_total_confidence": "high",
      "account_charge_amount": 21748.50,
      "prepaid_amount": 0.0,
      "due_date": "2026-05-15",
      "finance": {
        "finance_company": "CHS",
        "loan_number": "4010320100",
        "loan_year": 2026,
        "product_rate": "Accolade CP Syngenta 7.25%",
        "amount_to_retailer": 21748.50,
        "amount_to_producer": 0.00,
        "batch_number": "142269",
        "ach_date": "2026-04-14",
        "manufacturer_from_notes": "SYNGENTA",
        "notes": "VOUNDARY 6.5EC",
        "_confidence": "high"
      },
      "extraction_notes": "Optional: any ambiguities or warnings",
      "math_check": {
        "line_items_sum": 21748.50,
        "invoice_total_matches": true,
        "prepaid_plus_charge_matches": true
      }
    }
  ]
}

If a field is unknown, use null. If finance section is absent (no CHS page), use {"finance_company": null, "_confidence": "high"}.
"""


VERIFY_SYSTEM = """You are a verification agent for KAS invoice extraction.

The user message contains a JSON extraction (between ```json fences) followed by
the PDF page images. Your job: re-read every digit on the document and check
the JSON extraction for errors, typos, or missing fields.

CRITICAL RULES:
1. The "corrected" field MUST be a JSON object with the SAME schema as the input
   extraction — i.e. {"invoices": [...]} where each invoice has invoice_number,
   patron_number, sold_date, retailer_name, retailer_city, retailer_state,
   grower (with first_name, last_name, company_name, address1, address2, city,
   state, zip), line_items (each with item_number, description, unit, quantity,
   unit_price, ext_amount, manufacturer), invoice_total, account_charge_amount,
   prepaid_amount, finance (with finance_company, loan_number, loan_year,
   product_rate, amount_to_retailer, batch_number, ach_date,
   manufacturer_from_notes), extraction_notes.
2. If the input JSON is correct, return corrected = input JSON unchanged.
3. If you find errors, return corrected with the fixed values.
4. NEVER invent a different schema. NEVER omit required fields.

Output ONLY a JSON object (no markdown, no prose around it):

{
  "discrepancies": [
    {"path": "invoices[0].invoice_total", "extracted": 21748.50, "actual": 21478.50, "evidence": "bottom-right of page 1 reads '21,478.50'"}
  ],
  "missing_fields": ["invoices[0].grower.company_name"],
  "math_errors": ["sum of line items (21000.00) != invoice_total (21748.50)"],
  "corrected": { "invoices": [...] },
  "overall_confidence": "high"
}

If no discrepancies, return discrepancies=[], missing_fields=[], math_errors=[],
corrected=<input JSON unchanged>, overall_confidence="high".
"""


def _first_text_block(msg) -> str:
    """Return the first text block from a Claude response.

    Guards against empty content or non-text blocks (tool_use, refusals, etc.)
    that would otherwise IndexError on ``msg.content[0].text``.
    """
    for block in getattr(msg, "content", []) or []:
        text = getattr(block, "text", None)
        if text:
            return text
    raise ValueError("Claude response contained no text content. "
                     "The model may have returned a tool-use or refusal block.")


def _parse_json(text: str) -> dict:
    """Parse JSON from a model response that may include prose or fences."""
    s = text.strip()
    # strip code fences if present
    if "```" in s:
        m = re.search(r"```(?:json)?\s*(.*?)```", s, re.DOTALL)
        if m:
            s = m.group(1).strip()
    # find first { ... } balanced span
    if not s.startswith("{"):
        first = s.find("{")
        last = s.rfind("}")
        if first >= 0 and last > first:
            s = s[first:last + 1]
    return json.loads(s)


def extract_pdf(
    pdf_bytes: bytes,
    api_key: str,
    *,
    verify: bool = True,
    model: str = MODEL,
) -> dict:
    """Extract invoices from a PDF. Returns a dict with keys:

        {"pass1": <extractor JSON>,
         "pass2": <verifier JSON> | None,
         "final": <corrected JSON if verify else pass1>,
         "agreement": <per-field comparison stats>,
         "usage": {"input_tokens": N, "output_tokens": N}}
    """
    client = Anthropic(api_key=api_key)
    pages = pdf_render.render_pdf_pages(pdf_bytes)
    if not pages:
        return {"pass1": {"invoices": []}, "pass2": None, "final": {"invoices": []},
                "agreement": {}, "usage": {"input_tokens": 0, "output_tokens": 0}}

    # ---- Pass 1: extract ----
    user_blocks: list[dict] = list(pages) + [{
        "type": "text",
        "text": ("Extract every invoice and every line item from these PDF pages. "
                 "Return JSON only, conforming exactly to the schema in the system prompt."),
    }]
    msg1 = client.messages.create(
        model=model,
        max_tokens=8000,
        system=EXTRACT_SYSTEM,
        messages=[{"role": "user", "content": user_blocks}],
    )
    pass1_text = _first_text_block(msg1)
    pass1 = _parse_json(pass1_text)
    usage = {
        "input_tokens": msg1.usage.input_tokens,
        "output_tokens": msg1.usage.output_tokens,
    }

    pass2 = None
    final = pass1
    agreement: dict = {}

    if verify:
        # ---- Pass 2: verify against the same images ----
        # Put the JSON FIRST so it's visible before the images.
        verify_blocks: list[dict] = [{
            "type": "text",
            "text": (
                "Here is a JSON extraction produced by a previous agent. Verify it "
                "against the PDF page images that follow. Return the verifier JSON "
                "with `corrected` containing the same schema as below.\n\n"
                "EXTRACTION TO VERIFY:\n```json\n"
                + json.dumps(pass1, indent=2) + "\n```\n\n"
                "PDF page images follow:"
            ),
        }] + list(pages) + [{
            "type": "text",
            "text": "Now produce your verifier JSON. Schema: {discrepancies:[],missing_fields:[],math_errors:[],corrected:{invoices:[...]},overall_confidence:\"...\"}.",
        }]
        msg2 = client.messages.create(
            model=model,
            max_tokens=8000,
            system=VERIFY_SYSTEM,
            messages=[{"role": "user", "content": verify_blocks}],
        )
        pass2 = _parse_json(_first_text_block(msg2))
        usage["input_tokens"] += msg2.usage.input_tokens
        usage["output_tokens"] += msg2.usage.output_tokens
        final = pass2.get("corrected", pass1)
        agreement = _compare(pass1, final)

    return {
        "pass1": pass1,
        "pass2": pass2,
        "final": final,
        "agreement": agreement,
        "usage": usage,
    }


def _compare(a: dict, b: dict) -> dict:
    """Compute simple per-invoice agreement stats between two extractions."""
    a_invs = a.get("invoices", [])
    b_invs = b.get("invoices", [])
    by_inv: dict[str, dict] = {}
    for ai, bi in zip(a_invs, b_invs):
        inv_no = str(ai.get("invoice_number") or bi.get("invoice_number") or "?")
        diffs: list[str] = []
        for key in ("invoice_total", "account_charge_amount", "prepaid_amount",
                    "patron_number", "sold_date"):
            if ai.get(key) != bi.get(key):
                diffs.append(f"{key}: {ai.get(key)} -> {bi.get(key)}")
        # Compare line item count
        if len(ai.get("line_items", [])) != len(bi.get("line_items", [])):
            diffs.append(f"line_items count: {len(ai.get('line_items', []))} -> {len(bi.get('line_items', []))}")
        by_inv[inv_no] = {"discrepancies": diffs}
    return by_inv


# ---- Adapter: extracted JSON -> workbook dataclasses ----

def _parse_date(s: str | None) -> datetime | None:
    if not s:
        return None
    for fmt in ("%Y-%m-%d", "%m/%d/%Y", "%m/%d/%y", "%Y-%m-%dT%H:%M:%S"):
        try:
            return datetime.strptime(s.strip(), fmt)
        except (ValueError, AttributeError):
            continue
    return None


def _to_float(v: Any) -> float | None:
    if v is None or v == "":
        return None
    try:
        return float(str(v).replace(",", "").replace("$", "").strip())
    except (ValueError, TypeError):
        return None


def _to_int(v: Any) -> int | str | None:
    if v is None or v == "":
        return None
    s = str(v).strip()
    if s.isdigit():
        return int(s)
    return s


def to_invoice_bundles(extracted: dict, *, pdf_filename: str | None = None):
    """Convert extractor JSON to a list of workbook.InvoiceBundle objects.

    Applies normalization (units, manufacturer, retailer, finance company) on the
    way in. Historical rows in Sheet1 are never rewritten.
    """
    from .workbook import LineItem, FinanceDetail, InvoiceBundle  # local import avoids cycle

    bundles: list[InvoiceBundle] = []
    for inv in extracted.get("invoices", []):
        invoice_no = inv.get("invoice_number")
        finance = inv.get("finance") or {}
        grower = inv.get("grower") or {}
        retailer_name = norm.normalize_retailer(inv.get("retailer_name"))
        finance_co = norm.normalize_finance_company(finance.get("finance_company"))
        manufacturer = norm.normalize_manufacturer(finance.get("manufacturer_from_notes"))
        invoice_date = _parse_date(inv.get("sold_date"))

        line_items: list[LineItem] = []
        for li in inv.get("line_items", []) or []:
            line_items.append(LineItem(
                finance_company=finance_co,
                manufacturer_name=norm.normalize_manufacturer(li.get("manufacturer")) or manufacturer,
                retailer_name=retailer_name,
                retailer_city=(inv.get("retailer_city") or "").strip().upper() or None,
                retailer_state=norm.normalize_state(inv.get("retailer_state")),
                grower_id=_to_int(inv.get("patron_number")),
                grower_first_name=(grower.get("first_name") or "").strip().upper() or None,
                grower_last_name=(grower.get("last_name") or "").strip().upper() or None,
                grower_company_name=(grower.get("company_name") or "").strip().upper() or None,
                grower_address1=(grower.get("address1") or "").strip().upper() or None,
                grower_address2=(grower.get("address2") or "").strip().upper() or None,
                grower_city=(grower.get("city") or "").strip().upper() or None,
                grower_state=norm.normalize_state(grower.get("state")),
                grower_zip=norm.normalize_zip(grower.get("zip")),
                item_description=(li.get("description") or "").strip().upper() or None,
                invoice_date=invoice_date,
                invoice_number=_to_int(invoice_no),
                unit=norm.normalize_unit(li.get("unit")),
                quantity=_to_float(li.get("quantity")),
                sum_total_price=_to_float(li.get("ext_amount")),
            ))

        finance_detail = None
        if finance and any(v not in (None, "") for v in finance.values() if not str(v).startswith("_")):
            finance_detail = FinanceDetail(
                invoice_number=_to_int(invoice_no),
                patron_number=_to_int(inv.get("patron_number")),
                loan_number=str(finance.get("loan_number") or "").strip() or None,
                loan_year=_to_int(finance.get("loan_year")),
                finance_company=finance_co,
                product_rate=finance.get("product_rate"),
                batch_number=str(finance.get("batch_number") or "").strip() or None,
                ach_date=_parse_date(finance.get("ach_date")),
                invoice_total=_to_float(inv.get("invoice_total")),
                amount_to_retailer=_to_float(finance.get("amount_to_retailer")),
                prepaid_amount=_to_float(inv.get("prepaid_amount")),
                account_charge_amount=_to_float(inv.get("account_charge_amount")),
                merchandised_by=inv.get("merchandised_by"),
                pdf_source_file=pdf_filename,
                needs_review=not finance_co,
                notes=inv.get("extraction_notes"),
            )

        bundles.append(InvoiceBundle(
            invoice_number=invoice_no or "?",
            line_items=line_items,
            finance=finance_detail,
            pdf_source_file=pdf_filename,
        ))
    return bundles
