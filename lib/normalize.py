"""Normalize extracted values to match historical conventions in Sheet1.

Only applied to NEW rows. Historical rows are never rewritten.
"""
from __future__ import annotations
import re

UNIT_MAP = {
    "GAL": "GAL", "GAL ": "GAL", "GALLON": "GAL", "GALLONS": "GAL",
    "GALLOON": "GAL", "GALLOONS": "GAL", "GA": "GAL",
    "OZ": "OZ", "OUNCE": "OZ", "OUNCES": "OZ", "FL OZ": "OZ", "FLOZ": "OZ",
    "QT": "QUART", "QUART": "QUART", "QUARTS": "QUART",
    "LB": "POUND", "LBS": "POUND", "POUND": "POUND", "POUNDS": "POUND",
    "TON": "TON", "TONS": "TON",
    "BAG": "BAG", "BAGS": "BAG", "BG": "BAG",
    "UNIT": "UNIT", "UNITS": "UNIT", "UN": "UNIT", "EACH": "UNIT", "EA": "UNIT",
}

MANUFACTURER_MAP = {
    "NUFRAM": "NUFARM", "NU FARM": "NUFARM",
    "SYNGENTA NK": "SYNGENTA", "SYNGENTA CP": "SYNGENTA",
    "F & F": "F&F", "F AND F": "F&F",
}

RETAILER_MAP = {
    "KAS MAD": "KAS MAD", "KAS MADISONVILLE": "KAS MAD",
    "KAS HOP": "KAS HOP", "KAS HOPKINSVILLE": "KAS HOP",
    "KAS TX": "KAS TX", "KAS TEXAS": "KAS TX", "KAS FREDONIA": "KAS TX",
    "KAS ETOWN": "KAS ETOWN", "KAS ELIZABETHTOWN": "KAS ETOWN",
    "PATRIOT AG": "PATRIOT AG", "PATRIOTAG": "PATRIOT AG",
}

FINANCE_MAP = {
    "CHS": "CHS", "CHS CAPITAL": "CHS", "CHS ": "CHS",
    "CFA": "CFA", "CFAFS": "CFA",
    "CENTRAL FARM AND FIELD SERVICES": "CFA",
}


def _clean(s: str | None) -> str:
    if s is None:
        return ""
    return re.sub(r"\s+", " ", str(s).strip()).upper()


def normalize_unit(s: str | None) -> str | None:
    k = _clean(s)
    if not k:
        return None
    return UNIT_MAP.get(k, k)


def normalize_manufacturer(s: str | None) -> str | None:
    k = _clean(s)
    if not k:
        return None
    return MANUFACTURER_MAP.get(k, k)


def normalize_retailer(s: str | None) -> str | None:
    k = _clean(s)
    if not k:
        return None
    return RETAILER_MAP.get(k, k)


def normalize_finance_company(s: str | None) -> str | None:
    k = _clean(s)
    if not k:
        return None
    return FINANCE_MAP.get(k, k)


def normalize_state(s: str | None) -> str | None:
    if not s:
        return None
    s = str(s).strip().upper()
    return s if len(s) == 2 else s


def normalize_zip(s: str | int | None) -> int | str | None:
    if s is None or s == "":
        return None
    s = str(s).strip().split("-")[0]
    if s.isdigit():
        return int(s)
    return s


def split_grower_name(full_name: str | None) -> tuple[str | None, str | None]:
    """Split 'JOHN KUEGEL' or 'KUEGEL, JOHN' into (first, last)."""
    if not full_name:
        return None, None
    s = str(full_name).strip().upper()
    if "," in s:
        last, first = [p.strip() for p in s.split(",", 1)]
        return first or None, last or None
    parts = s.split()
    if len(parts) == 1:
        return None, parts[0]
    return parts[0], " ".join(parts[1:])


def is_company_name(name: str | None) -> bool:
    """Detect if a 'grower name' is actually a farm/company."""
    if not name:
        return False
    s = str(name).upper()
    indicators = ["FARM", "FARMS", "INC", "LLC", "L.L.C", "CORP",
                  "COMPANY", " CO.", "CO ", "ENTERPRISES", "PARTNERS",
                  "BROTHERS", "BROS", "& SONS", "AND SONS", "TRUST",
                  "PROPERTIES", "RANCH"]
    return any(ind in s for ind in indicators)
