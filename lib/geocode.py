"""Cached geocoder.

Primary: Google Maps Geocoding API (when GOOGLE_MAPS_API_KEY is configured).
Fallback: Nominatim (free, OSM-based) — used only if Google is unavailable.

Cache lives in ``geocache.json`` next to the project root. Each entry stores
the resolved coords plus the provider so we know whether to refresh stale
fallbacks once a paid key is added.
"""
from __future__ import annotations
import json
import time
import urllib.parse
import urllib.request
from pathlib import Path
from typing import Optional

import streamlit as st
from geopy.geocoders import Nominatim
from geopy.exc import GeocoderTimedOut, GeocoderServiceError


CACHE_FILE = Path(__file__).resolve().parent.parent / "geocache.json"
NOMINATIM_USER_AGENT = "kas_transactions_app (kentuckyamericanseeds.com)"


def _api_key() -> str | None:
    try:
        return st.secrets.get("GOOGLE_MAPS_API_KEY") or None
    except Exception:
        return None


def _load_cache() -> dict:
    if CACHE_FILE.exists():
        try:
            return json.loads(CACHE_FILE.read_text())
        except Exception:
            return {}
    return {}


def _save_cache(cache: dict) -> None:
    try:
        CACHE_FILE.write_text(json.dumps(cache, indent=2))
    except Exception:
        pass  # read-only filesystem (e.g. Streamlit Cloud) — cache lives in memory only


def _norm_addr(addr1: str | None, city: str | None,
               state: str | None, zp: str | int | None) -> str:
    parts = [str(p).strip() for p in (addr1, city, state, zp) if p]
    return ", ".join(parts).upper()


def _format_query(addr1, city, state, zp) -> str:
    parts = [str(p).strip() for p in (addr1, city, state, zp) if p]
    return ", ".join(parts) + ", USA"


def _google_geocode(query: str, key: str) -> Optional[tuple[float, float]]:
    url = (
        "https://maps.googleapis.com/maps/api/geocode/json?"
        + urllib.parse.urlencode({"address": query, "key": key, "region": "us"})
    )
    try:
        with urllib.request.urlopen(url, timeout=10) as resp:
            data = json.loads(resp.read())
    except Exception:
        return None
    if data.get("status") != "OK" or not data.get("results"):
        return None
    loc = data["results"][0]["geometry"]["location"]
    return float(loc["lat"]), float(loc["lng"])


def _nominatim_geocode(query: str) -> Optional[tuple[float, float]]:
    geocoder = Nominatim(user_agent=NOMINATIM_USER_AGENT, timeout=10)
    try:
        time.sleep(1.05)  # rate limit
        loc = geocoder.geocode(query, country_codes="us")
        if loc:
            return loc.latitude, loc.longitude
    except (GeocoderTimedOut, GeocoderServiceError):
        return None
    return None


def geocode_address(addr1: str | None, city: str | None,
                    state: str | None = "KY",
                    zp: str | int | None = None,
                    *, force: bool = False) -> Optional[tuple[float, float]]:
    """Return (lat, lon) for an address, or None if no match.

    Tries: full address → city+state → city only. Caches every call (including
    failures) so we don't keep hitting the API for known-bad addresses.
    """
    key = _norm_addr(addr1, city, state, zp)
    if not key:
        return None
    cache = _load_cache()
    if not force and key in cache:
        entry = cache[key]
        if isinstance(entry, list):
            return tuple(entry) if entry else None
        if isinstance(entry, dict):
            coords = entry.get("coords")
            return tuple(coords) if coords else None

    api_key = _api_key()
    queries = [_format_query(addr1, city, state, zp)]
    if city and state:
        queries.append(f"{city}, {state}, USA")
    if zp:
        queries.append(f"{zp}, USA")

    coords = None
    provider = None
    for q in queries:
        if api_key:
            coords = _google_geocode(q, api_key)
            if coords:
                provider = "google"
                break
        # Nominatim fallback (also primary if no Google key)
        coords = _nominatim_geocode(q)
        if coords:
            provider = provider or "nominatim"
            break

    cache[key] = {"coords": list(coords) if coords else None,
                  "provider": provider,
                  "ts": int(time.time())}
    _save_cache(cache)
    return coords


def batch_geocode(rows: list[dict]) -> dict[str, tuple[float, float] | None]:
    """Geocode many rows; returns map of normalized-key -> (lat,lon) or None."""
    out: dict[str, tuple[float, float] | None] = {}
    for r in rows:
        key = _norm_addr(r.get("addr1"), r.get("city"), r.get("state"), r.get("zip"))
        if key in out:
            continue
        out[key] = geocode_address(
            r.get("addr1"), r.get("city"), r.get("state"), r.get("zip")
        )
    return out


def cache_stats() -> dict:
    cache = _load_cache()
    total = len(cache)
    by_provider: dict[str, int] = {}
    hits = 0
    for v in cache.values():
        if isinstance(v, dict):
            p = v.get("provider") or "miss"
            by_provider[p] = by_provider.get(p, 0) + 1
            if v.get("coords"):
                hits += 1
        elif isinstance(v, list):
            by_provider["legacy"] = by_provider.get("legacy", 0) + 1
            if v:
                hits += 1
    return {"total": total, "hits": hits, "misses": total - hits, "by_provider": by_provider}


def clear_failed_cache() -> int:
    """Remove cached failures (None entries) so we retry next time."""
    cache = _load_cache()
    removed = 0
    keep: dict = {}
    for k, v in cache.items():
        coords = v.get("coords") if isinstance(v, dict) else v
        if coords:
            keep[k] = v
        else:
            removed += 1
    _save_cache(keep)
    return removed
