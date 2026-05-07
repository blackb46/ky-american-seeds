"""Shared-password gate with persistent cookie.

Uses extra-streamlit-components for a signed cookie. The user enters the password
once per browser; an HMAC-signed cookie keeps them logged in for 30 days.
"""
from __future__ import annotations
import hmac
import hashlib
import time
import json
import base64
from datetime import datetime, timedelta

import streamlit as st
import extra_streamlit_components as stx


COOKIE_NAME = "kas_auth"
COOKIE_DAYS = 30


def _cookie_manager():
    # Must NOT be cached: CookieManager renders a hidden widget on every run,
    # and Streamlit forbids widgets inside cache decorators. Use session_state
    # to keep one instance per session instead.
    if "_cm" not in st.session_state:
        st.session_state["_cm"] = stx.CookieManager(key="kas_cookie_mgr")
    return st.session_state["_cm"]


def _sign(payload: str, key: str) -> str:
    return hmac.new(key.encode(), payload.encode(), hashlib.sha256).hexdigest()


def _make_cookie(key: str) -> str:
    expires = int((datetime.utcnow() + timedelta(days=COOKIE_DAYS)).timestamp())
    payload = json.dumps({"v": 1, "exp": expires}, separators=(",", ":"))
    sig = _sign(payload, key)
    raw = base64.urlsafe_b64encode(payload.encode()).decode().rstrip("=")
    return f"{raw}.{sig}"


def _verify_cookie(value: str | None, key: str) -> bool:
    if not value or "." not in value:
        return False
    try:
        raw, sig = value.rsplit(".", 1)
        padded = raw + "=" * (-len(raw) % 4)
        payload = base64.urlsafe_b64decode(padded.encode()).decode()
        if not hmac.compare_digest(sig, _sign(payload, key)):
            return False
        data = json.loads(payload)
        return int(data.get("exp", 0)) > int(time.time())
    except Exception:
        return False


def require_login() -> bool:
    """Render password gate. Returns True once the user is authenticated."""
    if st.session_state.get("_authed"):
        return True

    # Validate required secrets up-front so missing config produces a clear
    # error instead of a silent KeyError that crashes the app on startup.
    missing = [k for k in ("COOKIE_SIGNING_KEY", "APP_PASSWORD") if k not in st.secrets]
    if missing:
        st.error(
            f"Configuration error: missing required secret(s): {', '.join(missing)}. "
            "Add these in the app's Streamlit Cloud secrets settings and reload."
        )
        st.stop()

    cookie_key = st.secrets["COOKIE_SIGNING_KEY"]
    expected_pw = st.secrets["APP_PASSWORD"]

    cm = _cookie_manager()
    # Read cookies once. If the iframe hasn't responded yet, we just show the
    # login form — repeated `st.rerun()` calls during cold-start can trip
    # Streamlit Cloud's startup watchdog ("Code: 1ST"). One single retry is
    # enough in practice; after that the user can just sign in again.
    cookies = cm.get_all() or {}
    existing = cookies.get(COOKIE_NAME) or cm.get(COOKIE_NAME)
    if _verify_cookie(existing, cookie_key):
        st.session_state["_authed"] = True
        return True

    # Short polling for the cookie iframe to post back. Splitting into a few
    # 80 ms ticks instead of one 400 ms blocking sleep lets the page paint
    # the login form sooner (avoids the visible login-flash → authed-view
    # bounce on cold start). Cap the total wait so we never trip the
    # Streamlit Cloud startup watchdog.
    if not st.session_state.get("_cookie_retried"):
        st.session_state["_cookie_retried"] = True
        for _ in range(5):
            time.sleep(0.08)
            cookies = cm.get_all() or {}
            existing = cookies.get(COOKIE_NAME) or cm.get(COOKIE_NAME)
            if _verify_cookie(existing, cookie_key):
                st.session_state["_authed"] = True
                return True
        st.rerun()

    st.markdown("### 🔒 Sign in")
    pw = st.text_input("Password", type="password", key="_pw_input",
                       help="Ask the admin if you don't have it.")
    col1, col2 = st.columns([1, 4])
    with col1:
        submit = st.button("Sign in", type="primary")
    with col2:
        remember = st.checkbox("Remember me on this browser (30 days)", value=True)

    if submit:
        if hmac.compare_digest(pw, expected_pw):
            st.session_state["_authed"] = True
            if remember:
                cm.set(
                    COOKIE_NAME,
                    _make_cookie(cookie_key),
                    expires_at=datetime.utcnow() + timedelta(days=COOKIE_DAYS),
                    key="set_cookie",
                )
                # The iframe needs a moment to actually persist the cookie
                # to document.cookie before the rerun, otherwise refreshes
                # within the next ~500ms won't see it.
                time.sleep(0.5)
            st.rerun()
        else:
            st.error("Incorrect password.")
    return False


def logout() -> None:
    cm = _cookie_manager()
    cm.delete(COOKIE_NAME, key="del_cookie")
    st.session_state["_authed"] = False
    st.rerun()
