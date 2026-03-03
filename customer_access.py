import hmac
import hashlib
import json
import time

import requests
import streamlit as st


def _safe_eq(a: str, b: str) -> bool:
    try:
        return hmac.compare_digest(a.encode("utf-8"), b.encode("utf-8"))
    except Exception:
        return False


def make_sig(secret: str, customer_id: str) -> str:
    mac = hmac.new(secret.encode("utf-8"), customer_id.encode("utf-8"), hashlib.sha256)
    return mac.hexdigest()


def require_customer_access() -> str:
    params = st.query_params
    c = (params.get("c") or "").strip()
    sig = (params.get("sig") or "").strip()

    secret = st.secrets.get("CUSTOMER_LINK_SECRET", "")
    if not secret:
        st.error("Missing CUSTOMER_LINK_SECRET.")
        st.stop()

    if not c or not sig:
        st.error("Invalid link. Please use the link provided by Equator.")
        st.stop()

    expected = make_sig(secret, c)
    if not _safe_eq(sig, expected):
        st.error("Invalid link. Please use the link provided by Equator.")
        st.stop()

    # Optional allowlist/revoke map: {"acme": true, "globex": false}
    access_json = st.secrets.get("CUSTOMER_ACCESS_JSON", "")
    if access_json:
        try:
            m = json.loads(access_json)
            if not bool(m.get(c, False)):
                st.error("Access disabled. Please contact Equator.")
                st.stop()
        except Exception:
            st.error("Invalid CUSTOMER_ACCESS_JSON in secrets.")
            st.stop()

    return c


def log_event(customer_id: str, event: str, meta: dict | None = None) -> None:
    url = st.secrets.get("EVENT_LOG_WEBHOOK_URL", "")
    if not url:
        return

    payload = {
        "ts": int(time.time()),
        "customer_id": customer_id,
        "event": event,
        "meta": meta or {},
        "utm_source": st.query_params.get("utm_source", ""),
        "utm_campaign": st.query_params.get("utm_campaign", ""),
        "utm_medium": st.query_params.get("utm_medium", ""),
    }

    try:
        requests.post(url, json=payload, timeout=3)
    except Exception:
        pass
