import os
import json
import hmac
import hashlib
import time
from pathlib import Path

import streamlit as st


def _get_secret(key: str, default: str = "") -> str:
    v = os.environ.get(key)
    if v is not None and str(v).strip() != "":
        return str(v)
    try:
        return str(st.secrets.get(key, default))
    except Exception:
        return default


def _get_access_json() -> dict:
    raw = _get_secret("CUSTOMER_ACCESS_JSON", "{}")
    try:
        return json.loads(raw) if raw else {}
    except Exception:
        return {}


def _sign(customer_id: str, secret: str) -> str:
    return hmac.new(
        secret.encode("utf-8"),
        customer_id.encode("utf-8"),
        hashlib.sha256,
    ).hexdigest()


def require_customer_access() -> str:
    qp = st.query_params
    customer_id = (qp.get("c") or "").strip()
    sig = (qp.get("sig") or "").strip()

    secret = _get_secret("CUSTOMER_LINK_SECRET", "")
    allow = _get_access_json()

    if not secret or not customer_id or not sig:
        st.error("Invalid link. Please use the link provided by Equator.")
        st.stop()

    expected = _sign(customer_id, secret)
    if not hmac.compare_digest(expected, sig):
        st.error("Invalid link. Please use the link provided by Equator.")
        st.stop()

    if isinstance(allow, dict) and len(allow) > 0:
        if not allow.get(customer_id, False):
            st.error("Invalid link. Please use the link provided by Equator.")
            st.stop()

    return customer_id


def log_event(customer_id: str, event: str, payload: dict | None = None) -> None:
    payload = payload or {}
    webhook = _get_secret("EVENT_LOG_WEBHOOK_URL", "").strip()

    data = {
        "ts": int(time.time()),
        "customer_id": customer_id,
        "event": event,
        "payload": payload,
    }

    if webhook:
        try:
            import requests

            requests.post(webhook, json=data, timeout=3)
            return
        except Exception:
            pass

    try:
        p = Path(__file__).parent / "events.log"
        with p.open("a", encoding="utf-8") as f:
            f.write(json.dumps(data) + "\n")
    except Exception:
        pass
