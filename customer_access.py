import os
import hmac
import hashlib
import json
from datetime import datetime, timezone

import streamlit as st


def _now_iso() -> str:
    return datetime.now(timezone.utc).isoformat()


def _get_query_params() -> dict:
    """
    Robustly read query params across Streamlit versions.

    Returns a plain dict: {key: [values...] } like the older API.
    """
    # Streamlit >= 1.30 has st.query_params (Mapping[str, str | list[str]])
    try:
        qp = st.query_params  # type: ignore[attr-defined]
        out = {}
        for k, v in qp.items():
            if isinstance(v, list):
                out[k] = v
            else:
                out[k] = [v]
        return out
    except Exception:
        pass

    # Older Streamlit
    try:
        return st.experimental_get_query_params()
    except Exception:
        return {}


def _hmac_sig(secret: str, customer_id: str) -> str:
    """
    Signature function used for signed links.

    IMPORTANT:
    - If your existing links were generated using a different scheme
      (e.g., including a timestamp, or using sha256(secret+customer_id)),
      you must match that exact scheme here.

    This implements: sig = hex(hmac_sha256(secret, customer_id)).
    """
    msg = customer_id.encode("utf-8")
    key = secret.encode("utf-8")
    return hmac.new(key, msg, hashlib.sha256).hexdigest()


def _constant_time_equal(a: str, b: str) -> bool:
    try:
        return hmac.compare_digest(a, b)
    except Exception:
        return False


def require_customer_access() -> str:
    """
    Enforce signed-link access:
      ?c=<customer_id>&sig=<signature>

    Stores verified customer_id in session_state so reruns never re-check.
    """
    if "customer_id" in st.session_state and st.session_state["customer_id"]:
        return str(st.session_state["customer_id"])

    secret = os.getenv("CUSTOMER_LINK_SECRET", "")
    if not secret:
        st.error("Server misconfigured: CUSTOMER_LINK_SECRET is not set.")
        st.stop()

    qp = _get_query_params()

    c_vals = qp.get("c", []) or qp.get("customer_id", []) or qp.get("customer", [])
    s_vals = qp.get("sig", []) or qp.get("signature", [])

    customer_id = (c_vals[0] if c_vals else "").strip()
    sig = (s_vals[0] if s_vals else "").strip()

    if not customer_id or not sig:
        # Debug: show what query params arrived (safe)
        st.error("Invalid link: missing parameters.")
        st.caption("Debug (what the app received):")
        st.code(json.dumps(qp, indent=2))
        st.stop()

    expected = _hmac_sig(secret, customer_id)

    if not _constant_time_equal(sig, expected):
        # Debug: show partial comparison safely (do NOT show secret)
        st.error("Invalid link: signature mismatch.")
        st.caption("Debug (what the app received):")
        st.code(json.dumps({"c": customer_id, "sig_prefix": sig[:8], "expected_prefix": expected[:8]}, indent=2))
        st.stop()

    st.session_state["customer_id"] = customer_id
    return customer_id


def log_event(customer_id: str, event_name: str, meta: dict | None = None) -> None:
    """
    Lightweight logging. Replace with your real logger if needed.
    Keeps failures non-fatal.
    """
    try:
        payload = {
            "ts": _now_iso(),
            "customer_id": customer_id,
            "event": event_name,
            "meta": meta or {},
        }
        # If you have a log sink, write it here.
        # For now: print goes to Render logs.
        print(json.dumps(payload))
    except Exception:
        pass
