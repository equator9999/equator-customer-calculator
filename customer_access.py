import os
import hmac
import hashlib
import json
import base64
from datetime import datetime, timezone

import streamlit as st


def _now_iso() -> str:
    return datetime.now(timezone.utc).isoformat()


def _get_query_params() -> dict:
    """
    Robustly read query params across Streamlit versions.

    Returns a plain dict: {key: [values...]}
    """
    # New Streamlit API
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

    # Older Streamlit API
    try:
        return st.experimental_get_query_params()
    except Exception:
        return {}


def _sha256_hex(s: str) -> str:
    return hashlib.sha256(s.encode("utf-8")).hexdigest()


def _hmac_sha256_hex(secret: str, msg: str) -> str:
    return hmac.new(secret.encode("utf-8"), msg.encode("utf-8"), hashlib.sha256).hexdigest()


def _hmac_sha256_b64url(secret: str, msg: str) -> str:
    raw = hmac.new(secret.encode("utf-8"), msg.encode("utf-8"), hashlib.sha256).digest()
    return base64.urlsafe_b64encode(raw).decode("utf-8").rstrip("=")


def _constant_time_equal(a: str, b: str) -> bool:
    try:
        return hmac.compare_digest(a, b)
    except Exception:
        return False


def _candidate_sigs(secret: str, customer_id: str) -> list[str]:
    """
    Backwards-compatible signature candidates.

    We don't know which scheme your old link generator used, so we accept
    a small whitelist of common schemes.
    """
    c = customer_id

    candidates = []

    # HMAC variants (hex)
    candidates.append(_hmac_sha256_hex(secret, c))
    candidates.append(_hmac_sha256_hex(secret, f"c={c}"))
    candidates.append(_hmac_sha256_hex(secret, f"{c}:{secret}"))  # uncommon but seen

    # HMAC variants (base64url)
    candidates.append(_hmac_sha256_b64url(secret, c))
    candidates.append(_hmac_sha256_b64url(secret, f"c={c}"))

    # Plain SHA256 variants (hex)
    candidates.append(_sha256_hex(secret + c))
    candidates.append(_sha256_hex(c + secret))
    candidates.append(_sha256_hex(f"{c}:{secret}"))
    candidates.append(_sha256_hex(f"{secret}:{c}"))

    # Normalise: some generators output uppercase hex
    out = []
    for s in candidates:
        out.append(s)
        out.append(s.lower())
        out.append(s.upper())
    # Deduplicate while preserving order
    seen = set()
    uniq = []
    for s in out:
        if s not in seen:
            seen.add(s)
            uniq.append(s)
    return uniq


def require_customer_access() -> str:
    """
    Enforce signed-link access:
      ?c=<customer_id>&sig=<signature>

    Stores verified customer_id in session_state so reruns never re-check.
    """
    if st.session_state.get("customer_id"):
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
        st.error("Invalid link: missing parameters.")
        st.caption("Debug (what the app received):")
        st.code(json.dumps(qp, indent=2))
        st.stop()

    # Validate against known candidate schemes
    candidates = _candidate_sigs(secret, customer_id)

    if not any(_constant_time_equal(sig, exp) for exp in candidates):
        # Show safe debug prefixes to help identify which scheme is in use
        st.error("Invalid link: signature mismatch.")
        st.caption("Debug (what the app received):")
        st.code(
            json.dumps(
                {
                    "c": customer_id,
                    "sig_prefix": sig[:8],
                    "expected_prefixes": [c[:8] for c in candidates[:6]],
                },
                indent=2,
            )
        )
        st.stop()

    st.session_state["customer_id"] = customer_id
    return customer_id


def log_event(customer_id: str, event_name: str, meta: dict | None = None) -> None:
    """
    Lightweight logging to stdout (Render logs).
    """
    try:
        payload = {
            "ts": _now_iso(),
            "customer_id": customer_id,
            "event": event_name,
            "meta": meta or {},
        }
        print(json.dumps(payload))
    except Exception:
        pass
