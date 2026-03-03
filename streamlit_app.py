import sys
import json
import tempfile
from pathlib import Path
import io

import streamlit as st
import pandas as pd
import openpyxl

import run_pricing
from customer_access import require_customer_access, log_event


# ===============================
# CUSTOMER ACCESS + TRACKING
# ===============================
customer_id = require_customer_access()
log_event(customer_id, "page_view")
# ===============================


APP_DIR = Path(__file__).parent
RATES_FILE = APP_DIR / "RETOOL ALL COST UPLOAD 2026 WITH FUEL TYPE.xlsx"
TEMPLATE_FILE = APP_DIR / "Shipment Template.xlsx"

TRANSIT_BASE = APP_DIR / "TRANSIT_BASE.xlsx"
TRANSIT_OVERRIDES = APP_DIR / "TRANSIT_CITY_OVERRIDES.xlsx"
CITYCLASS_MASTER = APP_DIR / "CITYCLASS_MASTER.xlsx"

LOGO_FILE = APP_DIR / "logo.png"

st.set_page_config(page_title="Equator Calculator", layout="wide")


@st.cache_data(show_spinner=False)
def read_fuel_table_cached(rates_path_str: str) -> pd.DataFrame | None:
    rates_path = Path(rates_path_str)
    try:
        wb = openpyxl.load_workbook(rates_path, data_only=True)
        if "FUEL" not in wb.sheetnames:
            return None
        ws = wb["FUEL"]
        headers = [ws.cell(1, c).value for c in range(1, ws.max_column + 1)]
        headers = [str(h).strip() if h is not None else "" for h in headers]
        rows = []
        for r in range(2, ws.max_row + 1):
            row = [ws.cell(r, c).value for c in range(1, ws.max_column + 1)]
            if any(v is not None and str(v).strip() != "" for v in row):
                rows.append(row)
        return pd.DataFrame(rows, columns=headers)
    except Exception:
        return None


def normalise_pct_input(x: float) -> float:
    v = float(x)
    return v / 100.0 if v > 1.0 else v


def ensure_to_city_in_template_bytes(template_bytes: bytes) -> bytes:
    bio = io.BytesIO(template_bytes)
    wb = openpyxl.load_workbook(bio)
    ws = wb.active
    headers = [ws.cell(1, c).value for c in range(1, ws.max_column + 1)]
    headers_str = [str(h).strip() if h else "" for h in headers]

    if "To City" not in headers_str:
        try:
            idx = headers_str.index("To Country") + 1
        except ValueError:
            idx = len(headers_str)
        ws.insert_cols(idx + 1, 1)
        ws.cell(1, idx + 1).value = "To City"

    out = io.BytesIO()
    wb.save(out)
    return out.getvalue()


def safe_show_logo():
    if not LOGO_FILE.exists():
        return
    try:
        # Read bytes first so Streamlit/PIL validates; catch failures
        b = LOGO_FILE.read_bytes()
        st.image(b, use_container_width=True)
    except Exception:
        # Don’t crash the app if the image is bad
        pass


# ---------- Header ----------
header_cols = st.columns([1.2, 6, 1])
with header_cols[0]:
    safe_show_logo()
with header_cols[1]:
    st.markdown("## Pricing Calculator")
    st.caption("Customer calculator")

st.divider()


# ---------------------------
# Required file checks
# ---------------------------
for f in [RATES_FILE, TRANSIT_BASE, TRANSIT_OVERRIDES, CITYCLASS_MASTER]:
    if not f.exists():
        st.error(f"Missing required file: {f.name}")
        st.stop()


# ---------------------------
# Template download
# ---------------------------
st.subheader("Shipment template")

if TEMPLATE_FILE.exists():
    templ_bytes = ensure_to_city_in_template_bytes(TEMPLATE_FILE.read_bytes())
    st.download_button(
        label="Download Shipment Template",
        data=templ_bytes,
        file_name="Shipment Template.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )

st.divider()


# ---------------------------
# Inputs form
# ---------------------------
with st.form("inputs_form"):
    carriers_selected = st.multiselect(
        "Carriers",
        ["UPS", "DHL", "FEDEX"],
        default=["UPS", "DHL", "FEDEX"],
    )

    types_selected = st.multiselect(
        "Service types",
        ["EXPRESS", "ECONOMY", "DOMESTIC", "FREIGHT"],
        default=["EXPRESS", "ECONOMY", "DOMESTIC", "FREIGHT"],
    )

    include_customs = st.checkbox("Include duties & taxes", value=False)

    default_origin = st.text_input("Default Origin", value="GB")
    default_currency = st.selectbox("Currency", ["GBP", "EUR", "USD"])
    default_incoterm = st.selectbox("Incoterm", ["DAP", "DDP"])
    default_hs = st.text_input("Default HS", value="49111090")
    default_item_value = st.number_input("Default item value", 0.0, 100000.0, 10.0)

    divisor = st.number_input("Volumetric divisor", 1000, 20000, 5000)
    max_parcels = st.number_input("Max parcels", 1, 5, 5)

    upload = st.file_uploader("Upload shipment file", type=["xlsx", "xlsm"])
    run_btn = st.form_submit_button("Calculate Prices")


# ---------------------------
# Engine
# ---------------------------
def run_engine(df: pd.DataFrame, tmpdir: Path):
    ship_path = tmpdir / "shipments.xlsx"
    out_path = tmpdir / "output.xlsx"

    df.to_excel(ship_path, index=False)

    old_argv = sys.argv[:]
    try:
        sys.argv = [
            "run_pricing.py",
            "--rates", str(RATES_FILE),
            "--shipments", str(ship_path),
            "--out", str(out_path),
            "--divisor", str(int(divisor)),
            "--max_parcels", str(int(max_parcels)),
            "--carriers", ",".join(carriers_selected),
            "--types", ",".join(types_selected),
            "--include_customs", "1" if include_customs else "0",
        ]
        run_pricing.main()
    finally:
        sys.argv = old_argv

    return pd.read_excel(out_path, sheet_name="Prices_All_Services"), out_path


# ---------------------------
# Run handler
# ---------------------------
if run_btn:
    log_event(customer_id, "price_run")

    if upload is None:
        st.error("Please upload a shipment file.")
        st.stop()

    with st.spinner("Running pricing..."):
        with tempfile.TemporaryDirectory() as td:
            td = Path(td)
            df = pd.read_excel(upload, dtype=str)
            results, out_path = run_engine(df, td)

            st.success("Done")
            st.dataframe(results, use_container_width=True, height=500)

            st.download_button(
                "Download pricing_output.xlsx",
                data=out_path.read_bytes(),
                file_name="pricing_output.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
