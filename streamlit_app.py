import sys
import json
import tempfile
import subprocess
from pathlib import Path
import io

import streamlit as st
import pandas as pd

from customer_access import require_customer_access, log_event


# --- Ensure required packages exist in THIS environment (Streamlit Cloud can be flaky) ---
def ensure(pkg: str) -> None:
    try:
        __import__(pkg)
    except Exception:
        subprocess.check_call([sys.executable, "-m", "pip", "install", pkg])


ensure("openpyxl")
ensure("pandas")
ensure("numpy")
ensure("requests")

import openpyxl
import run_pricing


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


# ---------------------------
# Caching helpers
# ---------------------------
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
        if not any(headers):
            return None
        rows = []
        for r in range(2, ws.max_row + 1):
            row = [ws.cell(r, c).value for c in range(1, ws.max_column + 1)]
            if all(v is None or str(v).strip() == "" for v in row):
                continue
            rows.append(row)
        return pd.DataFrame(rows, columns=headers)
    except Exception:
        return None


def normalise_pct_input(x: float) -> float:
    if x is None:
        return 0.0
    v = float(x)
    if v > 1.0:
        v = v / 100.0
    return v


def ensure_to_city_in_template_bytes(template_bytes: bytes) -> bytes:
    bio = io.BytesIO(template_bytes)
    wb = openpyxl.load_workbook(bio)
    ws = wb.active
    headers = [ws.cell(1, c).value for c in range(1, ws.max_column + 1)]
    headers_str = [str(h).strip() if h is not None else "" for h in headers]

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


def style_hs_risk(df: pd.DataFrame):
    if df is None or df.empty:
        return df.style

    def _risk_style(val):
        v = str(val).strip().upper()
        if v == "LOW":
            return "background-color: #d7f7d7; color: #0b4f0b; font-weight: 700;"
        if v in ("MED", "MEDIUM"):
            return "background-color: #fff2cc; color: #7a4b00; font-weight: 700;"
        if v == "HIGH":
            return "background-color: #ffd6d6; color: #7a0000; font-weight: 700;"
        return ""

    if "HS_Risk" in df.columns:
        return df.style.applymap(_risk_style, subset=["HS_Risk"])
    return df.style


# ---------- Header ----------
header_cols = st.columns([1.2, 6, 1])
with header_cols[0]:
    if LOGO_FILE.exists():
        st.image(str(LOGO_FILE), use_container_width=True)
with header_cols[1]:
    st.markdown("## Pricing Calculator")
    st.caption("Customer calculator")
with header_cols[2]:
    st.write("")


# ---------------------------
# Required file checks
# ---------------------------
missing = []
for f in [RATES_FILE, TRANSIT_BASE, TRANSIT_OVERRIDES, CITYCLASS_MASTER]:
    if not f.exists():
        missing.append(f.name)

if missing:
    st.error("Missing required files in repo root: " + ", ".join(missing))
    st.stop()


# ---------------------------
# 1) Template download
# ---------------------------
st.subheader("1) Shipment template")

colA, colB = st.columns([2, 3])
with colA:
    if TEMPLATE_FILE.exists():
        templ_bytes = ensure_to_city_in_template_bytes(TEMPLATE_FILE.read_bytes())
        st.download_button(
            label="Download Shipment Template (includes To City)",
            data=templ_bytes,
            file_name="Shipment Template.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )
    else:
        st.warning('Template file not found. Ensure "Shipment Template.xlsx" exists in the repo root.')

with colB:
    st.info("Use the template to upload shipments for pricing.", icon="ℹ️")

st.divider()


# ---------------------------
# Common inputs form
# ---------------------------
with st.form("inputs_form", clear_on_submit=False):
    st.subheader("2) Filters")

    f1, f2 = st.columns([2, 2])
    with f1:
        carriers_selected = st.multiselect(
            "Carriers",
            options=["UPS", "DHL", "FEDEX"],
            default=["UPS", "DHL", "FEDEX"],
        )
    with f2:
        types_selected = st.multiselect(
            "Service types",
            options=["EXPRESS", "ECONOMY", "DOMESTIC", "FREIGHT"],
            default=["EXPRESS", "ECONOMY", "DOMESTIC", "FREIGHT"],
        )

    st.subheader("3) Customs (optional)")
    include_customs = st.checkbox("Include duties & taxes (Easyship)", value=False)
    only_cheapest_customs = False
    cheapest_n = 5
    if include_customs:
        only_cheapest_customs = st.checkbox("Only call duties/taxes for cheapest options", value=True)
        if only_cheapest_customs:
            cheapest_n = st.number_input("Cheapest N per shipment (customs calls)", 1, 25, 5, 1)

    st.subheader("4) Defaults (for missing fields)")
    d1, d2, d3, d4 = st.columns([1.2, 1.2, 1.2, 2.2])
    with d1:
        default_origin = st.text_input("Default Origin (ISO2)", value="GB")
    with d2:
        default_currency = st.selectbox("Default Currency", ["GBP", "EUR", "USD"], index=0)
    with d3:
        default_incoterm = st.selectbox("Default Incoterm", ["DAP", "DDP"], index=0)
    with d4:
        default_hs = st.text_input("Default HS Code", value="49111090")
    default_item_value = st.number_input("Default item value (GBP)", min_value=0.0, value=10.0, step=1.0)

    st.subheader("5) Fuel overrides (optional)")
    fuel_df = read_fuel_table_cached(str(RATES_FILE))
    with st.expander("Show current fuel table", expanded=False):
        if fuel_df is None:
            st.info("No readable FUEL sheet found.")
        else:
            st.dataframe(fuel_df, use_container_width=True, height=220)

    override_fuel = st.checkbox("Override fuel % (this run only)", value=False)
    fuel_overrides = {}
    if override_fuel:
        carriers = ["UPS", "DHL", "FEDEX"]
        types = ["EXPRESS", "ECONOMY", "DOMESTIC", "FREIGHT"]

        prefill = {}
        if fuel_df is not None and {"Carrier", "Type", "FuelPct"}.issubset(set(fuel_df.columns)):
            for _, r in fuel_df.iterrows():
                c = str(r.get("Carrier", "")).strip().upper()
                t = str(r.get("Type", "")).strip().upper()
                p = r.get("FuelPct", None)
                try:
                    if p is not None and str(p).strip() != "":
                        p = float(p)
                        if p > 1:
                            p = p / 100.0
                        prefill[(c, t)] = p
                except Exception:
                    pass

        g = st.columns(5)
        g[0].markdown("**Carrier**")
        for i, t in enumerate(types, start=1):
            g[i].markdown(f"**{t[:4]}**")

        for c in carriers:
            row_cols = st.columns(5)
            row_cols[0].markdown(f"**{c}**")
            for i, t in enumerate(types, start=1):
                default_val = prefill.get((c, t), None)
                default_display = (default_val * 100.0) if isinstance(default_val, (int, float)) else 0.0
                val = row_cols[i].number_input(
                    label=f"{c}_{t}",
                    min_value=0.0,
                    max_value=200.0,
                    value=float(default_display),
                    step=0.1,
                    key=f"fuel_{c}_{t}",
                    label_visibility="collapsed",
                )
                if val and float(val) > 0:
                    fuel_overrides[f"{c}|{t}"] = normalise_pct_input(val)

    st.subheader("6) Upload + run")
    divisor = st.number_input("Volumetric divisor (cm)", 1000, 20000, 5000, 100)
    max_parcels = st.number_input("Max parcels per row", 1, 5, 5, 1)

    upload = st.file_uploader("Shipment file (xlsx/xlsm)", type=["xlsx", "xlsm"])
    run_btn = st.form_submit_button("Calculate Prices", type="primary")


# ---------------------------
# Engine runner
# ---------------------------
def run_engine_on_shipments(df: pd.DataFrame, tmpdir: Path):
    ship_path = tmpdir / "shipments_normalised.xlsx"
    out_path = tmpdir / "pricing_output.xlsx"

    for col in ["Origin Country", "Currency", "Incoterm", "To City"]:
        if col not in df.columns:
            df[col] = ""

    def _fill_default(col, default_val):
        df[col] = df[col].fillna("").astype(str)
        df.loc[df[col].str.strip().isin(["", "nan", "None"]), col] = str(default_val)

    _fill_default("Origin Country", default_origin)
    _fill_default("Currency", default_currency)
    _fill_default("Incoterm", default_incoterm)

    for i in range(1, int(max_parcels) + 1):
        hs_col = f"P{i}_HSCode"
        val_col = f"P{i}_Value"
        if hs_col not in df.columns:
            df[hs_col] = ""
        if val_col not in df.columns:
            df[val_col] = ""

        df[hs_col] = df[hs_col].fillna("").astype(str)
        df[val_col] = df[val_col].fillna("").astype(str)

        df.loc[df[hs_col].str.strip().isin(["", "nan", "None"]), hs_col] = str(default_hs).strip()
        df.loc[df[val_col].str.strip().isin(["", "nan", "None"]), val_col] = str(float(default_item_value))

    df.to_excel(ship_path, index=False)

    fuel_overrides_json = json.dumps(fuel_overrides) if fuel_overrides else ""

    old_argv = sys.argv[:]
    try:
        sys.argv = [
            "run_pricing.py",
            "--rates", str(RATES_FILE),
            "--shipments", str(ship_path),
            "--out", str(out_path),
            "--divisor", str(int(divisor)),
            "--max_parcels", str(int(max_parcels)),
            "--carriers", ",".join(carriers_selected) if carriers_selected else "",
            "--types", ",".join(types_selected) if types_selected else "",
            "--include_customs", "1" if include_customs else "0",
            "--customs_only_cheapest", "1" if (include_customs and only_cheapest_customs) else "0",
            "--customs_top_n", str(int(cheapest_n)),
            "--fuel_overrides_json", fuel_overrides_json,
        ]
        run_pricing.main()
    finally:
        sys.argv = old_argv

    results = pd.read_excel(out_path, sheet_name="Prices_All_Services")
    return out_path, results


# ---------------------------
# Run button handler
# ---------------------------
if run_btn:
    log_event(customer_id, "price_run")

    with st.spinner("Running…"):
        with tempfile.TemporaryDirectory() as td:
            td = Path(td)

            if upload is None:
                st.error("Please upload a shipment file.")
                st.stop()

            df = pd.read_excel(upload, dtype=str)
            out_path, results = run_engine_on_shipments(df, td)

            num_cols = [
                "ChargeableKg", "BaseFreight", "OversizeAmount", "NonConvAmount", "OverGirthAmount",
                "SubtotalNonFuel", "SubtotalBeforeFuel", "FuelAmount", "FinalTotal",
                "GoodsValueTotal", "DutyTotal", "TaxTotal", "LandedCostTotal",
            ]
            for c in num_cols:
                if c in results.columns:
                    results[c] = pd.to_numeric(results[c], errors="coerce").round(2)

            show_cols = [
                "Shipment ID", "Country", "ToCity",
                "CarrierKey", "Service", "Type",
                "FinalTotal", "IsLowestRatePerShipment",
                "TransitDays_Min", "TransitDays_Max", "TransitSource", "CityClass",
                "HSCode_Used", "HS_Risk",
                "FuelPctSource",
            ]
            show_cols = [c for c in show_cols if c in results.columns]
            preview = results[show_cols].copy()

            st.success("Done! Preview below and download.")
            st.dataframe(style_hs_risk(preview), use_container_width=True, height=520)

            st.download_button(
                "Download pricing_output.xlsx",
                data=out_path.read_bytes(),
                file_name="pricing_output.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )
