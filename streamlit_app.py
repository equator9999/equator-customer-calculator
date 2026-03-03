import sys
import tempfile
from pathlib import Path
import io

import streamlit as st
import pandas as pd
import openpyxl

import run_pricing
from customer_access import require_customer_access, log_event


# ----------------------------
# Page config FIRST
# ----------------------------
APP_DIR = Path(__file__).parent
LOGO_FILE = APP_DIR / "logo.png"

st.set_page_config(page_title="Equator Portal", layout="wide", page_icon="📦")

# ---------- UI polish ----------
st.markdown(
    """
<style>
.block-container { padding-top: 1.4rem; padding-bottom: 2.6rem; max-width: 1200px; }
h1, h2, h3 { letter-spacing: -0.02em; }
div[data-testid="stToolbar"] { visibility: hidden; height: 0px; }

div[data-baseweb="input"] input,
div[data-baseweb="select"] > div {
  border-radius: 12px !important;
}

.eq-card {
  background: white;
  border: 1px solid rgba(15, 23, 42, 0.08);
  border-radius: 16px;
  padding: 18px 18px 14px 18px;
  box-shadow: 0 1px 2px rgba(15, 23, 42, 0.05);
}
.eq-title {
  font-size: 12px;
  font-weight: 900;
  letter-spacing: 0.06em;
  color: rgba(15, 23, 42, 0.65);
  text-transform: uppercase;
  margin-bottom: 10px;
}
.eq-muted { color: rgba(15,23,42,0.62); font-size: 0.92rem; line-height: 1.2rem; }

.eq-best {
  display: inline-block;
  padding: 4px 10px;
  border-radius: 999px;
  background: rgba(34, 197, 94, 0.12);
  color: rgb(22, 101, 52);
  font-weight: 900;
  font-size: 12px;
}

.eq-pill {
  display: inline-block;
  padding: 4px 10px;
  border-radius: 999px;
  background: rgba(15, 23, 42, 0.06);
  color: rgba(15, 23, 42, 0.80);
  font-weight: 800;
  font-size: 12px;
  margin-right: 6px;
}

.eq-kpi {
  background: rgba(15, 23, 42, 0.03);
  border: 1px solid rgba(15, 23, 42, 0.06);
  border-radius: 14px;
  padding: 10px 12px;
}

div[data-testid="stDataFrame"] {
  border-radius: 14px;
  overflow: hidden;
  border: 1px solid rgba(15, 23, 42, 0.08);
}
</style>
""",
    unsafe_allow_html=True,
)

# ----------------------------
# CUSTOMER ACCESS + TRACKING
# Verify ONCE, then use session_state forever (prevents "invalid link" on reruns)
# ----------------------------
if "customer_id" not in st.session_state:
    customer_id = require_customer_access()
    st.session_state["customer_id"] = customer_id
    log_event(customer_id, "page_view")
else:
    customer_id = st.session_state["customer_id"]


# ----------------------------
# Files / constants
# ----------------------------
RATES_FILE = APP_DIR / "RETOOL ALL COST UPLOAD 2026 WITH FUEL TYPE.xlsx"
TEMPLATE_FILE = APP_DIR / "Shipment Template.xlsx"

TRANSIT_BASE = APP_DIR / "TRANSIT_BASE.xlsx"
TRANSIT_OVERRIDES = APP_DIR / "TRANSIT_CITY_OVERRIDES.xlsx"
CITYCLASS_MASTER = APP_DIR / "CITYCLASS_MASTER.xlsx"

DEFAULT_DIVISOR = 5000
DEFAULT_MAX_PARCELS = 10


# ---------------------------
# Helpers
# ---------------------------
def safe_show_logo():
    if not LOGO_FILE.exists():
        return
    try:
        st.image(LOGO_FILE.read_bytes(), use_container_width=True)
    except Exception:
        pass


@st.cache_data(show_spinner=False)
def get_template_headers(template_path_str: str) -> list[str]:
    p = Path(template_path_str)
    if not p.exists():
        return []
    bio = io.BytesIO(p.read_bytes())
    wb = openpyxl.load_workbook(bio)
    ws = wb.active
    headers = [ws.cell(1, c).value for c in range(1, ws.max_column + 1)]
    headers = [str(h).strip() if h else "" for h in headers]
    return [h for h in headers if h]


def build_single_shipment_df(
    template_headers: list[str],
    *,
    from_country: str,
    to_country: str,
    to_city: str,
    currency: str,
    incoterm: str,
    hs_code: str,
    declared_value: float,
    parcels_df: pd.DataFrame,
    max_parcels: int,
) -> pd.DataFrame:
    row = {h: "" for h in template_headers}

    def set_if_present(key: str, value):
        if key in row:
            row[key] = value

    set_if_present("From Country", from_country)
    set_if_present("To Country", to_country)
    set_if_present("To City", to_city)
    set_if_present("Currency", currency)
    set_if_present("Incoterm", incoterm)

    set_if_present("HS", hs_code)
    set_if_present("HS Code", hs_code)
    set_if_present("Item Value", declared_value)
    set_if_present("Declared Value", declared_value)

    parcels_df = parcels_df.fillna(0)
    expanded = []
    for _, r in parcels_df.iterrows():
        qty = int(r.get("Qty", 1) or 1)
        for _ in range(qty):
            expanded.append(
                {
                    "WeightKg": float(r.get("WeightKg", 0) or 0),
                    "Lcm": float(r.get("Lcm", 0) or 0),
                    "Wcm": float(r.get("Wcm", 0) or 0),
                    "Hcm": float(r.get("Hcm", 0) or 0),
                }
            )
    expanded = expanded[:max_parcels]

    for i, p in enumerate(expanded, start=1):
        weight_keys = [f"Parcel {i} Weight", f"Parcel{i} Weight", f"Parcel {i} Weight (kg)", f"Parcel {i} WeightKg"]
        length_keys = [f"Parcel {i} Length", f"Parcel{i} Length", f"Parcel {i} L", f"Parcel {i} Length (cm)"]
        width_keys = [f"Parcel {i} Width", f"Parcel{i} Width", f"Parcel {i} W", f"Parcel {i} Width (cm)"]
        height_keys = [f"Parcel {i} Height", f"Parcel{i} Height", f"Parcel {i} H", f"Parcel {i} Height (cm)"]

        for k in weight_keys:
            if k in row:
                row[k] = p["WeightKg"]
                break
        for k in length_keys:
            if k in row:
                row[k] = p["Lcm"]
                break
        for k in width_keys:
            if k in row:
                row[k] = p["Wcm"]
                break
        for k in height_keys:
            if k in row:
                row[k] = p["Hcm"]
                break

    set_if_present("Parcels", len(expanded))
    set_if_present("Parcel Count", len(expanded))

    return pd.DataFrame([row])


def run_engine(df: pd.DataFrame, tmpdir: Path, carriers: list[str], types: list[str], include_customs_flag: bool):
    ship_path = tmpdir / "shipments.xlsx"
    out_path = tmpdir / "output.xlsx"

    df.to_excel(ship_path, index=False)

    old_argv = sys.argv[:]
    try:
        sys.argv = [
            "run_pricing.py",
            "--rates",
            str(RATES_FILE),
            "--shipments",
            str(ship_path),
            "--out",
            str(out_path),
            "--divisor",
            str(int(DEFAULT_DIVISOR)),
            "--max_parcels",
            str(int(DEFAULT_MAX_PARCELS)),
            "--carriers",
            ",".join(carriers),
            "--types",
            ",".join(types),
            "--include_customs",
            "1" if include_customs_flag else "0",
        ]
        run_pricing.main()
    finally:
        sys.argv = old_argv

    results_df = pd.read_excel(out_path, sheet_name="Prices_All_Services")
    return results_df, out_path


def pick_best_row(results: pd.DataFrame) -> pd.Series | None:
    if not isinstance(results, pd.DataFrame) or results.empty:
        return None
    if "Total" in results.columns:
        tmp = results.copy()
        tmp["_TotalNum"] = pd.to_numeric(tmp["Total"], errors="coerce")
        tmp = tmp.dropna(subset=["_TotalNum"]).sort_values("_TotalNum", ascending=True)
        if not tmp.empty:
            return tmp.iloc[0]
    return results.iloc[0]


def format_results_for_display(results: pd.DataFrame) -> pd.DataFrame:
    if not isinstance(results, pd.DataFrame) or results.empty:
        return results
    df = results.copy()

    cols = list(df.columns)
    front = []
    for c in ["Carrier", "Service", "Service Type", "Product", "Total", "TransitDays", "Transit Days"]:
        if c in cols:
            front.append(c)
    rest = [c for c in cols if c not in front]
    df = df[front + rest]

    if "Total" in df.columns:
        try:
            df["_TotalNum"] = pd.to_numeric(df["Total"], errors="coerce")
            df = df.sort_values(["_TotalNum"], ascending=True).drop(columns=["_TotalNum"])
        except Exception:
            pass
    return df


# ---------------------------
# Required file checks
# ---------------------------
for f in [RATES_FILE, TRANSIT_BASE, TRANSIT_OVERRIDES, CITYCLASS_MASTER, TEMPLATE_FILE]:
    if not f.exists():
        st.error(f"Missing required file: {f.name}")
        st.stop()

template_headers = get_template_headers(str(TEMPLATE_FILE))
if not template_headers:
    st.error("Shipment Template.xlsx is missing or unreadable (needed to map inputs).")
    st.stop()


# ---------------------------
# Header
# ---------------------------
header_cols = st.columns([1.1, 6.0, 1.2])
with header_cols[0]:
    safe_show_logo()
with header_cols[1]:
    st.markdown("## Shipping Quote")
    st.markdown(
        '<div class="eq-muted">Enter your shipment details to compare services. You’ll see the lowest price clearly and all alternatives below.</div>',
        unsafe_allow_html=True,
    )
with header_cols[2]:
    st.markdown('<span class="eq-pill">Secure link</span>', unsafe_allow_html=True)

st.write("")


# ---------------------------
# UI: two columns
# ---------------------------
left, right = st.columns([1.05, 0.95], gap="large")

carriers_default = ["UPS", "DHL", "FEDEX"]
types_default = ["EXPRESS", "ECONOMY"]

with left:
    st.markdown('<div class="eq-card">', unsafe_allow_html=True)
    st.markdown('<div class="eq-title">Shipment</div>', unsafe_allow_html=True)

    with st.form("single_shipment_form", border=False):
        c1, c2, c3 = st.columns([1, 1, 1])
        with c1:
            from_country = st.text_input("Origin country", value="GB", help="2-letter country code, e.g. GB, US, DE")
        with c2:
            to_country = st.text_input("Destination country", value="", help="2-letter country code, e.g. US")
        with c3:
            to_city = st.text_input("Destination city", value="", placeholder="e.g. New York")

        c4, c5 = st.columns(2)
        with c4:
            currency = st.selectbox("Currency", ["GBP", "EUR", "USD"], index=0)
        with c5:
            incoterm = st.selectbox("Incoterm", ["DAP", "DDP"], index=0)

        st.divider()
        st.markdown('<div class="eq-title">Parcels</div>', unsafe_allow_html=True)

        parcel_lines = st.slider("Parcel lines", 1, 6, 1)
        base = pd.DataFrame([{"Qty": 1, "WeightKg": 1.0, "Lcm": 10.0, "Wcm": 10.0, "Hcm": 10.0}] * parcel_lines)

        parcels = st.data_editor(
            base,
            use_container_width=True,
            hide_index=True,
            column_config={
                "Qty": st.column_config.NumberColumn(min_value=1, step=1),
                "WeightKg": st.column_config.NumberColumn(min_value=0.01, step=0.1, format="%.2f"),
                "Lcm": st.column_config.NumberColumn(min_value=1.0, step=1.0, format="%.1f"),
                "Wcm": st.column_config.NumberColumn(min_value=1.0, step=1.0, format="%.1f"),
                "Hcm": st.column_config.NumberColumn(min_value=1.0, step=1.0, format="%.1f"),
            },
        )

        expanded_pieces = int((parcels["Qty"].fillna(0)).sum())
        total_weight = float((parcels["Qty"] * parcels["WeightKg"]).fillna(0).sum())
        total_cm3 = float((parcels["Qty"] * parcels["Lcm"] * parcels["Wcm"] * parcels["Hcm"]).fillna(0).sum())
        volumetric_kg = (total_cm3 / float(DEFAULT_DIVISOR)) if DEFAULT_DIVISOR else 0.0

        st.markdown(
            f"""
<div class="eq-kpi">
<span class="eq-pill">Pieces: {expanded_pieces}</span>
<span class="eq-pill">Actual: {total_weight:.2f} kg</span>
<span class="eq-pill">Volumetric: {volumetric_kg:.2f} kg</span>
</div>
""",
            unsafe_allow_html=True,
        )

        st.write("")

        with st.expander("Customs (optional)"):
            include_customs = st.checkbox("Include duties & taxes", value=False)
            hs_code = st.text_input("HS code", value="49111090")
            declared_value = st.number_input("Declared value", 0.0, 100000.0, 10.0, step=10.0)

        with st.expander("More filters"):
            carriers_selected = st.multiselect("Carriers", ["UPS", "DHL", "FEDEX"], default=carriers_default)
            types_selected = st.multiselect(
                "Service types",
                ["EXPRESS", "ECONOMY", "DOMESTIC", "FREIGHT"],
                default=types_default,
            )

        run_btn = st.form_submit_button("Get prices", use_container_width=True)

    st.markdown("</div>", unsafe_allow_html=True)

with right:
    st.markdown('<div class="eq-card">', unsafe_allow_html=True)
    st.markdown('<div class="eq-title">Results</div>', unsafe_allow_html=True)

    if not st.session_state.get("latest_results_ready"):
        st.info("Complete the form and click **Get prices** to see the cheapest option.")
        st.markdown("</div>", unsafe_allow_html=True)
    else:
        results = st.session_state["latest_results"]
        results_disp = format_results_for_display(results)
        best = pick_best_row(results_disp)

        if best is not None:
            st.markdown('<span class="eq-best">Best option</span>', unsafe_allow_html=True)
            carrier_val = best.get("Carrier", "—")
            service_val = best.get("Service", best.get("Service Type", "—"))
            total_val = best.get("Total", "—")
            transit_val = best.get("TransitDays", best.get("Transit Days", best.get("Transit", "—")))

            k1, k2 = st.columns([1.2, 1.0])
            with k1:
                st.markdown(f"### {total_val}")
                st.markdown(f"<div class='eq-muted'>{carrier_val} · {service_val}</div>", unsafe_allow_html=True)
            with k2:
                st.markdown("")
                st.markdown("")
                st.markdown(f"<div class='eq-muted'>Transit: <b>{transit_val}</b></div>", unsafe_allow_html=True)

            st.divider()

        show_details = st.toggle("Show full breakdown columns", value=False)
        if show_details:
            table_df = results_disp
        else:
            keep = []
            for c in ["Carrier", "Service", "Service Type", "Total", "TransitDays", "Transit Days", "FuelPctSource", "Fuel %", "FuelPct"]:
                if c in results_disp.columns:
                    keep.append(c)
            keep = keep or list(results_disp.columns[:8])
            table_df = results_disp[keep]

        st.dataframe(table_df, use_container_width=True, height=520)

        out_path_str = st.session_state.get("latest_out_path")
        if out_path_str:
            out_path = Path(out_path_str)
            if out_path.exists():
                st.download_button(
                    "Download full output (Excel)",
                    data=out_path.read_bytes(),
                    file_name="pricing_output.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                )

        st.markdown("</div>", unsafe_allow_html=True)


# ---------------------------
# Run handler
# ---------------------------
if "run_btn" in locals() and run_btn:
    log_event(customer_id, "price_run")

    if not to_country.strip():
        st.error("Please enter a destination country (2-letter code, e.g. US).")
        st.stop()
    if not to_city.strip():
        st.error("Please enter a destination city.")
        st.stop()

    expanded_pieces = int((parcels["Qty"].fillna(0)).sum())
    total_weight = float((parcels["Qty"] * parcels["WeightKg"]).fillna(0).sum())
    if expanded_pieces <= 0 or total_weight <= 0:
        st.error("Please enter valid parcel quantities and weights.")
        st.stop()

    if "include_customs" not in locals():
        include_customs = False
    if "hs_code" not in locals():
        hs_code = "49111090"
    if "declared_value" not in locals():
        declared_value = 0.0
    if "carriers_selected" not in locals():
        carriers_selected = carriers_default
    if "types_selected" not in locals():
        types_selected = types_default

    shipment_df = build_single_shipment_df(
        template_headers,
        from_country=from_country.strip(),
        to_country=to_country.strip(),
        to_city=to_city.strip(),
        currency=currency,
        incoterm=incoterm,
        hs_code=hs_code.strip(),
        declared_value=float(declared_value),
        parcels_df=parcels,
        max_parcels=DEFAULT_MAX_PARCELS,
    )

    with st.spinner("Calculating…"):
        with tempfile.TemporaryDirectory() as td:
            td = Path(td)
            results_df, out_path = run_engine(
                shipment_df,
                td,
                carriers=carriers_selected,
                types=types_selected,
                include_customs_flag=include_customs,
            )

            st.session_state["latest_results"] = results_df
            st.session_state["latest_results_ready"] = True
            st.session_state["latest_out_path"] = str(out_path)

            st.rerun()
