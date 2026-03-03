import sys
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

# Hidden operational defaults (not shown in customer UI)
DEFAULT_DIVISOR = 5000
DEFAULT_MAX_PARCELS = 10

st.set_page_config(page_title="Equator Calculator", layout="wide", page_icon="📦")

# ---------- Minimal CSS to make it look like a product ----------
st.markdown(
    """
<style>
.block-container { padding-top: 1.6rem; padding-bottom: 2.5rem; max-width: 1200px; }
h1, h2, h3 { letter-spacing: -0.02em; }

div[data-baseweb="input"] input,
div[data-baseweb="select"] > div {
  border-radius: 12px !important;
}

.eq-card {
  background: white;
  border: 1px solid rgba(15, 23, 42, 0.08);
  border-radius: 16px;
  padding: 18px 18px 12px 18px;
  box-shadow: 0 1px 2px rgba(15, 23, 42, 0.04);
}

.eq-title {
  font-size: 13px;
  font-weight: 800;
  letter-spacing: 0.04em;
  color: rgba(15, 23, 42, 0.75);
  text-transform: uppercase;
  margin-bottom: 10px;
}

.eq-best {
  display: inline-block;
  padding: 4px 10px;
  border-radius: 999px;
  background: rgba(34, 197, 94, 0.12);
  color: rgb(22, 101, 52);
  font-weight: 800;
  font-size: 12px;
}

.eq-muted { color: rgba(15,23,42,0.65); font-size: 0.9rem; }
</style>
""",
    unsafe_allow_html=True,
)


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
    include_customs: bool,
    hs_code: str,
    declared_value: float,
    parcels_df: pd.DataFrame,
    max_parcels: int,
) -> pd.DataFrame:
    """
    Build a 1-row dataframe matching the Shipment Template headers as closely as possible.
    We populate common fields + attempt to fill Parcel 1..N dimension columns using common naming patterns.
    """

    # Start with all template columns blank so we never miss required columns
    row = {h: "" for h in template_headers}

    def set_if_present(key: str, value):
        if key in row:
            row[key] = value

    # Basic fields (adjust these if your template uses different column names)
    set_if_present("From Country", from_country)
    set_if_present("To Country", to_country)
    set_if_present("To City", to_city)
    set_if_present("Currency", currency)
    set_if_present("Incoterm", incoterm)

    # Customs (simple)
    # We set values regardless; engine can ignore if include_customs is false
    set_if_present("HS", hs_code)
    set_if_present("HS Code", hs_code)
    set_if_present("Item Value", declared_value)
    set_if_present("Declared Value", declared_value)

    # Expand parcel lines by Qty into a flat list of actual parcels
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

    # Fill Parcel 1..N using common header naming patterns
    for i, p in enumerate(expanded, start=1):
        weight_keys = [
            f"Parcel {i} Weight",
            f"Parcel{i} Weight",
            f"Parcel {i} Weight (kg)",
            f"Parcel {i} WeightKg",
        ]
        length_keys = [
            f"Parcel {i} Length",
            f"Parcel{i} Length",
            f"Parcel {i} L",
            f"Parcel {i} Length (cm)",
        ]
        width_keys = [
            f"Parcel {i} Width",
            f"Parcel{i} Width",
            f"Parcel {i} W",
            f"Parcel {i} Width (cm)",
        ]
        height_keys = [
            f"Parcel {i} Height",
            f"Parcel{i} Height",
            f"Parcel {i} H",
            f"Parcel {i} Height (cm)",
        ]

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


def run_engine(
    df: pd.DataFrame,
    tmpdir: Path,
    carriers: list[str],
    types: list[str],
    include_customs_flag: bool,
):
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


# ---------------------------
# Required file checks
# ---------------------------
for f in [RATES_FILE, TRANSIT_BASE, TRANSIT_OVERRIDES, CITYCLASS_MASTER, TEMPLATE_FILE]:
    if not f.exists():
        st.error(f"Missing required file: {f.name}")
        st.stop()


# ---------------------------
# Header (clean)
# ---------------------------
header_cols = st.columns([1.1, 6.0, 1.2])
with header_cols[0]:
    safe_show_logo()
with header_cols[1]:
    st.markdown("## Pricing Calculator")
    st.markdown('<div class="eq-muted">Instant quote for a single shipment</div>', unsafe_allow_html=True)
st.write("")


# ---------------------------
# Template headers for mapping
# ---------------------------
template_headers = get_template_headers(str(TEMPLATE_FILE))
if not template_headers:
    st.error("Shipment Template.xlsx is missing or unreadable (needed to map inputs).")
    st.stop()


# ---------------------------
# UI (2-column: form + results)
# ---------------------------
left, right = st.columns([1.05, 0.95], gap="large")

# Defaults
carriers_default = ["UPS", "DHL", "FEDEX"]
types_default = ["EXPRESS", "ECONOMY"]

with left:
    st.markdown('<div class="eq-card">', unsafe_allow_html=True)
    st.markdown('<div class="eq-title">Shipment details</div>', unsafe_allow_html=True)

    with st.form("single_shipment_form", border=False):
        c1, c2 = st.columns(2)

        with c1:
            from_country = st.text_input("Origin country", value="GB")
            to_country = st.text_input("Destination country", value="")
            to_city = st.text_input("Destination city", value="", placeholder="e.g. New York")

        with c2:
            currency = st.selectbox("Currency", ["GBP", "EUR", "USD"], index=0)
            incoterm = st.selectbox("Incoterm", ["DAP", "DDP"], index=0)

        st.divider()
        st.markdown('<div class="eq-title">Parcels</div>', unsafe_allow_html=True)

        parcel_lines = st.slider("Parcel lines", 1, 6, 1)
        base = pd.DataFrame(
            [{"Qty": 1, "WeightKg": 1.0, "Lcm": 10.0, "Wcm": 10.0, "Hcm": 10.0}] * parcel_lines
        )

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
        st.caption(f"Total pieces: **{expanded_pieces}** · Total weight: **{total_weight:.2f} kg**")

        with st.expander("Customs (optional)"):
            include_customs = st.checkbox("Include duties & taxes", value=False)
            hs_code = st.text_input("HS code", value="49111090")
            declared_value = st.number_input("Declared value", 0.0, 100000.0, 10.0, step=10.0)

        with st.expander("More filters"):
            carriers_selected = st.multiselect(
                "Carriers",
                ["UPS", "DHL", "FEDEX"],
                default=carriers_default,
            )
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
        st.info("Fill in the shipment details and click **Get prices**.")
    else:
        results = st.session_state["latest_results"]

        best = None
        if isinstance(results, pd.DataFrame) and len(results):
            if "Total" in results.columns:
                best = results.sort_values("Total", ascending=True).iloc[0]

        if best is not None:
            st.markdown('<span class="eq-best">Best option</span>', unsafe_allow_html=True)
            m1, m2, m3 = st.columns(3)
            m1.metric("Total", f"{best.get('Total', '')}")
            m2.metric("Carrier", f"{best.get('Carrier', '')}")
            m3.metric("Service", f"{best.get('Service', '')}")
            st.divider()

        st.dataframe(results, use_container_width=True, height=520)

        out_path_str = st.session_state.get("latest_out_path")
        if out_path_str:
            out_path = Path(out_path_str)
            if out_path.exists():
                st.download_button(
                    "Download pricing_output.xlsx",
                    data=out_path.read_bytes(),
                    file_name="pricing_output.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                )

    st.markdown("</div>", unsafe_allow_html=True)


# ---------------------------
# Run handler
# ---------------------------
if run_btn:
    log_event(customer_id, "price_run")

    # Validation
    if not to_country.strip():
        st.error("Please enter a destination country.")
        st.stop()
    if not to_city.strip():
        st.error("Please enter a destination city.")
        st.stop()

    expanded_pieces = int((parcels["Qty"].fillna(0)).sum())
    total_weight = float((parcels["Qty"] * parcels["WeightKg"]).fillna(0).sum())

    if expanded_pieces <= 0 or total_weight <= 0:
        st.error("Please enter valid parcel quantities and weights.")
        st.stop()

    # Ensure customs vars exist even if expander not used
    if "hs_code" not in locals():
        hs_code = "49111090"
    if "declared_value" not in locals():
        declared_value = 0.0
    if "include_customs" not in locals():
        include_customs = False
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
        include_customs=include_customs,
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
