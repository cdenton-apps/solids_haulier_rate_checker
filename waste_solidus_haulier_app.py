# app.py
import os
import math
import json
from datetime import date, datetime
from typing import Optional, List, Dict, Tuple

import pandas as pd
import streamlit as st
from PIL import Image

st.set_page_config(page_title="Solidus Haulier Rate Checker", layout="wide")

st.markdown(
    """
    <style>
      #MainMenu { visibility: hidden; }
      footer { visibility: hidden; }
    </style>
    """,
    unsafe_allow_html=True
)

col_logo, col_text = st.columns([1, 3], gap="medium")

with col_logo:
    logo_path = "assets/solidus_logo.png"
    try:
        st.image(Image.open(logo_path), width=150)
    except Exception:
        st.warning(f"Could not load logo at '{logo_path}'.")

with col_text:
    st.markdown(
        "<h1 style='color:#0D4B6A; margin-bottom:0.2em;'>Solidus Haulier Rate Checker</h1>",
        unsafe_allow_html=True
    )
    st.markdown(
        """
        V3.0.1    
        Enter a UK postcode area, select a service type, set pallets and surcharges,
        and optionally add AM/PM, Tail Lift or Timed Delivery. Dual Collection splits the load.

    **Version 3 BETA**    
    **NEW:** PO Building in-app.
        """,
        unsafe_allow_html=True
    )

# -------------------------
# Config / constants
# -------------------------
DATA_FILE = "joda_surcharge.json"
RATE_XLSX = "haulier prices 2.xlsx"

WAREHOUSE_NAME_FIXED = "101 - Skipton"

# Supplier account codes
JODA_ACC = "J040"
MCD_ACC = "M127"

# PO grouping
JODA_PO_GROUP = 1
MCD_PO_GROUP = 2

# Optional template file candidates (repo / deployed runtime)
EXPORT_TEMPLATE_CANDIDATES = [
    "PO Import Example File.csv",
    "po_import_example_file.csv",
    "PO_Import_Example_File.csv",
]

# Fallback columns (minimum safe set to keep the app running)
# If you upload the template, we'll use its exact columns instead.
FALLBACK_EXPORT_COLUMNS = [
    "Purchase Order Import Type",
    "Purchase Order Number",
    "Purchase Order Supplier Acc Code",
    "Purchase Order Document Date",
    "Purchase Order Header Requested Date",
    "Purchase Order Discount Percent",
    "Warehouse Name",
    "Free Text Item Description",
    "Item Quantity",
    "Unit Buying Price",
    "Unit Discount Percent",
    "Purchase Order Line Requested Date",
]

# -------------------------
# Sidebar: optional template upload
# -------------------------
st.sidebar.header("Export Template")
uploaded_template = st.sidebar.file_uploader(
    "Upload PO Import Example File.csv (optional but recommended)",
    type=["csv"],
    help="If provided, export will follow the exact column structure/order from your Sage import template."
)

# -------------------------
# Joda stored surcharge
# -------------------------
def load_joda_surcharge() -> float:
    today_str = date.today().isoformat()
    if not os.path.exists(DATA_FILE):
        with open(DATA_FILE, "w") as f:
            json.dump({"surcharge": 0.0, "last_updated": today_str}, f)
        return 0.0

    try:
        with open(DATA_FILE, "r") as f:
            data = json.load(f)
    except Exception:
        data = {"surcharge": 0.0, "last_updated": today_str}

    # reset on Wednesdays
    if date.today().weekday() == 2 and data.get("last_updated") != today_str:
        data = {"surcharge": 0.0, "last_updated": today_str}
        with open(DATA_FILE, "w") as f:
            json.dump(data, f)
        return 0.0

    try:
        return float(data.get("surcharge", 0.0))
    except Exception:
        return 0.0

def save_joda_surcharge(new_pct: float):
    today_str = date.today().isoformat()
    with open(DATA_FILE, "w") as f:
        json.dump({"surcharge": float(new_pct), "last_updated": today_str}, f)

joda_stored_pct = load_joda_surcharge()

# -------------------------
# Rates load
# -------------------------
@st.cache_data
def load_rate_table(excel_path: str, _mtime: float) -> pd.DataFrame:
    raw = pd.read_excel(excel_path, header=1)
    raw = raw.rename(columns={
        raw.columns[0]: "PostcodeArea",
        raw.columns[1]: "Service",
        raw.columns[2]: "Vendor"
    })

    raw["PostcodeArea"] = raw["PostcodeArea"].ffill()
    raw["Service"]      = raw["Service"].ffill()
    raw["Vendor"]       = raw["Vendor"].ffill()
    raw = raw[raw["Vendor"] != "Vendor"].copy()

    pallet_cols = [
        c for c in raw.columns
        if isinstance(c, (int, float)) or (isinstance(c, str) and str(c).isdigit())
    ]

    melted = raw.melt(
        id_vars=["PostcodeArea", "Service", "Vendor"],
        value_vars=pallet_cols,
        var_name="Pallets",
        value_name="BaseRate"
    )
    melted["Pallets"] = melted["Pallets"].astype(int)
    melted["BaseRate"] = pd.to_numeric(melted["BaseRate"], errors="coerce")
    melted = melted.dropna(subset=["BaseRate"]).copy()

    melted["PostcodeArea"] = melted["PostcodeArea"].astype(str).str.strip().str.upper()
    melted["Service"]      = melted["Service"].astype(str).str.strip().str.title()
    melted["Vendor"]       = melted["Vendor"].astype(str).str.strip().str.title()

    return melted.reset_index(drop=True)

mtime = os.path.getmtime(RATE_XLSX)
rate_df = load_rate_table(RATE_XLSX, mtime)
unique_areas = sorted(rate_df["PostcodeArea"].unique())

# -------------------------
# Export template columns
# -------------------------
def _find_export_template_path() -> Optional[str]:
    for p in EXPORT_TEMPLATE_CANDIDATES:
        if os.path.exists(p):
            return p
    return None

@st.cache_data
def load_export_template_columns_from_path(path: str) -> List[str]:
    df = pd.read_csv(path)
    return df.columns.tolist()

@st.cache_data
def load_export_template_columns_from_upload(file_bytes: bytes) -> List[str]:
    from io import BytesIO
    df = pd.read_csv(BytesIO(file_bytes))
    return df.columns.tolist()

EXPORT_COLUMNS: List[str]
if uploaded_template is not None:
    try:
        EXPORT_COLUMNS = load_export_template_columns_from_upload(uploaded_template.getvalue())
        st.sidebar.success("Using uploaded template columns.")
    except Exception:
        EXPORT_COLUMNS = FALLBACK_EXPORT_COLUMNS
        st.sidebar.warning("Could not read uploaded template; using fallback columns.")
else:
    template_path = _find_export_template_path()
    if template_path:
        try:
            EXPORT_COLUMNS = load_export_template_columns_from_path(template_path)
            st.sidebar.info(f"Using template found in repo: {template_path}")
        except Exception:
            EXPORT_COLUMNS = FALLBACK_EXPORT_COLUMNS
            st.sidebar.warning("Template found but could not be read; using fallback columns.")
    else:
        EXPORT_COLUMNS = FALLBACK_EXPORT_COLUMNS
        st.sidebar.warning("No template found. Upload it to match the exact Sage import format.")

def _blank_export_row() -> Dict[str, object]:
    return {c: "" for c in EXPORT_COLUMNS}

def _ddmmyyyy(d: date) -> str:
    return d.strftime("%d/%m/%Y")

def _export_line(
    po_group: int,
    supplier_acc: str,
    so_number: str,
    area_code: str,
    service: str,
    label: str,
    qty: float,
    unit_price: float,
    doc_date: Optional[date] = None,
    req_date: Optional[date] = None,
) -> Dict[str, object]:
    doc_date = doc_date or date.today()
    req_date = req_date or date.today()

    r = _blank_export_row()

    # Only set fields that exist in the template (works with fallback too)
    def set_if(col: str, val):
        if col in r:
            r[col] = val

    set_if("Purchase Order Import Type", 1)
    set_if("Purchase Order Number", int(po_group))
    set_if("Purchase Order Supplier Acc Code", str(supplier_acc).strip())
    set_if("Purchase Order Document Date", _ddmmyyyy(doc_date))
    set_if("Purchase Order Header Requested Date", _ddmmyyyy(req_date))
    set_if("Purchase Order Discount Percent", 0)

    set_if("Warehouse Name", WAREHOUSE_NAME_FIXED)
    set_if("Unit Discount Percent", 0)
    set_if("Purchase Order Line Requested Date", _ddmmyyyy(req_date))

    so_number = str(so_number).strip()
    so_suffix = f" - SO{so_number}" if so_number else ""
    svc_suffix = f" ({service})" if str(service).strip() else ""
    set_if("Free Text Item Description", f"{area_code} {label}{svc_suffix}{so_suffix}".strip())

    set_if("Item Quantity", float(qty))
    set_if("Unit Buying Price", float(unit_price))

    return r

# -------------------------
# Session defaults
# -------------------------
def _ensure_defaults():
    st.session_state.setdefault("area", "")
    st.session_state.setdefault("service", "Economy")
    st.session_state.setdefault("pallets", 1)
    st.session_state.setdefault("joda_pct", round(joda_stored_pct, 2))
    st.session_state.setdefault("mcd_pct", 0.0)
    st.session_state.setdefault("ampm", False)
    st.session_state.setdefault("tail", False)
    st.session_state.setdefault("dual", False)
    st.session_state.setdefault("timed", False)
    st.session_state.setdefault("split1", 1)
    st.session_state.setdefault("split2", 1)

    st.session_state.setdefault("export_basket", [])
    st.session_state.setdefault("so_number", "")

    # multi-select state
    st.session_state.setdefault("export_selected_keys", [])

_ensure_defaults()

# -------------------------
# Apply pending load (History -> inputs)
# -------------------------
def _apply_pending_load():
    payload = st.session_state.pop("__pending_load", None)
    if not payload:
        return

    saved_area = payload.get("Area", "")
    if isinstance(saved_area, list):
        saved_area = saved_area[0] if saved_area else ""
    saved_area = str(saved_area).upper().strip()

    saved_service = payload.get("Service", "Economy")
    if isinstance(saved_service, list):
        saved_service = saved_service[0] if saved_service else "Economy"
    saved_service = str(saved_service).strip()

    if saved_area in unique_areas:
        st.session_state.area = saved_area
    if saved_service in ["Economy", "Next Day"]:
        st.session_state.service = saved_service

    if payload.get("Dual"):
        st.session_state.dual = True
        st.session_state.split1 = int(payload.get("Split1", 1) or 1)
        st.session_state.split2 = int(payload.get("Split2", 1) or 1)
        st.session_state.pallets = int(st.session_state.split1 + st.session_state.split2)
    else:
        st.session_state.dual = False
        st.session_state.split1 = 1
        st.session_state.split2 = 1
        try:
            st.session_state.pallets = int(str(payload.get("Pallets", "1")).split("+")[0])
        except Exception:
            st.session_state.pallets = 1

    st.session_state.ampm  = bool(payload.get("AM/PM", False))
    st.session_state.timed = bool(payload.get("Timed", False))
    st.session_state.tail  = bool(payload.get("Tail", False))

    try:
        st.session_state.joda_pct = float(payload.get("JodaPct", st.session_state.joda_pct))
    except Exception:
        pass
    try:
        st.session_state.mcd_pct = float(payload.get("McdPct", st.session_state.mcd_pct))
    except Exception:
        pass

_apply_pending_load()

# -------------------------
# Helpers
# -------------------------
def get_base_rate(df, area, service, vendor, pallets):
    subset = df[
        (df["PostcodeArea"] == area) &
        (df["Service"] == service) &
        (df["Vendor"] == vendor) &
        (df["Pallets"] == pallets)
    ]
    return None if subset.empty else float(subset["BaseRate"].iloc[0])

def joda_effective_pct(pallet_count: int, input_pct: float) -> float:
    return 0.0 if pallet_count < 7 else float(input_pct)

def mcd_smallload_extra(pallet_count: int) -> float:
    return (5.0 * min(pallet_count, 4)) if pallet_count < 5 else 0.0

def _add_to_basket(rows: List[Dict[str, object]]):
    st.session_state.export_basket.extend(rows)

# -------------------------
# UI: Inputs
# -------------------------
st.header("1. Input Parameters")
col_a, col_b, col_c, col_d, col_e, col_f = st.columns([1, 1, 1, 1, 1, 1], gap="medium")

with col_a:
    st.selectbox(
        "Postcode Area",
        options=[""] + unique_areas,
        index=([""] + unique_areas).index(st.session_state.area) if st.session_state.area in ([""] + unique_areas) else 0,
        key="area",
        format_func=lambda x: x if x else "— Select area —"
    )
    if st.session_state.area == "":
        st.info("Please select a postcode area to continue.")
        st.stop()

with col_b:
    st.selectbox("Service Type", options=["Economy", "Next Day"], key="service")

with col_c:
    st.number_input("Number of Pallets", min_value=1, max_value=26, step=1, key="pallets")

with col_d:
    st.number_input(
        "Joda Fuel Surcharge (%)",
        min_value=0.0, max_value=100.0,
        step=0.1, format="%.2f",
        key="joda_pct"
    )
    if st.button("Save Joda Surcharge"):
        save_joda_surcharge(float(st.session_state.joda_pct))
        st.success(f"Saved Joda surcharge at {float(st.session_state.joda_pct):.2f}%")

with col_e:
    st.number_input(
        "McDowells Fuel Surcharge (%)",
        min_value=0.0, max_value=100.0,
        step=0.1, format="%.2f",
        key="mcd_pct"
    )

with col_f:
    st.markdown(" ")

st.markdown("---")
postcode_area = st.session_state.area

st.subheader("2. Optional Extras")
col1, col2, col3, col4 = st.columns(4, gap="large")

with col1:
    st.checkbox("AM/PM Delivery", key="ampm")
with col2:
    st.checkbox("Tail Lift", key="tail")
with col3:
    st.checkbox("Dual Collection", key="dual")
with col4:
    st.checkbox("Timed Delivery", key="timed")

if st.session_state.dual and st.session_state.pallets == 1:
    st.error("Dual Collection requires at least 2 pallets.")
    st.stop()

split1 = split2 = None
if st.session_state.dual:
    st.markdown("**Split pallets into two despatches (e.g., ESL & U4).**")
    sp1, sp2 = st.columns(2, gap="large")
    with sp1:
        split1 = st.number_input("First Pallet Group", 1, st.session_state.pallets - 1, key="split1")
    with sp2:
        split2 = st.number_input("Second Pallet Group", 1, st.session_state.pallets - 1, key="split2")
    if st.session_state.split1 + st.session_state.split2 != st.session_state.pallets:
        st.error("Pallet Split values must add up to total pallets.")
        st.stop()

st.markdown("---")
st.subheader("3. Add to Export List")
st.text_input("SO Number (manual)", key="so_number", placeholder="e.g. 020502")
st.caption(f"Warehouse is fixed to: **{WAREHOUSE_NAME_FIXED}**")

# -------------------------
# Calculations
# -------------------------
# Joda
joda_base = None
joda_final = None
joda_charge_fixed = (7.5 if st.session_state.ampm else 0) + (20 if st.session_state.timed else 0)

if st.session_state.dual:
    b1 = get_base_rate(rate_df, postcode_area, st.session_state.service, "Joda", st.session_state.split1)
    b2 = get_base_rate(rate_df, postcode_area, st.session_state.service, "Joda", st.session_state.split2)
    if b1 is not None and b2 is not None:
        eff1 = joda_effective_pct(st.session_state.split1, float(st.session_state.joda_pct))
        eff2 = joda_effective_pct(st.session_state.split2, float(st.session_state.joda_pct))
        g1 = b1 * (1 + eff1 / 100.0) + joda_charge_fixed
        g2 = b2 * (1 + eff2 / 100.0) + joda_charge_fixed
        joda_base = b1 + b2
        joda_final = g1 + g2
else:
    base = get_base_rate(rate_df, postcode_area, st.session_state.service, "Joda", st.session_state.pallets)
    if base is not None:
        joda_base = base
        eff = joda_effective_pct(st.session_state.pallets, float(st.session_state.joda_pct))
        joda_final = base * (1 + eff / 100.0) + joda_charge_fixed

# McDowells (small-load extra is part of base)
mcd_base = get_base_rate(rate_df, postcode_area, st.session_state.service, "Mcdowells", st.session_state.pallets)
mcd_final = None

mcd_charge_fixed = (10 if st.session_state.ampm else 0) + (19 if st.session_state.timed else 0)
mcd_tail_lift_per_pallet = 3.90 if st.session_state.tail else 0.0
mcd_tail_lift_total = mcd_tail_lift_per_pallet * st.session_state.pallets
mcd_small_extra = mcd_smallload_extra(st.session_state.pallets)

if mcd_base is not None:
    mcd_base_for_calc = mcd_base + mcd_small_extra
    mcd_final = (
        mcd_base_for_calc * (1 + float(st.session_state.mcd_pct) / 100.0)
        + mcd_charge_fixed
        + mcd_tail_lift_total
    )

# -------------------------
# Summary table
# -------------------------
summary_rows = []

if joda_base is None:
    summary_rows.append({
        "Haulier": "Joda",
        "Base Rate": "No rate",
        "Fuel Surcharge (%)": f"{float(st.session_state.joda_pct):.2f}%",
        "Delivery Charge": "N/A",
        "Final Rate": "N/A"
    })
else:
    shown_pct = (
        joda_effective_pct(st.session_state.pallets, float(st.session_state.joda_pct))
        if not st.session_state.dual else float(st.session_state.joda_pct)
    )
    summary_rows.append({
        "Haulier": "Joda",
        "Base Rate": f"£{joda_base:,.2f}",
        "Fuel Surcharge (%)": f"{shown_pct:.2f}%",
        "Delivery Charge": f"£{joda_charge_fixed:,.2f}",
        "Final Rate": f"£{joda_final:,.2f}"
    })

if mcd_base is None:
    summary_rows.append({
        "Haulier": "McDowells",
        "Base Rate": "No rate",
        "Fuel Surcharge (%)": f"{float(st.session_state.mcd_pct):.2f}%",
        "Delivery Charge": "N/A",
        "Final Rate": "N/A"
    })
else:
    mcd_base_for_display = mcd_base + mcd_small_extra
    summary_rows.append({
        "Haulier": "McDowells",
        "Base Rate": f"£{mcd_base_for_display:,.2f}",
        "Fuel Surcharge (%)": f"{float(st.session_state.mcd_pct):.2f}%",
        "Delivery Charge": f"£{(mcd_charge_fixed + mcd_tail_lift_total):,.2f}",
        "Final Rate": f"£{mcd_final:,.2f}"
    })

summary_df = pd.DataFrame(summary_rows).set_index("Haulier")

def highlight_cheapest(row):
    fr = row["Final Rate"]
    if isinstance(fr, str) and fr.startswith("£"):
        val = float(fr.strip("£").replace(",", ""))
        j_r = round(joda_final, 2) if joda_final is not None else float("inf")
        m_r = round(mcd_final, 2) if mcd_final is not None else float("inf")
        if math.isclose(round(val, 2), min(j_r, m_r), rel_tol=1e-9):
            return ["background-color: #b3e6b3"] * len(row)
    return [""] * len(row)

# -------------------------
# History (kept)
# -------------------------
HISTORY_FILE = "rate_search_history.json"
if "rate_history" not in st.session_state:
    try:
        if os.path.exists(HISTORY_FILE):
            with open(HISTORY_FILE, "r") as f:
                st.session_state.rate_history = json.load(f)
        else:
            st.session_state.rate_history = []
    except Exception:
        st.session_state.rate_history = []

def _add_history_entry():
    if (joda_final is None) and (mcd_final is None):
        return

    pallets_repr = (
        f"{st.session_state.split1}+{st.session_state.split2}"
        if st.session_state.dual else f"{st.session_state.pallets}"
    )

    cheapest = None
    if joda_final is not None and mcd_final is not None:
        cheapest = "Joda" if joda_final <= mcd_final else "McDowells"
    elif joda_final is not None:
        cheapest = "Joda"
    elif mcd_final is not None:
        cheapest = "McDowells"

    entry = {
        "Time": datetime.now().strftime("%Y-%m-%d %H:%M"),
        "Area": st.session_state.area,
        "Service": st.session_state.service,
        "Pallets": pallets_repr,
        "AM/PM": st.session_state.ampm,
        "Timed": st.session_state.timed,
        "Tail": st.session_state.tail,
        "Dual": st.session_state.dual,
        "Split1": st.session_state.split1,
        "Split2": st.session_state.split2,
        "JodaPct": float(st.session_state.joda_pct),
        "McdPct": float(st.session_state.mcd_pct),
        "JodaFinal": joda_final,
        "McdFinal": mcd_final,
        "Cheapest": cheapest or "—",
    }
    st.session_state.rate_history.insert(0, entry)
    st.session_state.rate_history = st.session_state.rate_history[:10]
    try:
        with open(HISTORY_FILE, "w") as f:
            json.dump(st.session_state.rate_history, f, indent=2)
    except Exception:
        pass

_add_history_entry()

# -------------------------
# Export line builder (fuel baked into delivery unit price ✅)
# -------------------------
def build_export_lines_for_haulier(haulier: str) -> List[Dict[str, object]]:
    so = str(st.session_state.so_number).strip()
    area = str(st.session_state.area).strip().upper()
    svc = str(st.session_state.service).strip()

    if not so:
        raise ValueError("SO Number is required before adding lines.")

    out: List[Dict[str, object]] = []

    if haulier.lower() == "joda":
        if joda_base is None or joda_final is None:
            raise ValueError("No Joda rate available to add.")

        if st.session_state.dual:
            for n, base_n in [
                (st.session_state.split1, get_base_rate(rate_df, area, svc, "Joda", st.session_state.split1)),
                (st.session_state.split2, get_base_rate(rate_df, area, svc, "Joda", st.session_state.split2)),
            ]:
                if base_n is None:
                    continue

                eff = joda_effective_pct(int(n), float(st.session_state.joda_pct))
                base_after_fuel_total = base_n * (1 + eff / 100.0)

                unit = base_after_fuel_total / max(int(n), 1)
                out.append(_export_line(JODA_PO_GROUP, JODA_ACC, so, area, svc, "Delivery", int(n), unit))

                if st.session_state.ampm:
                    out.append(_export_line(JODA_PO_GROUP, JODA_ACC, so, area, svc, "AM Charge", 1, 7.5))
                if st.session_state.timed:
                    out.append(_export_line(JODA_PO_GROUP, JODA_ACC, so, area, svc, "Timed Charge", 1, 20.0))
        else:
            n = int(st.session_state.pallets)
            eff = joda_effective_pct(n, float(st.session_state.joda_pct))
            base_after_fuel_total = float(joda_base) * (1 + eff / 100.0)

            unit = base_after_fuel_total / max(n, 1)
            out.append(_export_line(JODA_PO_GROUP, JODA_ACC, so, area, svc, "Delivery", n, unit))

            if st.session_state.ampm:
                out.append(_export_line(JODA_PO_GROUP, JODA_ACC, so, area, svc, "AM Charge", 1, 7.5))
            if st.session_state.timed:
                out.append(_export_line(JODA_PO_GROUP, JODA_ACC, so, area, svc, "Timed Charge", 1, 20.0))

        return out

    if haulier.lower() in ["mcdowells", "mcd", "mcdowell"]:
        if mcd_base is None or mcd_final is None:
            raise ValueError("No McDowells rate available to add.")

        n = int(st.session_state.pallets)

        base_total_for_calc = float(mcd_base) + float(mcd_small_extra)
        base_after_fuel_total = base_total_for_calc * (1 + float(st.session_state.mcd_pct) / 100.0)

        unit = base_after_fuel_total / max(n, 1)
        out.append(_export_line(MCD_PO_GROUP, MCD_ACC, so, area, svc, "Delivery", n, unit))

        if st.session_state.ampm:
            out.append(_export_line(MCD_PO_GROUP, MCD_ACC, so, area, svc, "AM Charge", 1, 10.0))
        if st.session_state.timed:
            out.append(_export_line(MCD_PO_GROUP, MCD_ACC, so, area, svc, "Timed Charge", 1, 19.0))
        if st.session_state.tail:
            out.append(_export_line(MCD_PO_GROUP, MCD_ACC, so, area, svc, "Tail Lift", n, 3.90))

        return out

    raise ValueError(f"Unknown haulier: {haulier}")

# -------------------------
# Tabs
# -------------------------
st.header("4. Calculated Rates")
tab_table, tab_export, tab_history, tab_map = st.tabs(["Table", "Export List", "History", "Map (Beta)"])

with tab_table:
    if all(r["Final Rate"] == "N/A" for r in summary_rows):
        st.warning("No rates found for that area/service/pallet combination.")
    else:
        st.table(summary_df.style.apply(highlight_cheapest, axis=1))
        st.markdown(
            "<i style='color:gray;'>Rows in green are the cheapest available. "
            "Joda fuel surcharge is waived for pallet counts below 7 (per group when split).</i>",
            unsafe_allow_html=True
        )

    bcols = st.columns([1, 1, 2])
    with bcols[0]:
        if st.button("Add Joda to Export List", use_container_width=True):
            try:
                _add_to_basket(build_export_lines_for_haulier("Joda"))
                st.success("Added Joda lines to Export List.")
            except Exception as e:
                st.error(str(e))
    with bcols[1]:
        if st.button("Add McDowells to Export List", use_container_width=True):
            try:
                _add_to_basket(build_export_lines_for_haulier("McDowells"))
                st.success("Added McDowells lines to Export List.")
            except Exception as e:
                st.error(str(e))
    with bcols[2]:
        st.caption("Adds Delivery (fuel baked into unit price) + extras as separate lines into the Export List tab.")

# -------------------------
# Export List tab (multi-select removal ✅)
# -------------------------
with tab_export:
    st.subheader("Saved Lines (ready for CSV export)")
    basket = st.session_state.get("export_basket", [])

    if not basket:
        st.info("Nothing saved yet. Go to the Table tab and click 'Add Joda…' or 'Add McDowells…'.")
    else:
        view_df = pd.DataFrame(basket)

        show_cols = [
            "Purchase Order Number",
            "Purchase Order Supplier Acc Code",
            "Warehouse Name",
            "Free Text Item Description",
            "Item Quantity",
            "Unit Buying Price",
            "Purchase Order Document Date",
        ]
        show_cols = [c for c in show_cols if c in view_df.columns]

        # Build selectable labels with stable indices
        options: List[str] = []
        idx_map: Dict[str, int] = {}
        for idx, r in view_df.iterrows():
            po = r.get("Purchase Order Number", "")
            acc = r.get("Purchase Order Supplier Acc Code", "")
            desc = r.get("Free Text Item Description", "")
            opt = f"{idx} | PO{po} | {acc} | {desc}"
            options.append(opt)
            idx_map[opt] = int(idx)

        csel1, csel2, csel3 = st.columns([3.4, 1.3, 1.6], gap="medium")

        with csel1:
            st.multiselect(
                "Select lines to remove",
                options=options,
                key="export_selected_keys",
                placeholder="Pick one or more lines…"
            )

        with csel2:
            if st.button("Select all", use_container_width=True):
                st.session_state.export_selected_keys = options
                st.rerun()
            if st.button("Clear selection", use_container_width=True):
                st.session_state.export_selected_keys = []
                st.rerun()

        with csel3:
            if st.button("Remove selected", use_container_width=True):
                selected = st.session_state.get("export_selected_keys", [])
                if not selected:
                    st.warning("No lines selected.")
                else:
                    indices = sorted({idx_map[s] for s in selected if s in idx_map}, reverse=True)
                    for i in indices:
                        try:
                            st.session_state.export_basket.pop(i)
                        except Exception:
                            pass
                    st.session_state.export_selected_keys = []
                    st.success(f"Removed {len(indices)} line(s).")
                    st.rerun()

        st.markdown("---")
        st.dataframe(view_df[show_cols], use_container_width=True, hide_index=True)

        c1, c2 = st.columns([1, 4])
        with c1:
            if st.button("Clear all"):
                st.session_state.export_basket = []
                st.session_state.export_selected_keys = []
                st.rerun()
        with c2:
            st.caption("Fuel is baked into the Delivery unit price. AM/Timed/Tail Lift export as separate lines.")

        export_df = pd.DataFrame(st.session_state.export_basket).reindex(columns=EXPORT_COLUMNS)
        csv_bytes = export_df.to_csv(index=False).encode("utf-8")

        st.download_button(
            label="Download CSV (PO Import Format)",
            data=csv_bytes,
            file_name=f"PO_Import_Export_{date.today().strftime('%Y%m%d')}.csv",
            mime="text/csv",
            use_container_width=True
        )

# -------------------------
# History tab (existing)
# -------------------------
with tab_history:
    hist = st.session_state.get("rate_history", [])
    if not hist:
        st.info("No history yet. Run a calculation to populate this list.")
    else:
        for i, h in enumerate(hist):
            with st.container():
                cols = st.columns([2, 2, 1.3, 1, 1, 1, 1, 1])
                cols[0].markdown(f"**{h.get('Time','')}**")

                saved_area = h.get("Area", "")
                if isinstance(saved_area, list):
                    saved_area = saved_area[0] if saved_area else ""
                saved_area = str(saved_area).upper().strip()

                saved_service = h.get("Service", "Economy")
                if isinstance(saved_service, list):
                    saved_service = saved_service[0] if saved_service else "Economy"
                saved_service = str(saved_service).strip()

                cols[1].markdown(f"**{saved_area}** — {saved_service}")
                cols[2].markdown(f"Pallets: {h.get('Pallets','')}")
                cols[3].markdown(f"AMP/PM: {'Yes' if h.get('AM/PM') else 'No'}")
                cols[4].markdown(f"Timed: {'Yes' if h.get('Timed') else 'No'}")
                cols[5].markdown(f"Tail: {'Yes' if h.get('Tail') else 'No'}")
                cols[6].markdown(f"Cheapest: **{h.get('Cheapest','—')}**")

                if cols[7].button("Load", key=f"load_{i}"):
                    st.session_state["__pending_load"] = {
                        "Area": saved_area,
                        "Service": saved_service,
                        "Pallets": h.get("Pallets", "1"),
                        "AM/PM": h.get("AM/PM", False),
                        "Timed": h.get("Timed", False),
                        "Tail":  h.get("Tail", False),
                        "Dual":  h.get("Dual", False),
                        "Split1": h.get("Split1", 1),
                        "Split2": h.get("Split2", 1),
                        "JodaPct": h.get("JodaPct", st.session_state.get("joda_pct", 0.0)),
                        "McdPct":  h.get("McdPct",  st.session_state.get("mcd_pct", 0.0)),
                    }
                    st.rerun()

        st.markdown("---")
        table_rows = []
        for h in hist:
            table_rows.append({
                "Time": h.get("Time", ""),
                "Area": h.get("Area", ""),
                "Service": h.get("Service", ""),
                "Pallets": h.get("Pallets", ""),
                "AM/PM": "Yes" if h.get("AM/PM") else "No",
                "Timed": "Yes" if h.get("Timed") else "No",
                "Tail lift": "Yes" if h.get("Tail") else "No",
                "Joda final": f"£{h['JodaFinal']:,.2f}" if isinstance(h.get("JodaFinal"), (int, float)) else "—",
                "McD final": f"£{h['McdFinal']:,.2f}" if isinstance(h.get("McdFinal"), (int, float)) else "—",
                "Cheapest": h.get("Cheapest", "—"),
            })
        st.table(pd.DataFrame(table_rows))

# -------------------------
# Map tab (existing)
# -------------------------
with tab_map:
    centroids_path_candidates = [
        "postcode_area_centroids.csv",
        "postcode_area_centroids_filled.csv",
        "/mnt/data/postcode_area_centroids_filled.csv",
    ]

    def _find_col(df: pd.DataFrame, candidates: List[str]) -> Optional[str]:
        cols = {str(c).lower().strip(): c for c in df.columns}
        for key in candidates:
            if key in cols:
                return cols[key]
        return None

    centroid_df = None
    for p in centroids_path_candidates:
        if os.path.exists(p):
            try:
                tmp = pd.read_csv(p)
                tmp.columns = [str(c).strip() for c in tmp.columns]

                area_col = _find_col(tmp, [
                    "area", "postcodearea", "postcode_area", "postcode area", "code", "district", "pc_area"
                ])
                lat_col = _find_col(tmp, ["lat", "latitude", "y"])
                lon_col = _find_col(tmp, ["lon", "lng", "longitude", "x"])

                if not area_col or not lat_col or not lon_col:
                    continue

                centroid_df = tmp.rename(columns={
                    area_col: "Area",
                    lat_col: "lat",
                    lon_col: "lon"
                }).copy()

                centroid_df["Area"] = centroid_df["Area"].astype(str).str.upper().str.strip()
                centroid_df["lat"] = pd.to_numeric(centroid_df["lat"], errors="coerce")
                centroid_df["lon"] = pd.to_numeric(centroid_df["lon"], errors="coerce")
                centroid_df = centroid_df.dropna(subset=["lat", "lon"])
                break
            except Exception:
                continue

    if centroid_df is None:
        st.warning(
            "No usable centroid file found. Ensure your CSV has columns like "
            "`Area` (or `PostcodeArea`), `lat` (or `latitude`) and `lon` (or `longitude`)."
        )
    else:
        def calc_for_area(area_code: str) -> Tuple[Optional[float], Optional[float]]:
            # Joda
            jf = None
            if st.session_state.dual:
                b1 = get_base_rate(rate_df, area_code, st.session_state.service, "Joda", st.session_state.split1)
                b2 = get_base_rate(rate_df, area_code, st.session_state.service, "Joda", st.session_state.split2)
                if b1 is not None and b2 is not None:
                    p1 = joda_effective_pct(st.session_state.split1, float(st.session_state.joda_pct))
                    p2 = joda_effective_pct(st.session_state.split2, float(st.session_state.joda_pct))
                    jf = b1 * (1 + p1 / 100.0) + joda_charge_fixed
                    jf += b2 * (1 + p2 / 100.0) + joda_charge_fixed
            else:
                jb = get_base_rate(rate_df, area_code, st.session_state.service, "Joda", st.session_state.pallets)
                if jb is not None:
                    ep = joda_effective_pct(st.session_state.pallets, float(st.session_state.joda_pct))
                    jf = jb * (1 + ep / 100.0) + joda_charge_fixed

            # McDowells
            mf = None
            mb = get_base_rate(rate_df, area_code, st.session_state.service, "Mcdowells", st.session_state.pallets)
            if mb is not None:
                tl_total = (3.90 if st.session_state.tail else 0.0) * st.session_state.pallets
                small_extra = mcd_smallload_extra(st.session_state.pallets)
                mb_for_calc = mb + small_extra
                mf = (
                    mb_for_calc * (1 + float(st.session_state.mcd_pct) / 100.0)
                    + ((10 if st.session_state.ampm else 0) + (19 if st.session_state.timed else 0))
                    + tl_total
                )

            return jf, mf

        areas = rate_df["PostcodeArea"].unique()
        map_rows = []
        for a in areas:
            jf, mf = calc_for_area(a)
            if jf is None and mf is None:
                continue
            crow = centroid_df.loc[centroid_df["Area"] == a]
            if crow.empty:
                continue
            map_rows.append({
                "Area": a,
                "lat": float(crow.iloc[0]["lat"]),
                "lon": float(crow.iloc[0]["lon"]),
                "JodaFinal": jf if jf is not None else float("nan"),
                "McDFinal": mf if mf is not None else float("nan"),
            })

        if not map_rows:
            st.info("No mappable rates for the selected inputs.")
        else:
            mdf = pd.DataFrame(map_rows)
            mdf["cheaper"] = mdf[["JodaFinal", "McDFinal"]].idxmin(axis=1)
            mdf["size"] = 16

            import pydeck as pdk
            tooltip = {
                "html": """
                <div style="padding:4px 6px">
                  <b>{Area}</b><br/>
                  Joda final: £{JodaFinal}<br/>
                  McDowells final: £{McDFinal}
                </div>
                """,
                "style": {"backgroundColor": "rgba(30,30,30,0.9)", "color": "white"}
            }

            layer = pdk.Layer(
                "ScatterplotLayer",
                data=mdf,
                get_position='[lon, lat]',
                get_radius="size",
                radius_units="pixels",
                pickable=True,
                get_fill_color="""
                    [cheaper == 'JodaFinal' ? 255 : 90,
                     cheaper == 'JodaFinal' ? 64  : 90,
                     cheaper == 'JodaFinal' ? 160 : 255, 200]
                """,
            )

            view_state = pdk.ViewState(latitude=54.5, longitude=-2.5, zoom=4.8)
            st.pydeck_chart(pdk.Deck(layers=[layer], initial_view_state=view_state, tooltip=tooltip))
