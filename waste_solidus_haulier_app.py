# app.py
import os
import math
import json
import uuid
import csv as csvlib
from datetime import date
from typing import Optional, List, Dict

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
        V3.2.2    
        Enter a UK postcode area, select a service type, set pallets and surcharges,
        and optionally add AM/PM, Tail Lift or Timed Delivery. Dual Collection splits the load.

        **What’s NEW?**    
        **NEW:** All fuel surcharges are saveable.    
        **FIX:** Joda base rates always round up to £0dp before surcharges.    
        """,
        unsafe_allow_html=True
    )

# -------------------------
# Config / constants
# -------------------------
JODA_DATA_FILE = "joda_surcharge.json"
MCD_DATA_FILE = "mcd_surcharge.json"
PCH_DATA_FILE = "pch_surcharge.json"

RATE_XLSX_MAIN = "haulier prices 2.xlsx"   # Joda + McDowells
RATE_XLSX_PCH  = "pch_rates_app.xlsx"      # PC Howard (converted for app)

TEMPLATE_PATH = "PO Import Example File.csv"

# Supplier account codes
JODA_ACC = "J040"
MCD_ACC  = "M127"
PCH_ACC  = "P031"  # PC Howard

# Warehouse options (dropdown)
WAREHOUSE_OPTIONS = ["101 - Skipton", "201 - Skipton 2", "102 - Corby"]

# Which hauliers are allowed for which warehouse
WAREHOUSE_HAULIERS = {
    "101 - Skipton": ["Joda", "Mcdowells"],
    "201 - Skipton 2": ["Joda", "Mcdowells"],
    "102 - Corby": ["Pc Howard"],
}

# Each unique (haulier, warehouse) must have a unique PO Number
PO_NUMBER_MAP = {
    ("Joda", "101 - Skipton"): 1,
    ("Joda", "201 - Skipton 2"): 2,
    ("Mcdowells", "101 - Skipton"): 3,
    ("Mcdowells", "201 - Skipton 2"): 4,
    ("Pc Howard", "102 - Corby"): 5,
}


def po_number_for(haulier: str, warehouse: str) -> int:
    key = (haulier.strip().title(), warehouse.strip())
    if key not in PO_NUMBER_MAP:
        raise KeyError(f"No PO number mapping for {key}")
    return int(PO_NUMBER_MAP[key])


def available_hauliers() -> List[str]:
    wh = st.session_state.warehouse_name
    return WAREHOUSE_HAULIERS.get(wh, [])


def display_haulier(name: str) -> str:
    n = str(name).strip()
    if n.lower() in {"pc howard", "pc", "pch", "p.c. howard", "pc howard "}:
        return "PC Howard"
    if n.lower() == "mcdowells":
        return "McDowells"
    if n.lower() == "joda":
        return "Joda"
    return n


# Fallback columns if template file is missing
DEFAULT_EXPORT_COLUMNS: List[str] = [
    "Purchase Order Import Type",
    "Purchase Order Number",
    "Purchase Order Supplier Acc Code",
    "Purchase Order Document Date",
    "Purchase Order Header Requested Date",
    "Purchase Order Discount Percent",
    "Purchase Order Supplier Document No.",
    "Item Code",
    "Warehouse Name",
    "Purchase Order Line Requested Date",
    "Free Text Item Description",
    "Item Quantity",
    "Unit Buying Price",
]


@st.cache_data
def load_export_template_columns(path: str) -> List[str]:
    if not os.path.exists(path):
        return DEFAULT_EXPORT_COLUMNS

    with open(path, "r", encoding="utf-8-sig", newline="") as f:
        reader = csvlib.reader(f, delimiter=",")
        header = next(reader, None)

    if not header:
        return DEFAULT_EXPORT_COLUMNS

    return [h.strip() for h in header if str(h).strip()]


EXPORT_COLUMNS = load_export_template_columns(TEMPLATE_PATH)

# -------------------------
# Surcharge persistence
# -------------------------
def load_joda_surcharge() -> float:
    """
    Joda has special weekly reset logic (Wednesday).
    """
    today_str = date.today().isoformat()
    if not os.path.exists(JODA_DATA_FILE):
        with open(JODA_DATA_FILE, "w") as f:
            json.dump({"surcharge": 0.0, "last_updated": today_str}, f)
        return 0.0

    try:
        with open(JODA_DATA_FILE, "r") as f:
            data = json.load(f)
    except Exception:
        data = {"surcharge": 0.0, "last_updated": today_str}

    # reset on Wednesdays
    if date.today().weekday() == 2 and data.get("last_updated") != today_str:
        data = {"surcharge": 0.0, "last_updated": today_str}
        with open(JODA_DATA_FILE, "w") as f:
            json.dump(data, f)
        return 0.0

    try:
        return float(data.get("surcharge", 0.0))
    except Exception:
        return 0.0


def save_joda_surcharge(new_pct: float):
    today_str = date.today().isoformat()
    with open(JODA_DATA_FILE, "w") as f:
        json.dump({"surcharge": float(new_pct), "last_updated": today_str}, f)


def load_simple_surcharge(path: str) -> float:
    """Loads a stored surcharge % from a json file. No weekly reset logic."""
    today_str = date.today().isoformat()
    if not os.path.exists(path):
        with open(path, "w") as f:
            json.dump({"surcharge": 0.0, "last_updated": today_str}, f)
        return 0.0
    try:
        with open(path, "r") as f:
            data = json.load(f)
        return float(data.get("surcharge", 0.0))
    except Exception:
        return 0.0


def save_simple_surcharge(path: str, new_pct: float):
    today_str = date.today().isoformat()
    with open(path, "w") as f:
        json.dump({"surcharge": float(new_pct), "last_updated": today_str}, f)


joda_stored_pct = load_joda_surcharge()
mcd_stored_pct = load_simple_surcharge(MCD_DATA_FILE)
pch_stored_pct = load_simple_surcharge(PCH_DATA_FILE)

# -------------------------
# Rates load
# -------------------------
@st.cache_data
def load_rate_table(excel_path: str, _mtime: float) -> pd.DataFrame:
    raw = pd.read_excel(excel_path, header=1)
    raw = raw.rename(
        columns={
            raw.columns[0]: "PostcodeArea",
            raw.columns[1]: "Service",
            raw.columns[2]: "Vendor",
        }
    )

    raw["PostcodeArea"] = raw["PostcodeArea"].ffill()
    raw["Service"] = raw["Service"].ffill()
    raw["Vendor"] = raw["Vendor"].ffill()
    raw = raw[raw["Vendor"] != "Vendor"].copy()

    pallet_cols = [
        c for c in raw.columns
        if isinstance(c, (int, float)) or (isinstance(c, str) and str(c).isdigit())
    ]

    melted = raw.melt(
        id_vars=["PostcodeArea", "Service", "Vendor"],
        value_vars=pallet_cols,
        var_name="Pallets",
        value_name="BaseRate",
    )
    melted["Pallets"] = melted["Pallets"].astype(int)
    melted["BaseRate"] = pd.to_numeric(melted["BaseRate"], errors="coerce")
    melted = melted.dropna(subset=["BaseRate"]).copy()

    melted["PostcodeArea"] = melted["PostcodeArea"].astype(str).str.strip().str.upper()
    melted["Service"] = melted["Service"].astype(str).str.strip().str.title()
    melted["Vendor"] = melted["Vendor"].astype(str).str.strip().str.title()

    return melted.reset_index(drop=True)


mtime_main = os.path.getmtime(RATE_XLSX_MAIN)
rate_df_main = load_rate_table(RATE_XLSX_MAIN, mtime_main)
unique_areas_main = sorted(rate_df_main["PostcodeArea"].unique())

rate_df_pch = pd.DataFrame(columns=["PostcodeArea", "Service", "Vendor", "Pallets", "BaseRate"])
unique_areas_pch: List[str] = []
if os.path.exists(RATE_XLSX_PCH):
    mtime_pch = os.path.getmtime(RATE_XLSX_PCH)
    rate_df_pch = load_rate_table(RATE_XLSX_PCH, mtime_pch)
    unique_areas_pch = sorted(rate_df_pch["PostcodeArea"].unique())

# -------------------------
# Session defaults
# -------------------------
def _ensure_defaults():
    st.session_state.setdefault("warehouse_name", WAREHOUSE_OPTIONS[0])
    st.session_state.setdefault("area", "")
    st.session_state.setdefault("service", "Economy")
    st.session_state.setdefault("pallets", 1)

    st.session_state.setdefault("joda_pct", round(joda_stored_pct, 2))
    st.session_state.setdefault("mcd_pct", round(mcd_stored_pct, 2))
    st.session_state.setdefault("pch_pct", round(pch_stored_pct, 2))

    st.session_state.setdefault("ampm", False)
    st.session_state.setdefault("tail", False)
    st.session_state.setdefault("dual", False)
    st.session_state.setdefault("timed", False)
    st.session_state.setdefault("split1", 1)
    st.session_state.setdefault("split2", 1)

    st.session_state.setdefault("so_number", "")
    st.session_state.setdefault("export_basket", [])


_ensure_defaults()

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


def joda_round_base_up(x: float) -> float:
    # Always round UP to whole pounds
    return float(math.ceil(float(x)))


def joda_effective_pct(pallet_count: int, input_pct: float) -> float:
    return 0.0 if pallet_count < 7 else float(input_pct)


def mcd_smallload_extra(pallet_count: int) -> float:
    return (5.0 * min(pallet_count, 4)) if pallet_count < 5 else 0.0


def _ddmmyyyy(d: date) -> str:
    return d.strftime("%d/%m/%Y")


def _blank_export_row() -> Dict[str, object]:
    return {c: "" for c in EXPORT_COLUMNS}


def _export_line(
    po_number: int,
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
    r["_row_id"] = uuid.uuid4().hex  # stable ID for UI removes (not exported)

    r["Purchase Order Import Type"] = 1
    r["Purchase Order Number"] = int(po_number)
    r["Purchase Order Supplier Acc Code"] = str(supplier_acc).strip()
    r["Purchase Order Document Date"] = _ddmmyyyy(doc_date)
    r["Purchase Order Header Requested Date"] = _ddmmyyyy(req_date)
    r["Purchase Order Discount Percent"] = 0

    wh = st.session_state.warehouse_name
    r["Warehouse Name"] = wh

    # Copy warehouse into column G: "Purchase Order Supplier Document No."
    if "Purchase Order Supplier Document No." in r:
        r["Purchase Order Supplier Document No."] = wh

    if "Purchase Order Line Requested Date" in r:
        r["Purchase Order Line Requested Date"] = _ddmmyyyy(req_date)

    so_number = str(so_number).strip()
    so_suffix = f" - SO{so_number}" if so_number else ""
    svc_suffix = f" ({service})" if str(service).strip() else ""
    desc = f"{area_code} {label}{svc_suffix}{so_suffix}".strip()
    if "Free Text Item Description" in r:
        r["Free Text Item Description"] = desc

    if "Item Quantity" in r:
        r["Item Quantity"] = float(qty)
    if "Unit Buying Price" in r:
        r["Unit Buying Price"] = round(float(unit_price), 5)

    return r


def _add_to_basket(rows: List[Dict[str, object]]):
    for r in rows:
        r["_row_id"] = r.get("_row_id") or uuid.uuid4().hex
    st.session_state.export_basket.extend(rows)


# -------------------------
# UI: Inputs
# -------------------------
st.header("1. Input Parameters")
col_a, col_b, col_c, col_d, col_e, col_f, col_g, col_h, col_i = st.columns(
    [1, 1, 1, 1, 1, 1, 1, 1, 1], gap="medium"
)

with col_a:
    st.selectbox("Warehouse", options=WAREHOUSE_OPTIONS, key="warehouse_name")

allowed = set(available_hauliers())
pc_only = (allowed == {"Pc Howard"})

area_options = unique_areas_pch if pc_only else unique_areas_main

if st.session_state.area and st.session_state.area not in area_options:
    st.session_state.area = ""

with col_b:
    options_for_select = [""] + area_options
    st.selectbox(
        "Postcode Area",
        options=options_for_select,
        index=options_for_select.index(st.session_state.area) if st.session_state.area in options_for_select else 0,
        key="area",
        format_func=lambda x: x if x else "— Select area —",
    )
    if st.session_state.area == "":
        st.info("Please select a postcode area to continue.")
        st.stop()

with col_c:
    st.selectbox("Service Type", options=["Economy", "Next Day"], key="service")

with col_d:
    st.number_input("Number of Pallets", min_value=1, max_value=26, step=1, key="pallets")

with col_e:
    st.number_input(
        "Joda Fuel Surcharge (%)",
        min_value=0.0, max_value=100.0,
        step=0.1, format="%.2f",
        key="joda_pct",
        disabled=pc_only
    )
    if st.button("Save Joda", disabled=pc_only):
        save_joda_surcharge(float(st.session_state.joda_pct))
        st.success(f"Saved Joda at {float(st.session_state.joda_pct):.2f}%")

with col_f:
    st.number_input(
        "McDowells Fuel Surcharge (%)",
        min_value=0.0, max_value=100.0,
        step=0.1, format="%.2f",
        key="mcd_pct",
        disabled=pc_only
    )
    if st.button("Save McDowells", disabled=pc_only):
        save_simple_surcharge(MCD_DATA_FILE, float(st.session_state.mcd_pct))
        st.success(f"Saved McDowells at {float(st.session_state.mcd_pct):.2f}%")

with col_g:
    st.number_input(
        "PC Howard Fuel Surcharge (%)",
        min_value=0.0, max_value=100.0,
        step=0.1, format="%.2f",
        key="pch_pct",
        disabled=not pc_only
    )
    if st.button("Save PC Howard", disabled=not pc_only):
        save_simple_surcharge(PCH_DATA_FILE, float(st.session_state.pch_pct))
        st.success(f"Saved PC Howard at {float(st.session_state.pch_pct):.2f}%")

with col_h:
    st.markdown("**Available hauliers**")
    st.write(", ".join(display_haulier(x) for x in sorted(allowed)) if allowed else "—")

with col_i:
    st.markdown("&nbsp;", unsafe_allow_html=True)

st.markdown("---")
postcode_area = st.session_state.area

if pc_only:
    st.session_state.dual = False
    st.session_state.tail = False

st.subheader("2. Optional Extras")
col1, col2, col3, col4 = st.columns(4, gap="large")

with col1:
    st.checkbox("AM/PM Delivery", key="ampm")
with col2:
    st.checkbox("Tail Lift", key="tail", disabled=pc_only)
with col3:
    st.checkbox("Dual Collection", key="dual", disabled=pc_only)
with col4:
    st.checkbox("Timed Delivery", key="timed")

if st.session_state.dual and st.session_state.pallets == 1:
    st.error("Dual Collection requires at least 2 pallets.")
    st.stop()

if st.session_state.dual:
    st.markdown("**Split pallets into two despatches (e.g., ESL & U4).**")
    sp1, sp2 = st.columns(2, gap="large")
    with sp1:
        st.number_input("First Pallet Group", 1, st.session_state.pallets - 1, key="split1")
    with sp2:
        st.number_input("Second Pallet Group", 1, st.session_state.pallets - 1, key="split2")
    if st.session_state.split1 + st.session_state.split2 != st.session_state.pallets:
        st.error("Pallet Split values must add up to total pallets.")
        st.stop()


# -------------------------
# Core calculations for selected area
# -------------------------
def calc_for_area(area_code: str):
    svc = st.session_state.service
    allowed_local = set(available_hauliers())

    # fixed extras
    joda_charge_fixed = (7.5 if st.session_state.ampm else 0) + (20 if st.session_state.timed else 0)
    mcd_charge_fixed = (10 if st.session_state.ampm else 0) + (19 if st.session_state.timed else 0)
    pch_charge_fixed = (15.0 if st.session_state.ampm else 0) + (17.5 if st.session_state.timed else 0)

    # Joda (round base UP to whole pounds BEFORE fuel)
    jb = jf = None
    if "Joda" in allowed_local:
        if st.session_state.dual:
            b1 = get_base_rate(rate_df_main, area_code, svc, "Joda", st.session_state.split1)
            b2 = get_base_rate(rate_df_main, area_code, svc, "Joda", st.session_state.split2)
            if b1 is not None:
                b1 = joda_round_base_up(b1)
            if b2 is not None:
                b2 = joda_round_base_up(b2)

            if b1 is not None and b2 is not None:
                p1 = joda_effective_pct(st.session_state.split1, float(st.session_state.joda_pct))
                p2 = joda_effective_pct(st.session_state.split2, float(st.session_state.joda_pct))
                jf = b1 * (1 + p1 / 100.0) + joda_charge_fixed
                jf += b2 * (1 + p2 / 100.0) + joda_charge_fixed
                jb = b1 + b2
        else:
            jb = get_base_rate(rate_df_main, area_code, svc, "Joda", st.session_state.pallets)
            if jb is not None:
                jb = joda_round_base_up(jb)
                ep = joda_effective_pct(st.session_state.pallets, float(st.session_state.joda_pct))
                jf = jb * (1 + ep / 100.0) + joda_charge_fixed

    # McDowells
    mb = mf = None
    if "Mcdowells" in allowed_local:
        mb = get_base_rate(rate_df_main, area_code, svc, "Mcdowells", st.session_state.pallets)
        if mb is not None:
            small_extra = mcd_smallload_extra(st.session_state.pallets)
            tl_total = (3.90 if st.session_state.tail else 0.0) * st.session_state.pallets
            mb_calc = mb + small_extra
            mf = (
                mb_calc * (1 + float(st.session_state.mcd_pct) / 100.0)
                + mcd_charge_fixed
                + tl_total
            )

    # PC Howard (fuel applies to base delivery only; AM/Timed are fixed)
    pb = pf = None
    if "Pc Howard" in allowed_local and not rate_df_pch.empty:
        pb = get_base_rate(rate_df_pch, area_code, svc, "Pc Howard", st.session_state.pallets)
        if pb is not None:
            pb_after_fuel = pb * (1 + float(st.session_state.pch_pct) / 100.0)
            pf = pb_after_fuel + pch_charge_fixed

    return jb, jf, mb, mf, pb, pf


joda_base, joda_final, mcd_base, mcd_final, pch_base, pch_final = calc_for_area(postcode_area)

# derive components for display
joda_charge_fixed = (7.5 if st.session_state.ampm else 0) + (20 if st.session_state.timed else 0)
mcd_charge_fixed = (10 if st.session_state.ampm else 0) + (19 if st.session_state.timed else 0)
mcd_tail_lift_total = (3.90 if st.session_state.tail else 0.0) * st.session_state.pallets
mcd_small_extra = mcd_smallload_extra(st.session_state.pallets)
pch_charge_fixed = (15.0 if st.session_state.ampm else 0) + (17.5 if st.session_state.timed else 0)

# -------------------------
# Summary table
# -------------------------
summary_rows = []

if "Joda" in allowed:
    if joda_base is None:
        summary_rows.append({"Haulier": "Joda", "Base Rate": "No rate", "Fuel Surcharge (%)": f"{float(st.session_state.joda_pct):.2f}%", "Delivery Charge": "N/A", "Final Rate": "N/A"})
    else:
        shown_pct = (joda_effective_pct(st.session_state.pallets, float(st.session_state.joda_pct)) if not st.session_state.dual else float(st.session_state.joda_pct))
        summary_rows.append({"Haulier": "Joda", "Base Rate": f"£{joda_base:,.0f}", "Fuel Surcharge (%)": f"{shown_pct:.2f}%", "Delivery Charge": f"£{joda_charge_fixed:,.2f}", "Final Rate": f"£{joda_final:,.2f}"})

if "Mcdowells" in allowed:
    if mcd_base is None:
        summary_rows.append({"Haulier": "McDowells", "Base Rate": "No rate", "Fuel Surcharge (%)": f"{float(st.session_state.mcd_pct):.2f}%", "Delivery Charge": "N/A", "Final Rate": "N/A"})
    else:
        mcd_base_for_display = mcd_base + mcd_small_extra
        summary_rows.append({"Haulier": "McDowells", "Base Rate": f"£{mcd_base_for_display:,.2f}", "Fuel Surcharge (%)": f"{float(st.session_state.mcd_pct):.2f}%", "Delivery Charge": f"£{(mcd_charge_fixed + mcd_tail_lift_total):,.2f}", "Final Rate": f"£{mcd_final:,.2f}"})

if "Pc Howard" in allowed:
    if pch_base is None:
        summary_rows.append({"Haulier": "PC Howard", "Base Rate": "No rate", "Fuel Surcharge (%)": f"{float(st.session_state.pch_pct):.2f}%", "Delivery Charge": "N/A", "Final Rate": "N/A"})
    else:
        summary_rows.append({"Haulier": "PC Howard", "Base Rate": f"£{pch_base:,.2f}", "Fuel Surcharge (%)": f"{float(st.session_state.pch_pct):.2f}%", "Delivery Charge": f"£{pch_charge_fixed:,.2f}", "Final Rate": f"£{pch_final:,.2f}"})

summary_df = pd.DataFrame(summary_rows).set_index("Haulier") if summary_rows else pd.DataFrame()

def highlight_cheapest(row):
    fr = row.get("Final Rate", "")
    if isinstance(fr, str) and fr.startswith("£"):
        val = float(fr.strip("£").replace(",", ""))
        candidates = []
        if isinstance(joda_final, (int, float)): candidates.append(round(float(joda_final), 2))
        if isinstance(mcd_final, (int, float)):  candidates.append(round(float(mcd_final), 2))
        if isinstance(pch_final, (int, float)):  candidates.append(round(float(pch_final), 2))
        if candidates and math.isclose(round(val, 2), min(candidates), rel_tol=1e-9):
            return ["background-color: #b3e6b3"] * len(row)
    return [""] * len(row)

# -------------------------
# Export line builder
# -------------------------
def build_export_lines_for_haulier(haulier: str) -> List[Dict[str, object]]:
    so = str(st.session_state.so_number).strip()
    area = str(st.session_state.area).strip().upper()
    svc = str(st.session_state.service).strip()
    wh = str(st.session_state.warehouse_name).strip()

    if not so:
        raise ValueError("SO Number is required before adding lines.")

    allowed_local = set(available_hauliers())
    h_norm = haulier.strip().title()
    if h_norm not in allowed_local:
        raise ValueError(f"{display_haulier(haulier)} is not available for warehouse {wh}.")

    out: List[Dict[str, object]] = []

    if h_norm == "Joda":
        if joda_base is None or joda_final is None:
            raise ValueError("No Joda rate available to add.")
        po_no = po_number_for("Joda", wh)

        if st.session_state.dual:
            for n, base_n in [
                (st.session_state.split1, get_base_rate(rate_df_main, area, svc, "Joda", st.session_state.split1)),
                (st.session_state.split2, get_base_rate(rate_df_main, area, svc, "Joda", st.session_state.split2)),
            ]:
                if base_n is None:
                    continue

                base_n = joda_round_base_up(base_n)

                eff = joda_effective_pct(int(n), float(st.session_state.joda_pct))
                base_after_fuel_total = base_n * (1 + eff / 100.0)
                unit = base_after_fuel_total / max(int(n), 1)
                out.append(_export_line(po_no, JODA_ACC, so, area, svc, "Delivery", int(n), unit))
                if st.session_state.ampm:
                    out.append(_export_line(po_no, JODA_ACC, so, area, svc, "AM Charge", 1, 7.5))
                if st.session_state.timed:
                    out.append(_export_line(po_no, JODA_ACC, so, area, svc, "Timed Charge", 1, 20.0))
        else:
            n = int(st.session_state.pallets)
            base_whole = joda_round_base_up(float(joda_base))
            eff = joda_effective_pct(n, float(st.session_state.joda_pct))
            base_after_fuel_total = base_whole * (1 + eff / 100.0)
            unit = base_after_fuel_total / max(n, 1)
            out.append(_export_line(po_no, JODA_ACC, so, area, svc, "Delivery", n, unit))
            if st.session_state.ampm:
                out.append(_export_line(po_no, JODA_ACC, so, area, svc, "AM Charge", 1, 7.5))
            if st.session_state.timed:
                out.append(_export_line(po_no, JODA_ACC, so, area, svc, "Timed Charge", 1, 20.0))
        return out

    if h_norm in ["Mcdowells", "Mcdowell", "Mcd"]:
        if mcd_base is None or mcd_final is None:
            raise ValueError("No McDowells rate available to add.")
        po_no = po_number_for("Mcdowells", wh)

        n = int(st.session_state.pallets)
        base_total_for_calc = float(mcd_base) + float(mcd_small_extra)
        base_after_fuel_total = base_total_for_calc * (1 + float(st.session_state.mcd_pct) / 100.0)
        unit = base_after_fuel_total / max(n, 1)
        out.append(_export_line(po_no, MCD_ACC, so, area, svc, "Delivery", n, unit))
        if st.session_state.ampm:
            out.append(_export_line(po_no, MCD_ACC, so, area, svc, "AM Charge", 1, 10.0))
        if st.session_state.timed:
            out.append(_export_line(po_no, MCD_ACC, so, area, svc, "Timed Charge", 1, 19.0))
        if st.session_state.tail:
            out.append(_export_line(po_no, MCD_ACC, so, area, svc, "Tail Lift", n, 3.90))
        return out

    if h_norm == "Pc Howard":
        if rate_df_pch.empty:
            raise ValueError("PC Howard rate file missing. Place 'pch_rates_app.xlsx' alongside app.py.")
        if pch_base is None or pch_final is None:
            raise ValueError("No PC Howard rate available to add.")

        po_no = po_number_for("Pc Howard", wh)
        n = int(st.session_state.pallets)

        base_after_fuel = float(pch_base) * (1 + float(st.session_state.pch_pct) / 100.0)
        unit = base_after_fuel / max(n, 1)
        out.append(_export_line(po_no, PCH_ACC, so, area, svc, "Delivery", n, unit))
        if st.session_state.ampm:
            out.append(_export_line(po_no, PCH_ACC, so, area, svc, "AM Charge", 1, 15.0))
        if st.session_state.timed:
            out.append(_export_line(po_no, PCH_ACC, so, area, svc, "Timed Charge", 1, 17.5))
        return out

    raise ValueError(f"Unknown haulier: {haulier}")

# -------------------------
# Tabs
# -------------------------
st.header("2. Calculated Rates")
tab_table, tab_export = st.tabs(["Table", "Export List"])

with tab_table:
    if summary_df.empty or all(r.get("Final Rate") == "N/A" for r in summary_rows):
        st.warning("No rates found for that area/service/pallet combination.")
    else:
        st.table(summary_df.style.apply(highlight_cheapest, axis=1))

    st.markdown("---")
    st.subheader("Add to Export List")

    c_exp1, c_exp2 = st.columns([1, 1], gap="medium")
    with c_exp1:
        st.text_input("SO Number (manual)", key="so_number", placeholder="e.g. 020502")
    with c_exp2:
        st.write(f"Warehouse: **{st.session_state.warehouse_name}**")

    buttons = st.columns([1, 1, 1, 2])

    if "Joda" in allowed:
        if buttons[0].button("Add Joda", use_container_width=True):
            try:
                _add_to_basket(build_export_lines_for_haulier("Joda"))
                st.success("Added Joda lines to Export List.")
            except Exception as e:
                st.error(str(e))

    if "Mcdowells" in allowed:
        if buttons[1].button("Add McDowells", use_container_width=True):
            try:
                _add_to_basket(build_export_lines_for_haulier("Mcdowells"))
                st.success("Added McDowells lines to Export List.")
            except Exception as e:
                st.error(str(e))

    if "Pc Howard" in allowed:
        if buttons[2].button("Add PC Howard", use_container_width=True):
            try:
                _add_to_basket(build_export_lines_for_haulier("Pc Howard"))
                st.success("Added PC Howard lines to Export List.")
            except Exception as e:
                st.error(str(e))

with tab_export:
    st.subheader("Saved Lines (ready for export)")

    if not os.path.exists(TEMPLATE_PATH):
        st.warning(f"Template file '{TEMPLATE_PATH}' not found next to app.py. Using fallback columns.")

    basket = st.session_state.get("export_basket", [])
    if not basket:
        st.info("Nothing saved yet. Go to the Table tab and click an 'Add …' button.")
    else:
        h = st.columns([0.7, 1.2, 1.6, 4.8, 1.0, 1.2, 1.5, 0.9])
        h[0].markdown("**PO**")
        h[1].markdown("**Supplier**")
        h[2].markdown("**Warehouse**")
        h[3].markdown("**Description**")
        h[4].markdown("**Qty**")
        h[5].markdown("**Unit £**")
        h[6].markdown("**Doc Date**")
        h[7].markdown("**Remove**")
        st.divider()

        remove_id = None
        for r in basket:
            rid = r.get("_row_id", "")
            cols = st.columns([0.7, 1.2, 1.6, 4.8, 1.0, 1.2, 1.5, 0.9])

            cols[0].write(r.get("Purchase Order Number", ""))
            cols[1].write(r.get("Purchase Order Supplier Acc Code", ""))
            cols[2].write(r.get("Warehouse Name", ""))
            cols[3].write(r.get("Free Text Item Description", ""))
            cols[4].write(r.get("Item Quantity", ""))
            cols[5].write(r.get("Unit Buying Price", ""))
            cols[6].write(r.get("Purchase Order Document Date", ""))

            if cols[7].button("🗑", key=f"rm_{rid}", help="Remove this line"):
                remove_id = rid

        if remove_id:
            st.session_state.export_basket = [x for x in st.session_state.export_basket if x.get("_row_id") != remove_id]
            st.rerun()

        st.markdown("---")
        c1, c2 = st.columns([1, 4])
        with c1:
            if st.button("Clear all"):
                st.session_state.export_basket = []
                st.rerun()
        with c2:
            st.caption("Blank cells must be truly blank on export.")

        export_df = pd.DataFrame(st.session_state.export_basket).reindex(columns=EXPORT_COLUMNS)
        export_df = export_df.where(pd.notnull(export_df), "")

        csv_bytes = export_df.to_csv(
            index=False,
            sep=",",
            na_rep="",
            lineterminator="\n",
            quoting=csvlib.QUOTE_MINIMAL,
        ).encode("utf-8")

        st.download_button(
            label="Download PO Import File (.csv)",
            data=csv_bytes,
            file_name=f"PO_Import_Export_{date.today().strftime('%Y%m%d')}.csv",
            mime="text/csv",
            use_container_width=True,
        )
