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
        V3.3.0    
        **Changes in this build**
        - Pallets can exceed 26; pricing caps at the maximum pallet band in the rate sheet (e.g. 26)
        - Fuel surcharges export as their own line (qty=1, unit £ = total surcharge)
        - Tax Code column always outputs 1
        - SO Number auto-clears after successfully adding to Export List
        - Added a **second export**: **McDowells Portal CSV** (Service codes: Economy=2D, Next Day=ND; Depots fixed to 008)
        - Fuel surcharge inputs live above Warehouse selection so they never disappear / reset
        - Joda base rate rounds up to £0dp before surcharges (display keeps .00)
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
RATE_XLSX_PCH = "pch_rates_app.xlsx"       # PC Howard (converted for app)

TEMPLATE_SAGE_PATH = "PO Import Example File.csv"
TEMPLATE_MCD_PATH = "Reference.csv"  # McDowells portal template header (your example)

# Supplier account codes
JODA_ACC = "J040"
MCD_ACC = "M127"
PCH_ACC = "P031"  # PC Howard

# Warehouse options (dropdown)
WAREHOUSE_OPTIONS = ["101 - Skipton", "201 - Skipton 2", "102 - Corby"]

# Which hauliers are allowed for which warehouse
WAREHOUSE_HAULIERS = {
    "101 - Skipton": ["Joda", "Mcdowells"],
    "201 - Skipton 2": ["Joda", "Mcdowells"],
    "102 - Corby": ["Pc Howard"],  # internal key stays "Pc Howard"
}

# Each unique (haulier, warehouse) must have a unique PO Number
PO_NUMBER_MAP = {
    ("Joda", "101 - Skipton"): 1,
    ("Joda", "201 - Skipton 2"): 2,
    ("Mcdowells", "101 - Skipton"): 3,
    ("Mcdowells", "201 - Skipton 2"): 4,
    ("Pc Howard", "102 - Corby"): 5,
}

# McDowells portal constants (per your note)
MCD_REQ_DEPOT = "008"
MCD_COLL_DEPOT = "008"
MCD_DEL_DEPOT = "008"
MCD_SERVICE_MAP = {"Economy": "2D", "Next Day": "ND"}


def po_number_for(haulier: str, warehouse: str) -> int:
    key = (haulier.strip().title(), warehouse.strip())
    if key not in PO_NUMBER_MAP:
        raise KeyError(f"No PO number mapping for {key}")
    return int(PO_NUMBER_MAP[key])


def available_hauliers() -> List[str]:
    wh = st.session_state.get("warehouse_name", WAREHOUSE_OPTIONS[0])
    return WAREHOUSE_HAULIERS.get(wh, [])


def display_haulier(name: str) -> str:
    n = str(name).strip()
    if n.lower() in {"pc howard", "pc", "pch", "p.c. howard"}:
        return "PC Howard"
    if n.lower() == "mcdowells":
        return "McDowells"
    if n.lower() == "joda":
        return "Joda"
    return n


# -------------------------
# Template columns
# -------------------------
DEFAULT_SAGE_EXPORT_COLUMNS: List[str] = [
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
    "Tax Code",
    "Item Quantity",
    "Unit Buying Price",
]


@st.cache_data
def load_csv_header_columns(path: str, fallback: List[str]) -> List[str]:
    if not os.path.exists(path):
        return fallback
    try:
        with open(path, "r", encoding="utf-8-sig", newline="") as f:
            reader = csvlib.reader(f, delimiter=",")
            header = next(reader, None)
        if not header:
            return fallback
        return [h.strip() for h in header if str(h).strip()]
    except Exception:
        return fallback


SAGE_EXPORT_COLUMNS = load_csv_header_columns(TEMPLATE_SAGE_PATH, DEFAULT_SAGE_EXPORT_COLUMNS)

# McDowells portal columns (fallback from your Reference.csv example)
DEFAULT_MCD_PORTAL_COLUMNS: List[str] = [
    "Docket",
    "Order_No",
    "Despatch Date",
    "Requesting Depot",
    "Collect Depot",
    "Consignor Name",
    "ConsignorPostCode",
    "Consignee Name",
    "Consignee Address 1",
    "Consignee Address 2",
    "Consignee Address 3",
    "Consignee Address 4",
    "Consignee Postcode",
    "Delivery Depot",
    "Trunk",
    "Service",
    "Delivery Time",
    "Half Pallets",
    "Half Weight",
    "Full Pallets",
    "Full Weight",
    "Half Oversize Pallets",
    "Half Oversize Weight",
    "Full Oversize Pallets",
    "Full Oversize Weight",
    "Remarks 1",
    "Remarks 2",
    "Delivery Date ",
    "Revenue",
    "Insure Value",
    "Manifest Date",
    "Quarter Pallets",
    "Quarter Weight",
    "Customer Own Paperwork",
    "Consignor Account",
    "Consignee Contact",
    "Consignee Tel",
    "Day Time Freight",
    "Insurance Charge",
    "Insured Name",
    "Insured Email",
    "Entered By",
    "OOG3 Pallets",
    "OOG3 Weight",
    "OOG4 Pallets",
    "OOG4 Weight",
    "Not Used 1",
    "Not Used 2",
    "Hazchem",
    "Customer Reference",
    "UN Number",
    "Hazchem Weight",
    "Consignor Email",
    "7.5t",
]

MCD_PORTAL_COLUMNS = load_csv_header_columns(TEMPLATE_MCD_PATH, DEFAULT_MCD_PORTAL_COLUMNS)

# -------------------------
# Surcharge persistence
# -------------------------
def load_joda_surcharge() -> float:
    """Joda has special weekly reset logic (Wednesday)."""
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


def refresh_surcharges_from_disk():
    st.session_state["joda_pct"] = round(load_joda_surcharge(), 2)
    st.session_state["mcd_pct"] = round(load_simple_surcharge(MCD_DATA_FILE), 2)
    st.session_state["pch_pct"] = round(load_simple_surcharge(PCH_DATA_FILE), 2)

# -------------------------
# Rates load
# -------------------------
@st.cache_data
def load_rate_table(excel_path: str, _mtime: float) -> pd.DataFrame:
    raw = pd.read_excel(excel_path, header=1)
    raw = raw.rename(columns={raw.columns[0]: "PostcodeArea", raw.columns[1]: "Service", raw.columns[2]: "Vendor"})

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

    st.session_state.setdefault("joda_pct", 0.0)
    st.session_state.setdefault("mcd_pct", 0.0)
    st.session_state.setdefault("pch_pct", 0.0)

    st.session_state.setdefault("ampm", False)
    st.session_state.setdefault("tail", False)
    st.session_state.setdefault("dual", False)
    st.session_state.setdefault("timed", False)
    st.session_state.setdefault("split1", 1)
    st.session_state.setdefault("split2", 1)

    st.session_state.setdefault("so_number", "")
    st.session_state.setdefault("export_basket", [])

    # McDowells portal export state + defaults
    st.session_state.setdefault("mcd_portal_rows", [])
    st.session_state.setdefault("mcd_consignor_name", "")
    st.session_state.setdefault("mcd_consignor_postcode", "")
    st.session_state.setdefault("mcd_consignor_account", "")
    st.session_state.setdefault("mcd_consignor_email", "")
    st.session_state.setdefault("mcd_entered_by", "")
    st.session_state.setdefault("mcd_cons_name", "")
    st.session_state.setdefault("mcd_cons_addr1", "")
    st.session_state.setdefault("mcd_cons_addr2", "")
    st.session_state.setdefault("mcd_cons_addr3", "")
    st.session_state.setdefault("mcd_cons_addr4", "")
    st.session_state.setdefault("mcd_cons_postcode", "")
    st.session_state.setdefault("mcd_cons_contact", "")
    st.session_state.setdefault("mcd_cons_tel", "")
    st.session_state.setdefault("mcd_weight_per_pallet", 0.0)  # optional; total = pallets*weight_per_pallet
    st.session_state.setdefault("mcd_remarks1", "")
    st.session_state.setdefault("mcd_remarks2", "")


_ensure_defaults()

if "surcharges_loaded" not in st.session_state:
    refresh_surcharges_from_disk()
    st.session_state["surcharges_loaded"] = True

# -------------------------
# Helpers
# -------------------------
def get_base_rate(df, area, service, vendor, pallets):
    subset = df[
        (df["PostcodeArea"] == area)
        & (df["Service"] == service)
        & (df["Vendor"] == vendor)
        & (df["Pallets"] == pallets)
    ]
    return None if subset.empty else float(subset["BaseRate"].iloc[0])


def get_max_pallets_for(df: pd.DataFrame, vendor: str) -> int:
    sub = df[df["Vendor"] == vendor]
    if sub.empty:
        return 26
    try:
        return int(sub["Pallets"].max())
    except Exception:
        return 26


def get_base_rate_capped(df: pd.DataFrame, area: str, service: str, vendor: str, pallets: int) -> Optional[float]:
    max_p = get_max_pallets_for(df, vendor)
    lookup_p = min(int(pallets), int(max_p))
    return get_base_rate(df, area, service, vendor, lookup_p)


def joda_round_base_up(x: float) -> float:
    return float(math.ceil(float(x)))


def joda_effective_pct(pallet_count: int, input_pct: float) -> float:
    return 0.0 if pallet_count < 7 else float(input_pct)


def mcd_smallload_extra(pallet_count: int) -> float:
    return (5.0 * min(pallet_count, 4)) if pallet_count < 5 else 0.0


def _ddmmyyyy(d: date) -> str:
    return d.strftime("%d/%m/%Y")


def _ddmmyyyy_compact(d: date) -> str:
    # McDowells portal example uses 15042026 style
    return d.strftime("%d%m%Y")


def _blank_sage_row() -> Dict[str, object]:
    return {c: "" for c in SAGE_EXPORT_COLUMNS}


def _blank_mcd_row() -> Dict[str, object]:
    return {c: "" for c in MCD_PORTAL_COLUMNS}


def _export_line_sage(
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

    r = _blank_sage_row()
    r["_row_id"] = uuid.uuid4().hex  # not exported

    r["Purchase Order Import Type"] = 1
    r["Purchase Order Number"] = int(po_number)
    r["Purchase Order Supplier Acc Code"] = str(supplier_acc).strip()
    r["Purchase Order Document Date"] = _ddmmyyyy(doc_date)
    r["Purchase Order Header Requested Date"] = _ddmmyyyy(req_date)
    r["Purchase Order Discount Percent"] = 0

    wh = st.session_state["warehouse_name"]
    r["Warehouse Name"] = wh
    if "Purchase Order Supplier Document No." in r:
        r["Purchase Order Supplier Document No."] = wh

    if "Purchase Order Line Requested Date" in r:
        r["Purchase Order Line Requested Date"] = _ddmmyyyy(req_date)

    # Tax Code always 1 (AK)
    if "Tax Code" in r:
        r["Tax Code"] = 1

    so_number = str(so_number).strip()
    so_suffix = f" - SO{so_number}" if so_number else ""
    svc_suffix = f" ({service})" if str(service).strip() else ""
    desc = f"{area_code} {label}{svc_suffix}{so_suffix}".strip()
    if "Free Text Item Description" in r:
        r["Free Text Item Description"] = desc

    if "Item Quantity" in r:
        r["Item Quantity"] = float(qty)
    if "Unit Buying Price" in r:
        r["Unit Buying Price"] = round(float(unit_price), 5)  # max 5 dp

    return r


def _add_to_sage_basket(rows: List[Dict[str, object]]):
    for r in rows:
        r["_row_id"] = r.get("_row_id") or uuid.uuid4().hex
    st.session_state["export_basket"].extend(rows)


def _add_to_mcd_portal(rows: List[Dict[str, object]]):
    for r in rows:
        r["_row_id"] = r.get("_row_id") or uuid.uuid4().hex
    st.session_state["mcd_portal_rows"].extend(rows)


def _clear_so_on_next_run():
    if st.session_state.pop("_clear_so_next", False):
        st.session_state["so_number"] = ""


def _mcd_delivery_time() -> str:
    # Keep simple: only set when relevant; otherwise blank
    if st.session_state.get("timed"):
        return "TIMED"
    if st.session_state.get("ampm"):
        return "AM"
    return ""


def build_mcd_portal_row() -> Dict[str, object]:
    """
    One row per McDowells consignment (not per charge line).
    Depots fixed 008; Service mapping Economy=2D, Next Day=ND.
    """
    so = str(st.session_state["so_number"]).strip()
    if not so:
        raise ValueError("SO Number is required before adding McDowells portal row.")

    pallets = int(st.session_state["pallets"])
    svc_ui = str(st.session_state["service"]).strip()
    svc_code = MCD_SERVICE_MAP.get(svc_ui, "")

    r = _blank_mcd_row()
    r["_row_id"] = uuid.uuid4().hex

    # Core refs
    if "Order_No" in r:
        r["Order_No"] = so
    if "Customer Reference" in r:
        r["Customer Reference"] = so

    # Dates
    if "Despatch Date" in r:
        r["Despatch Date"] = _ddmmyyyy_compact(date.today())

    # Depots
    if "Requesting Depot" in r:
        r["Requesting Depot"] = MCD_REQ_DEPOT
    if "Collect Depot" in r:
        r["Collect Depot"] = MCD_COLL_DEPOT
    if "Delivery Depot" in r:
        r["Delivery Depot"] = MCD_DEL_DEPOT

    # Service / time
    if "Service" in r:
        r["Service"] = svc_code
    if "Delivery Time" in r:
        r["Delivery Time"] = _mcd_delivery_time()

    # Pallets/weights (full pallets only)
    if "Full Pallets" in r:
        r["Full Pallets"] = pallets

    wpp = float(st.session_state.get("mcd_weight_per_pallet", 0.0) or 0.0)
    total_weight = (pallets * wpp) if wpp > 0 else ""
    if "Full Weight" in r and total_weight != "":
        r["Full Weight"] = round(float(total_weight), 3)

    # Consignor
    if "Consignor Name" in r:
        r["Consignor Name"] = str(st.session_state.get("mcd_consignor_name", "")).strip()
    if "ConsignorPostCode" in r:
        r["ConsignorPostCode"] = str(st.session_state.get("mcd_consignor_postcode", "")).strip()
    if "Consignor Account" in r:
        r["Consignor Account"] = str(st.session_state.get("mcd_consignor_account", "")).strip()
    if "Consignor Email" in r:
        r["Consignor Email"] = str(st.session_state.get("mcd_consignor_email", "")).strip()
    if "Entered By" in r:
        r["Entered By"] = str(st.session_state.get("mcd_entered_by", "")).strip()

    # Consignee
    if "Consignee Name" in r:
        r["Consignee Name"] = str(st.session_state.get("mcd_cons_name", "")).strip()
    if "Consignee Address 1" in r:
        r["Consignee Address 1"] = str(st.session_state.get("mcd_cons_addr1", "")).strip()
    if "Consignee Address 2" in r:
        r["Consignee Address 2"] = str(st.session_state.get("mcd_cons_addr2", "")).strip()
    if "Consignee Address 3" in r:
        r["Consignee Address 3"] = str(st.session_state.get("mcd_cons_addr3", "")).strip()
    if "Consignee Address 4" in r:
        r["Consignee Address 4"] = str(st.session_state.get("mcd_cons_addr4", "")).strip()
    if "Consignee Postcode" in r:
        r["Consignee Postcode"] = str(st.session_state.get("mcd_cons_postcode", "")).strip()
    if "Consignee Contact" in r:
        r["Consignee Contact"] = str(st.session_state.get("mcd_cons_contact", "")).strip()
    if "Consignee Tel" in r:
        r["Consignee Tel"] = str(st.session_state.get("mcd_cons_tel", "")).strip()

    # Remarks
    if "Remarks 1" in r:
        r["Remarks 1"] = str(st.session_state.get("mcd_remarks1", "")).strip()
    if "Remarks 2" in r:
        r["Remarks 2"] = str(st.session_state.get("mcd_remarks2", "")).strip()

    return r

# -------------------------
# Fuel surcharges (always visible)
# -------------------------
st.subheader("Fuel Surcharges")
cfs1, cfs2, cfs3 = st.columns(3, gap="medium")

with cfs1:
    st.number_input("Joda Fuel Surcharge (%)", 0.0, 100.0, step=0.1, format="%.2f", key="joda_pct")
    if st.button("Save Joda", use_container_width=True):
        save_joda_surcharge(float(st.session_state["joda_pct"]))
        st.success(f"Saved Joda at {float(st.session_state['joda_pct']):.2f}%")

with cfs2:
    st.number_input("McDowells Fuel Surcharge (%)", 0.0, 100.0, step=0.1, format="%.2f", key="mcd_pct")
    if st.button("Save McDowells", use_container_width=True):
        save_simple_surcharge(MCD_DATA_FILE, float(st.session_state["mcd_pct"]))
        st.success(f"Saved McDowells at {float(st.session_state['mcd_pct']):.2f}%")

with cfs3:
    st.number_input("PC Howard Fuel Surcharge (%)", 0.0, 100.0, step=0.1, format="%.2f", key="pch_pct")
    if st.button("Save PC Howard", use_container_width=True):
        save_simple_surcharge(PCH_DATA_FILE, float(st.session_state["pch_pct"]))
        st.success(f"Saved PC Howard at {float(st.session_state['pch_pct']):.2f}%")

st.markdown("---")

# -------------------------
# Input parameters
# -------------------------
st.header("1. Input Parameters")
col_a, col_b, col_c, col_d, col_h = st.columns([1, 1, 1, 1, 1], gap="medium")

with col_a:
    st.selectbox("Warehouse", options=WAREHOUSE_OPTIONS, key="warehouse_name")

allowed = set(available_hauliers())
pc_only = (allowed == {"Pc Howard"})

area_options = unique_areas_pch if pc_only else unique_areas_main
if st.session_state["area"] and st.session_state["area"] not in area_options:
    st.session_state["area"] = ""

with col_b:
    options_for_select = [""] + area_options
    st.selectbox(
        "Postcode Area",
        options=options_for_select,
        index=options_for_select.index(st.session_state["area"]) if st.session_state["area"] in options_for_select else 0,
        key="area",
        format_func=lambda x: x if x else "— Select area —",
    )
    if st.session_state["area"] == "":
        st.info("Please select a postcode area to continue.")
        st.stop()

with col_c:
    st.selectbox("Service Type", options=["Economy", "Next Day"], key="service")

with col_d:
    st.number_input("Number of Pallets", min_value=1, step=1, key="pallets")

with col_h:
    st.markdown("**Available hauliers**")
    st.write(", ".join(display_haulier(x) for x in sorted(allowed)) if allowed else "—")

st.markdown("---")

# Optional extras
st.subheader("2. Optional Extras")
col1, col2, col3, col4 = st.columns(4, gap="large")

if pc_only:
    st.session_state["dual"] = False
    st.session_state["tail"] = False

with col1:
    st.checkbox("AM/PM Delivery", key="ampm")
with col2:
    st.checkbox("Tail Lift", key="tail", disabled=pc_only)
with col3:
    st.checkbox("Dual Collection", key="dual", disabled=pc_only)
with col4:
    st.checkbox("Timed Delivery", key="timed")

if st.session_state["dual"] and st.session_state["pallets"] == 1:
    st.error("Dual Collection requires at least 2 pallets.")
    st.stop()

if st.session_state["dual"]:
    st.markdown("**Split pallets into two despatches.**")
    sp1, sp2 = st.columns(2, gap="large")
    with sp1:
        st.number_input("First Pallet Group", 1, max(int(st.session_state["pallets"]) - 1, 1), key="split1")
    with sp2:
        st.number_input("Second Pallet Group", 1, max(int(st.session_state["pallets"]) - 1, 1), key="split2")
    if int(st.session_state["split1"]) + int(st.session_state["split2"]) != int(st.session_state["pallets"]):
        st.error("Pallet Split values must add up to total pallets.")
        st.stop()

postcode_area = st.session_state["area"]

# -------------------------
# Core calculations (display only)
# -------------------------
def calc_for_area(area_code: str):
    svc = st.session_state["service"]
    allowed_local = set(available_hauliers())

    joda_charge_fixed = (7.5 if st.session_state["ampm"] else 0) + (20 if st.session_state["timed"] else 0)
    mcd_charge_fixed = (10 if st.session_state["ampm"] else 0) + (19 if st.session_state["timed"] else 0)
    pch_charge_fixed = (15.0 if st.session_state["ampm"] else 0) + (17.5 if st.session_state["timed"] else 0)

    jb = jf = None
    if "Joda" in allowed_local:
        if st.session_state["dual"]:
            b1 = get_base_rate_capped(rate_df_main, area_code, svc, "Joda", int(st.session_state["split1"]))
            b2 = get_base_rate_capped(rate_df_main, area_code, svc, "Joda", int(st.session_state["split2"]))
            if b1 is not None:
                b1 = float(math.ceil(float(b1)))
            if b2 is not None:
                b2 = float(math.ceil(float(b2)))
            if b1 is not None and b2 is not None:
                p1 = joda_effective_pct(int(st.session_state["split1"]), float(st.session_state["joda_pct"]))
                p2 = joda_effective_pct(int(st.session_state["split2"]), float(st.session_state["joda_pct"]))
                jf = (b1 * (1 + p1 / 100.0)) + (b2 * (1 + p2 / 100.0)) + (2 * joda_charge_fixed)
                jb = b1 + b2
        else:
            jb = get_base_rate_capped(rate_df_main, area_code, svc, "Joda", int(st.session_state["pallets"]))
            if jb is not None:
                jb = float(math.ceil(float(jb)))
                ep = joda_effective_pct(int(st.session_state["pallets"]), float(st.session_state["joda_pct"]))
                jf = jb * (1 + ep / 100.0) + joda_charge_fixed

    mb = mf = None
    if "Mcdowells" in allowed_local:
        mb = get_base_rate_capped(rate_df_main, area_code, svc, "Mcdowells", int(st.session_state["pallets"]))
        if mb is not None:
            n = int(st.session_state["pallets"])
            small_extra = mcd_smallload_extra(n)
            tl_total = (3.90 if st.session_state["tail"] else 0.0) * n
            mb_calc = float(mb) + small_extra
            mf = (mb_calc * (1 + float(st.session_state["mcd_pct"]) / 100.0)) + mcd_charge_fixed + tl_total

    pb = pf = None
    if "Pc Howard" in allowed_local and not rate_df_pch.empty:
        pb = get_base_rate_capped(rate_df_pch, area_code, svc, "Pc Howard", int(st.session_state["pallets"]))
        if pb is not None:
            pb_after_fuel = float(pb) * (1 + float(st.session_state["pch_pct"]) / 100.0)
            pf = pb_after_fuel + pch_charge_fixed

    return jb, jf, mb, mf, pb, pf


joda_base, joda_final, mcd_base, mcd_final, pch_base, pch_final = calc_for_area(postcode_area)

joda_charge_fixed = (7.5 if st.session_state["ampm"] else 0) + (20 if st.session_state["timed"] else 0)
mcd_charge_fixed = (10 if st.session_state["ampm"] else 0) + (19 if st.session_state["timed"] else 0)
mcd_tail_lift_total = (3.90 if st.session_state["tail"] else 0.0) * int(st.session_state["pallets"])
mcd_small_extra = mcd_smallload_extra(int(st.session_state["pallets"]))
pch_charge_fixed = (15.0 if st.session_state["ampm"] else 0) + (17.5 if st.session_state["timed"] else 0)

summary_rows = []
allowed_now = set(available_hauliers())

if "Joda" in allowed_now:
    if joda_base is None:
        summary_rows.append({"Haulier": "Joda", "Base Rate": "No rate", "Fuel Surcharge (%)": f"{float(st.session_state['joda_pct']):.2f}%", "Delivery Charge": "N/A", "Final Rate": "N/A"})
    else:
        shown_pct = joda_effective_pct(int(st.session_state["pallets"]), float(st.session_state["joda_pct"])) if not st.session_state["dual"] else float(st.session_state["joda_pct"])
        summary_rows.append({"Haulier": "Joda", "Base Rate": f"£{float(joda_base):,.2f}", "Fuel Surcharge (%)": f"{shown_pct:.2f}%", "Delivery Charge": f"£{joda_charge_fixed:,.2f}", "Final Rate": f"£{float(joda_final):,.2f}"})

if "Mcdowells" in allowed_now:
    if mcd_base is None:
        summary_rows.append({"Haulier": "McDowells", "Base Rate": "No rate", "Fuel Surcharge (%)": f"{float(st.session_state['mcd_pct']):.2f}%", "Delivery Charge": "N/A", "Final Rate": "N/A"})
    else:
        mcd_base_for_display = float(mcd_base) + mcd_small_extra
        summary_rows.append({"Haulier": "McDowells", "Base Rate": f"£{mcd_base_for_display:,.2f}", "Fuel Surcharge (%)": f"{float(st.session_state['mcd_pct']):.2f}%", "Delivery Charge": f"£{(mcd_charge_fixed + mcd_tail_lift_total):,.2f}", "Final Rate": f"£{float(mcd_final):,.2f}"})

if "Pc Howard" in allowed_now:
    if pch_base is None:
        summary_rows.append({"Haulier": "PC Howard", "Base Rate": "No rate", "Fuel Surcharge (%)": f"{float(st.session_state['pch_pct']):.2f}%", "Delivery Charge": "N/A", "Final Rate": "N/A"})
    else:
        summary_rows.append({"Haulier": "PC Howard", "Base Rate": f"£{float(pch_base):,.2f}", "Fuel Surcharge (%)": f"{float(st.session_state['pch_pct']):.2f}%", "Delivery Charge": f"£{pch_charge_fixed:,.2f}", "Final Rate": f"£{float(pch_final):,.2f}"})

summary_df = pd.DataFrame(summary_rows).set_index("Haulier") if summary_rows else pd.DataFrame()


def highlight_cheapest(row):
    fr = row.get("Final Rate", "")
    if isinstance(fr, str) and fr.startswith("£"):
        val = float(fr.strip("£").replace(",", ""))
        candidates = []
        if isinstance(joda_final, (int, float)):
            candidates.append(round(float(joda_final), 2))
        if isinstance(mcd_final, (int, float)):
            candidates.append(round(float(mcd_final), 2))
        if isinstance(pch_final, (int, float)):
            candidates.append(round(float(pch_final), 2))
        if candidates and math.isclose(round(val, 2), min(candidates), rel_tol=1e-9):
            return ["background-color: #b3e6b3"] * len(row)
    return [""] * len(row)

# -------------------------
# Sage export line builder
# -------------------------
def build_export_lines_for_haulier_sage(haulier: str) -> List[Dict[str, object]]:
    """
    Delivery line uses BASE price (no fuel).
    Fuel surcharge exports as its OWN line (qty=1, £=total).
    Pricing caps at max pallet band while qty can be > max.
    """
    so = str(st.session_state["so_number"]).strip()
    area = str(st.session_state["area"]).strip().upper()
    svc = str(st.session_state["service"]).strip()
    wh = str(st.session_state["warehouse_name"]).strip()
    n = int(st.session_state["pallets"])

    if not so:
        raise ValueError("SO Number is required before adding lines.")

    allowed_local = set(available_hauliers())
    h_norm = haulier.strip().title()
    if h_norm not in allowed_local:
        raise ValueError(f"{display_haulier(haulier)} is not available for warehouse {wh}.")

    out: List[Dict[str, object]] = []

    if h_norm == "Joda":
        po_no = po_number_for("Joda", wh)

        if st.session_state["dual"]:
            for group_n in [int(st.session_state["split1"]), int(st.session_state["split2"])]:
                base_n = get_base_rate_capped(rate_df_main, area, svc, "Joda", group_n)
                if base_n is None:
                    continue
                base_n = float(math.ceil(float(base_n)))

                unit_base = base_n / max(group_n, 1)
                out.append(_export_line_sage(po_no, JODA_ACC, so, area, svc, "Delivery", group_n, unit_base))

                eff = joda_effective_pct(group_n, float(st.session_state["joda_pct"]))
                fuel_total = base_n * (eff / 100.0)
                if fuel_total > 0:
                    out.append(_export_line_sage(po_no, JODA_ACC, so, area, svc, "Fuel Surcharge", 1, fuel_total))

                if st.session_state["ampm"]:
                    out.append(_export_line_sage(po_no, JODA_ACC, so, area, svc, "AM Charge", 1, 7.5))
                if st.session_state["timed"]:
                    out.append(_export_line_sage(po_no, JODA_ACC, so, area, svc, "Timed Charge", 1, 20.0))
        else:
            base = get_base_rate_capped(rate_df_main, area, svc, "Joda", n)
            if base is None:
                raise ValueError("No Joda rate available to add.")
            base = float(math.ceil(float(base)))

            unit_base = base / max(n, 1)
            out.append(_export_line_sage(po_no, JODA_ACC, so, area, svc, "Delivery", n, unit_base))

            eff = joda_effective_pct(n, float(st.session_state["joda_pct"]))
            fuel_total = base * (eff / 100.0)
            if fuel_total > 0:
                out.append(_export_line_sage(po_no, JODA_ACC, so, area, svc, "Fuel Surcharge", 1, fuel_total))

            if st.session_state["ampm"]:
                out.append(_export_line_sage(po_no, JODA_ACC, so, area, svc, "AM Charge", 1, 7.5))
            if st.session_state["timed"]:
                out.append(_export_line_sage(po_no, JODA_ACC, so, area, svc, "Timed Charge", 1, 20.0))

        return out

    if h_norm in ["Mcdowells", "Mcdowell", "Mcd"]:
        po_no = po_number_for("Mcdowells", wh)
        base = get_base_rate_capped(rate_df_main, area, svc, "Mcdowells", n)
        if base is None:
            raise ValueError("No McDowells rate available to add.")

        base_for_calc = float(base) + float(mcd_smallload_extra(n))

        unit_base = base_for_calc / max(n, 1)
        out.append(_export_line_sage(po_no, MCD_ACC, so, area, svc, "Delivery", n, unit_base))

        fuel_total = base_for_calc * (float(st.session_state["mcd_pct"]) / 100.0)
        if fuel_total > 0:
            out.append(_export_line_sage(po_no, MCD_ACC, so, area, svc, "Fuel Surcharge", 1, fuel_total))

        if st.session_state["ampm"]:
            out.append(_export_line_sage(po_no, MCD_ACC, so, area, svc, "AM Charge", 1, 10.0))
        if st.session_state["timed"]:
            out.append(_export_line_sage(po_no, MCD_ACC, so, area, svc, "Timed Charge", 1, 19.0))
        if st.session_state["tail"]:
            out.append(_export_line_sage(po_no, MCD_ACC, so, area, svc, "Tail Lift", n, 3.90))

        return out

    if h_norm == "Pc Howard":
        if rate_df_pch.empty:
            raise ValueError("PC Howard rate file missing. Place 'pch_rates_app.xlsx' alongside app.py.")

        po_no = po_number_for("Pc Howard", wh)
        base = get_base_rate_capped(rate_df_pch, area, svc, "Pc Howard", n)
        if base is None:
            raise ValueError("No PC Howard rate available to add.")

        unit_base = float(base) / max(n, 1)
        out.append(_export_line_sage(po_no, PCH_ACC, so, area, svc, "Delivery", n, unit_base))

        fuel_total = float(base) * (float(st.session_state["pch_pct"]) / 100.0)
        if fuel_total > 0:
            out.append(_export_line_sage(po_no, PCH_ACC, so, area, svc, "Fuel Surcharge", 1, fuel_total))

        if st.session_state["ampm"]:
            out.append(_export_line_sage(po_no, PCH_ACC, so, area, svc, "AM Charge", 1, 15.0))
        if st.session_state["timed"]:
            out.append(_export_line_sage(po_no, PCH_ACC, so, area, svc, "Timed Charge", 1, 17.5))

        return out

    raise ValueError(f"Unknown haulier: {haulier}")

# -------------------------
# Tabs
# -------------------------
st.header("3. Calculated Rates")
tab_table, tab_export, tab_mcd = st.tabs(["Table", "Export List", "McDowells Portal"])

with tab_table:
    if summary_df.empty or all(r.get("Final Rate") == "N/A" for r in summary_rows):
        st.warning("No rates found for that area/service/pallet combination.")
    else:
        st.table(summary_df.style.apply(highlight_cheapest, axis=1))

    st.markdown("---")
    st.subheader("Add to Export List")

    _clear_so_on_next_run()

    c_exp1, c_exp2 = st.columns([1, 1], gap="medium")
    with c_exp1:
        st.text_input("SO Number (manual)", key="so_number", placeholder="e.g. 020502")
    with c_exp2:
        st.write(f"Warehouse: **{st.session_state['warehouse_name']}**")

    buttons = st.columns([1, 1, 1, 2])

    if "Joda" in allowed_now:
        if buttons[0].button("Add Joda", use_container_width=True):
            try:
                _add_to_sage_basket(build_export_lines_for_haulier_sage("Joda"))
                st.session_state["_clear_so_next"] = True
                st.success("Added Joda lines to Export List.")
                st.rerun()
            except Exception as e:
                st.error(str(e))

    if "Mcdowells" in allowed_now:
        if buttons[1].button("Add McDowells", use_container_width=True):
            try:
                _add_to_sage_basket(build_export_lines_for_haulier_sage("Mcdowells"))

                # Also add McDowells portal row (one per consignment)
                _add_to_mcd_portal([build_mcd_portal_row()])

                st.session_state["_clear_so_next"] = True
                st.success("Added McDowells lines to Export List (+ McDowells portal row).")
                st.rerun()
            except Exception as e:
                st.error(str(e))

    if "Pc Howard" in allowed_now:
        if buttons[2].button("Add PC Howard", use_container_width=True):
            try:
                _add_to_sage_basket(build_export_lines_for_haulier_sage("Pc Howard"))
                st.session_state["_clear_so_next"] = True
                st.success("Added PC Howard lines to Export List.")
                st.rerun()
            except Exception as e:
                st.error(str(e))

with tab_export:
    st.subheader("Saved Lines (Sage PO import)")

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

            if cols[7].button("🗑", key=f"rm_sage_{rid}", help="Remove this line"):
                remove_id = rid

        if remove_id:
            st.session_state["export_basket"] = [x for x in st.session_state["export_basket"] if x.get("_row_id") != remove_id]
            st.rerun()

        st.markdown("---")
        c1, c2 = st.columns([1, 4])
        with c1:
            if st.button("Clear all (Sage)"):
                st.session_state["export_basket"] = []
                st.rerun()
        with c2:
            st.caption("Blank cells must be truly blank on export.")

        export_df = pd.DataFrame(st.session_state["export_basket"]).reindex(columns=SAGE_EXPORT_COLUMNS)
        export_df = export_df.where(pd.notnull(export_df), "")

        csv_bytes = export_df.to_csv(
            index=False,
            sep=",",
            na_rep="",
            lineterminator="\n",
            quoting=csvlib.QUOTE_MINIMAL,
        ).encode("utf-8")

        st.download_button(
            label="Download Sage PO Import File (.csv)",
            data=csv_bytes,
            file_name=f"PO_Import_Export_{date.today().strftime('%Y%m%d')}.csv",
            mime="text/csv",
            use_container_width=True,
        )

with tab_mcd:
    st.subheader("McDowells Portal Export")

    st.info("Depots fixed to 008. Service codes: Economy=2D, Next Day=ND.")

    st.markdown("### Consignor (Sender)")
    cc1, cc2, cc3, cc4 = st.columns(4)
    with cc1:
        st.text_input("Consignor Name", key="mcd_consignor_name")
    with cc2:
        st.text_input("Consignor Postcode", key="mcd_consignor_postcode")
    with cc3:
        st.text_input("Consignor Account", key="mcd_consignor_account")
    with cc4:
        st.text_input("Consignor Email", key="mcd_consignor_email")

    st.markdown("### Consignee (Receiver)")
    c1, c2, c3, c4 = st.columns(4)
    with c1:
        st.text_input("Consignee Name", key="mcd_cons_name")
    with c2:
        st.text_input("Address 1", key="mcd_cons_addr1")
    with c3:
        st.text_input("Address 2", key="mcd_cons_addr2")
    with c4:
        st.text_input("Consignee Postcode", key="mcd_cons_postcode")

    c5, c6, c7, c8 = st.columns(4)
    with c5:
        st.text_input("Address 3", key="mcd_cons_addr3")
    with c6:
        st.text_input("Address 4", key="mcd_cons_addr4")
    with c7:
        st.text_input("Consignee Contact", key="mcd_cons_contact")
    with c8:
        st.text_input("Consignee Tel", key="mcd_cons_tel")

    st.markdown("### Weights / Notes")
    w1, w2, w3 = st.columns([1, 1, 2])
    with w1:
        st.number_input("Weight per pallet (optional)", min_value=0.0, step=0.1, format="%.1f", key="mcd_weight_per_pallet")
    with w2:
        st.text_input("Entered By (optional)", key="mcd_entered_by")
    with w3:
        st.text_input("Remarks 1 (optional)", key="mcd_remarks1")
        st.text_input("Remarks 2 (optional)", key="mcd_remarks2")

    st.markdown("---")

    rows = st.session_state.get("mcd_portal_rows", [])
    if not rows:
        st.info("No McDowells portal rows yet. Add McDowells from the Table tab.")
    else:
        preview_cols = ["Order_No", "Despatch Date", "Service", "Full Pallets", "Full Weight", "Consignee Name", "Consignee Postcode"]
        df_prev = pd.DataFrame(rows).reindex(columns=[c for c in preview_cols if c in MCD_PORTAL_COLUMNS])
        st.dataframe(df_prev, use_container_width=True, hide_index=True)

        # Remove single row buttons (simple)
        st.markdown("#### Remove a row")
        rid_to_remove = None
        for r in rows:
            rid = r.get("_row_id", "")
            label = f"{r.get('Order_No','')} — {r.get('Consignee Postcode','')}"
            if st.button(f"Remove {label}", key=f"rm_mcd_{rid}"):
                rid_to_remove = rid
                break
        if rid_to_remove:
            st.session_state["mcd_portal_rows"] = [x for x in st.session_state["mcd_portal_rows"] if x.get("_row_id") != rid_to_remove]
            st.rerun()

        st.markdown("---")
        c1, c2 = st.columns([1, 3])
        with c1:
            if st.button("Clear all (McDowells Portal)"):
                st.session_state["mcd_portal_rows"] = []
                st.rerun()
        with c2:
            st.caption("Dates export as ddmmyyyy (no slashes), matching your example.")

        export_df = pd.DataFrame(st.session_state["mcd_portal_rows"]).reindex(columns=MCD_PORTAL_COLUMNS)
        export_df = export_df.where(pd.notnull(export_df), "")

        csv_bytes = export_df.to_csv(
            index=False,
            sep=",",
            na_rep="",
            lineterminator="\n",
            quoting=csvlib.QUOTE_MINIMAL,
        ).encode("utf-8")

        st.download_button(
            label="Download McDowells Portal CSV",
            data=csv_bytes,
            file_name=f"McDowells_Portal_{date.today().strftime('%Y%m%d')}.csv",
            mime="text/csv",
            use_container_width=True,
        )
