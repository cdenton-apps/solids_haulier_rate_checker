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

# -------------------------
# Streamlit config / style
# -------------------------
st.set_page_config(page_title="Solidus Haulier Rate Checker", layout="wide")

st.markdown(
    """
    <style>
      #MainMenu { visibility: hidden; }
      footer { visibility: hidden; }
      .small-help { color:#666; font-size:0.85rem; }
      .tight  { margin-top:-0.6rem; }
    </style>
    """,
    unsafe_allow_html=True
)

# -------------------------
# Header
# -------------------------
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
        V3.5.0  
        **Updates**
        - Address book delete buttons are now shown **next to each row**
        - Export screen cleaned up: **Sage PO CSV + McDowells Portal CSV** download from the **same tab** (using expanders)
        - customers.xlsx is the “central” address book; edits write back to it
        - Pallets can exceed 26; pricing caps at the maximum pallet band in each rate sheet
        - Fuel surcharges export as their own line (qty=1, unit £ = total surcharge)
        - Tax Code = 1 on every Sage export line
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
RATE_XLSX_PCH = "pch_rates_app.xlsx"       # PC Howard

TEMPLATE_SAGE_PATH = "PO Import Example File.csv"
TEMPLATE_MCD_PATH = "Reference.csv"        # McDowells portal template header

CUSTOMERS_XLSX = "customers.xlsx"
CUSTOMERS_SHEET = "Customers"

# Sage supplier account codes
JODA_ACC = "J040"
MCD_ACC = "M127"
PCH_ACC = "P031"

# Warehouses
WAREHOUSE_OPTIONS = ["101 - Skipton", "201 - Skipton 2", "102 - Corby"]
WAREHOUSE_HAULIERS = {
    "101 - Skipton": ["Joda", "Mcdowells"],
    "201 - Skipton 2": ["Joda", "Mcdowells"],
    "102 - Corby": ["Pc Howard"],  # internal casing
}

# Each unique (haulier, warehouse) must have unique PO number
PO_NUMBER_MAP = {
    ("Joda", "101 - Skipton"): 1,
    ("Joda", "201 - Skipton 2"): 2,
    ("Mcdowells", "101 - Skipton"): 3,
    ("Mcdowells", "201 - Skipton 2"): 4,
    ("Pc Howard", "102 - Corby"): 5,
}

# McDowells portal constants
MCD_REQ_DEPOT = "008"
MCD_COLL_DEPOT = "008"
MCD_DEL_DEPOT = "008"
MCD_SERVICE_MAP = {"Economy": "2D", "Next Day": "ND"}

# Customers.xlsx columns
CUSTOMER_COLS = [
    "ID",
    "CustomerCode",
    "CustomerName",
    "Address1",
    "Address2",
    "Address3",
    "Address4",
    "Postcode",
    "Contact",
    "Tel",
    "Email",
]

# -------------------------
# Small helpers
# -------------------------
def _norm(s: str) -> str:
    """Uppercase + trim + collapse whitespace."""
    return " ".join((s or "").upper().split())

def _ddmmyyyy(d: date) -> str:
    return d.strftime("%d/%m/%Y")

def _ddmmyyyy_compact(d: date) -> str:
    # McDowells portal example uses ddmmyyyy without slashes
    return d.strftime("%d%m%Y")

def display_haulier(name: str) -> str:
    n = str(name).strip()
    if n.lower() in {"pc howard", "pc", "pch", "p.c. howard"}:
        return "PC Howard"
    if n.lower() == "mcdowells":
        return "McDowells"
    if n.lower() == "joda":
        return "Joda"
    return n

def available_hauliers() -> List[str]:
    wh = st.session_state.get("warehouse_name", WAREHOUSE_OPTIONS[0])
    return WAREHOUSE_HAULIERS.get(wh, [])

def po_number_for(haulier: str, warehouse: str) -> int:
    key = (haulier.strip().title(), warehouse.strip())
    if key not in PO_NUMBER_MAP:
        raise KeyError(f"No PO number mapping for {key}")
    return int(PO_NUMBER_MAP[key])

def customers_to_label(row: pd.Series) -> str:
    code = str(row.get("CustomerCode", "")).strip()
    name = str(row.get("CustomerName", "")).strip()
    pc = str(row.get("Postcode", "")).strip()
    left = code or name or "Customer"
    if code and name:
        left = f"{code} — {name}"
    return f"{left} — {pc}".strip(" —")

# -------------------------
# Template column loading
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

DEFAULT_MCD_PORTAL_COLUMNS: List[str] = [
    "Docket", "Order_No", "Despatch Date", "Requesting Depot", "Collect Depot",
    "Consignor Name", "ConsignorPostCode", "Consignee Name",
    "Consignee Address 1", "Consignee Address 2", "Consignee Address 3", "Consignee Address 4",
    "Consignee Postcode", "Delivery Depot", "Trunk", "Service", "Delivery Time",
    "Half Pallets", "Half Weight", "Full Pallets", "Full Weight",
    "Half Oversize Pallets", "Half Oversize Weight", "Full Oversize Pallets", "Full Oversize Weight",
    "Remarks 1", "Remarks 2", "Delivery Date ", "Revenue", "Insure Value",
    "Manifest Date", "Quarter Pallets", "Quarter Weight", "Customer Own Paperwork",
    "Consignor Account", "Consignee Contact", "Consignee Tel", "Day Time Freight",
    "Insurance Charge", "Insured Name", "Insured Email", "Entered By",
    "OOG3 Pallets", "OOG3 Weight", "OOG4 Pallets", "OOG4 Weight",
    "Not Used 1", "Not Used 2", "Hazchem", "Customer Reference", "UN Number",
    "Hazchem Weight", "Consignor Email", "7.5t",
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
# customers.xlsx persistence (read/write)
# -------------------------
def _ensure_customers_file_exists():
    if os.path.exists(CUSTOMERS_XLSX):
        return
    df = pd.DataFrame(columns=CUSTOMER_COLS)
    with pd.ExcelWriter(CUSTOMERS_XLSX, engine="openpyxl") as w:
        df.to_excel(w, index=False, sheet_name=CUSTOMERS_SHEET)

def load_customers_df() -> pd.DataFrame:
    _ensure_customers_file_exists()
    df = pd.read_excel(CUSTOMERS_XLSX, sheet_name=CUSTOMERS_SHEET, dtype=str).fillna("")
    for c in CUSTOMER_COLS:
        if c not in df.columns:
            df[c] = ""
    df = df[CUSTOMER_COLS].copy()

    missing = df["ID"].astype(str).str.strip() == ""
    if missing.any():
        df.loc[missing, "ID"] = [uuid.uuid4().hex for _ in range(missing.sum())]
        save_customers_df(df)
    return df

def save_customers_df(df: pd.DataFrame) -> None:
    df = df.copy()
    for c in CUSTOMER_COLS:
        if c not in df.columns:
            df[c] = ""
    df = df[CUSTOMER_COLS].fillna("")
    with pd.ExcelWriter(CUSTOMERS_XLSX, engine="openpyxl") as w:
        df.to_excel(w, index=False, sheet_name=CUSTOMERS_SHEET)

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
    st.session_state.setdefault("mcd_portal_rows", [])

    # McDowells portal settings
    st.session_state.setdefault("mcd_consignor_name", "")
    st.session_state.setdefault("mcd_consignor_postcode", "")
    st.session_state.setdefault("mcd_consignor_account", "")
    st.session_state.setdefault("mcd_consignor_email", "")
    st.session_state.setdefault("mcd_entered_by", "")
    st.session_state.setdefault("mcd_weight_per_pallet", 0.0)
    st.session_state.setdefault("mcd_remarks1", "")
    st.session_state.setdefault("mcd_remarks2", "")

    # Customer select/search (McD portal row)
    st.session_state.setdefault("mcd_customer_search", "")
    st.session_state.setdefault("mcd_selected_customer_id", "")

_ensure_defaults()

if "surcharges_loaded" not in st.session_state:
    refresh_surcharges_from_disk()
    st.session_state["surcharges_loaded"] = True

# -------------------------
# Pricing helpers
# -------------------------
def get_base_rate(df, area, service, vendor, pallets) -> Optional[float]:
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

# -------------------------
# Sage export row builder
# -------------------------
def _blank_sage_row() -> Dict[str, object]:
    return {c: "" for c in SAGE_EXPORT_COLUMNS}

def _export_line_sage(
    po_number: int,
    supplier_acc: str,
    so_number: str,
    area_code: str,
    service: str,
    label: str,
    qty: float,
    unit_price: float,
) -> Dict[str, object]:
    r = _blank_sage_row()
    r["_row_id"] = uuid.uuid4().hex

    r["Purchase Order Import Type"] = 1
    r["Purchase Order Number"] = int(po_number)
    r["Purchase Order Supplier Acc Code"] = str(supplier_acc).strip()
    r["Purchase Order Document Date"] = _ddmmyyyy(date.today())
    r["Purchase Order Header Requested Date"] = _ddmmyyyy(date.today())
    r["Purchase Order Discount Percent"] = 0

    wh = st.session_state["warehouse_name"]
    r["Warehouse Name"] = wh
    if "Purchase Order Supplier Document No." in r:
        r["Purchase Order Supplier Document No."] = wh

    if "Purchase Order Line Requested Date" in r:
        r["Purchase Order Line Requested Date"] = _ddmmyyyy(date.today())

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
        r["Unit Buying Price"] = round(float(unit_price), 5)

    return r

def _add_to_sage_basket(rows: List[Dict[str, object]]):
    for r in rows:
        r["_row_id"] = r.get("_row_id") or uuid.uuid4().hex
    st.session_state["export_basket"].extend(rows)

def _clear_so_on_next_run():
    if st.session_state.pop("_clear_so_next", False):
        st.session_state["so_number"] = ""

# -------------------------
# McDowells portal row builder
# -------------------------
def _blank_mcd_row() -> Dict[str, object]:
    return {c: "" for c in MCD_PORTAL_COLUMNS}

def _mcd_delivery_time() -> str:
    if st.session_state.get("timed"):
        return "TIMED"
    if st.session_state.get("ampm"):
        return "AM"
    return ""

def build_mcd_portal_row(customer_row: pd.Series) -> Dict[str, object]:
    so = str(st.session_state["so_number"]).strip()
    if not so:
        raise ValueError("SO Number is required before adding McDowells portal row.")

    pallets = int(st.session_state["pallets"])
    svc_ui = str(st.session_state["service"]).strip()
    svc_code = MCD_SERVICE_MAP.get(svc_ui, "")

    r = _blank_mcd_row()
    r["_row_id"] = uuid.uuid4().hex
    r["_consignee_label"] = customers_to_label(customer_row)

    if "Order_No" in r:
        r["Order_No"] = so
    if "Customer Reference" in r:
        r["Customer Reference"] = so

    if "Despatch Date" in r:
        r["Despatch Date"] = _ddmmyyyy_compact(date.today())

    if "Requesting Depot" in r:
        r["Requesting Depot"] = MCD_REQ_DEPOT
    if "Collect Depot" in r:
        r["Collect Depot"] = MCD_COLL_DEPOT
    if "Delivery Depot" in r:
        r["Delivery Depot"] = MCD_DEL_DEPOT

    if "Service" in r:
        r["Service"] = svc_code
    if "Delivery Time" in r:
        r["Delivery Time"] = _mcd_delivery_time()

    if "Full Pallets" in r:
        r["Full Pallets"] = pallets

    wpp = float(st.session_state.get("mcd_weight_per_pallet", 0.0) or 0.0)
    if wpp > 0 and "Full Weight" in r:
        r["Full Weight"] = round(float(pallets * wpp), 3)

    # Consignor (optional)
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

    # Consignee from customers.xlsx
    if "Consignee Name" in r:
        r["Consignee Name"] = str(customer_row.get("CustomerName", "")).strip() or str(customer_row.get("CustomerCode", "")).strip()
    if "Consignee Address 1" in r:
        r["Consignee Address 1"] = str(customer_row.get("Address1", "")).strip()
    if "Consignee Address 2" in r:
        r["Consignee Address 2"] = str(customer_row.get("Address2", "")).strip()
    if "Consignee Address 3" in r:
        r["Consignee Address 3"] = str(customer_row.get("Address3", "")).strip()
    if "Consignee Address 4" in r:
        r["Consignee Address 4"] = str(customer_row.get("Address4", "")).strip()
    if "Consignee Postcode" in r:
        r["Consignee Postcode"] = str(customer_row.get("Postcode", "")).strip()
    if "Consignee Contact" in r:
        r["Consignee Contact"] = str(customer_row.get("Contact", "")).strip()
    if "Consignee Tel" in r:
        r["Consignee Tel"] = str(customer_row.get("Tel", "")).strip()

    if "Remarks 1" in r:
        r["Remarks 1"] = str(st.session_state.get("mcd_remarks1", "")).strip()
    if "Remarks 2" in r:
        r["Remarks 2"] = str(st.session_state.get("mcd_remarks2", "")).strip()

    return r

def _add_to_mcd_portal(rows: List[Dict[str, object]]):
    for r in rows:
        r["_row_id"] = r.get("_row_id") or uuid.uuid4().hex
    st.session_state["mcd_portal_rows"].extend(rows)

# -------------------------
# Core calculations (summary only)
# -------------------------
def calc_for_area(area_code: str):
    svc = st.session_state["service"]
    allowed_local = set(available_hauliers())
    n = int(st.session_state["pallets"])

    joda_charge_fixed = (7.5 if st.session_state["ampm"] else 0) + (20 if st.session_state["timed"] else 0)
    mcd_charge_fixed = (10 if st.session_state["ampm"] else 0) + (19 if st.session_state["timed"] else 0)
    pch_charge_fixed = (15.0 if st.session_state["ampm"] else 0) + (17.5 if st.session_state["timed"] else 0)

    # Joda
    jb = jf = None
    if "Joda" in allowed_local:
        base = get_base_rate_capped(rate_df_main, area_code, svc, "Joda", n)
        if base is not None:
            base = joda_round_base_up(base)
            eff = joda_effective_pct(n, float(st.session_state["joda_pct"]))
            jb = base
            jf = base * (1 + eff / 100.0) + joda_charge_fixed

    # McDowells
    mb = mf = None
    if "Mcdowells" in allowed_local:
        base = get_base_rate_capped(rate_df_main, area_code, svc, "Mcdowells", n)
        if base is not None:
            small_extra = mcd_smallload_extra(n)
            tl_total = (3.90 if st.session_state["tail"] else 0.0) * n
            base_calc = float(base) + float(small_extra)
            mb = base_calc
            mf = base_calc * (1 + float(st.session_state["mcd_pct"]) / 100.0) + mcd_charge_fixed + tl_total

    # PC Howard
    pb = pf = None
    if "Pc Howard" in allowed_local and not rate_df_pch.empty:
        base = get_base_rate_capped(rate_df_pch, area_code, svc, "Pc Howard", n)
        if base is not None:
            pb = float(base)
            pf = float(base) * (1 + float(st.session_state["pch_pct"]) / 100.0) + pch_charge_fixed

    return jb, jf, mb, mf, pb, pf

# -------------------------
# Sage export line builder (all hauliers)
# -------------------------
def build_export_lines_for_haulier_sage(haulier: str) -> List[Dict[str, object]]:
    """
    Delivery line uses BASE price (no fuel).
    Fuel surcharge exports as its own line (qty=1, £=total surcharge).
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

    # --- Joda ---
    if h_norm == "Joda":
        po_no = po_number_for("Joda", wh)
        base = get_base_rate_capped(rate_df_main, area, svc, "Joda", n)
        if base is None:
            raise ValueError("No Joda rate available to add.")
        base = joda_round_base_up(base)

        # Delivery (no fuel)
        out.append(_export_line_sage(po_no, JODA_ACC, so, area, svc, "Delivery", n, base / max(n, 1)))

        # Fuel line
        eff = joda_effective_pct(n, float(st.session_state["joda_pct"]))
        fuel_total = base * (eff / 100.0)
        if fuel_total > 0:
            out.append(_export_line_sage(po_no, JODA_ACC, so, area, svc, "Fuel Surcharge", 1, fuel_total))

        # Extras
        if st.session_state["ampm"]:
            out.append(_export_line_sage(po_no, JODA_ACC, so, area, svc, "AM Charge", 1, 7.5))
        if st.session_state["timed"]:
            out.append(_export_line_sage(po_no, JODA_ACC, so, area, svc, "Timed Charge", 1, 20.0))
        return out

    # --- McDowells ---
    if h_norm in ["Mcdowells", "Mcdowell", "Mcd"]:
        po_no = po_number_for("Mcdowells", wh)
        base = get_base_rate_capped(rate_df_main, area, svc, "Mcdowells", n)
        if base is None:
            raise ValueError("No McDowells rate available to add.")

        base_for_calc = float(base) + float(mcd_smallload_extra(n))

        # Delivery (no fuel)
        out.append(_export_line_sage(po_no, MCD_ACC, so, area, svc, "Delivery", n, base_for_calc / max(n, 1)))

        # Fuel line
        fuel_total = base_for_calc * (float(st.session_state["mcd_pct"]) / 100.0)
        if fuel_total > 0:
            out.append(_export_line_sage(po_no, MCD_ACC, so, area, svc, "Fuel Surcharge", 1, fuel_total))

        # Extras
        if st.session_state["ampm"]:
            out.append(_export_line_sage(po_no, MCD_ACC, so, area, svc, "AM Charge", 1, 10.0))
        if st.session_state["timed"]:
            out.append(_export_line_sage(po_no, MCD_ACC, so, area, svc, "Timed Charge", 1, 19.0))
        if st.session_state["tail"]:
            out.append(_export_line_sage(po_no, MCD_ACC, so, area, svc, "Tail Lift", n, 3.90))
        return out

    # --- PC Howard ---
    if h_norm == "Pc Howard":
        if rate_df_pch.empty:
            raise ValueError("PC Howard rate file missing. Place 'pch_rates_app.xlsx' alongside app.py.")
        po_no = po_number_for("Pc Howard", wh)
        base = get_base_rate_capped(rate_df_pch, area, svc, "Pc Howard", n)
        if base is None:
            raise ValueError("No PC Howard rate available to add.")

        base = float(base)

        # Delivery (no fuel)
        out.append(_export_line_sage(po_no, PCH_ACC, so, area, svc, "Delivery", n, base / max(n, 1)))

        # Fuel line
        fuel_total = base * (float(st.session_state["pch_pct"]) / 100.0)
        if fuel_total > 0:
            out.append(_export_line_sage(po_no, PCH_ACC, so, area, svc, "Fuel Surcharge", 1, fuel_total))

        # Extras (PC Howard: AM + Timed only)
        if st.session_state["ampm"]:
            out.append(_export_line_sage(po_no, PCH_ACC, so, area, svc, "AM Charge", 1, 15.0))
        if st.session_state["timed"]:
            out.append(_export_line_sage(po_no, PCH_ACC, so, area, svc, "Timed Charge", 1, 17.5))
        return out

    raise ValueError(f"Unknown haulier: {haulier}")

# =============================================================================
# UI
# =============================================================================

# Fuel surcharges (tidy)
with st.expander("Fuel Surcharges", expanded=False):
    cfs1, cfs2, cfs3 = st.columns(3, gap="medium")
    with cfs1:
        st.number_input("Joda (%)", 0.0, 100.0, step=0.1, format="%.2f", key="joda_pct")
        if st.button("Save Joda", use_container_width=True):
            save_joda_surcharge(float(st.session_state["joda_pct"]))
            st.success(f"Saved Joda at {float(st.session_state['joda_pct']):.2f}%")
    with cfs2:
        st.number_input("McDowells (%)", 0.0, 100.0, step=0.1, format="%.2f", key="mcd_pct")
        if st.button("Save McDowells", use_container_width=True):
            save_simple_surcharge(MCD_DATA_FILE, float(st.session_state["mcd_pct"]))
            st.success(f"Saved McDowells at {float(st.session_state['mcd_pct']):.2f}%")
    with cfs3:
        st.number_input("PC Howard (%)", 0.0, 100.0, step=0.1, format="%.2f", key="pch_pct")
        if st.button("Save PC Howard", use_container_width=True):
            save_simple_surcharge(PCH_DATA_FILE, float(st.session_state["pch_pct"]))
            st.success(f"Saved PC Howard at {float(st.session_state['pch_pct']):.2f}%")

st.markdown("---")

# Tabs
tab_table, tab_export, tab_mcd = st.tabs(["Table", "Export", "McDowells Address Book"])

# -------------------------
# TABLE TAB
# -------------------------
with tab_table:
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
        # Allow > 26; charges cap by sheet maximum band
        st.number_input("Number of Pallets", min_value=1, step=1, key="pallets")

    with col_h:
        st.markdown("**Available hauliers**")
        st.write(", ".join(display_haulier(x) for x in sorted(allowed)) if allowed else "—")

    st.markdown("---")

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

    if st.session_state["dual"] and int(st.session_state["pallets"]) == 1:
        st.error("Dual Collection requires at least 2 pallets.")
        st.stop()

    if st.session_state["dual"]:
        st.markdown("<div class='small-help'>Split pallets into two despatches.</div>", unsafe_allow_html=True)
        sp1, sp2 = st.columns(2, gap="large")
        with sp1:
            st.number_input("First Pallet Group", 1, max(int(st.session_state["pallets"]) - 1, 1), key="split1")
        with sp2:
            st.number_input("Second Pallet Group", 1, max(int(st.session_state["pallets"]) - 1, 1), key="split2")
        if int(st.session_state["split1"]) + int(st.session_state["split2"]) != int(st.session_state["pallets"]):
            st.error("Pallet Split values must add up to total pallets.")
            st.stop()

    st.markdown("---")

    # Summary
    jb, jf, mb, mf, pb, pf = calc_for_area(st.session_state["area"])

    summary_rows = []
    if "Joda" in allowed:
        summary_rows.append({
            "Haulier": "Joda",
            "Base Rate": "No rate" if jb is None else f"£{float(jb):,.2f}",
            "Fuel Surcharge (%)": f"{float(st.session_state['joda_pct']):.2f}%",
            "Final Rate": "N/A" if jf is None else f"£{float(jf):,.2f}",
        })
    if "Mcdowells" in allowed:
        summary_rows.append({
            "Haulier": "McDowells",
            "Base Rate": "No rate" if mb is None else f"£{float(mb):,.2f}",
            "Fuel Surcharge (%)": f"{float(st.session_state['mcd_pct']):.2f}%",
            "Final Rate": "N/A" if mf is None else f"£{float(mf):,.2f}",
        })
    if "Pc Howard" in allowed:
        summary_rows.append({
            "Haulier": "PC Howard",
            "Base Rate": "No rate" if pb is None else f"£{float(pb):,.2f}",
            "Fuel Surcharge (%)": f"{float(st.session_state['pch_pct']):.2f}%",
            "Final Rate": "N/A" if pf is None else f"£{float(pf):,.2f}",
        })

    if summary_rows:
        st.subheader("3. Calculated Rates")
        st.table(pd.DataFrame(summary_rows).set_index("Haulier"))

    st.markdown("---")
    st.subheader("Add to Export Lists")

    # SO clear mechanics (must happen before widget is created)
    _clear_so_on_next_run()

    top1, top2, top3 = st.columns([1, 1, 2], gap="medium")
    with top1:
        st.text_input("SO Number", key="so_number", placeholder="e.g. 020502")
    with top2:
        st.write(f"Warehouse: **{st.session_state['warehouse_name']}**")
    with top3:
        customers_df = load_customers_df()

        q = _norm(st.text_input("McDowells customer search", key="mcd_customer_search", placeholder="code / name / postcode…"))
        q_compact = q.replace(" ", "")

        blobs = (
            customers_df["CustomerCode"].astype(str).map(_norm) + " " +
            customers_df["CustomerName"].astype(str).map(_norm) + " " +
            customers_df["Postcode"].astype(str).map(_norm)
        )
        pc_compact = customers_df["Postcode"].astype(str).map(_norm).str.replace(" ", "", regex=False)

        if q:
            mask = blobs.str.contains(q, na=False) | pc_compact.str.contains(q_compact, na=False)
            filtered = customers_df[mask].copy()
        else:
            filtered = customers_df.copy()

        st.caption(f"Matches: {len(filtered):,}" + (" (showing first 200)" if len(filtered) > 200 else ""))
        filtered = filtered.head(200)

        options = [""] + filtered["ID"].tolist()
        label_map = {row["ID"]: customers_to_label(row) for _, row in filtered.iterrows()}

        st.selectbox(
            "Consignee (for McDowells portal row)",
            options=options,
            key="mcd_selected_customer_id",
            format_func=lambda x: "— Select —" if x == "" else label_map.get(x, x),
            disabled=("Mcdowells" not in allowed),
        )

    btns = st.columns([1, 1, 1, 2])
    if "Joda" in allowed:
        if btns[0].button("Add Joda", use_container_width=True):
            try:
                _add_to_sage_basket(build_export_lines_for_haulier_sage("Joda"))
                st.session_state["_clear_so_next"] = True
                st.success("Added Joda lines.")
                st.rerun()
            except Exception as e:
                st.error(str(e))

    if "Mcdowells" in allowed:
        if btns[1].button("Add McDowells", use_container_width=True):
            try:
                # Sage lines
                _add_to_sage_basket(build_export_lines_for_haulier_sage("Mcdowells"))

                # Portal row (one row per click) using selected customer
                cid = str(st.session_state.get("mcd_selected_customer_id", "")).strip()
                if not cid:
                    raise ValueError("Select a consignee (customer) for the McDowells portal row.")
                crow = customers_df.loc[customers_df["ID"] == cid]
                if crow.empty:
                    raise ValueError("Selected customer not found in customers.xlsx.")
                _add_to_mcd_portal([build_mcd_portal_row(crow.iloc[0])])

                st.session_state["_clear_so_next"] = True
                st.success("Added McDowells lines (+ portal row).")
                st.rerun()
            except Exception as e:
                st.error(str(e))

    if "Pc Howard" in allowed:
        if btns[2].button("Add PC Howard", use_container_width=True):
            try:
                _add_to_sage_basket(build_export_lines_for_haulier_sage("Pc Howard"))
                st.session_state["_clear_so_next"] = True
                st.success("Added PC Howard lines.")
                st.rerun()
            except Exception as e:
                st.error(str(e))

# -------------------------
# EXPORT TAB (ONE SCREEN, TWO EXPORTS)
# -------------------------
with tab_export:
    st.header("Exports")

    # ========== Sage PO export ==========
    with st.expander("Sage PO Export (PO Import CSV)", expanded=True):
        basket = st.session_state.get("export_basket", [])
        if not basket:
            st.info("No Sage PO lines saved yet. Use the Table tab to add lines.")
        else:
            # Preview + per-line delete
            h = st.columns([0.7, 1.2, 1.6, 4.8, 1.0, 1.2, 0.9])
            h[0].markdown("**PO**")
            h[1].markdown("**Supplier**")
            h[2].markdown("**Warehouse**")
            h[3].markdown("**Description**")
            h[4].markdown("**Qty**")
            h[5].markdown("**Unit £**")
            h[6].markdown("**Remove**")
            st.divider()

            remove_id = None
            for r in basket:
                rid = r.get("_row_id", "")
                cols = st.columns([0.7, 1.2, 1.6, 4.8, 1.0, 1.2, 0.9])
                cols[0].write(r.get("Purchase Order Number", ""))
                cols[1].write(r.get("Purchase Order Supplier Acc Code", ""))
                cols[2].write(r.get("Warehouse Name", ""))
                cols[3].write(r.get("Free Text Item Description", ""))
                cols[4].write(r.get("Item Quantity", ""))
                cols[5].write(r.get("Unit Buying Price", ""))
                if cols[6].button("🗑", key=f"rm_sage_{rid}", help="Remove this line"):
                    remove_id = rid

            if remove_id:
                st.session_state["export_basket"] = [x for x in st.session_state["export_basket"] if x.get("_row_id") != remove_id]
                st.rerun()

            st.markdown("---")
            c1, c2 = st.columns([1, 3])
            with c1:
                if st.button("Clear all Sage lines"):
                    st.session_state["export_basket"] = []
                    st.rerun()
            with c2:
                st.caption("Blank cells export as truly blank.")

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
                label="Download Sage PO Import CSV",
                data=csv_bytes,
                file_name=f"PO_Import_Export_{date.today().strftime('%Y%m%d')}.csv",
                mime="text/csv",
                use_container_width=True,
            )

    # ========== McDowells portal export ==========
    with st.expander("McDowells Portal Export (Portal CSV)", expanded=True):
        st.caption("Depots fixed to 008. Service codes: Economy=2D, Next Day=ND. One row per consignment click.")
        rows = st.session_state.get("mcd_portal_rows", [])
        if not rows:
            st.info("No McDowells portal rows saved yet. Use the Table tab → Add McDowells.")
        else:
            h = st.columns([1.6, 2.4, 1.2, 1.0, 0.9])
            h[0].markdown("**Order**")
            h[1].markdown("**Consignee**")
            h[2].markdown("**Postcode**")
            h[3].markdown("**Pallets**")
            h[4].markdown("**Remove**")
            st.divider()

            remove_id = None
            for r in rows:
                rid = r.get("_row_id", "")
                cols = st.columns([1.6, 2.4, 1.2, 1.0, 0.9])
                cols[0].write(r.get("Order_No", ""))
                cols[1].write(r.get("_consignee_label", "") or r.get("Consignee Name", ""))
                cols[2].write(r.get("Consignee Postcode", ""))
                cols[3].write(r.get("Full Pallets", ""))
                if cols[4].button("🗑", key=f"rm_mcd_{rid}", help="Remove this portal row"):
                    remove_id = rid

            if remove_id:
                st.session_state["mcd_portal_rows"] = [x for x in st.session_state["mcd_portal_rows"] if x.get("_row_id") != remove_id]
                st.rerun()

            st.markdown("---")
            c1, c2 = st.columns([1, 3])
            with c1:
                if st.button("Clear all McDowells portal rows"):
                    st.session_state["mcd_portal_rows"] = []
                    st.rerun()
            with c2:
                st.caption("Dates export as ddmmyyyy (no slashes), matching your portal example.")

            export_mcd_df = pd.DataFrame(st.session_state["mcd_portal_rows"]).reindex(columns=MCD_PORTAL_COLUMNS)
            export_mcd_df = export_mcd_df.where(pd.notnull(export_mcd_df), "")

            mcd_bytes = export_mcd_df.to_csv(
                index=False,
                sep=",",
                na_rep="",
                lineterminator="\n",
                quoting=csvlib.QUOTE_MINIMAL,
            ).encode("utf-8")

            st.download_button(
                label="Download McDowells Portal CSV",
                data=mcd_bytes,
                file_name=f"McDowells_Portal_{date.today().strftime('%Y%m%d')}.csv",
                mime="text/csv",
                use_container_width=True,
            )

# -------------------------
# McD Address Book tab (delete button per row)
# -------------------------
with tab_mcd:
    st.header("McDowells Address Book (customers.xlsx)")
    st.caption("Edits here write back to customers.xlsx (the central document for this deployment).")

    customers_df = load_customers_df()

    q = _norm(st.text_input("Search (code / name / postcode)", key="mcd_ab_search", placeholder="e.g. A0003 or BD7…"))
    q_compact = q.replace(" ", "")

    blobs = (
        customers_df["CustomerCode"].astype(str).map(_norm) + " " +
        customers_df["CustomerName"].astype(str).map(_norm) + " " +
        customers_df["Postcode"].astype(str).map(_norm)
    )
    pc_compact = customers_df["Postcode"].astype(str).map(_norm).str.replace(" ", "", regex=False)

    if q:
        mask = blobs.str.contains(q, na=False) | pc_compact.str.contains(q_compact, na=False)
        filtered = customers_df[mask].copy()
    else:
        filtered = customers_df.copy()

    st.caption(f"Matches: {len(filtered):,}" + (" (showing first 50)" if len(filtered) > 50 else ""))

    filtered = filtered.head(50)

    st.markdown("#### Results")
    h = st.columns([1.2, 2.8, 1.2, 0.9, 0.9])
    h[0].markdown("**Code**")
    h[1].markdown("**Name**")
    h[2].markdown("**Postcode**")
    h[3].markdown("**Edit**")
    h[4].markdown("**Delete**")
    st.divider()

    edit_id = None
    delete_id = None

    for _, r in filtered.iterrows():
        rid = str(r["ID"])
        cols = st.columns([1.2, 2.8, 1.2, 0.9, 0.9])
        cols[0].write(str(r.get("CustomerCode", "")).strip())
        cols[1].write(str(r.get("CustomerName", "")).strip())
        cols[2].write(str(r.get("Postcode", "")).strip())

        if cols[3].button("✏️", key=f"edit_c_{rid}", help="Edit this customer"):
            edit_id = rid
        if cols[4].button("🗑", key=f"del_c_{rid}", help="Delete this customer"):
            delete_id = rid

    if delete_id:
        customers_df = customers_df[customers_df["ID"] != delete_id].copy()
        save_customers_df(customers_df)
        # also clear selection if it was deleted
        if st.session_state.get("mcd_ab_selected_id") == delete_id:
            st.session_state["mcd_ab_selected_id"] = ""
        st.success("Deleted customer from customers.xlsx")
        st.rerun()

    if edit_id:
        st.session_state["mcd_ab_selected_id"] = edit_id
        st.rerun()

    st.markdown("---")
    st.markdown("#### Edit / Add customer")

    selected_id = st.session_state.get("mcd_ab_selected_id", "")
    selected_row = None
    if selected_id:
        match = customers_df.loc[customers_df["ID"] == selected_id]
        if not match.empty:
            selected_row = match.iloc[0]

    # Default values
    def _v(col: str) -> str:
        if selected_row is None:
            return ""
        return str(selected_row.get(col, "") or "")

    f1, f2, f3, f4 = st.columns(4)
    with f1:
        code = st.text_input("CustomerCode", value=_v("CustomerCode"), key="edit_code")
        name = st.text_input("CustomerName", value=_v("CustomerName"), key="edit_name")
    with f2:
        a1 = st.text_input("Address1", value=_v("Address1"), key="edit_a1")
        a2 = st.text_input("Address2", value=_v("Address2"), key="edit_a2")
    with f3:
        a3 = st.text_input("Address3", value=_v("Address3"), key="edit_a3")
        a4 = st.text_input("Address4", value=_v("Address4"), key="edit_a4")
    with f4:
        pc = st.text_input("Postcode", value=_v("Postcode"), key="edit_pc")
        contact = st.text_input("Contact", value=_v("Contact"), key="edit_contact")
        tel = st.text_input("Tel", value=_v("Tel"), key="edit_tel")
        email = st.text_input("Email", value=_v("Email"), key="edit_email")

    b1, b2, b3 = st.columns([1, 1, 2])
    with b1:
        if st.button("Save NEW", use_container_width=True):
            new = {
                "ID": uuid.uuid4().hex,
                "CustomerCode": str(code).strip(),
                "CustomerName": str(name).strip(),
                "Address1": str(a1).strip(),
                "Address2": str(a2).strip(),
                "Address3": str(a3).strip(),
                "Address4": str(a4).strip(),
                "Postcode": str(pc).strip().upper(),
                "Contact": str(contact).strip(),
                "Tel": str(tel).strip(),
                "Email": str(email).strip(),
            }
            customers_df = pd.concat([customers_df, pd.DataFrame([new])], ignore_index=True)
            save_customers_df(customers_df)
            st.success("Added to customers.xlsx")
            st.rerun()

    with b2:
        if st.button("Update SELECTED", use_container_width=True, disabled=(not selected_id)):
            customers_df.loc[customers_df["ID"] == selected_id, CUSTOMER_COLS] = [
                selected_id,
                str(code).strip(),
                str(name).strip(),
                str(a1).strip(),
                str(a2).strip(),
                str(a3).strip(),
                str(a4).strip(),
                str(pc).strip().upper(),
                str(contact).strip(),
                str(tel).strip(),
                str(email).strip(),
            ]
            save_customers_df(customers_df)
            st.success("Updated customers.xlsx")
            st.rerun()

    with b3:
        if st.button("Clear form", use_container_width=True):
            st.session_state["mcd_ab_selected_id"] = ""
            for k in ["edit_code","edit_name","edit_a1","edit_a2","edit_a3","edit_a4","edit_pc","edit_contact","edit_tel","edit_email"]:
                if k in st.session_state:
                    st.session_state[k] = ""
            st.rerun()
