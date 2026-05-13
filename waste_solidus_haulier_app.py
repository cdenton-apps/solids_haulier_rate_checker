# app.py
import os
import math
import json
import uuid
import csv as csvlib
from datetime import date
from typing import Optional, List, Dict, Any

import pandas as pd
import streamlit as st
from PIL import Image

st.set_page_config(page_title="Solidus Haulier Rate Checker", layout="wide")

st.markdown(
    """
    <style>
      #MainMenu { visibility: hidden; }
      footer { visibility: hidden; }
      .small-help { color: #666; font-size: 0.85rem; }
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
        V3.3.3    
        - McDowells portal: saved Consignee Address Book + choose consignee per line
        - Address Book search fixed (normalised + shows match count)
        - UI cleaned up: McDowells portal split into Address Book / Settings / Rows & Export
        - Pallets can exceed 26; pricing caps at max pallet band in sheet
        - Fuel surcharges export as their own line (qty=1, unit £ = total)
        - Tax Code = 1 on every Sage export line
        - SO Number auto-clears after successful add
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

MCD_CONSIGNEES_FILE = "mcd_consignees.json"
MCD_PREFILL_XLSX = "tmp2578.xlsx"          # customer codes + delivery postcode sheet (optional)

# Supplier account codes (Sage PO import)
JODA_ACC = "J040"
MCD_ACC = "M127"
PCH_ACC = "P031"

WAREHOUSE_OPTIONS = ["101 - Skipton", "201 - Skipton 2", "102 - Corby"]
WAREHOUSE_HAULIERS = {
    "101 - Skipton": ["Joda", "Mcdowells"],
    "201 - Skipton 2": ["Joda", "Mcdowells"],
    "102 - Corby": ["Pc Howard"],
}
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

# -------------------------
# Small utilities
# -------------------------
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

def _ddmmyyyy(d: date) -> str:
    return d.strftime("%d/%m/%Y")

def _ddmmyyyy_compact(d: date) -> str:
    return d.strftime("%d%m%Y")

def _safe_str(x: Any) -> str:
    return "" if x is None or (isinstance(x, float) and math.isnan(x)) else str(x)

def _norm(s: str) -> str:
    # uppercase + trim + collapse whitespace
    return " ".join((s or "").upper().split())

# -------------------------
# Load CSV headers (template-driven)
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
# McDowells Consignee Address Book (JSON + optional prefill from XLSX)
# -------------------------
def load_mcd_consignees() -> List[Dict[str, str]]:
    if not os.path.exists(MCD_CONSIGNEES_FILE):
        return []
    try:
        with open(MCD_CONSIGNEES_FILE, "r", encoding="utf-8") as f:
            data = json.load(f)
        if isinstance(data, list):
            out = []
            for x in data:
                if not isinstance(x, dict):
                    continue
                y = dict(x)
                y.setdefault("id", uuid.uuid4().hex)
                y.setdefault("label", y.get("customer_code", "") or y.get("name", "") or "Consignee")
                out.append(y)
            return out
        return []
    except Exception:
        return []

def save_mcd_consignees(items: List[Dict[str, str]]):
    try:
        with open(MCD_CONSIGNEES_FILE, "w", encoding="utf-8") as f:
            json.dump(items, f, indent=2)
    except Exception:
        pass

def _is_uk_postcode_like(s: str) -> bool:
    s = (s or "").strip().upper()
    if not s:
        return False
    import re
    return bool(re.match(r"^[A-Z]{1,2}\d", s))

@st.cache_data
def read_prefill_consignees_from_xlsx(xlsx_path: str, limit: int = 800) -> List[Dict[str, str]]:
    """
    tmp2578.xlsx layout (as provided):
      Row 1: column names
      Row 2+: data
    We treat:
      col0 = CustomerCode
      col1 = Addr3
      col2 = Addr4
      col3 = Addr5 (often postcode)
      col4 = Tel (optional)
      col6 = Email (optional)
    We pick postcode from Addr5 else Addr4 and keep UK-like postcodes only.
    """
    if not os.path.exists(xlsx_path):
        return []

    raw = pd.read_excel(xlsx_path, header=0)
    if raw.empty or len(raw) < 2:
        return []

    header = list(raw.iloc[0].tolist())
    df = raw.iloc[1:].copy()
    df.columns = header

    cols = list(df.columns)
    if len(cols) < 3:
        return []

    df = df.rename(columns={cols[0]: "CustomerCode", cols[1]: "Addr3", cols[2]: "Addr4"})
    if len(cols) > 3:
        df = df.rename(columns={cols[3]: "Addr5"})
    if len(cols) > 4:
        df = df.rename(columns={cols[4]: "Tel"})
    if len(cols) > 6:
        df = df.rename(columns={cols[6]: "Email"})

    df = df.dropna(subset=["CustomerCode"]).copy()
    df["CustomerCode"] = df["CustomerCode"].astype(str).str.strip()

    def pick_postcode(row) -> str:
        a5 = _safe_str(row.get("Addr5", "")).strip()
        a4 = _safe_str(row.get("Addr4", "")).strip()
        return a5 if a5 else a4

    df["Postcode"] = df.apply(pick_postcode, axis=1).astype(str).str.strip()
    df = df[df["Postcode"].apply(_is_uk_postcode_like)].copy()
    df = df.head(limit)

    items: List[Dict[str, str]] = []
    for _, r in df.iterrows():
        code = _safe_str(r.get("CustomerCode", "")).strip()
        pc = _safe_str(r.get("Postcode", "")).strip().upper()
        if not code or not pc:
            continue

        addr3 = _safe_str(r.get("Addr3", "")).strip()
        addr4 = _safe_str(r.get("Addr4", "")).strip()

        items.append({
            "id": uuid.uuid4().hex,
            "customer_code": code,
            "label": code,
            "name": code,
            "addr1": "",
            "addr2": "",
            "addr3": addr3,
            "addr4": addr4 if addr4 != pc else "",
            "postcode": pc,
            "contact": "",
            "tel": "",
        })
    return items

def merge_prefill(existing: List[Dict[str, str]], incoming: List[Dict[str, str]]) -> List[Dict[str, str]]:
    existing_keys = set()
    for e in existing:
        cc = (e.get("customer_code") or "").strip().upper()
        if cc:
            existing_keys.add(("CC", cc))
        else:
            existing_keys.add(("LP", (e.get("label","").strip().upper(), e.get("postcode","").strip().upper())))

    out = list(existing)
    for n in incoming:
        cc = (n.get("customer_code") or "").strip().upper()
        key = ("CC", cc) if cc else ("LP", (n.get("label","").strip().upper(), n.get("postcode","").strip().upper()))
        if key in existing_keys:
            continue
        out.append(n)
        existing_keys.add(key)
    return out

def get_consignee_by_id(cid: str) -> Optional[Dict[str, str]]:
    for c in st.session_state.get("mcd_consignees", []):
        if c.get("id") == cid:
            return c
    return None

def consignee_display(c: Dict[str, str]) -> str:
    code = (c.get("customer_code") or "").strip()
    label = (c.get("label") or "").strip()
    pc = (c.get("postcode") or "").strip()
    left = code or label or (c.get("name") or "").strip() or "Consignee"
    return f"{left} — {pc}".strip(" —")

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

    # McDowells portal settings (kept compact)
    st.session_state.setdefault("mcd_consignor_name", "")
    st.session_state.setdefault("mcd_consignor_postcode", "")
    st.session_state.setdefault("mcd_consignor_account", "")
    st.session_state.setdefault("mcd_consignor_email", "")
    st.session_state.setdefault("mcd_entered_by", "")
    st.session_state.setdefault("mcd_weight_per_pallet", 0.0)
    st.session_state.setdefault("mcd_remarks1", "")
    st.session_state.setdefault("mcd_remarks2", "")

    # Address book
    st.session_state.setdefault("mcd_consignees", load_mcd_consignees())
    st.session_state.setdefault("mcd_selected_consignee_id", "")
    st.session_state.setdefault("mcd_search", "")

_ensure_defaults()

if "surcharges_loaded" not in st.session_state:
    refresh_surcharges_from_disk()
    st.session_state["surcharges_loaded"] = True

# Prefill consignees once (if file exists) — only if user has none yet
if "consignees_prefilled" not in st.session_state:
    st.session_state["consignees_prefilled"] = True
    if len(st.session_state.get("mcd_consignees", [])) == 0 and os.path.exists(MCD_PREFILL_XLSX):
        pre = read_prefill_consignees_from_xlsx(MCD_PREFILL_XLSX, limit=800)
        if pre:
            st.session_state["mcd_consignees"] = merge_prefill([], pre)
            save_mcd_consignees(st.session_state["mcd_consignees"])

# -------------------------
# Pricing helpers
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

# -------------------------
# Export row builders (Sage)
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

def build_mcd_portal_row(consignee_id: str) -> Dict[str, object]:
    so = str(st.session_state["so_number"]).strip()
    if not so:
        raise ValueError("SO Number is required before adding McDowells portal row.")

    c = get_consignee_by_id(consignee_id)
    if not c:
        raise ValueError("Please select a saved consignee for the McDowells portal row.")

    pallets = int(st.session_state["pallets"])
    svc_ui = str(st.session_state["service"]).strip()
    svc_code = MCD_SERVICE_MAP.get(svc_ui, "")

    r = _blank_mcd_row()
    r["_row_id"] = uuid.uuid4().hex
    r["_consignee_label"] = consignee_display(c)

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

    # Consignor settings (optional)
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

    # Consignee from address book
    if "Consignee Name" in r:
        r["Consignee Name"] = (c.get("name") or c.get("label") or "").strip()
    if "Consignee Address 1" in r:
        r["Consignee Address 1"] = (c.get("addr1") or "").strip()
    if "Consignee Address 2" in r:
        r["Consignee Address 2"] = (c.get("addr2") or "").strip()
    if "Consignee Address 3" in r:
        r["Consignee Address 3"] = (c.get("addr3") or "").strip()
    if "Consignee Address 4" in r:
        r["Consignee Address 4"] = (c.get("addr4") or "").strip()
    if "Consignee Postcode" in r:
        r["Consignee Postcode"] = (c.get("postcode") or "").strip()
    if "Consignee Contact" in r:
        r["Consignee Contact"] = (c.get("contact") or "").strip()
    if "Consignee Tel" in r:
        r["Consignee Tel"] = (c.get("tel") or "").strip()

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
# Fuel surcharge UI (compact)
# -------------------------
with st.expander("Fuel Surcharges", expanded=True):
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

# -------------------------
# Optional extras (compact)
# -------------------------
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

postcode_area = st.session_state["area"]

# -------------------------
# Display calculations (summary)
# -------------------------
def calc_for_area(area_code: str):
    svc = st.session_state["service"]
    allowed_local = set(available_hauliers())

    joda_charge_fixed = (7.5 if st.session_state["ampm"] else 0) + (20 if st.session_state["timed"] else 0)
    mcd_charge_fixed = (10 if st.session_state["ampm"] else 0) + (19 if st.session_state["timed"] else 0)
    pch_charge_fixed = (15.0 if st.session_state["ampm"] else 0) + (17.5 if st.session_state["timed"] else 0)

    jb = jf = None
    if "Joda" in allowed_local:
        n = int(st.session_state["pallets"])
        jb = get_base_rate_capped(rate_df_main, area_code, svc, "Joda", n)
        if jb is not None:
            jb = joda_round_base_up(jb)
            ep = joda_effective_pct(n, float(st.session_state["joda_pct"]))
            jf = jb * (1 + ep / 100.0) + joda_charge_fixed

    mb = mf = None
    if "Mcdowells" in allowed_local:
        n = int(st.session_state["pallets"])
        mb = get_base_rate_capped(rate_df_main, area_code, svc, "Mcdowells", n)
        if mb is not None:
            small_extra = mcd_smallload_extra(n)
            tl_total = (3.90 if st.session_state["tail"] else 0.0) * n
            mb_calc = float(mb) + small_extra
            mf = (mb_calc * (1 + float(st.session_state["mcd_pct"]) / 100.0)) + mcd_charge_fixed + tl_total

    pb = pf = None
    if "Pc Howard" in allowed_local and not rate_df_pch.empty:
        n = int(st.session_state["pallets"])
        pb = get_base_rate_capped(rate_df_pch, area_code, svc, "Pc Howard", n)
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
        shown_pct = joda_effective_pct(int(st.session_state["pallets"]), float(st.session_state["joda_pct"]))
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
        if isinstance(joda_final, (int, float)): candidates.append(round(float(joda_final), 2))
        if isinstance(mcd_final, (int, float)): candidates.append(round(float(mcd_final), 2))
        if isinstance(pch_final, (int, float)): candidates.append(round(float(pch_final), 2))
        if candidates and math.isclose(round(val, 2), min(candidates), rel_tol=1e-9):
            return ["background-color: #b3e6b3"] * len(row)
    return [""] * len(row)

# -------------------------
# Sage export line builder (delivery base + fuel as own line; price capped)
# -------------------------
def build_export_lines_for_haulier_sage(haulier: str) -> List[Dict[str, object]]:
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
        base = get_base_rate_capped(rate_df_main, area, svc, "Joda", n)
        if base is None:
            raise ValueError("No Joda rate available to add.")
        base = joda_round_base_up(base)

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
    if summary_df.empty:
        st.warning("No rates found for that area/service/pallet combination.")
    else:
        st.table(summary_df.style.apply(highlight_cheapest, axis=1))

    st.markdown("---")
    st.subheader("Add to Export List")

    _clear_so_on_next_run()

    # Tight layout
    c1, c2, c3 = st.columns([1, 1, 2], gap="medium")
    with c1:
        st.text_input("SO Number", key="so_number", placeholder="e.g. 020502")
    with c2:
        st.write(f"Warehouse: **{st.session_state['warehouse_name']}**")
    with c3:
        consignees = st.session_state.get("mcd_consignees", [])
        search = st.text_input("McDowells consignee search", key="mcd_search", placeholder="type code or postcode…")
        q = _norm(search)
        q_compact = q.replace(" ", "")

        filtered = []
        for c in consignees:
            code = _norm(c.get("customer_code") or "")
            label = _norm(c.get("label") or "")
            name = _norm(c.get("name") or "")
            pc = _norm(c.get("postcode") or "")
            pc_compact = pc.replace(" ", "")
            blob = f"{code} {label} {name} {pc}"
            if (not q) or (q in blob) or (q_compact and q_compact in pc_compact):
                filtered.append(c)
            if len(filtered) >= 200:
                break

        options = [""] + [c.get("id", "") for c in filtered]
        label_map = {c.get("id", ""): consignee_display(c) for c in filtered}

        st.selectbox(
            "Consignee (McDowells portal row)",
            options=options,
            key="mcd_selected_consignee_id",
            format_func=lambda x: "— Select —" if x == "" else label_map.get(x, x),
            disabled=("Mcdowells" not in allowed_now),
        )

    buttons = st.columns([1, 1, 1, 2])

    if "Joda" in allowed_now:
        if buttons[0].button("Add Joda", use_container_width=True):
            try:
                _add_to_sage_basket(build_export_lines_for_haulier_sage("Joda"))
                st.session_state["_clear_so_next"] = True
                st.success("Added Joda lines.")
                st.rerun()
            except Exception as e:
                st.error(str(e))

    if "Mcdowells" in allowed_now:
        if buttons[1].button("Add McDowells", use_container_width=True):
            try:
                _add_to_sage_basket(build_export_lines_for_haulier_sage("Mcdowells"))
                cid = str(st.session_state.get("mcd_selected_consignee_id", "")).strip()
                _add_to_mcd_portal([build_mcd_portal_row(cid)])
                st.session_state["_clear_so_next"] = True
                st.success("Added McDowells lines (+ portal row).")
                st.rerun()
            except Exception as e:
                st.error(str(e))

    if "Pc Howard" in allowed_now:
        if buttons[2].button("Add PC Howard", use_container_width=True):
            try:
                _add_to_sage_basket(build_export_lines_for_haulier_sage("Pc Howard"))
                st.session_state["_clear_so_next"] = True
                st.success("Added PC Howard lines.")
                st.rerun()
            except Exception as e:
                st.error(str(e))

with tab_export:
    st.subheader("Saved Lines (Sage PO import)")
    basket = st.session_state.get("export_basket", [])
    if not basket:
        st.info("Nothing saved yet.")
    else:
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
            if cols[6].button("🗑", key=f"rm_sage_{rid}"):
                remove_id = rid

        if remove_id:
            st.session_state["export_basket"] = [x for x in st.session_state["export_basket"] if x.get("_row_id") != remove_id]
            st.rerun()

        st.markdown("---")
        c1, c2 = st.columns([1, 3])
        with c1:
            if st.button("Clear all (Sage)"):
                st.session_state["export_basket"] = []
                st.rerun()
        with c2:
            st.caption("Blanks export as truly blank cells.")

        export_df = pd.DataFrame(st.session_state["export_basket"]).reindex(columns=SAGE_EXPORT_COLUMNS)
        export_df = export_df.where(pd.notnull(export_df), "")
        csv_bytes = export_df.to_csv(index=False, sep=",", na_rep="", lineterminator="\n", quoting=csvlib.QUOTE_MINIMAL).encode("utf-8")

        st.download_button(
            label="Download Sage PO Import CSV",
            data=csv_bytes,
            file_name=f"PO_Import_Export_{date.today().strftime('%Y%m%d')}.csv",
            mime="text/csv",
            use_container_width=True,
        )

with tab_mcd:
    st.subheader("McDowells Portal Export")
    st.caption("Depots fixed to 008. Service codes: Economy=2D, Next Day=ND.")

    # Clean layout: split into sub-tabs
    t_ab, t_settings, t_export = st.tabs(["Address Book", "Settings", "Rows & Export"])

    with t_ab:
        st.markdown("#### Prefill from customer codes/postcodes")
        prefill_cols = st.columns([2, 1])
        with prefill_cols[0]:
            if os.path.exists(MCD_PREFILL_XLSX):
                st.caption(f"Found '{MCD_PREFILL_XLSX}'. You can import/top-up the address book.")
            else:
                st.caption(f"'{MCD_PREFILL_XLSX}' not found next to app.py (optional).")
        with prefill_cols[1]:
            if st.button("Import / top-up now", use_container_width=True, disabled=not os.path.exists(MCD_PREFILL_XLSX)):
                incoming = read_prefill_consignees_from_xlsx(MCD_PREFILL_XLSX, limit=800)
                st.session_state["mcd_consignees"] = merge_prefill(st.session_state.get("mcd_consignees", []), incoming)
                save_mcd_consignees(st.session_state["mcd_consignees"])
                st.success(f"Imported {len(incoming)} (deduped into address book).")
                st.rerun()

        st.markdown("---")
        st.markdown("#### Search & manage")

        consignees = st.session_state.get("mcd_consignees", [])
        q = st.text_input("Search (code / name / postcode)", key="mcd_ab_search", placeholder="e.g. A0003 or BD7…")

        q_norm = _norm(q)
        q_compact = q_norm.replace(" ", "")

        filtered = []
        for c in consignees:
            code = _norm(c.get("customer_code") or "")
            label = _norm(c.get("label") or "")
            name = _norm(c.get("name") or "")
            pc = _norm(c.get("postcode") or "")
            pc_compact = pc.replace(" ", "")

            blob = f"{code} {label} {name} {pc}"
            if (not q_norm) or (q_norm in blob) or (q_compact and q_compact in pc_compact):
                filtered.append(c)

        st.caption(
            f"Matches: {len(filtered):,}" + (" (showing first 200)" if len(filtered) > 200 else "")
        )

        filtered = filtered[:200]

        options = [""] + [c.get("id", "") for c in filtered]
        label_map = {c.get("id", ""): consignee_display(c) for c in filtered}

        selected_id = st.selectbox(
            "Select consignee to edit",
            options=options,
            key="mcd_ab_selected_id",
            format_func=lambda x: "— Select —" if x == "" else label_map.get(x, x),
        )

        # Form keys
        st.session_state.setdefault("mcd_ab_customer_code", "")
        st.session_state.setdefault("mcd_ab_label", "")
        st.session_state.setdefault("mcd_ab_name", "")
        st.session_state.setdefault("mcd_ab_addr1", "")
        st.session_state.setdefault("mcd_ab_addr2", "")
        st.session_state.setdefault("mcd_ab_addr3", "")
        st.session_state.setdefault("mcd_ab_addr4", "")
        st.session_state.setdefault("mcd_ab_postcode", "")
        st.session_state.setdefault("mcd_ab_contact", "")
        st.session_state.setdefault("mcd_ab_tel", "")

        current = get_consignee_by_id(selected_id) if selected_id else None

        load_cols = st.columns([1, 1, 1, 2])
        with load_cols[0]:
            if st.button("Load into form", use_container_width=True, disabled=(not current)):
                c = current
                st.session_state["mcd_ab_customer_code"] = c.get("customer_code", "")
                st.session_state["mcd_ab_label"] = c.get("label", "")
                st.session_state["mcd_ab_name"] = c.get("name", "")
                st.session_state["mcd_ab_addr1"] = c.get("addr1", "")
                st.session_state["mcd_ab_addr2"] = c.get("addr2", "")
                st.session_state["mcd_ab_addr3"] = c.get("addr3", "")
                st.session_state["mcd_ab_addr4"] = c.get("addr4", "")
                st.session_state["mcd_ab_postcode"] = c.get("postcode", "")
                st.session_state["mcd_ab_contact"] = c.get("contact", "")
                st.session_state["mcd_ab_tel"] = c.get("tel", "")
                st.rerun()

        with load_cols[1]:
            if st.button("New (clear)", use_container_width=True):
                st.session_state["mcd_ab_customer_code"] = ""
                st.session_state["mcd_ab_label"] = ""
                st.session_state["mcd_ab_name"] = ""
                st.session_state["mcd_ab_addr1"] = ""
                st.session_state["mcd_ab_addr2"] = ""
                st.session_state["mcd_ab_addr3"] = ""
                st.session_state["mcd_ab_addr4"] = ""
                area = (st.session_state.get("area", "") or "").strip().upper()
                st.session_state["mcd_ab_postcode"] = (area + " ") if area else ""
                st.session_state["mcd_ab_contact"] = ""
                st.session_state["mcd_ab_tel"] = ""
                st.rerun()

        with load_cols[2]:
            if st.button("Delete", use_container_width=True, disabled=(not current)):
                st.session_state["mcd_consignees"] = [x for x in consignees if x.get("id") != selected_id]
                save_mcd_consignees(st.session_state["mcd_consignees"])
                if st.session_state.get("mcd_selected_consignee_id") == selected_id:
                    st.session_state["mcd_selected_consignee_id"] = ""
                st.success("Deleted.")
                st.rerun()

        st.markdown("##### Consignee details")
        f1, f2, f3, f4 = st.columns(4)
        with f1:
            st.text_input("Customer Code (optional)", key="mcd_ab_customer_code")
            st.text_input("Label (display)", key="mcd_ab_label")
            st.text_input("Name", key="mcd_ab_name")
        with f2:
            st.text_input("Address 1", key="mcd_ab_addr1")
            st.text_input("Address 2", key="mcd_ab_addr2")
        with f3:
            st.text_input("Address 3", key="mcd_ab_addr3")
            st.text_input("Address 4", key="mcd_ab_addr4")
        with f4:
            st.text_input("Postcode", key="mcd_ab_postcode")
            st.text_input("Contact", key="mcd_ab_contact")
            st.text_input("Tel", key="mcd_ab_tel")

        save_cols = st.columns([1, 1, 2])
        with save_cols[0]:
            if st.button("Save NEW", use_container_width=True):
                new_item = {
                    "id": uuid.uuid4().hex,
                    "customer_code": str(st.session_state.get("mcd_ab_customer_code", "")).strip(),
                    "label": str(st.session_state.get("mcd_ab_label", "")).strip()
                             or str(st.session_state.get("mcd_ab_customer_code", "")).strip(),
                    "name": str(st.session_state.get("mcd_ab_name", "")).strip(),
                    "addr1": str(st.session_state.get("mcd_ab_addr1", "")).strip(),
                    "addr2": str(st.session_state.get("mcd_ab_addr2", "")).strip(),
                    "addr3": str(st.session_state.get("mcd_ab_addr3", "")).strip(),
                    "addr4": str(st.session_state.get("mcd_ab_addr4", "")).strip(),
                    "postcode": str(st.session_state.get("mcd_ab_postcode", "")).strip().upper(),
                    "contact": str(st.session_state.get("mcd_ab_contact", "")).strip(),
                    "tel": str(st.session_state.get("mcd_ab_tel", "")).strip(),
                }
                st.session_state["mcd_consignees"].append(new_item)
                save_mcd_consignees(st.session_state["mcd_consignees"])
                st.success("Saved new consignee.")
                st.rerun()

        with save_cols[1]:
            if st.button("Update SELECTED", use_container_width=True, disabled=(not current)):
                updated = []
                for x in st.session_state["mcd_consignees"]:
                    if x.get("id") != selected_id:
                        updated.append(x)
                        continue
                    y = dict(x)
                    y["customer_code"] = str(st.session_state.get("mcd_ab_customer_code", "")).strip()
                    y["label"] = str(st.session_state.get("mcd_ab_label", "")).strip() or y.get("customer_code", "")
                    y["name"] = str(st.session_state.get("mcd_ab_name", "")).strip()
                    y["addr1"] = str(st.session_state.get("mcd_ab_addr1", "")).strip()
                    y["addr2"] = str(st.session_state.get("mcd_ab_addr2", "")).strip()
                    y["addr3"] = str(st.session_state.get("mcd_ab_addr3", "")).strip()
                    y["addr4"] = str(st.session_state.get("mcd_ab_addr4", "")).strip()
                    y["postcode"] = str(st.session_state.get("mcd_ab_postcode", "")).strip().upper()
                    y["contact"] = str(st.session_state.get("mcd_ab_contact", "")).strip()
                    y["tel"] = str(st.session_state.get("mcd_ab_tel", "")).strip()
                    updated.append(y)
                st.session_state["mcd_consignees"] = updated
                save_mcd_consignees(updated)
                st.success("Updated.")
                st.rerun()

        with save_cols[2]:
            st.caption("Tip: keep Label short (e.g. customer code). Postcodes are normalised for searching.")

    with t_settings:
        st.markdown("#### Consignor / Weights / Notes")
        s1, s2, s3, s4 = st.columns(4)
        with s1:
            st.text_input("Consignor Name", key="mcd_consignor_name")
        with s2:
            st.text_input("Consignor Postcode", key="mcd_consignor_postcode")
        with s3:
            st.text_input("Consignor Account", key="mcd_consignor_account")
        with s4:
            st.text_input("Consignor Email", key="mcd_consignor_email")

        s5, s6, s7, s8 = st.columns(4)
        with s5:
            st.text_input("Entered By", key="mcd_entered_by")
        with s6:
            st.number_input("Weight per pallet (optional)", min_value=0.0, step=0.1, format="%.1f", key="mcd_weight_per_pallet")
        with s7:
            st.text_input("Remarks 1", key="mcd_remarks1")
        with s8:
            st.text_input("Remarks 2", key="mcd_remarks2")

        st.info("These settings are stamped onto each new McDowells portal row you add from the Table tab.")

    with t_export:
        rows = st.session_state.get("mcd_portal_rows", [])
        if not rows:
            st.info("No McDowells portal rows yet. Add McDowells from the Table tab.")
        else:
            st.markdown("#### Current rows")
            preview_cols = ["Order_No", "Despatch Date", "Service", "Full Pallets", "Full Weight", "_consignee_label", "Consignee Postcode"]
            df_prev = pd.DataFrame(rows).reindex(columns=[c for c in preview_cols if c in rows[0] or c in MCD_PORTAL_COLUMNS])
            st.dataframe(df_prev, use_container_width=True, hide_index=True)

            st.markdown("---")
            st.markdown("#### Remove a row")
            rid_to_remove = None
            for r in rows[:50]:
                rid = r.get("_row_id", "")
                label = f"{r.get('Order_No','')} — {r.get('_consignee_label','')}"
                if st.button(f"🗑 {label}", key=f"rm_mcd_{rid}"):
                    rid_to_remove = rid
                    break
            if len(rows) > 50:
                st.caption("Showing remove buttons for first 50 rows only (download will still include all).")

            if rid_to_remove:
                st.session_state["mcd_portal_rows"] = [x for x in st.session_state["mcd_portal_rows"] if x.get("_row_id") != rid_to_remove]
                st.rerun()

            st.markdown("---")
            c1, c2 = st.columns([1, 3])
            with c1:
                if st.button("Clear all (McDowells)"):
                    st.session_state["mcd_portal_rows"] = []
                    st.rerun()
            with c2:
                st.caption("Dates export as ddmmyyyy (no slashes), matching your example.")

            export_df = pd.DataFrame(st.session_state["mcd_portal_rows"]).reindex(columns=MCD_PORTAL_COLUMNS)
            export_df = export_df.where(pd.notnull(export_df), "")
            csv_bytes = export_df.to_csv(index=False, sep=",", na_rep="", lineterminator="\n", quoting=csvlib.QUOTE_MINIMAL).encode("utf-8")

            st.download_button(
                label="Download McDowells Portal CSV",
                data=csv_bytes,
                file_name=f"McDowells_Portal_{date.today().strftime('%Y%m%d')}.csv",
                mime="text/csv",
                use_container_width=True,
            )
