# waste_solidus_haulier_app.py
import os
import math
import json
import uuid
import csv as csvlib
from datetime import date
from typing import Optional, List, Dict, Tuple
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
      .muted { color:#666; font-size:0.9rem; }
    </style>
    """,
    unsafe_allow_html=True,
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
        unsafe_allow_html=True,
    )
    st.markdown(
        """
        V3.9.2  
        - Calculated Rates: Extras column restored
        - Export: Sage PO + Portal exports always show saved lines with per-line 🗑
        """,
        unsafe_allow_html=True,
    )

# -------------------------
# Config / constants
# -------------------------
JODA_DATA_FILE = "joda_surcharge.json"
MCD_DATA_FILE = "mcd_surcharge.json"
PCH_DATA_FILE = "pch_surcharge.json"
POREFS_FILE = "po_refs.json"
SO_DONE_FILE = "so_done.json"
RATE_XLSX_MAIN = "haulier prices 2.xlsx"   # Joda + McDowells
RATE_XLSX_PCH = "pch_rates_app.xlsx"       # PC Howard
TEMPLATE_SAGE_PATH = "PO Import Example File.csv"
TEMPLATE_PORTAL_PATH = "Reference.csv"     # ALL portal exports are based on this header for now.
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

# Default PO number mapping (these are the "true defaults")
PO_NUMBER_MAP = {
    ("Joda", "101 - Skipton"): 1,
    ("Joda", "201 - Skipton 2"): 2,
    ("Mcdowells", "101 - Skipton"): 3,
    ("Mcdowells", "201 - Skipton 2"): 4,
    ("Pc Howard", "102 - Corby"): 5,
}

# Portal constants (model on McDowells for now)
PORTAL_REQ_DEPOT = "008"
PORTAL_COLL_DEPOT = "008"
PORTAL_DEL_DEPOT = "008"
PORTAL_SERVICE_MAP = {"Economy": "2D", "Next Day": "ND"}  # apply to all portal exports for now

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
# Helpers
# -------------------------
def _norm(s: str) -> str:
    return " ".join((s or "").upper().split())

def _ddmmyyyy(d: date) -> str:
    return d.strftime("%d/%m/%Y")

def _ddmmyyyy_compact(d: date) -> str:
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

def _postcode_area(postcode: str) -> str:
    pc = _norm(str(postcode)).replace(" ", "")
    letters = ""
    for ch in pc:
        if ch.isalpha():
            letters += ch
        else:
            break
    return letters

def _safe_str(x) -> str:
    if x is None:
        return ""
    if isinstance(x, float) and pd.isna(x):
        return ""
    return str(x).strip()

def customer_label(row: pd.Series) -> str:
    code = str(row.get("CustomerCode", "")).strip()
    name = str(row.get("CustomerName", "")).strip()
    pc = str(row.get("Postcode", "")).strip()
    a1 = str(row.get("Address1", "")).strip()
    left = code or name or "Customer"
    if code and name:
        left = f"{code} — {name}"
    if a1:
        return f"{left} — {pc} — {a1}".strip(" —")
    return f"{left} — {pc}".strip(" —")

# ✅ NEW (only change): service rule based on pallets
def sync_service_from_pallets():
    try:
        n = int(st.session_state.get("pallets", 1))
    except Exception:
        n = 1
    st.session_state["service"] = "Next Day" if n > 6 else "Economy"

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
DEFAULT_PORTAL_COLUMNS: List[str] = [
    "Docket","Order_No","Despatch Date","Requesting Depot","Collect Depot",
    "Consignor Name","ConsignorPostCode","Consignee Name",
    "Consignee Address 1","Consignee Address 2","Consignee Address 3","Consignee Address 4",
    "Consignee Postcode","Delivery Depot","Trunk","Service","Delivery Time",
    "Half Pallets","Half Weight","Full Pallets","Full Weight",
    "Half Oversize Pallets","Half Oversize Weight","Full Oversize Pallets","Full Oversize Weight",
    "Remarks 1","Remarks 2","Delivery Date ","Revenue","Insure Value",
    "Manifest Date","Quarter Pallets","Quarter Weight","Customer Own Paperwork",
    "Consignor Account","Consignee Contact","Consignee Tel","Day Time Freight",
    "Insurance Charge","Insured Name","Insured Email","Entered By",
    "OOG3 Pallets","OOG3 Weight","OOG4 Pallets","OOG4 Weight",
    "Not Used 1","Not Used 2","Hazchem","Customer Reference","UN Number",
    "Hazchem Weight","Consignor Email","7.5t",
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
PORTAL_COLUMNS = load_csv_header_columns(TEMPLATE_PORTAL_PATH, DEFAULT_PORTAL_COLUMNS)

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
    st.session_state["mcd_pct"]  = round(load_simple_surcharge(MCD_DATA_FILE), 2)
    st.session_state["pch_pct"]  = round(load_simple_surcharge(PCH_DATA_FILE), 2)

# -------------------------
# Daily PO refs persistence
# -------------------------
def load_porefs_for_today() -> Dict[str, int]:
    today_str = date.today().isoformat()
    if not os.path.exists(POREFS_FILE):
        return {"date": today_str}
    try:
        with open(POREFS_FILE, "r") as f:
            data = json.load(f)
    except Exception:
        return {"date": today_str}
    if data.get("date") != today_str:
        return {"date": today_str}
    return data
def save_porefs_for_today(joda: int, mcd: int, pch: int) -> None:
    today_str = date.today().isoformat()
    payload = {"date": today_str, "joda": int(joda), "mcd": int(mcd), "pch": int(pch)}
    with open(POREFS_FILE, "w") as f:
        json.dump(payload, f)
def initialise_porefs_session_defaults():
    saved = load_porefs_for_today()
    st.session_state.setdefault("po_ref_joda", int(saved.get("joda", PO_NUMBER_MAP.get(("Joda", "101 - Skipton"), 1))))
    st.session_state.setdefault("po_ref_mcd", int(saved.get("mcd", PO_NUMBER_MAP.get(("Mcdowells", "101 - Skipton"), 1))))
    st.session_state.setdefault("po_ref_pch", int(saved.get("pch", PO_NUMBER_MAP.get(("Pc Howard", "102 - Corby"), 1))))

# -------------------------
# Daily SO "done" list
# -------------------------
def load_done_sos_for_today() -> Dict[str, object]:
    today_str = date.today().isoformat()
    if not os.path.exists(SO_DONE_FILE):
        return {"date": today_str, "done": []}
    try:
        with open(SO_DONE_FILE, "r") as f:
            data = json.load(f)
    except Exception:
        return {"date": today_str, "done": []}
    if data.get("date") != today_str:
        return {"date": today_str, "done": []}
    done = data.get("done", [])
    if not isinstance(done, list):
        done = []
    return {"date": today_str, "done": done}
def save_done_sos_for_today(done_list: List[str]) -> None:
    today_str = date.today().isoformat()
    payload = {"date": today_str, "done": list(dict.fromkeys([str(x) for x in done_list if str(x).strip()]))}
    with open(SO_DONE_FILE, "w") as f:
        json.dump(payload, f)
def mark_so_done(so_no: str) -> None:
    so_no = str(so_no).strip()
    if not so_no:
        return
    data = load_done_sos_for_today()
    done = data.get("done", [])
    if so_no not in done:
        done.append(so_no)
        save_done_sos_for_today(done)
    st.session_state["done_sos"] = done

# -------------------------
# customers.xlsx persistence
# -------------------------
def _ensure_customers_file_exists():
    if os.path.exists(CUSTOMERS_XLSX):
        return
    df = pd.DataFrame(columns=CUSTOMER_COLS)
    with pd.ExcelWriter(CUSTOMERS_XLSX, engine="openpyxl") as w:
        df.to_excel(w, index=False, sheet_name=CUSTOMERS_SHEET)
def save_customers_df(df: pd.DataFrame) -> None:
    df = df.copy()
    for c in CUSTOMER_COLS:
        if c not in df.columns:
            df[c] = ""
    df = df[CUSTOMER_COLS].fillna("")
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

# -------------------------
# Rates load (robust to Delivered Cost*)
# -------------------------
@st.cache_data
def load_rate_table(excel_path: str, _mtime: float) -> pd.DataFrame:
    preview = pd.read_excel(excel_path, sheet_name=0, header=None, nrows=10)
    header_row = None
    for i in range(len(preview)):
        row = preview.iloc[i].astype(str).str.strip().str.lower().tolist()
        if ("postcode" in row) and ("service" in row) and ("vendor" in row):
            header_row = i
            break
    if header_row is None:
        header_row = 1
    raw = pd.read_excel(excel_path, sheet_name=0, header=header_row)
    cols = list(raw.columns)
    if len(cols) < 3:
        raise ValueError(f"Rate sheet in {excel_path} does not have expected columns.")
    raw = raw.rename(columns={cols[0]: "PostcodeArea", cols[1]: "Service", cols[2]: "Vendor"})
    raw["PostcodeArea"] = raw["PostcodeArea"].ffill()
    raw["Service"] = raw["Service"].ffill()
    raw["Vendor"] = raw["Vendor"].ffill()
    if "Delivered Cost*" in raw.columns:
        pallet_start_col = "Delivered Cost*"
        after = raw.columns.tolist()[raw.columns.tolist().index(pallet_start_col):]
        raw = raw.rename(columns={c: i for i, c in enumerate(after, start=1)})
    pallet_cols = [c for c in raw.columns if isinstance(c, int) or (isinstance(c, str) and str(c).isdigit())]
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

# -------------------------
# Load rates
# -------------------------
mtime_main = os.path.getmtime(RATE_XLSX_MAIN)
rate_df_main = load_rate_table(RATE_XLSX_MAIN, mtime_main)
unique_areas_main = sorted(rate_df_main["PostcodeArea"].dropna().astype(str).unique())
rate_df_pch = pd.DataFrame(columns=["PostcodeArea", "Service", "Vendor", "Pallets", "BaseRate"])
unique_areas_pch: List[str] = []
if os.path.exists(RATE_XLSX_PCH):
    mtime_pch = os.path.getmtime(RATE_XLSX_PCH)
    rate_df_pch = load_rate_table(RATE_XLSX_PCH, mtime_pch)
    unique_areas_pch = sorted(rate_df_pch["PostcodeArea"].dropna().astype(str).unique())

# -------------------------
# Sage SO load helpers
# -------------------------
def load_sage_sales_export(uploaded_file) -> pd.DataFrame:
    raw = pd.read_excel(uploaded_file, header=None, dtype=str)
    if raw.shape[0] < 3:
        return pd.DataFrame()
    headers = [str(x).strip() for x in raw.iloc[1].tolist()]
    df = raw.iloc[2:].copy()
    df.columns = headers
    df = df.dropna(how="all").reset_index(drop=True)
    df.columns = [str(c).strip() for c in df.columns]
    return df
def col_pick(df: pd.DataFrame, candidates: List[str]) -> Optional[str]:
    cols = {str(c).strip().lower(): c for c in df.columns}
    for cand in candidates:
        k = cand.strip().lower()
        if k in cols:
            return cols[k]
    return None
def build_so_summary(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty:
        return pd.DataFrame()
    so_col = col_pick(df, ["SOPOrderReturns.DocumentNo", "DocumentNo"])
    pc_col = col_pick(df, ["SOPDocDelAddresses.PostCode", "PostCode", "Post Code"])
    cust_code_col = col_pick(df, ["SLCustomerAccounts.CustomerAccountNumber", "CustomerAccountNumber"])
    cust_name_col = col_pick(df, ["SLCustomerAccounts.CustomerAccountName", "CustomerAccountName", "Customer"])
    prom_col = col_pick(df, ["SOPOrderReturns.PromisedDeliveryDate", "PromisedDeliveryDate", "Promised Delivery Date"])
    qty_col = col_pick(df, ["SOPOrderReturnLines.LineQuantity", "LineQuantity", "Quantity"])
    ac18_col = col_pick(df, ["StockItems.AnalysisCode18", "AnalysisCode18"])
    if so_col is None:
        return pd.DataFrame()
    tmp = df.copy()
    tmp[so_col] = tmp[so_col].astype(str).str.strip()
    pallets_est = None
    if qty_col and ac18_col:
        q = pd.to_numeric(tmp[qty_col], errors="coerce")
        ac18 = pd.to_numeric(tmp[ac18_col], errors="coerce")
        denom = (ac18 / 1000.0).replace(0, pd.NA)
        pallets = (q / denom).replace([pd.NA, float("inf"), -float("inf")], pd.NA)
        pallets_est = pallets
    grp = tmp.groupby(so_col, dropna=False)
    def first_nonempty(s: pd.Series) -> str:
        s2 = s.dropna().astype(str).str.strip()
        s2 = s2[s2 != ""]
        return s2.iloc[0] if len(s2) else ""
    out = pd.DataFrame({
        "SO": grp[so_col].first().astype(str).str.strip(),
        "CustomerCode": grp[cust_code_col].apply(first_nonempty) if cust_code_col else "",
        "CustomerName": grp[cust_name_col].apply(first_nonempty) if cust_name_col else "",
        "Postcode": grp[pc_col].apply(first_nonempty) if pc_col else "",
        "PromisedDate": grp[prom_col].apply(first_nonempty) if prom_col else "",
    }).reset_index(drop=True)
    if pallets_est is not None:
        tmp2 = tmp[[so_col]].copy()
        tmp2["pallets_est"] = pallets_est
        pe = tmp2.groupby(so_col)["pallets_est"].sum(min_count=1)
        out = out.merge(pe.rename("PalletsEst"), left_on="SO", right_index=True, how="left")
    else:
        out["PalletsEst"] = pd.NA
    out["PostcodeArea"] = out["Postcode"].apply(_postcode_area)
    return out

def extract_consignee_from_so(df: pd.DataFrame, so_no: str) -> Dict[str, str]:
    if df.empty:
        return {}
    so_col = col_pick(df, ["SOPOrderReturns.DocumentNo", "DocumentNo"])
    if so_col is None:
        return {}
    postal_name_col = col_pick(df, ["SOPDocDelAddresses.PostalName", "PostalName", "Postal Name", "Delivery Name"])
    addr1_col = col_pick(df, ["SOPDocDelAddresses.AddressLine1", "AddressLine1", "Address Line 1"])
    addr2_col = col_pick(df, ["SOPDocDelAddresses.AddressLine2", "AddressLine2", "Address Line 2"])
    addr3_col = col_pick(df, ["SOPDocDelAddresses.AddressLine3", "AddressLine3", "Address Line 3"])
    addr4_col = col_pick(df, ["SOPDocDelAddresses.AddressLine4", "AddressLine4", "Address Line 4"])
    city_col = col_pick(df, ["SOPDocDelAddresses.City", "City", "Town"])
    county_col = col_pick(df, ["SOPDocDelAddresses.County", "County"])
    post_col = col_pick(df, ["SOPDocDelAddresses.PostCode", "PostCode", "Post Code"])
    contact_col = col_pick(df, ["SOPDocDelAddresses.Contact", "Contact"])
    tel_col = col_pick(df, ["SOPDocDelAddresses.TelephoneNo", "TelephoneNo", "Telephone No", "Telephone", "Phone"])
    email_col = col_pick(df, ["SOPDocDelAddresses.EmailAddress", "EmailAddress", "Email Address", "Email"])
    cust_code_col = col_pick(df, ["SLCustomerAccounts.CustomerAccountNumber", "CustomerAccountNumber"])
    cust_name_col = col_pick(df, ["SLCustomerAccounts.CustomerAccountName", "CustomerAccountName"])
    sub = df[df[so_col].astype(str).str.strip() == str(so_no)].copy()
    if sub.empty:
        return {}
    r0 = sub.iloc[0]
    def get(col):
        return _safe_str(r0[col]) if col and col in sub.columns else ""
    address4 = get(addr4_col)
    if not address4:
        parts = [get(city_col), get(county_col)]
        address4 = ", ".join([p for p in parts if p])
    out = {
        "CustomerCode": get(cust_code_col),
        "CustomerName": get(cust_name_col),
        "PostalName": get(postal_name_col),
        "Address1": get(addr1_col),
        "Address2": get(addr2_col),
        "Address3": get(addr3_col),
        "Address4": address4,
        "Postcode": get(post_col),
        "Contact": get(contact_col),
        "Tel": get(tel_col),
        "Email": get(email_col),
    }
    if not out["PostalName"]:
        out["PostalName"] = out["CustomerName"] or out["CustomerCode"]
    return out

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
    st.session_state.setdefault("so_number", "")
    st.session_state.setdefault("export_basket", [])
    st.session_state.setdefault("portal_rows_mcd", [])
    st.session_state.setdefault("portal_rows_joda", [])
    st.session_state.setdefault("portal_rows_pch", [])
    st.session_state.setdefault("portal_consignor_name", "")
    st.session_state.setdefault("portal_consignor_postcode", "")
    st.session_state.setdefault("portal_consignor_account", "")
    st.session_state.setdefault("portal_consignor_email", "")
    st.session_state.setdefault("portal_entered_by", "")
    st.session_state.setdefault("portal_weight_per_pallet", 0.0)
    st.session_state.setdefault("portal_remarks1", "")
    st.session_state.setdefault("portal_remarks2", "")
    st.session_state.setdefault("sage_so_uploaded", False)
    st.session_state.setdefault("sage_so_selected", "")
    st.session_state.setdefault("sage_so_search", "")
    st.session_state.setdefault("_last_so_applied", "")
    st.session_state.setdefault("_so_consignee", {})
    st.session_state.setdefault("cust_search", "")
    st.session_state.setdefault("cust_selected_id", "")
    st.session_state.setdefault("po_ref_joda", None)
    st.session_state.setdefault("po_ref_mcd", None)
    st.session_state.setdefault("po_ref_pch", None)
    st.session_state.setdefault("done_sos", [])
    st.session_state.setdefault("show_done_sos", False)

_ensure_defaults()
if "surcharges_loaded" not in st.session_state:
    refresh_surcharges_from_disk()
    st.session_state["surcharges_loaded"] = True
if "porefs_loaded" not in st.session_state:
    initialise_porefs_session_defaults()
    st.session_state["porefs_loaded"] = True
if "done_loaded" not in st.session_state:
    d = load_done_sos_for_today()
    st.session_state["done_sos"] = d.get("done", [])
    st.session_state["done_loaded"] = True

# -------------------------
# Pricing helpers
# -------------------------
def get_max_pallets_for(df: pd.DataFrame, vendor: str) -> int:
    sub = df[df["Vendor"] == vendor]
    if sub.empty:
        return 26
    try:
        return int(sub["Pallets"].max())
    except Exception:
        return 26

def get_base_rate(df, area, service, vendor, pallets) -> Optional[float]:
    subset = df[
        (df["PostcodeArea"] == area)
        & (df["Service"] == service)
        & (df["Vendor"] == vendor)
        & (df["Pallets"] == pallets)
    ]
    return None if subset.empty else float(subset["BaseRate"].iloc[0])

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

def extras_cost(haulier_key: str, pallets: int) -> float:
    """Used for the Calculated Rates Extras column (pure extras, not fuel)."""
    ampm = bool(st.session_state.get("ampm", False))
    timed = bool(st.session_state.get("timed", False))
    tail = bool(st.session_state.get("tail", False))
    hk = haulier_key.lower()
    if hk == "joda":
        return (7.5 if ampm else 0.0) + (20.0 if timed else 0.0)
    if hk == "mcdowells":
        return (10.0 if ampm else 0.0) + (19.0 if timed else 0.0) + ((3.90 * pallets) if tail else 0.0)
    if hk in {"pc howard", "pchodard", "pch", "pc"} or hk == "pc howard":
        return (15.0 if ampm else 0.0) + (17.5 if timed else 0.0)
    return 0.0

def calc_for_area(area_code: str):
    svc = st.session_state["service"]
    allowed_local = set(available_hauliers())
    n = int(st.session_state["pallets"])
    joda_fixed = (7.5 if st.session_state["ampm"] else 0) + (20 if st.session_state["timed"] else 0)
    mcd_fixed  = (10 if st.session_state["ampm"] else 0) + (19 if st.session_state["timed"] else 0)
    pch_fixed  = (15.0 if st.session_state["ampm"] else 0) + (17.5 if st.session_state["timed"] else 0)
    jb = jf = None
    if "Joda" in allowed_local:
        base = get_base_rate_capped(rate_df_main, area_code, svc, "Joda", n)
        if base is not None:
            base = joda_round_base_up(base)
            eff = joda_effective_pct(n, float(st.session_state["joda_pct"]))
            jb = base
            jf = base * (1 + eff / 100.0) + joda_fixed
    mb = mf = None
    if "Mcdowells" in allowed_local:
        base = get_base_rate_capped(rate_df_main, area_code, svc, "Mcdowells", n)
        if base is not None:
            small_extra = mcd_smallload_extra(n)
            tl_total = (3.90 if st.session_state["tail"] else 0.0) * n
            base_calc = float(base) + float(small_extra)
            mb = base_calc
            mf = base_calc * (1 + float(st.session_state["mcd_pct"]) / 100.0) + mcd_fixed + tl_total
    pb = pf = None
    if "Pc Howard" in allowed_local and not rate_df_pch.empty:
        base = get_base_rate_capped(rate_df_pch, area_code, svc, "Pc Howard", n)
        if base is not None:
            pb = float(base)
            pf = float(base) * (1 + float(st.session_state["pch_pct"]) / 100.0) + pch_fixed
    return jb, jf, mb, mf, pb, pf

def _parse_pounds(s: str) -> Optional[float]:
    if not isinstance(s, str) or not s.startswith("£"):
        return None
    try:
        return float(s.strip("£").replace(",", "").strip())
    except Exception:
        return None

def highlight_cheapest_factory():
    _, jf, _, mf, _, pf = calc_for_area(st.session_state["area"])
    candidates = []
    if isinstance(jf, (int, float)): candidates.append(round(float(jf), 2))
    if isinstance(mf, (int, float)): candidates.append(round(float(mf), 2))
    if isinstance(pf, (int, float)): candidates.append(round(float(pf), 2))
    cheapest = min(candidates) if candidates else None
    def _hl(row):
        v = _parse_pounds(row.get("Final Rate", ""))
        if cheapest is not None and v is not None and math.isclose(round(v, 2), cheapest, rel_tol=1e-9):
            return ["background-color: #b3e6b3"] * len(row)
        return [""] * len(row)
    return _hl

# -------------------------
# PO refs helper
# -------------------------
def _get_po_ref(haulier_title: str) -> int:
    if haulier_title == "Joda":
        return int(st.session_state.get("po_ref_joda") or PO_NUMBER_MAP.get(("Joda", "101 - Skipton"), 1))
    if haulier_title == "Mcdowells":
        return int(st.session_state.get("po_ref_mcd") or PO_NUMBER_MAP.get(("Mcdowells", "101 - Skipton"), 1))
    return int(st.session_state.get("po_ref_pch") or PO_NUMBER_MAP.get(("Pc Howard", "102 - Corby"), 1))

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
# Portal rows
# -------------------------
def _blank_portal_row() -> Dict[str, object]:
    return {c: "" for c in PORTAL_COLUMNS}

def _portal_delivery_time() -> str:
    if st.session_state.get("timed"):
        return "TIMED"
    if st.session_state.get("ampm"):
        return "AM"
    return ""

def _portal_service_code() -> str:
    return PORTAL_SERVICE_MAP.get(str(st.session_state.get("service", "")).strip(), "")

def build_portal_row(haulier_key: str, consignee: Dict[str, str]) -> Dict[str, object]:
    so = str(st.session_state["so_number"]).strip()
    pallets = int(st.session_state["pallets"])
    r = _blank_portal_row()
    r["_row_id"] = uuid.uuid4().hex
    r["_haulier"] = haulier_key
    label_bits = [consignee.get("CustomerCode", ""), consignee.get("CustomerName", ""), consignee.get("Postcode", ""), consignee.get("Address1", "")]
    r["_consignee_label"] = " — ".join([b for b in label_bits if b]).strip(" —")
    if "Order_No" in r:
        r["Order_No"] = so
    if "Customer Reference" in r:
        r["Customer Reference"] = so
    if "Despatch Date" in r:
        r["Despatch Date"] = _ddmmyyyy_compact(date.today())
    if "Requesting Depot" in r:
        r["Requesting Depot"] = PORTAL_REQ_DEPOT
    if "Collect Depot" in r:
        r["Collect Depot"] = PORTAL_COLL_DEPOT
    if "Delivery Depot" in r:
        r["Delivery Depot"] = PORTAL_DEL_DEPOT
    if "Service" in r:
        r["Service"] = _portal_service_code()
    if "Delivery Time" in r:
        r["Delivery Time"] = _portal_delivery_time()
    if "Full Pallets" in r:
        r["Full Pallets"] = pallets
    wpp = float(st.session_state.get("portal_weight_per_pallet", 0.0) or 0.0)
    if wpp > 0 and "Full Weight" in r:
        r["Full Weight"] = round(float(pallets * wpp), 3)
    if "Consignor Name" in r:
        r["Consignor Name"] = str(st.session_state.get("portal_consignor_name", "")).strip()
    if "ConsignorPostCode" in r:
        r["ConsignorPostCode"] = str(st.session_state.get("portal_consignor_postcode", "")).strip()
    if "Consignor Account" in r:
        r["Consignor Account"] = str(st.session_state.get("portal_consignor_account", "")).strip()
    if "Consignor Email" in r:
        r["Consignor Email"] = str(st.session_state.get("portal_consignor_email", "")).strip()
    if "Entered By" in r:
        r["Entered By"] = str(st.session_state.get("portal_entered_by", "")).strip()
    if "Consignee Name" in r:
        r["Consignee Name"] = consignee.get("PostalName") or consignee.get("CustomerName") or consignee.get("CustomerCode")
    if "Consignee Address 1" in r:
        r["Consignee Address 1"] = consignee.get("Address1", "")
    if "Consignee Address 2" in r:
        r["Consignee Address 2"] = consignee.get("Address2", "")
    if "Consignee Address 3" in r:
        r["Consignee Address 3"] = consignee.get("Address3", "")
    if "Consignee Address 4" in r:
        r["Consignee Address 4"] = consignee.get("Address4", "")
    if "Consignee Postcode" in r:
        r["Consignee Postcode"] = consignee.get("Postcode", "")
    if "Consignee Contact" in r:
        r["Consignee Contact"] = consignee.get("Contact", "")
    if "Consignee Tel" in r:
        r["Consignee Tel"] = consignee.get("Tel", "")
    if "Remarks 1" in r:
        r["Remarks 1"] = str(st.session_state.get("portal_remarks1", "")).strip()
    if "Remarks 2" in r:
        r["Remarks 2"] = str(st.session_state.get("portal_remarks2", "")).strip()
    return r

def _add_portal_row(haulier_key: str, row: Dict[str, object]):
    hk = haulier_key.lower().strip()
    if hk == "mcdowells":
        st.session_state["portal_rows_mcd"].append(row)
    elif hk == "joda":
        st.session_state["portal_rows_joda"].append(row)
    else:
        st.session_state["portal_rows_pch"].append(row)

# -------------------------
# Sage export line builder (all hauliers)
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
        po_no = _get_po_ref("Joda")
        base = get_base_rate_capped(rate_df_main, area, svc, "Joda", n)
        if base is None:
            raise ValueError("No Joda rate available to add.")
        base = joda_round_base_up(base)
        out.append(_export_line_sage(po_no, JODA_ACC, so, area, svc, "Delivery", n, base / max(n, 1)))
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
        po_no = _get_po_ref("Mcdowells")
        base = get_base_rate_capped(rate_df_main, area, svc, "Mcdowells", n)
        if base is None:
            raise ValueError("No McDowells rate available to add.")
        base_for_calc = float(base) + float(mcd_smallload_extra(n))
        out.append(_export_line_sage(po_no, MCD_ACC, so, area, svc, "Delivery", n, base_for_calc / max(n, 1)))
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
        po_no = _get_po_ref("Pc Howard")
        base = get_base_rate_capped(rate_df_pch, area, svc, "Pc Howard", n)
        if base is None:
            raise ValueError("No PC Howard rate available to add.")
        base = float(base)
        out.append(_export_line_sage(po_no, PCH_ACC, so, area, svc, "Delivery", n, base / max(n, 1)))
        fuel_total = base * (float(st.session_state["pch_pct"]) / 100.0)
        if fuel_total > 0:
            out.append(_export_line_sage(po_no, PCH_ACC, so, area, svc, "Fuel Surcharge", 1, fuel_total))
        if st.session_state["ampm"]:
            out.append(_export_line_sage(po_no, PCH_ACC, so, area, svc, "AM Charge", 1, 15.0))
        if st.session_state["timed"]:
            out.append(_export_line_sage(po_no, PCH_ACC, so, area, svc, "Timed Charge", 1, 17.5))
        return out
    raise ValueError(f"Unknown haulier: {haulier}")

# =============================================================================
# UI
# =============================================================================
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
tab_table, tab_export, tab_customers = st.tabs(["Table", "Export", "Customers"])

# -------------------------
# TABLE TAB
# -------------------------
with tab_table:
    st.header("Sales Orders (preferred)")
    wh_now = st.session_state.get("warehouse_name", WAREHOUSE_OPTIONS[0])
    allowed_now = set(WAREHOUSE_HAULIERS.get(wh_now, []))
    pc_only_now = allowed_now == {"Pc Howard"}
    area_options_now = unique_areas_pch if pc_only_now else unique_areas_main

    upl = st.file_uploader(
        "Upload Sage Sales Order export (.xlsx)",
        type=["xlsx"],
        help="Upload your standard Sage export (with Title row + header row).",
        key="sage_so_file",
    )

    so_summary = pd.DataFrame()
    so_df_full = pd.DataFrame()
    if upl is not None:
        try:
            so_df_full = load_sage_sales_export(upl)
            so_summary = build_so_summary(so_df_full)
            st.session_state["sage_so_uploaded"] = True
        except Exception as e:
            st.error(f"Could not read upload: {e}")

    done_sos = set(st.session_state.get("done_sos", []))
    show_done = st.checkbox("Show completed SOs", key="show_done_sos", value=False)

    if not so_summary.empty:
        so_search = st.text_input("Search SOs", key="sage_so_search", placeholder="SO / customer / postcode")
        ss = _norm(so_search)

        shown = so_summary.copy()
        if not show_done and done_sos:
            shown = shown[~shown["SO"].astype(str).isin(done_sos)].copy()

        if ss:
            mask = (
                shown["SO"].astype(str).str.upper().str.contains(ss, na=False)
                | shown["CustomerName"].astype(str).str.upper().str.contains(ss, na=False)
                | shown["CustomerCode"].astype(str).str.upper().str.contains(ss, na=False)
                | shown["Postcode"].astype(str).str.upper().str.contains(ss, na=False)
            )
            shown = shown[mask].copy()

        shown = shown.head(200)
        so_options = [""] + shown["SO"].astype(str).tolist()

        def _so_fmt(x: str) -> str:
            if not x:
                return "— Select SO —"
            row = shown.loc[shown["SO"].astype(str) == str(x)]
            if row.empty:
                return str(x)
            r0 = row.iloc[0]
            pc = str(r0.get("Postcode", "")).strip()
            nm = str(r0.get("CustomerName", "")).strip()
            dt = str(r0.get("PromisedDate", "")).strip()
            pe = r0.get("PalletsEst", "")
            pe_s = ""
            try:
                if pd.notna(pe):
                    pe_s = f" — est pallets {float(pe):.1f}"
            except Exception:
                pe_s = ""
            return f"{x} — {nm} — {pc} — {dt}{pe_s}".strip()

        picked = st.selectbox("Select Sales Order", options=so_options, key="sage_so_selected", format_func=_so_fmt)

        if picked and picked != st.session_state.get("_last_so_applied", ""):
            r0 = so_summary.loc[so_summary["SO"].astype(str) == str(picked)].iloc[0]
            pre_area = str(r0.get("PostcodeArea", "")).strip().upper()
            if pre_area and pre_area in area_options_now:
                st.session_state["area"] = pre_area

            st.session_state["so_number"] = str(picked)

            try:
                pe = r0.get("PalletsEst", pd.NA)
                if pd.notna(pe):
                    st.session_state["pallets"] = max(1, int(math.ceil(float(pe))))
                    # ✅ NEW (only change): apply service rule when pallets are set from SO
                    sync_service_from_pallets()
            except Exception:
                pass

            st.session_state["_so_consignee"] = extract_consignee_from_so(so_df_full, str(picked))
            st.session_state["_last_so_applied"] = str(picked)

    st.markdown("---")

    use_so = bool(st.session_state.get("sage_so_selected"))
    unlock_overrides = st.checkbox(
        "Unlock manual overrides",
        value=False,
        help="If an SO is selected, inputs are prefilled. Tick to override manually.",
        disabled=not use_so,
    )
    inputs_disabled = use_so and not unlock_overrides

    st.header("1. Input Parameters")
    col_a, col_b, col_c, col_d, col_h = st.columns([1, 1, 1, 1, 1], gap="medium")

    with col_a:
        st.selectbox("Warehouse", options=WAREHOUSE_OPTIONS, key="warehouse_name", disabled=inputs_disabled)

    allowed = set(available_hauliers())
    pc_only = allowed == {"Pc Howard"}
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
            disabled=inputs_disabled,
        )

    if st.session_state["area"] == "" and not use_so:
        st.info("Please select a postcode area (or upload/select an SO above).")
        st.stop()

    with col_c:
        st.selectbox("Service Type", options=["Economy", "Next Day"], key="service", disabled=inputs_disabled)

    with col_d:
        # ✅ NEW (only change): apply service rule when pallets are changed manually
        st.number_input(
            "Number of Pallets",
            min_value=1,
            step=1,
            key="pallets",
            disabled=inputs_disabled,
            on_change=sync_service_from_pallets,
        )

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

    st.markdown("---")

    jb, jf, mb, mf, pb, pf = calc_for_area(st.session_state["area"])
    n = int(st.session_state["pallets"])

    summary_rows = []
    if "Joda" in allowed:
        summary_rows.append({
            "Haulier": "Joda",
            "Base Rate": "No rate" if jb is None else f"£{float(jb):,.2f}",
            "Fuel Surcharge (%)": f"{float(st.session_state['joda_pct']):.2f}%",
            "Extras": f"£{extras_cost('joda', n):,.2f}",
            "Final Rate": "N/A" if jf is None else f"£{float(jf):,.2f}",
        })
    if "Mcdowells" in allowed:
        summary_rows.append({
            "Haulier": "McDowells",
            "Base Rate": "No rate" if mb is None else f"£{float(mb):,.2f}",
            "Fuel Surcharge (%)": f"{float(st.session_state['mcd_pct']):.2f}%",
            "Extras": f"£{extras_cost('mcdowells', n):,.2f}",
            "Final Rate": "N/A" if mf is None else f"£{float(mf):,.2f}",
        })
    if "Pc Howard" in allowed:
        summary_rows.append({
            "Haulier": "PC Howard",
            "Base Rate": "No rate" if pb is None else f"£{float(pb):,.2f}",
            "Fuel Surcharge (%)": f"{float(st.session_state['pch_pct']):.2f}%",
            "Extras": f"£{extras_cost('pc howard', n):,.2f}",
            "Final Rate": "N/A" if pf is None else f"£{float(pf):,.2f}",
        })

    st.subheader("3. Calculated Rates")
    if summary_rows:
        df = pd.DataFrame(summary_rows).set_index("Haulier")
        st.table(df.style.apply(highlight_cheapest_factory(), axis=1))
    else:
        st.info("No hauliers available for this warehouse.")

    st.markdown("---")

    st.markdown("---")
    st.subheader("Add to Export Lists")
    _clear_so_on_next_run()

    # SO number (should already be set by SO picker, but can be typed)
    st.text_input("SO Number", key="so_number", placeholder="e.g. 020502")

    # Consignee: prefer SO-derived delivery address, else fallback to address book
    so_con = st.session_state.get("_so_consignee", {}) or {}
    consignee_ok = bool(_safe_str(so_con.get("PostalName")) and _safe_str(so_con.get("Postcode")))
    consignee_obj = so_con.copy()

    customers_df = load_customers_df()

    if not consignee_ok:
        st.warning(
            "Consignee details not found in SO export (need at least Name + Postcode). "
            "Select from the address book below."
        )

        # Seed search from current postcode area if blank
        if not st.session_state.get("cust_search", "").strip():
            st.session_state["cust_search"] = str(st.session_state.get("area", "")).strip()

        q = _norm(st.text_input("Customer search", key="cust_search", placeholder="code / name / postcode…"))
        q_compact = q.replace(" ", "")

        blobs = (
            customers_df["CustomerCode"].astype(str).map(_norm)
            + " "
            + customers_df["CustomerName"].astype(str).map(_norm)
            + " "
            + customers_df["Postcode"].astype(str).map(_norm)
        )
        pc_compact = customers_df["Postcode"].astype(str).map(_norm).str.replace(" ", "", regex=False)

        if q:
            mask = blobs.str.contains(q, na=False) | pc_compact.str.contains(q_compact, na=False)
            filtered = customers_df[mask].copy()
        else:
            filtered = customers_df.copy()

        # remove obvious duplicates to make selection clearer
        filtered = filtered.drop_duplicates(
            subset=["CustomerCode", "CustomerName", "Postcode", "Address1"], keep="first"
        ).head(200)

        st.caption(f"Matches: {len(filtered):,}" + (" (showing first 200)" if len(filtered) > 200 else ""))

        options = [""] + filtered["ID"].tolist()
        label_map = {row["ID"]: customer_label(row) for _, row in filtered.iterrows()}

        st.selectbox(
            "Consignee (fallback)",
            options=options,
            key="cust_selected_id",
            format_func=lambda x: "— Select —" if x == "" else label_map.get(x, x),
        )

        cid = str(st.session_state.get("cust_selected_id", "")).strip()
        if cid:
            crow = customers_df.loc[customers_df["ID"] == cid]
            if not crow.empty:
                c0 = crow.iloc[0]
                consignee_obj = {
                    "CustomerCode": _safe_str(c0.get("CustomerCode")),
                    "CustomerName": _safe_str(c0.get("CustomerName")),
                    "PostalName": _safe_str(c0.get("CustomerName")) or _safe_str(c0.get("CustomerCode")),
                    "Address1": _safe_str(c0.get("Address1")),
                    "Address2": _safe_str(c0.get("Address2")),
                    "Address3": _safe_str(c0.get("Address3")),
                    "Address4": _safe_str(c0.get("Address4")),
                    "Postcode": _safe_str(c0.get("Postcode")),
                    "Contact": _safe_str(c0.get("Contact")),
                    "Tel": _safe_str(c0.get("Tel")),
                    "Email": _safe_str(c0.get("Email")),
                }
                consignee_ok = bool(consignee_obj["PostalName"] and consignee_obj["Postcode"])

    btns = st.columns([1, 1, 1, 2])

    def _add_all_for(haulier_key: str):
        _add_to_sage_basket(build_export_lines_for_haulier_sage(haulier_key))
        prow = build_portal_row(haulier_key, consignee_obj)
        _add_portal_row(haulier_key, prow)
        st.session_state["_clear_so_next"] = True
        mark_so_done(st.session_state.get("so_number", ""))

    if "Joda" in allowed:
        if btns[0].button("Add Joda", use_container_width=True):
            try:
                if not consignee_ok:
                    raise ValueError("Consignee not available (from SO or fallback).")
                _add_all_for("Joda")
                st.success("Added Joda lines (+ portal row).")
                st.rerun()
            except Exception as e:
                st.error(str(e))

    if "Mcdowells" in allowed:
        if btns[1].button("Add McDowells", use_container_width=True):
            try:
                if not consignee_ok:
                    raise ValueError("Consignee not available (from SO or fallback).")
                _add_all_for("Mcdowells")
                st.success("Added McDowells lines (+ portal row).")
                st.rerun()
            except Exception as e:
                st.error(str(e))

    if "Pc Howard" in allowed:
        if btns[2].button("Add PC Howard", use_container_width=True):
            try:
                if not consignee_ok:
                    raise ValueError("Consignee not available (from SO or fallback).")
                _add_all_for("Pc Howard")
                st.success("Added PC Howard lines (+ portal row).")
                st.rerun()
            except Exception as e:
                st.error(str(e))

# -------------------------
# EXPORT TAB
# -------------------------
with tab_export:
    st.header("Exports")
    # Daily PO refs
    with st.expander("Daily PO refs (today only)", expanded=True):
        st.caption("These apply to ALL lines you add today. They reset automatically after midnight.")
        j_101 = PO_NUMBER_MAP.get(("Joda", "101 - Skipton"), 1)
        j_201 = PO_NUMBER_MAP.get(("Joda", "201 - Skipton 2"), 2)
        m_101 = PO_NUMBER_MAP.get(("Mcdowells", "101 - Skipton"), 3)
        m_201 = PO_NUMBER_MAP.get(("Mcdowells", "201 - Skipton 2"), 4)
        p_102 = PO_NUMBER_MAP.get(("Pc Howard", "102 - Corby"), 5)
        c1, c2, c3, c4 = st.columns([1, 1, 1, 1.2], gap="medium")
        with c1:
            st.number_input("Joda PO Number", min_value=1, step=1, key="po_ref_joda")
            st.caption(f"defaults: 101→{j_101}, 201→{j_201}")
        with c2:
            st.number_input("McDowells PO Number", min_value=1, step=1, key="po_ref_mcd")
            st.caption(f"defaults: 101→{m_101}, 201→{m_201}")
        with c3:
            st.number_input("PC Howard PO Number", min_value=1, step=1, key="po_ref_pch")
            st.caption(f"default: 102→{p_102}")
        with c4:
            if st.button("Save PO refs for today", use_container_width=True):
                save_porefs_for_today(
                    int(st.session_state["po_ref_joda"]),
                    int(st.session_state["po_ref_mcd"]),
                    int(st.session_state["po_ref_pch"]),
                )
                st.success("Saved for today.")
    # Sage PO export + line list
    with st.expander("Sage PO Export (PO Import CSV)", expanded=True):
        basket = st.session_state.get("export_basket", [])
        export_df = pd.DataFrame(basket).reindex(columns=SAGE_EXPORT_COLUMNS) if basket else pd.DataFrame(columns=SAGE_EXPORT_COLUMNS)
        export_df = export_df.where(pd.notnull(export_df), "")
        sage_bytes = export_df.to_csv(index=False, sep=",", na_rep="", lineterminator="\n", quoting=csvlib.QUOTE_MINIMAL).encode("utf-8")
        top = st.columns([1.4, 1.0, 3.6])
        with top[0]:
            st.download_button(
                label="Download Sage PO Import CSV",
                data=sage_bytes,
                file_name=f"PO_Import_Export_{date.today().strftime('%Y%m%d')}.csv",
                mime="text/csv",
                use_container_width=True,
                disabled=(len(basket) == 0),
            )
        with top[1]:
            if st.button("Clear all", use_container_width=True, disabled=(len(basket) == 0), key="clear_sage"):
                st.session_state["export_basket"] = []
                st.rerun()
        with top[2]:
            st.caption(f"{len(basket)} line(s)" if basket else "No lines yet")
        st.divider()
        if not basket:
            st.info("No Sage PO lines saved yet.")
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
                if cols[6].button("🗑", key=f"rm_sage_{rid}", help="Remove this line"):
                    remove_id = rid
            if remove_id:
                st.session_state["export_basket"] = [x for x in st.session_state["export_basket"] if x.get("_row_id") != remove_id]
                st.rerun()

    def _portal_export_block(title: str, rows_key: str, filename_prefix: str, caption: str):
        rows = st.session_state.get(rows_key, [])
        export_df = pd.DataFrame(rows).reindex(columns=PORTAL_COLUMNS) if rows else pd.DataFrame(columns=PORTAL_COLUMNS)
        export_df = export_df.where(pd.notnull(export_df), "")
        bytes_ = export_df.to_csv(index=False, sep=",", na_rep="", lineterminator="\n", quoting=csvlib.QUOTE_MINIMAL).encode("utf-8")
        with st.expander(title, expanded=True):
            top = st.columns([1.4, 1.0, 3.6])
            with top[0]:
                st.download_button(
                    label=f"Download {filename_prefix} Portal CSV",
                    data=bytes_,
                    file_name=f"{filename_prefix}_Portal_{date.today().strftime('%Y%m%d')}.csv",
                    mime="text/csv",
                    use_container_width=True,
                    disabled=(len(rows) == 0),
                )
            with top[1]:
                if st.button("Clear all", use_container_width=True, disabled=(len(rows) == 0), key=f"clear_{rows_key}"):
                    st.session_state[rows_key] = []
                    st.rerun()
            with top[2]:
                st.caption(f"{len(rows)} row(s)" if rows else "No rows yet")
            st.caption(caption)
            st.divider()
            if not rows:
                st.info("No portal rows saved yet.")
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
                    if cols[4].button("🗑", key=f"rm_{rows_key}_{rid}", help="Remove this portal row"):
                        remove_id = rid
                if remove_id:
                    st.session_state[rows_key] = [x for x in st.session_state[rows_key] if x.get("_row_id") != remove_id]
                    st.rerun()

    _portal_export_block(
        "Portal Export — McDowells (CSV)",
        "portal_rows_mcd",
        "McDowells",
        "McDowells portal export (live format: Reference.csv).",
    )
    _portal_export_block(
        "Portal Export — Joda (CSV)",
        "portal_rows_joda",
        "Joda",
        "Placeholder: currently modelled on McDowells Reference.csv until Joda portal spec is provided.",
    )
    _portal_export_block(
        "Portal Export — PC Howard (CSV)",
        "portal_rows_pch",
        "PC_Howard",
        "Placeholder: currently modelled on McDowells Reference.csv until PC Howard portal spec is provided.",
    )

# -------------------------
# CUSTOMERS TAB
# -------------------------
with tab_customers:
    st.header("Customers (customers.xlsx)")
    st.caption("Edits here write back to customers.xlsx (shared across portal exports).")
    customers_df = load_customers_df()
    q = _norm(st.text_input("Search (code / name / postcode)", key="ab_search", placeholder="e.g. A0003 or BD7…"))
    q_compact = q.replace(" ", "")
    blobs = (
        customers_df["CustomerCode"].astype(str).map(_norm)
        + " "
        + customers_df["CustomerName"].astype(str).map(_norm)
        + " "
        + customers_df["Postcode"].astype(str).map(_norm)
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
        if cols[3].button("✏️", key=f"ab_edit_{rid}", help="Edit"):
            edit_id = rid
        if cols[4].button("🗑", key=f"ab_del_{rid}", help="Delete"):
            delete_id = rid
    if delete_id:
        customers_df = customers_df[customers_df["ID"] != delete_id].copy()
        save_customers_df(customers_df)
        st.success("Deleted customer from customers.xlsx")
        st.rerun()
    if edit_id:
        row = customers_df.loc[customers_df["ID"] == edit_id].iloc[0]
        st.session_state["ab_selected_id"] = edit_id
        st.session_state["ab_code"] = str(row.get("CustomerCode", "") or "")
        st.session_state["ab_name"] = str(row.get("CustomerName", "") or "")
        st.session_state["ab_a1"] = str(row.get("Address1", "") or "")
        st.session_state["ab_a2"] = str(row.get("Address2", "") or "")
        st.session_state["ab_a3"] = str(row.get("Address3", "") or "")
        st.session_state["ab_a4"] = str(row.get("Address4", "") or "")
        st.session_state["ab_pc"] = str(row.get("Postcode", "") or "")
        st.session_state["ab_contact"] = str(row.get("Contact", "") or "")
        st.session_state["ab_tel"] = str(row.get("Tel", "") or "")
        st.session_state["ab_email"] = str(row.get("Email", "") or "")
        st.rerun()
    st.markdown("---")
    st.markdown("#### Edit / Add customer")
    selected_id = st.session_state.get("ab_selected_id", "")
    selected_row = None
    if selected_id:
        match = customers_df.loc[customers_df["ID"] == selected_id]
        if not match.empty:
            selected_row = match.iloc[0]
    if selected_row is not None and st.session_state.get("_loaded_ab_id") != selected_id:
        st.session_state["_loaded_ab_id"] = selected_id
        st.session_state.setdefault("ab_code", str(selected_row.get("CustomerCode", "") or ""))
        st.session_state.setdefault("ab_name", str(selected_row.get("CustomerName", "") or ""))
        st.session_state.setdefault("ab_a1", str(selected_row.get("Address1", "") or ""))
        st.session_state.setdefault("ab_a2", str(selected_row.get("Address2", "") or ""))
        st.session_state.setdefault("ab_a3", str(selected_row.get("Address3", "") or ""))
        st.session_state.setdefault("ab_a4", str(selected_row.get("Address4", "") or ""))
        st.session_state.setdefault("ab_pc", str(selected_row.get("Postcode", "") or ""))
        st.session_state.setdefault("ab_contact", str(selected_row.get("Contact", "") or ""))
        st.session_state.setdefault("ab_tel", str(selected_row.get("Tel", "") or ""))
        st.session_state.setdefault("ab_email", str(selected_row.get("Email", "") or ""))
    f1, f2, f3, f4 = st.columns(4)
    with f1:
        code = st.text_input("CustomerCode", key="ab_code")
        name = st.text_input("CustomerName", key="ab_name")
    with f2:
        a1 = st.text_input("Address1", key="ab_a1")
        a2 = st.text_input("Address2", key="ab_a2")
    with f3:
        a3 = st.text_input("Address3", key="ab_a3")
        a4 = st.text_input("Address4", key="ab_a4")
    with f4:
        pc = st.text_input("Postcode", key="ab_pc")
        contact = st.text_input("Contact", key="ab_contact")
        tel = st.text_input("Tel", key="ab_tel")
        email = st.text_input("Email", key="ab_email")
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
            st.session_state["ab_selected_id"] = ""
            st.session_state["_loaded_ab_id"] = ""
            for k in ["ab_code","ab_name","ab_a1","ab_a2","ab_a3","ab_a4","ab_pc","ab_contact","ab_tel","ab_email"]:
                st.session_state[k] = ""
            st.rerun()
