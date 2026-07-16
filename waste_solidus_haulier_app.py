# app.py
import os
import math
import json
import uuid
import re
import csv as csvlib
from datetime import date, timedelta
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
        V3.9.1  
        **Haulier exports and portal imports**
        - Upload the Sage sales order export to pre-fill SO, postcode, consignee, promised date, notes and weight
        - Add Joda, McDowells or PC Howard jobs from the Table tab, then download the Sage and portal files from Export
        - Joda/Qargo exports include daily job numbers, combine/split tools, collection site details, extras and KG weight
        - McDowells and PC Howard portal exports use the imported consignee details where available
        - The portal delivery date defaults from the promised date and can be changed before adding jobs
        - Tick Pre-Booked only when pallets must be delivered on a specific agreed day
        - Sage PO numbers can be changed each day and are applied to existing export lines when saved
        """,
        unsafe_allow_html=True
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
TEMPLATE_MCD_PATH = "Reference.csv"        # McDowells portal template header
TEMPLATE_JODA_QARGO_PATH = "Qargo Import Template.xlsx"  # Joda/Qargo portal template header

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

# McDowells portal constants (haulier-specific exporter)
MCD_REQ_DEPOT = "034"
MCD_COLL_DEPOT = "034"
MCD_DEL_DEPOT = "008"
MCD_CONSIGNOR_ACCOUNT = "SOLIDU"
MCD_SERVICE_MAP = {"Economy": "EC", "Next Day": "ND"}

# McDowells combined service codes. The Service column should carry the
# delivery service and extras together where McDowells provides a specific code.
MCD_SERVICE_CODES = {
    "next_day": "ND",
    "next_day_tail_lift": "NDTL",
    "next_day_timed": "BKSL",
    "next_day_am": "AM",
    "next_day_am_tail_lift": "AMTL",
    "next_day_specific_time": "TIME",
    "economy": "EC",
    "economy_tail_lift": "ECTL",
    "book_in": "BKIN",
    "dedicated_day": "DDAY",
    "dedicated_day_am": "DDAM",
    "dedicated_day_timed": "DDBS",
}

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
    return d.strftime("%d%m%Y")

def _yyyymmdd(d: date) -> str:
    return d.strftime("%Y%m%d")


def _parse_date_or_none(value):
    if value is None:
        return None
    if isinstance(value, date):
        return value
    try:
        if pd.isna(value):
            return None
    except Exception:
        pass

    text = str(value).strip()
    if not text:
        return None

    # Sage sometimes provides ISO-like datetimes such as 2026-10-07 00:00:00.
    # Parse these explicitly as YYYY-MM-DD before using dayfirst parsing; otherwise
    # pandas can interpret 2026-10-07 as 10 July instead of 7 October.
    iso_head = text[:10]
    if (
        len(iso_head) == 10
        and iso_head[0:4].isdigit()
        and iso_head[4] in {"-", "/"}
        and iso_head[5:7].isdigit()
        and iso_head[7] in {"-", "/"}
        and iso_head[8:10].isdigit()
    ):
        try:
            return date(int(iso_head[0:4]), int(iso_head[5:7]), int(iso_head[8:10]))
        except Exception:
            pass

    # Qargo/date exports may already be stored as yyyymmdd.
    digits = ''.join(ch for ch in text if ch.isdigit())
    if len(digits) == 8:
        try:
            return date(int(digits[0:4]), int(digits[4:6]), int(digits[6:8]))
        except Exception:
            pass

    # Sage exports are usually UK-style, but allow pandas to handle Excel/date strings too.
    for dayfirst in (True, False):
        try:
            parsed = pd.to_datetime(text, dayfirst=dayfirst, errors="coerce")
            if pd.notna(parsed):
                return parsed.date()
        except Exception:
            pass

    return None


def _joda_automatic_delivery_date(service_value: str = "") -> date:
    svc = str(service_value or "").strip().upper().replace(" ", "")
    days = 1 if svc in {"ND", "NEXTDAY"} else 2
    return date.today() + timedelta(days=days)


def _joda_delivery_date_value(service_value: str = "") -> date:
    # Generic portal delivery date used by Joda/Qargo, McDowells and PC Howard.
    # Keep the old joda_delivery_date key as a fallback so existing sessions still work.
    selected = _parse_date_or_none(st.session_state.get("portal_delivery_date"))
    if selected is None:
        selected = _parse_date_or_none(st.session_state.get("joda_delivery_date"))
    return selected or _joda_automatic_delivery_date(service_value)


def _joda_delivery_date_str(service_value: str = "") -> str:
    """Joda/Qargo delivery date. Defaults from selected SO promised date where available, otherwise service-based automatic date."""
    return _yyyymmdd(_joda_delivery_date_value(service_value))


def _joda_prebooked_enabled(service_value: str = "", delivery_date_value=None) -> bool:
    # Pre-Booked is now deliberately user-selected by exception.
    # Changing the portal delivery date alone should not automatically add Pre-Booked.
    return bool(st.session_state.get("portal_prebooked", False))

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

def customer_label(row: pd.Series) -> str:
    code = str(row.get("CustomerCode", "")).strip()
    name = str(row.get("CustomerName", "")).strip()
    pc = str(row.get("Postcode", "")).strip()
    left = code or name or "Customer"
    if code and name:
        left = f"{code} — {name}"
    return f"{left} — {pc}".strip(" —")

def sync_service_from_pallets():
    """Default to Economy; switch to Next Day when pallets are over 6."""
    try:
        n = int(st.session_state.get("pallets", 1))
    except Exception:
        n = 1
    st.session_state["service"] = "Next Day" if n > 6 else "Economy"


def _safe_int(v, default: int) -> int:
    try:
        if v is None:
            return int(default)
        s = str(v).strip()
        if s == "":
            return int(default)
        return int(float(s))
    except Exception:
        return int(default)


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


def save_porefs_for_today(
    joda: int,
    mcd: int,
    pch: int,
    joda_job_combined: str = "",
    joda_job_101: str = "",
    joda_job_201: str = "",
) -> None:
    today_str = date.today().isoformat()
    with open(POREFS_FILE, "w") as f:
        json.dump({
            "date": today_str,
            "joda": int(joda),
            "mcd": int(mcd),
            "pch": int(pch),
            "joda_job_combined": str(joda_job_combined).strip(),
            "joda_job_101": str(joda_job_101).strip(),
            "joda_job_201": str(joda_job_201).strip(),
        }, f)


def initialise_porefs_session_defaults():
    saved = load_porefs_for_today()
    defaults = {
        "po_ref_joda": int(saved.get("joda", PO_NUMBER_MAP.get(("Joda", "101 - Skipton"), 1))),
        "po_ref_mcd": int(saved.get("mcd", PO_NUMBER_MAP.get(("Mcdowells", "101 - Skipton"), 3))),
        "po_ref_pch": int(saved.get("pch", PO_NUMBER_MAP.get(("Pc Howard", "102 - Corby"), 5))),
        "joda_job_combined": str(saved.get("joda_job_combined", "")).strip(),
        "joda_job_101": str(saved.get("joda_job_101", "")).strip(),
        "joda_job_201": str(saved.get("joda_job_201", "")).strip(),
    }
    for key, value in defaults.items():
        current = st.session_state.get(key)
        if current is None or str(current).strip() == "":
            st.session_state[key] = value


def _get_joda_job_number(collection_warehouse: str = "") -> str:
    wh = str(collection_warehouse or st.session_state.get("warehouse_name", "")).strip()
    if st.session_state.get("dual"):
        if wh == "101 - Skipton":
            return str(st.session_state.get("joda_job_101", "")).strip()
        if wh == "201 - Skipton 2":
            return str(st.session_state.get("joda_job_201", "")).strip()
    return str(st.session_state.get("joda_job_combined", "")).strip()


def _apply_joda_job_numbers_to_existing_rows():
    # Backwards-compatible helper: now just ensures missing Joda job numbers are allocated.
    _ensure_joda_job_numbers()


def _joda_job_prefix() -> str:
    # Two-digit year + month + day keeps the number short while still grouping by day.
    # Example: 260624001, 260624002, etc.
    return date.today().strftime("%y%m%d")


def _joda_existing_numeric_jobs(rows=None) -> List[int]:
    nums: List[int] = []
    prefix = _joda_job_prefix()
    source = rows if rows is not None else st.session_state.get("portal_rows_joda", [])
    for row in source:
        val = str(row.get("Job Number", "")).strip()
        if val.startswith(prefix) and val[len(prefix):].isdigit():
            nums.append(int(val[len(prefix):]))
    return nums


def _next_joda_job_number(rows=None) -> str:
    nums = _joda_existing_numeric_jobs(rows)
    next_no = (max(nums) if nums else 0) + 1
    return f"{_joda_job_prefix()}{next_no:03d}"


def _ensure_joda_job_numbers(rows=None) -> None:
    target = rows if rows is not None else st.session_state.get("portal_rows_joda", [])
    for row in target:
        if not str(row.get("Job Number", "")).strip():
            row["Job Number"] = _next_joda_job_number(target)


def _selected_joda_row_ids() -> List[str]:
    selected: List[str] = []
    for row in st.session_state.get("portal_rows_joda", []):
        rid = str(row.get("_row_id", "")).strip()
        if rid and st.session_state.get(f"sel_joda_{rid}"):
            selected.append(rid)
    return selected


def _clear_joda_selection(row_ids: Optional[List[str]] = None) -> None:
    """Queue selection clearing for the next rerun.

    Streamlit does not allow changing a checkbox key after that checkbox
    has been created in the same run. Queueing avoids the
    "cannot be modified after the widget ... is instantiated" error.
    """
    ids = row_ids or [str(r.get("_row_id", "")).strip() for r in st.session_state.get("portal_rows_joda", [])]
    st.session_state["_joda_clear_selection_ids"] = [rid for rid in ids if rid]


def _apply_pending_joda_selection_clear() -> None:
    ids = st.session_state.pop("_joda_clear_selection_ids", [])
    if not ids:
        return
    for rid in ids:
        if rid:
            st.session_state[f"sel_joda_{rid}"] = False


def _combine_selected_joda_rows(row_ids: List[str]) -> None:
    if len(row_ids) < 2:
        raise ValueError("Select at least two Joda rows to combine.")
    rows = st.session_state.get("portal_rows_joda", [])
    new_job = _next_joda_job_number(rows)
    selected = set(row_ids)
    for row in rows:
        if str(row.get("_row_id", "")) in selected:
            row["Job Number"] = new_job
    _clear_joda_selection(row_ids)


def _split_weight(total_weight, part_pallets: int, total_pallets: int):
    try:
        if str(total_weight).strip() == "" or total_pallets <= 0:
            return ""
        return round(float(total_weight) * int(part_pallets) / int(total_pallets), 3)
    except Exception:
        return ""


def _clone_joda_row_for_split(source: Dict[str, object], warehouse: str, pallets: int, weight_value) -> Dict[str, object]:
    new_row = dict(source)
    new_row["_row_id"] = uuid.uuid4().hex
    so = _extract_so_from_joda_row(new_row)
    if so:
        new_row["_so_number"] = so
        if "Job Order Number" in new_row:
            new_row["Job Order Number"] = _joda_po_so_ref(so)
    if "REF 2 Link" in new_row:
        new_row["REF 2 Link"] = _get_po_ref("Joda")
    new_row["_joda_collection_warehouse"] = warehouse
    new_row["Job Number"] = _next_joda_job_number(st.session_state.get("portal_rows_joda", []))
    new_row["Full"] = int(pallets)
    if "Spaces" in new_row:
        new_row["Spaces"] = int(pallets)
    if "Weight" in new_row:
        new_row["Weight"] = weight_value
    for k, v in _joda_collection_details(warehouse).items():
        if k in new_row:
            new_row[k] = v
    return new_row


def _apply_joda_split_from_inputs(row_ids: List[str]) -> None:
    if not row_ids:
        raise ValueError("Select at least one Joda row to split.")

    rows = st.session_state.get("portal_rows_joda", [])
    selected = set(row_ids)
    replacement: List[Dict[str, object]] = []

    for row in rows:
        rid = str(row.get("_row_id", ""))
        if rid not in selected:
            replacement.append(row)
            continue

        original_pallets = _safe_int(row.get("Full", 0), 0)
        split_101 = _safe_int(st.session_state.get(f"split_joda_{rid}_101"), 0)
        split_201 = _safe_int(st.session_state.get(f"split_joda_{rid}_201"), 0)

        if original_pallets <= 0:
            raise ValueError("Cannot split a Joda row with no pallet quantity.")
        if split_101 + split_201 != original_pallets:
            order_no = str(row.get("Job Order Number", "")).strip()
            raise ValueError(f"Split for {order_no or rid} must equal {original_pallets} pallets.")

        if split_101 > 0:
            replacement.append(_clone_joda_row_for_split(row, "101 - Skipton", split_101, _split_weight(row.get("Weight", ""), split_101, original_pallets)))
        if split_201 > 0:
            replacement.append(_clone_joda_row_for_split(row, "201 - Skipton 2", split_201, _split_weight(row.get("Weight", ""), split_201, original_pallets)))

    st.session_state["portal_rows_joda"] = replacement
    _clear_joda_selection(row_ids)
    st.session_state["_joda_split_mode"] = False
    st.session_state["_joda_split_ids"] = []


def _get_po_ref(haulier_title: str) -> int:
    if haulier_title == "Joda":
        return _safe_int(st.session_state.get("po_ref_joda"), PO_NUMBER_MAP.get(("Joda", "101 - Skipton"), 1))
    if haulier_title == "Mcdowells":
        return _safe_int(st.session_state.get("po_ref_mcd"), PO_NUMBER_MAP.get(("Mcdowells", "101 - Skipton"), 3))
    return _safe_int(st.session_state.get("po_ref_pch"), PO_NUMBER_MAP.get(("Pc Howard", "102 - Corby"), 5))


def _digits_only(value) -> str:
    s = str(value or "").strip()
    digits = "".join(ch for ch in s if ch.isdigit())
    if not digits:
        return ""
    stripped = digits.lstrip("0")
    return stripped or "0"


def _normalise_so_number(value) -> str:
    """Remove leading zeroes from an SO number for display/export while keeping non-numeric refs safe."""
    s = str(value or "").strip()
    digits = _digits_only(s)
    if digits:
        return digits
    return s.lstrip("0") or s


def _joda_po_so_ref(so_number: str = "", po_number: Optional[int] = None) -> str:
    po = _safe_int(po_number if po_number is not None else st.session_state.get("po_ref_joda"), _get_po_ref("Joda"))
    so = _digits_only(so_number)
    if so:
        return f"PO{po}/SO{so}"
    return f"PO{po}"


def _extract_so_from_joda_row(row: Dict[str, object]) -> str:
    so = str(row.get("_so_number", "")).strip()
    if so:
        return so
    order_no = str(row.get("Job Order Number", "")).strip()
    if "/SO" in order_no.upper():
        return order_no.upper().split("/SO", 1)[1].strip()
    if order_no.upper().startswith("SO"):
        return order_no[2:].strip()
    return order_no.strip()


def _joda_weight_for_so(so_number: str, pallets: Optional[int] = None):
    so = str(so_number or "").strip()
    weight_map = st.session_state.get("_so_weight_by_so", {}) or {}

    candidates = []
    if so:
        candidates.append(weight_map.get(so))
        candidates.append(weight_map.get(_digits_only(so)))
    current_so = str(st.session_state.get("so_number", "")).strip()
    if so and current_so and _digits_only(so) == _digits_only(current_so):
        candidates.append(st.session_state.get("_so_weight", ""))

    for val in candidates:
        try:
            if str(val).strip() != "" and pd.notna(val):
                return round(float(val), 3)
        except Exception:
            pass

    try:
        wpp = float(st.session_state.get("portal_weight_per_pallet", 0.0) or 0.0)
        if wpp > 0 and pallets is not None:
            return round(float(int(pallets) * wpp), 3)
    except Exception:
        pass

    return ""


def _ensure_joda_refs_and_weights(rows=None) -> None:
    target = rows if rows is not None else st.session_state.get("portal_rows_joda", [])
    for row in target:
        so = _extract_so_from_joda_row(row)
        if so:
            row["_so_number"] = so
            if "Job Order Number" in row:
                row["Job Order Number"] = _joda_po_so_ref(so)
        if "REF 2 Link" in row:
            row["REF 2 Link"] = _get_po_ref("Joda")
        if "Delivery Date" in row:
            row_date = row.get("_joda_delivery_date", "") or row.get("Delivery Date", "")
            parsed_row_date = _parse_date_or_none(row_date)
            if parsed_row_date is not None:
                row["Delivery Date"] = _yyyymmdd(parsed_row_date)
            elif str(row.get("Delivery Date", "")).strip() == "":
                row["Delivery Date"] = _joda_delivery_date_str(row.get("Service", ""))
if "Extras" in row:
    row_date = row.get("_joda_delivery_date", "") or row.get("Delivery Date", "")
    row["Extras"] = _merge_qargo_extras(
        row.get("Extras", ""),
        _qargo_extras(row.get("Service", ""), row_date),
    )
if "Delivery Time" in row:
    row["Delivery Time"] = ""
        if "Weight" in row and str(row.get("Weight", "")).strip() == "":
            wt = _joda_weight_for_so(so, _safe_int(row.get("Full", 0), 0))
            if str(wt).strip() != "":
                row["Weight"] = wt


def apply_po_refs_to_existing_lines():
    j = _get_po_ref("Joda")
    m = _get_po_ref("Mcdowells")
    p = _get_po_ref("Pc Howard")

    # Sage PO import rows
    for line in st.session_state.get("export_basket", []):
        acc = str(line.get("Purchase Order Supplier Acc Code", "")).strip().upper()
        if acc == JODA_ACC:
            line["Purchase Order Number"] = j
        elif acc == MCD_ACC:
            line["Purchase Order Number"] = m
        elif acc == PCH_ACC:
            line["Purchase Order Number"] = p

    # Portal exports that also need the same daily PO/reference number
    for row in st.session_state.get("portal_rows_joda", []):
        so = _extract_so_from_joda_row(row)
        if so and "Job Order Number" in row:
            row["Job Order Number"] = _joda_po_so_ref(so, j)
            row["_so_number"] = so
        if "REF 2 Link" in row:
            row["REF 2 Link"] = j

    for row in st.session_state.get("portal_rows_mcd", []):
        if "Order_No" in row:
            row["Order_No"] = m


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
    done = list(dict.fromkeys([str(x).strip() for x in done_list if str(x).strip()]))
    with open(SO_DONE_FILE, "w") as f:
        json.dump({"date": today_str, "done": done}, f)


def mark_so_done(so_no: str) -> None:
    so_no = _normalise_so_number(so_no)
    if not so_no:
        return
    current = list(st.session_state.get("done_sos", []))
    if so_no not in current:
        current.append(so_no)
    st.session_state["done_sos"] = current
    save_done_sos_for_today(current)

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

DEFAULT_JODA_QARGO_COLUMNS: List[str] = [
    "Job Number", "Job Order Number", "Export Dater", "Account Code", "Vehicle Code",
    "Collection Name", "Collection Address Line 1", "Collection Address Line 2",
    "Collection Address Line 4", "Collection Address Line 5", "Collection Post Code",
    "Collection Phone Number", "Collection Date", "Delivery Name",
    "Delivery Address Line 1", "Delivery Address Line 2", "Delivery Address Line 4",
    "Delivery Address Line 5", "Delivery Post Code", "Delivery Mobile",
    "Delivery Phone", "Delivery Date", "Full", "Half", "Quarter", "Oversized",
    "Weight", "Notes Line 1", "Notes Line 2", "Notes Line 3", "Notes Line 4",
    "Service", "Extras", "Nett", "Delivery Time", "Spaces", "Collection Country",
    "Delivery Country", "REF 2 Link", "Pallet Type",
]

@st.cache_data
def load_excel_header_columns(path: str, fallback: List[str]) -> List[str]:
    if not os.path.exists(path):
        return fallback
    try:
        df = pd.read_excel(path, sheet_name=0, nrows=0)
        cols = [str(c).strip() for c in df.columns if str(c).strip()]
        return cols or fallback
    except Exception:
        return fallback

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
JODA_QARGO_COLUMNS = load_excel_header_columns(TEMPLATE_JODA_QARGO_PATH, DEFAULT_JODA_QARGO_COLUMNS)

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
# customers.xlsx persistence (read/write)
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
# Rates load
# -------------------------
@st.cache_data
def load_rate_table(excel_path: str, _mtime: float) -> pd.DataFrame:
    # Read without headers first so we can detect the real header row
    preview = pd.read_excel(excel_path, sheet_name=0, header=None, nrows=10)

    header_row = None
    for i in range(len(preview)):
        row = preview.iloc[i].astype(str).str.strip().str.lower().tolist()
        if ("postcode" in row) and ("service" in row) and ("vendor" in row):
            header_row = i
            break

    # Fallback to previous behaviour if we can’t detect it
    if header_row is None:
        header_row = 1

    raw = pd.read_excel(excel_path, sheet_name=0, header=header_row)

    # Normalise first three columns to the app’s expected names
    cols = list(raw.columns)
    if len(cols) < 3:
        raise ValueError(f"Rate sheet in {excel_path} does not have expected columns.")

    raw = raw.rename(columns={
        cols[0]: "PostcodeArea",
        cols[1]: "Service",
        cols[2]: "Vendor",
    })

    # Clean / forward-fill the key columns
    raw["PostcodeArea"] = raw["PostcodeArea"].ffill()
    raw["Service"] = raw["Service"].ffill()
    raw["Vendor"] = raw["Vendor"].ffill()

    # If the sheet uses "Delivered Cost*" + unnamed columns, map them to pallet numbers
    if "Delivered Cost*" in raw.columns:
        pallet_start_col = "Delivered Cost*"
        after = raw.columns.tolist()[raw.columns.tolist().index(pallet_start_col):]

        # Map Delivered Cost* -> 1, next col -> 2, etc.
        pallet_map = {}
        for idx, c in enumerate(after, start=1):
            pallet_map[c] = idx

        raw = raw.rename(columns=pallet_map)

    # Identify pallet columns (either ints already, or digit strings)
    pallet_cols = [
        c for c in raw.columns
        if isinstance(c, int) or (isinstance(c, str) and str(c).isdigit())
    ]

    # Melt to long format
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
# Load rates into dataframes
# -------------------------
# Main (Joda + McDowells)
mtime_main = os.path.getmtime(RATE_XLSX_MAIN)
rate_df_main = load_rate_table(RATE_XLSX_MAIN, mtime_main)
unique_areas_main = sorted(rate_df_main["PostcodeArea"].dropna().astype(str).unique())

# PC Howard
rate_df_pch = pd.DataFrame(columns=["PostcodeArea", "Service", "Vendor", "Pallets", "BaseRate"])
unique_areas_pch: List[str] = []

if os.path.exists(RATE_XLSX_PCH):
    mtime_pch = os.path.getmtime(RATE_XLSX_PCH)
    rate_df_pch = load_rate_table(RATE_XLSX_PCH, mtime_pch)
    unique_areas_pch = sorted(rate_df_pch["PostcodeArea"].dropna().astype(str).unique())




def _apply_pending_order_options_reset():
    """Apply queued SO/order selection resets before related widgets are rendered.

    This keeps reruns safe when an SO is marked complete or the available
    Sales Order options change after filtering/completion. It is intentionally
    safe to call more than once per run.
    """
    if st.session_state.pop("_clear_so_next", False) or st.session_state.pop("_reset_sage_so_selected_next", False):
        st.session_state["so_number"] = ""
        st.session_state["sage_so_selected"] = ""
        st.session_state["_last_so_applied"] = ""
        st.session_state["_last_so_area_applied"] = ""

# -------------------------
# Session defaults (generic)
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
    st.session_state.setdefault("portal_delivery_date", st.session_state.get("joda_delivery_date", date.today() + timedelta(days=2)))
    st.session_state.setdefault("joda_delivery_date", st.session_state.get("portal_delivery_date"))
    st.session_state.setdefault("portal_prebooked", False)
    st.session_state.setdefault("split1", 1)
    st.session_state.setdefault("split2", 1)

    st.session_state.setdefault("so_number", "")
    st.session_state.setdefault("export_basket", [])

    # Daily PO refs and SO completion tracking
    st.session_state.setdefault("po_ref_joda", PO_NUMBER_MAP.get(("Joda", "101 - Skipton"), 1))
    st.session_state.setdefault("po_ref_mcd", PO_NUMBER_MAP.get(("Mcdowells", "101 - Skipton"), 3))
    st.session_state.setdefault("po_ref_pch", PO_NUMBER_MAP.get(("Pc Howard", "102 - Corby"), 5))
    st.session_state.setdefault("joda_job_combined", "")
    st.session_state.setdefault("joda_job_101", "")
    st.session_state.setdefault("joda_job_201", "")
    st.session_state.setdefault("done_sos", [])
    st.session_state.setdefault("show_done_sos", False)

    # Portal rows per haulier
    st.session_state.setdefault("portal_rows_mcd", [])
    st.session_state.setdefault("portal_rows_joda", [])
    st.session_state.setdefault("portal_rows_pch", [])
    st.session_state.setdefault("_joda_split_mode", False)
    st.session_state.setdefault("_joda_split_ids", [])

    # Portal settings (generic keys but currently used by McDowells exporter)
    st.session_state.setdefault("portal_consignor_name", "")
    st.session_state.setdefault("portal_consignor_postcode", "")
    st.session_state.setdefault("portal_consignor_account", MCD_CONSIGNOR_ACCOUNT)
    st.session_state.setdefault("portal_consignor_email", "")
    st.session_state.setdefault("portal_entered_by", "")
    st.session_state.setdefault("portal_weight_per_pallet", 0.0)
    st.session_state.setdefault("portal_remarks1", "")
    st.session_state.setdefault("portal_remarks2", "")

    # Customer search / selection (generic)
    st.session_state.setdefault("cust_search", "")
    st.session_state.setdefault("cust_selected_id", "")

    # Address book selection / edit form keys (generic)
    st.session_state.setdefault("ab_selected_id", "")
    st.session_state.setdefault("_loaded_ab_id", "")

_ensure_defaults()

if "surcharges_loaded" not in st.session_state:
    refresh_surcharges_from_disk()
    st.session_state["surcharges_loaded"] = True

if "porefs_loaded" not in st.session_state:
    initialise_porefs_session_defaults()
    st.session_state["porefs_loaded"] = True

if "done_loaded" not in st.session_state:
    done_data = load_done_sos_for_today()
    st.session_state["done_sos"] = done_data.get("done", [])
    st.session_state["done_loaded"] = True

_apply_pending_order_options_reset()

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
    # McDowells small-load extra no longer applies.
    return 0.0

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

    so_number = _normalise_so_number(so_number)
    svc_suffix = f" ({service})" if str(service).strip() else ""

    if so_number:
        desc = f"SO{so_number} - {area_code} {label}{svc_suffix}".strip()
    else:
        desc = f"{area_code} {label}{svc_suffix}".strip()

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
    _apply_pending_order_options_reset()

# -------------------------
# Portal row builder (McDowells exporter consumes generic customer row)
# -------------------------
def _blank_mcd_row() -> Dict[str, object]:
    return {c: "" for c in MCD_PORTAL_COLUMNS}

def _mcd_delivery_time(customer_row=None) -> str:
    # McDowells combined service codes carry AM/TIMED/Tail Lift in the Service column.
    # Only populate Delivery Time where there is a genuine specific-time request, e.g. Pre-10 / by 3pm.
    return _mcd_specific_time_label(customer_row)


def _mcd_note_text(customer_row=None) -> str:
    pieces: List[str] = []
    if customer_row is not None:
        for key in ["DeliveryNote1", "DeliveryNote2", "DeliveryNote3", "DeliveryNote4"]:
            val = _row_value(customer_row, key)
            if val:
                pieces.append(val)
    pieces.extend([
        st.session_state.get("portal_remarks1", ""),
        st.session_state.get("portal_remarks2", ""),
    ])
    return _norm(" | ".join(str(x) for x in pieces if str(x).strip()))


def _mcd_notes_request_book_in(customer_row=None) -> bool:
    text = _mcd_note_text(customer_row)
    compact = text.replace(" ", "").replace("-", "")
    return any(token in text for token in ["BOOK IN", "BOOK-IN", "BOOKING IN"]) or "BOOKIN" in compact


def _mcd_specific_time_label(customer_row=None) -> str:
    text = _mcd_note_text(customer_row)
    if not text:
        return ""

    # Match entries such as Pre-10, Pre 10am, by 3pm, before 10:30.
    patterns = [
        r"\bPRE\s*-?\s*(\d{1,2})(?::([0-5]\d))?\s*(AM|PM)?\b",
        r"\bBY\s+(\d{1,2})(?::([0-5]\d))?\s*(AM|PM)?\b",
        r"\bBEFORE\s+(\d{1,2})(?::([0-5]\d))?\s*(AM|PM)?\b",
        r"\b@(\d{1,2})(?::([0-5]\d))?\s*(AM|PM)?\b",
    ]
    for pattern in patterns:
        m = re.search(pattern, text, flags=re.IGNORECASE)
        if not m:
            continue
        hour = int(m.group(1))
        minute = m.group(2) or ""
        suffix = (m.group(3) or "AM").upper()

        # For morning delivery notes like Pre-10, make the AM explicit for McDowells.
        if minute:
            return f"{hour}:{minute}{suffix}"
        return f"{hour}{suffix}"

    return ""


def _mcd_notes_request_specific_time(customer_row=None) -> bool:
    text = _mcd_note_text(customer_row)
    compact = text.replace(" ", "").replace("-", "")
    return bool(_mcd_specific_time_label(customer_row)) or "SPECIFICTIME" in compact


def _mcd_service_code(service_value: str = "", customer_row=None) -> str:
    """Return the combined McDowells service code for the portal Service column."""
    svc = str(service_value or "").strip()
    svc_norm = _norm(svc).replace(" ", "").replace("/", "").replace("-", "")

    dedicated = bool(st.session_state.get("portal_prebooked", False))
    book_in = _mcd_notes_request_book_in(customer_row)
    specific_time = _mcd_notes_request_specific_time(customer_row)
    timed = bool(st.session_state.get("timed", False))
    am = bool(st.session_state.get("ampm", False))
    tail = bool(st.session_state.get("tail", False))

    if book_in:
        return MCD_SERVICE_CODES["book_in"]

    if dedicated:
        if timed or specific_time:
            return MCD_SERVICE_CODES["dedicated_day_timed"]
        if am:
            return MCD_SERVICE_CODES["dedicated_day_am"]
        return MCD_SERVICE_CODES["dedicated_day"]

    # If an imported note accidentally puts a delivery flag into the service value,
    # treat it as a Next Day special service rather than exporting a bad code.
    if svc_norm in {"AM", "AMPM", "PRE10", "PRE10AM"}:
        am = True
        svc = "Next Day"
    elif svc_norm in {"TIMED", "TIME"}:
        timed = True
        svc = "Next Day"

    if svc == "Economy":
        if tail:
            return MCD_SERVICE_CODES["economy_tail_lift"]
        return MCD_SERVICE_CODES["economy"]

    # Default to next day for Next Day and for any special/timed delivery flags.
    if specific_time:
        return MCD_SERVICE_CODES["next_day_specific_time"]
    if timed:
        return MCD_SERVICE_CODES["next_day_timed"]
    if am and tail:
        return MCD_SERVICE_CODES["next_day_am_tail_lift"]
    if am:
        return MCD_SERVICE_CODES["next_day_am"]
    if tail:
        return MCD_SERVICE_CODES["next_day_tail_lift"]
    return MCD_SERVICE_CODES["next_day"] if svc == "Next Day" else MCD_SERVICE_MAP.get(svc, MCD_SERVICE_CODES["economy"])


def _merge_qargo_extras(*values) -> str:
    extras: List[str] = []
    seen = set()

    canonical = {
        "AM": "AM",
        "A.M": "AM",
        "A.M.": "AM",
        "AM DELIVERY": "AM",
        "AM/PM": "AM",
        "AM-PM": "AM",
        "TIMED": "TIMED",
        "TIME": "TIMED",
        "TIMED DELIVERY": "TIMED",
        "TAIL LIFT": "Tail Lift",
        "TAILLIFT": "Tail Lift",
        "TAIL-LIFT": "Tail Lift",
        "PRE-BOOKED": "Pre-Booked",
        "PRE BOOKED": "Pre-Booked",
        "PREBOOKED": "Pre-Booked",
    }

    for value in values:
        for part in str(value or "").replace(";", "|").split("|"):
            raw = part.strip()
            if not raw:
                continue

            key = _norm(raw).replace("_", " ")
            label = canonical.get(key, raw)
            dedupe_key = _norm(label)

            if dedupe_key not in seen:
                extras.append(label)
                seen.add(dedupe_key)

    return " | ".join(extras)


def _qargo_extras(service_value: str = "", delivery_date_value=None) -> str:
    """Qargo wants delivery extras in the Extras column, not Delivery Time."""
    extras: List[str] = []
    if st.session_state.get("ampm"):
        extras.append("AM")
    if st.session_state.get("timed"):
        extras.append("TIMED")
    if st.session_state.get("tail"):
        extras.append("Tail Lift")
    if _joda_prebooked_enabled(service_value, delivery_date_value):
        extras.append("Pre-Booked")
    return _merge_qargo_extras(*extras)

def build_portal_row_mcd(customer_row: pd.Series) -> Dict[str, object]:
    so = _normalise_so_number(st.session_state["so_number"])
    if not so:
        raise ValueError("SO Number is required before adding a portal row.")

    pallets = int(st.session_state["pallets"])
    svc_ui = str(st.session_state["service"]).strip()
    svc_code = _mcd_service_code(svc_ui, customer_row)

    r = _blank_mcd_row()
    r["_row_id"] = uuid.uuid4().hex
    r["_consignee_label"] = customer_label(customer_row)

    if "Order_No" in r:
        r["Order_No"] = _get_po_ref("Mcdowells")
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
        r["Delivery Time"] = _mcd_delivery_time(customer_row)

    if "Delivery Date " in r:
        r["Delivery Date "] = _ddmmyyyy_compact(_joda_delivery_date_value(svc_ui))

    if "Manifest Date" in r:
        r["Manifest Date"] = _ddmmyyyy_compact(date.today())

    if "Full Pallets" in r:
        r["Full Pallets"] = pallets

    wt = _joda_weight_for_so(so, pallets)
    if str(wt).strip() != "" and "Full Weight" in r:
        r["Full Weight"] = wt
    else:
        wpp = float(st.session_state.get("portal_weight_per_pallet", 0.0) or 0.0)
        if wpp > 0 and "Full Weight" in r:
            r["Full Weight"] = round(float(pallets * wpp), 3)

    collection = _joda_collection_details()
    if "Consignor Name" in r:
        r["Consignor Name"] = str(st.session_state.get("portal_consignor_name", "")).strip() or collection.get("Collection Name", "")
    if "ConsignorPostCode" in r:
        r["ConsignorPostCode"] = str(st.session_state.get("portal_consignor_postcode", "")).strip() or collection.get("Collection Post Code", "")
    if "Consignor Account" in r:
        r["Consignor Account"] = str(st.session_state.get("portal_consignor_account", MCD_CONSIGNOR_ACCOUNT)).strip() or MCD_CONSIGNOR_ACCOUNT
    if "Consignor Email" in r:
        r["Consignor Email"] = str(st.session_state.get("portal_consignor_email", "")).strip()
    if "Entered By" in r:
        r["Entered By"] = str(st.session_state.get("portal_entered_by", "")).strip()

    if "Consignee Name" in r:
        r["Consignee Name"] = (
            _row_value(customer_row, "PostalName")
            or _row_value(customer_row, "CustomerName")
            or _row_value(customer_row, "CustomerCode")
        )
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

    # Keep McDowells remarks mapped to the intended Sage fields:
    # Remarks 1 = first delivery note, Remarks 2 = AnalysisCode2 / other booking notes / extras.
    remark1_note = _clean_delivery_note(_row_value(customer_row, "DeliveryNote1"))
    remarks2_notes = [
        _clean_delivery_note(_row_value(customer_row, "DeliveryNote2")),  # SOPOrderReturns.AnalysisCode2, e.g. Pre 10am
        _clean_delivery_note(_row_value(customer_row, "DeliveryNote3")),
        _clean_delivery_note(_row_value(customer_row, "DeliveryNote4")),
    ]
    remarks2_notes = [x for x in remarks2_notes if x]

    if not remark1_note and not remarks2_notes:
        manual_notes = _notes_or_manual(
            customer_row,
            st.session_state.get("portal_remarks1", ""),
            st.session_state.get("portal_remarks2", ""),
        )
        remark1_note = manual_notes[0] if len(manual_notes) > 0 else ""
        remarks2_notes = manual_notes[1:] if len(manual_notes) > 1 else []

    # McDowells service extras are now represented by the combined Service code
    # (AM, AMTL, BKSL, ECTL, DDAY, etc.), so do not duplicate them in Remarks 2.
    extra_notes: List[str] = []

    if "Remarks 1" in r:
        r["Remarks 1"] = remark1_note
    if "Remarks 2" in r:
        r["Remarks 2"] = " | ".join(remarks2_notes + extra_notes) if (remarks2_notes or extra_notes) else ""

    return r

def _add_to_portal_rows_mcd(rows: List[Dict[str, object]]):
    for r in rows:
        r["_row_id"] = r.get("_row_id") or uuid.uuid4().hex
    st.session_state["portal_rows_mcd"].extend(rows)

# -------------------------
# Portal row builder (Joda/Qargo)
# -------------------------
def _blank_joda_qargo_row() -> Dict[str, object]:
    return {c: "" for c in JODA_QARGO_COLUMNS}

def _row_value(row, key: str) -> str:
    try:
        if isinstance(row, dict):
            val = row.get(key, "")
        else:
            val = row.get(key, "")
        if val is None or (isinstance(val, float) and pd.isna(val)):
            return ""
        return str(val).strip()
    except Exception:
        return ""

def _has_usable_consignee(row) -> bool:
    """True when an imported SO/customer row has enough delivery detail to use directly.

    Do not require postcode here: some Sage rows can still have useful
    name/address/contact details even when postcode is blank or unmapped.
    """
    if not row:
        return False
    keys = [
        "PostalName", "CustomerName", "CustomerCode",
        "Address1", "Address2", "Address3", "Address4",
        "Contact", "Tel", "Email",
    ]
    return any(_row_value(row, k) for k in keys)

def _clean_delivery_note(value: str) -> str:
    val = str(value or "").strip()
    if not val or val.lower() in {"nan", "none", "nat"}:
        return ""
    return val


def _is_ampm_delivery_note(value: str) -> bool:
    """Return True for SO-upload notes that mean AM/AM-PM delivery.

    These should drive the AM/PM Delivery checkbox, not be repeated as
    free-text portal notes/remarks.
    """
    text = _norm(str(value or "")).replace(".", "")
    compact = text.replace(" ", "").replace("/", "").replace("-", "")
    if compact in {"AM", "AMPM"}:
        return True
    return text.startswith("AM ") or "AM/PM" in text or "AM-PM" in text


def _delivery_note_requests_ampm(row) -> bool:
    return any(
        _is_ampm_delivery_note(_row_value(row, key))
        for key in ["DeliveryNote1", "DeliveryNote2", "DeliveryNote3", "DeliveryNote4"]
    )


def _delivery_notes_from_customer(row) -> List[str]:
    """Delivery notes pulled from named Sage fields, excluding AM/PM flags.

    AM/PM from the SO upload auto-ticks the Optional Extras checkbox instead,
    so it does not need to be repeated in Joda Notes or McDowells Remarks.
    """
    notes: List[str] = []
    for key in ["DeliveryNote1", "DeliveryNote2", "DeliveryNote3", "DeliveryNote4"]:
        val = _clean_delivery_note(_row_value(row, key))
        if val and not _is_ampm_delivery_note(val):
            notes.append(val)
    return notes

def _notes_or_manual(row, manual1: str = "", manual2: str = "") -> List[str]:
    notes = _delivery_notes_from_customer(row)
    if notes:
        return notes
    out = []
    if str(manual1).strip():
        out.append(str(manual1).strip())
    if str(manual2).strip():
        out.append(str(manual2).strip())
    return out

def _qargo_customer_label(row) -> str:
    code = _row_value(row, "CustomerCode")
    name = _row_value(row, "PostalName") or _row_value(row, "CustomerName")
    pc = _row_value(row, "Postcode")
    left = code or name or "Customer"
    if code and name:
        left = f"{code} — {name}"
    return f"{left} — {pc}".strip(" —")

def _joda_collection_details(collection_warehouse: str = "") -> Dict[str, str]:
    wh = str(collection_warehouse or st.session_state.get("warehouse_name", "")).strip()
    if wh == "201 - Skipton 2":
        return {
            "Collection Name": "Solidus Solutions",
            "Collection Address Line 1": "Snaygill Industrial Estate",
            "Collection Address Line 2": "Keighley Road",
            "Collection Address Line 4": "Skipton",
            "Collection Address Line 5": "",
            "Collection Post Code": "BD23 2QR",
        }
    return {
        "Collection Name": "Solidus Solutions",
        "Collection Address Line 1": "Engine Shed Lane",
        "Collection Address Line 2": "Skipton",
        "Collection Address Line 4": "North Yorkshire",
        "Collection Address Line 5": "United Kingdom",
        "Collection Post Code": "BD23 1TX",
    }

def build_portal_row_joda(
    customer_row,
    collection_warehouse: str = "",
    pallets_override: Optional[int] = None,
    weight_override: Optional[float] = None,
) -> Dict[str, object]:
    so = _normalise_so_number(st.session_state["so_number"])
    if not so:
        raise ValueError("SO Number is required before adding a Joda/Qargo portal row.")

    pallets = int(pallets_override if pallets_override is not None else st.session_state["pallets"])
    svc_ui = str(st.session_state["service"]).strip()
    svc_code = "ND" if svc_ui == "Next Day" else "EC"

    r = _blank_joda_qargo_row()
    r["_row_id"] = uuid.uuid4().hex
    r["_so_number"] = so
    r["_consignee_label"] = _qargo_customer_label(customer_row)
    r["_joda_collection_warehouse"] = str(collection_warehouse or st.session_state.get("warehouse_name", "")).strip()

    delivery_name = _row_value(customer_row, "PostalName") or _row_value(customer_row, "CustomerName") or _row_value(customer_row, "CustomerCode")
    collection = _joda_collection_details(r["_joda_collection_warehouse"])

    if weight_override is not None:
        weight_value = round(float(weight_override), 3)
    else:
        weight_value = _joda_weight_for_so(so, pallets)

    delivery_notes = _notes_or_manual(
        customer_row,
        st.session_state.get("portal_remarks1", ""),
        st.session_state.get("portal_remarks2", ""),
    )

    delivery_date_str = _joda_delivery_date_str(svc_code)
    r["_joda_delivery_date"] = delivery_date_str

    values = {
        # Blank here; _add_to_portal_rows_joda allocates the next available job number.
        "Job Number": "",
        "Job Order Number": _joda_po_so_ref(so),
        "Export Dater": _yyyymmdd(date.today()),
        "Account Code": "NPB",
        "Collection Date": _yyyymmdd(date.today()),
        "Delivery Name": delivery_name,
        "Delivery Address Line 1": _row_value(customer_row, "Address1"),
        "Delivery Address Line 2": _row_value(customer_row, "Address2"),
        "Delivery Address Line 4": _row_value(customer_row, "Address3"),
        "Delivery Address Line 5": _row_value(customer_row, "Address4"),
        "Delivery Post Code": _row_value(customer_row, "Postcode"),
        "Delivery Mobile": _row_value(customer_row, "Tel"),
        "Delivery Phone": _row_value(customer_row, "Tel"),
        "Delivery Date": delivery_date_str,
        "Full": pallets,
        "Weight": weight_value,
        "Notes Line 1": delivery_notes[0] if len(delivery_notes) > 0 else "",
        "Notes Line 2": delivery_notes[1] if len(delivery_notes) > 1 else "",
        "Notes Line 3": delivery_notes[2] if len(delivery_notes) > 2 else "",
        "Service": svc_code,
        "Extras": _qargo_extras(svc_code, delivery_date_str),
        "Delivery Time": "",
        "Spaces": pallets,
        "Collection Country": "GB",
        "Delivery Country": "GB",
        "REF 2 Link": _get_po_ref("Joda"),
        "Pallet Type": "P",
    }
    values.update(collection)

    for k, v in values.items():
        if k in r:
            r[k] = v

    return r

def _add_to_portal_rows_joda(rows: List[Dict[str, object]]):
    existing = st.session_state.get("portal_rows_joda", [])
    working = list(existing)
    for r in rows:
        r["_row_id"] = r.get("_row_id") or uuid.uuid4().hex
        if not str(r.get("Job Number", "")).strip():
            r["Job Number"] = _next_joda_job_number(working)
        working.append(r)
    st.session_state["portal_rows_joda"].extend(rows)

def build_portal_row_pch(customer_row: pd.Series) -> Dict[str, object]:
    return build_portal_row_mcd(customer_row)

def _add_to_portal_rows_pch(rows: List[Dict[str, object]]):
    for r in rows:
        r["_row_id"] = r.get("_row_id") or uuid.uuid4().hex
    st.session_state["portal_rows_pch"].extend(rows)

# -------------------------
# Cheapest highlighting
# -------------------------
def _parse_pounds(s: str) -> Optional[float]:
    if not isinstance(s, str) or not s.startswith("£"):
        return None
    try:
        return float(s.strip("£").replace(",", "").strip())
    except Exception:
        return None

def calc_for_area(area_code: str):
    svc = st.session_state["service"]
    allowed_local = set(available_hauliers())
    n = int(st.session_state["pallets"])

    joda_charge_fixed = (7.5 if st.session_state["ampm"] else 0) + (20 if st.session_state["timed"] else 0)
    mcd_charge_fixed = (10 if st.session_state["ampm"] else 0) + (19 if st.session_state["timed"] else 0)
    pch_charge_fixed = (15.0 if st.session_state["ampm"] else 0) + (17.5 if st.session_state["timed"] else 0)

    jb = jf = None
    if "Joda" in allowed_local:
        base = get_base_rate_capped(rate_df_main, area_code, svc, "Joda", n)
        if base is not None:
            base = joda_round_base_up(base)
            eff = joda_effective_pct(n, float(st.session_state["joda_pct"]))
            jb = base
            jf = base * (1 + eff / 100.0) + joda_charge_fixed

    mb = mf = None
    if "Mcdowells" in allowed_local:
        base = get_base_rate_capped(rate_df_main, area_code, svc, "Mcdowells", n)
        if base is not None:
            small_extra = mcd_smallload_extra(n)
            tl_total = (3.90 if st.session_state["tail"] else 0.0) * n
            base_calc = float(base) + float(small_extra)
            mb = base_calc
            mf = base_calc * (1 + float(st.session_state["mcd_pct"]) / 100.0) + mcd_charge_fixed + tl_total

    pb = pf = None
    if "Pc Howard" in allowed_local and not rate_df_pch.empty:
        base = get_base_rate_capped(rate_df_pch, area_code, svc, "Pc Howard", n)
        if base is not None:
            pb = float(base)
            pf = float(base) * (1 + float(st.session_state["pch_pct"]) / 100.0) + pch_charge_fixed

    return jb, jf, mb, mf, pb, pf

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
# Sage export line builder (all hauliers)
# -------------------------
def build_export_lines_for_haulier_sage(haulier: str) -> List[Dict[str, object]]:
    """
    Delivery line uses BASE price (no fuel).
    Fuel surcharge exports as its own line (qty=1, £=total surcharge).
    Pricing caps at max pallet band while qty can be > max.
    """
    so = _normalise_so_number(st.session_state["so_number"])
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

    # Imported total weight for portal exports.
    # Prefer column headers so the calculation survives column order changes:
    # SOPOrderReturnLines.LineQuantity * StockItems.Weight * 1000 (tonnes to kg).
    cols_lower = {str(c).strip().lower(): c for c in df.columns}
    qty_col = cols_lower.get("soporderreturnlines.linequantity") or cols_lower.get("linequantity")
    weight_col = cols_lower.get("stockitems.weight") or cols_lower.get("weight")

    if qty_col and weight_col:
        qty = pd.to_numeric(df[qty_col], errors="coerce")
        weight = pd.to_numeric(df[weight_col], errors="coerce")
        df["_PortalWeightCalc"] = qty * weight * 1000
    elif df.shape[1] > 32:
        # Fallback only: Excel AC * AG if headers are unavailable.
        ac = pd.to_numeric(df.iloc[:, 28], errors="coerce")
        ag = pd.to_numeric(df.iloc[:, 32], errors="coerce")
        df["_PortalWeightCalc"] = ac * ag * 1000

    return df


def col_pick(df: pd.DataFrame, candidates: List[str]) -> Optional[str]:
    cols = {str(c).strip().lower(): c for c in df.columns}

    for cand in candidates:
        k = cand.strip().lower()
        if k in cols:
            return cols[k]

    return None

def _postcode_area(postcode: str) -> str:
    pc = _norm(str(postcode)).replace(" ", "")
    letters = ""
    for ch in pc:
        if ch.isalpha():
            letters += ch
        else:
            break
    return letters


def _postcode_letters_and_district(postcode: str):
    """Return ('PE', 12) from a postcode/outward code such as PE12 6JR.

    Important: use the outward postcode only. If we simply remove spaces,
    PE12 6JR becomes PE126JR and the district is wrongly read as 126.
    """
    text = _norm(str(postcode)).upper()
    compact = text.replace(" ", "")

    if not compact:
        return "", None

    # UK full postcodes have a 3-character inward code at the end.
    # With a space, take the first part. Without a space, strip the last 3 chars.
    if " " in text:
        outward = text.split()[0].replace(" ", "")
    elif len(compact) > 3:
        outward = compact[:-3]
    else:
        outward = compact

    letters = ""
    i = 0
    while i < len(outward) and outward[i].isalpha():
        letters += outward[i]
        i += 1

    digits = ""
    while i < len(outward) and outward[i].isdigit():
        digits += outward[i]
        i += 1

    try:
        district = int(digits) if digits else None
    except Exception:
        district = None

    return letters, district


def _postcode_area_matches_option(postcode: str, option: str) -> bool:
    """Match a postcode to rate-sheet options like PE 1-20, PE21-29, PE 30 or PR5."""
    letters, district = _postcode_letters_and_district(postcode)
    if not letters:
        return False

    opt = _norm(str(option))
    opt_compact = opt.replace(" ", "")

    # Exact outward-code style match, e.g. PR56AJ/PR5 or PE12.
    if district is not None and opt_compact == f"{letters}{district}":
        return True

    # Broad area match, e.g. PE.
    if opt_compact == letters:
        return True

    # Range/single district format after the area letters, e.g. PE 1-20, PE21-29, PE 30.
    if not opt_compact.startswith(letters):
        return False

    rest = opt_compact[len(letters):]
    if not rest or district is None:
        return False

    if "-" in rest:
        start_s, end_s = rest.split("-", 1)
        try:
            return int(start_s) <= district <= int(end_s)
        except Exception:
            return False

    try:
        return int(rest) == district
    except Exception:
        return False


def _resolve_postcode_area_option(postcode: str, area_options: List[str]) -> str:
    """Resolve PE12 6JR to the actual rate-sheet option, e.g. PE 1-20."""
    letters, district = _postcode_letters_and_district(postcode)
    if not letters:
        return ""

    # Prefer exact outward-code options first.
    if district is not None:
        exact = f"{letters}{district}"
        for option in area_options:
            if _norm(str(option)).replace(" ", "") == exact:
                return str(option)

    # Then rate-sheet range/single options.
    for option in area_options:
        if _postcode_area_matches_option(postcode, str(option)):
            return str(option)

    # Fallback to the old behaviour so the app still shows a useful area if no range exists.
    return letters

def _safe_str(x) -> str:
    if x is None:
        return ""
    if isinstance(x, float) and pd.isna(x):
        return ""
    return str(x).strip()

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

    if "_PortalWeightCalc" in tmp.columns:
        tmpw = tmp[[so_col, "_PortalWeightCalc"]].copy()
        tmpw["_PortalWeightCalc"] = pd.to_numeric(tmpw["_PortalWeightCalc"], errors="coerce")
        wt = tmpw.groupby(so_col)["_PortalWeightCalc"].sum(min_count=1)
        out = out.merge(wt.rename("Weight"), left_on="SO", right_index=True, how="left")
    else:
        out["Weight"] = pd.NA

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

    # Delivery/booking instructions should be picked by Sage header name, not by column position.
    # AnalysisCode2 is used for entries such as "Pre 10am" and should feed Remarks 2.
    delivery_note1_col = col_pick(df, ["SOPOrderReturns.AnalysisCode3", "AnalysisCode3"])
    delivery_note2_col = col_pick(df, ["SOPOrderReturns.AnalysisCode2", "AnalysisCode2"])
    delivery_note3_col = col_pick(df, ["SOPOrderReturns.AnalysisCode5", "AnalysisCode5"])
    delivery_note4_col = col_pick(df, ["SOPOrderReturns.AnalysisCode4", "AnalysisCode4"])

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

    def get_by_pos(pos: int) -> str:
        try:
            if len(r0) > pos:
                return _safe_str(r0.iloc[pos])
        except Exception:
            pass
        return ""

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
        # Delivery information from the uploaded Sage sheet.
        # Use named Sage fields only. Do not fall back to column positions,
        # because those can accidentally pull address fields such as City/Town into notes.
        "DeliveryNote1": get(delivery_note1_col),
        "DeliveryNote2": get(delivery_note2_col),
        "DeliveryNote3": get(delivery_note3_col),
        "DeliveryNote4": get(delivery_note4_col),
    }

    if not out["PostalName"]:
        out["PostalName"] = out["CustomerName"] or out["CustomerCode"]

    return out

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

# Use explicit navigation rather than Streamlit tabs so Export/Customers
# cannot render underneath the Table page if the browser keeps tab content in view.
selected_page = st.radio(
    "Section",
    options=["Table", "Export", "Customers"],
    horizontal=True,
    label_visibility="collapsed",
    key="main_page",
)

# -------------------------
# TABLE TAB
# -------------------------
if selected_page == "Table":
    st.header("Sales Orders (preferred)")

    wh_now = st.session_state.get("warehouse_name", WAREHOUSE_OPTIONS[0])
    allowed_now = set(WAREHOUSE_HAULIERS.get(wh_now, []))
    pc_only_now = allowed_now == {"Pc Howard"}
    area_options_now = unique_areas_pch if pc_only_now else unique_areas_main

    upl = st.file_uploader(
        "Upload Sage Sales Order export (.xlsx)",
        type=["xlsx"],
        help="Upload the Sage sales order export. The app uses it to pre-fill SO number, postcode, consignee, promised date, notes, pallets and weight.",
        key="sage_so_file",
    )

    # Keep the parsed Sage upload in session state so switching to Export/Customers
    # does not force the user to re-upload it when returning to the Table page.
    so_summary = st.session_state.get("_sage_so_summary", pd.DataFrame())
    so_df_full = st.session_state.get("_sage_so_df_full", pd.DataFrame())

    if upl is not None:
        try:
            so_df_full = load_sage_sales_export(upl)
            so_summary = build_so_summary(so_df_full)

            st.session_state["_sage_so_df_full"] = so_df_full
            st.session_state["_sage_so_summary"] = so_summary
            st.session_state["_sage_so_file_name"] = getattr(upl, "name", "Uploaded Sage SO export")

            try:
                weight_by_so = {}
                for _, r in so_summary.iterrows():
                    raw_so = str(r.get("SO", "")).strip()
                    if raw_so and pd.notna(r.get("Weight", pd.NA)):
                        val = round(float(r.get("Weight", 0)), 3)
                        weight_by_so[raw_so] = val
                        weight_by_so[_normalise_so_number(raw_so)] = val
                st.session_state["_so_weight_by_so"] = weight_by_so
            except Exception:
                st.session_state["_so_weight_by_so"] = {}
            st.session_state["sage_so_uploaded"] = True
        except Exception as e:
            st.error(f"Could not read upload: {e}")

    if not so_summary.empty:
        cached_name = str(st.session_state.get("_sage_so_file_name", "Uploaded Sage SO export")).strip()
        st.caption(f"Using uploaded SO file: {cached_name} ({len(so_summary):,} sales order(s) loaded).")
        if st.button("Clear uploaded SO file", key="clear_sage_so_upload"):
            for k in [
                "_sage_so_df_full",
                "_sage_so_summary",
                "_sage_so_file_name",
                "_so_weight_by_so",
                "_so_weight",
                "_so_consignee",
                "_last_so_applied",
                "_last_so_area_applied",
                "sage_so_selected",
            ]:
                st.session_state.pop(k, None)
            st.session_state["sage_so_uploaded"] = False
            st.rerun()

    done_sos = set(st.session_state.get("done_sos", []))
    show_done = st.checkbox("Show completed SOs", key="show_done_sos", value=False)

    if not so_summary.empty:
        so_search = st.text_input(
            "Search SOs",
            key="sage_so_search",
            placeholder="SO / customer / postcode"
        )
        ss = _norm(so_search)

        shown = so_summary.copy()

        if not show_done and done_sos:
            shown = shown[~shown["SO"].astype(str).map(_normalise_so_number).isin(done_sos)].copy()

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

        picked = st.selectbox(
            "Select Sales Order",
            options=so_options,
            key="sage_so_selected",
            format_func=_so_fmt
        )

        if picked:
            r0 = so_summary.loc[so_summary["SO"].astype(str) == str(picked)].iloc[0]
            pre_area = _resolve_postcode_area_option(str(r0.get("Postcode", "")), area_options_now)

            # Re-apply when either the selected SO changes OR the resolved rate-sheet area changes.
            # This fixes older sessions where the same SO was already marked as applied before
            # postcode ranges such as "PE 1-20" were understood. Manual postcode changes still
            # remain possible after the SO has been applied once for the current resolved area.
            apply_selected_so = (
                str(picked) != st.session_state.get("_last_so_applied", "")
                or str(pre_area) != st.session_state.get("_last_so_area_applied", "")
            )

            if apply_selected_so:
                if pre_area and pre_area in area_options_now:
                    st.session_state["area"] = pre_area

                st.session_state["so_number"] = _normalise_so_number(picked)

                try:
                    pe = r0.get("PalletsEst", pd.NA)
                    if pd.notna(pe):
                        st.session_state["pallets"] = max(1, int(math.ceil(float(pe))))
                        sync_service_from_pallets()
                except Exception:
                    pass

                try:
                    wt = r0.get("Weight", pd.NA)
                    st.session_state["_so_weight"] = round(float(wt), 3) if pd.notna(wt) else ""
                except Exception:
                    st.session_state["_so_weight"] = ""

                promised_dt = _parse_date_or_none(r0.get("PromisedDate", ""))
                if promised_dt is not None:
                    st.session_state["portal_delivery_date"] = promised_dt
                    st.session_state["joda_delivery_date"] = promised_dt
                else:
                    default_delivery_dt = _joda_automatic_delivery_date(st.session_state.get("service", ""))
                    st.session_state["portal_delivery_date"] = default_delivery_dt
                    st.session_state["joda_delivery_date"] = default_delivery_dt

                so_consignee = extract_consignee_from_so(
                    so_df_full,
                    str(picked)
                )
                st.session_state["_so_consignee"] = so_consignee

                # If Sage says AM/AM-PM on the SO, tick the proper Optional Extras flag.
                # The raw AM text is filtered out of portal notes/remarks to avoid duplicates.
                st.session_state["ampm"] = bool(_delivery_note_requests_ampm(so_consignee))

                st.session_state["_last_so_applied"] = str(picked)
                st.session_state["_last_so_area_applied"] = str(pre_area)

    st.markdown("---")
    
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
        # Allow > 26; pricing caps to max band
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

    if allowed:
        st.markdown("**Portal delivery date**")
        if _parse_date_or_none(st.session_state.get("portal_delivery_date")) is None:
            st.session_state["portal_delivery_date"] = _joda_automatic_delivery_date(st.session_state.get("service", ""))
        date_col, prebook_col = st.columns([1.4, 1.0], gap="medium")
        with date_col:
            st.date_input(
                "Delivery date for portal exports",
                key="portal_delivery_date",
                help="Starts with the promised delivery date from Sage where available. Change it here before adding the job if the delivery needs booking for a different day.",
            )
        with prebook_col:
            st.checkbox(
                "Pre-Booked",
                key="portal_prebooked",
                help="Tick Pre-Booked when pallets must be delivered on a specific agreed day. It adds Pre-Booked to the portal export notes/extras.",
            )
        st.session_state["joda_delivery_date"] = st.session_state.get("portal_delivery_date")
        if st.session_state.get("portal_prebooked", False):
            st.caption("Pre-Booked is on: the portal exports will flag that this delivery is booked for a specific agreed day.")
        else:
            st.caption("Delivery date starts from the selected SO promised date where available. Use Pre-Booked only when the delivery date has been specifically agreed.")

    if st.session_state["dual"] and int(st.session_state["pallets"]) == 1:
        st.error("Dual Collection requires at least 2 pallets.")
        st.stop()

    if st.session_state["dual"] and not pc_only:
        st.caption("Dual collection split for Joda/Qargo export")
        total_pallets = int(st.session_state.get("pallets", 1))
        s1, s2 = st.columns(2)
        with s1:
            st.number_input("101 - Skipton pallets", min_value=0, max_value=total_pallets, step=1, key="split1")
        with s2:
            st.number_input("201 - Skipton 2 pallets", min_value=0, max_value=total_pallets, step=1, key="split2")
        if int(st.session_state.get("split1", 0)) + int(st.session_state.get("split2", 0)) != total_pallets:
            st.warning("For split collections, 101 + 201 pallets should equal the total number of pallets.")

    st.markdown("---")

    # Calculated rates
    jb, jf, mb, mf, pb, pf = calc_for_area(st.session_state["area"])
    summary_rows = []
    if "Joda" in allowed:
        summary_rows.append({"Haulier": "Joda", "Base Rate": "No rate" if jb is None else f"£{float(jb):,.2f}",
                             "Fuel Surcharge (%)": f"{float(st.session_state['joda_pct']):.2f}%",
                             "Final Rate": "N/A" if jf is None else f"£{float(jf):,.2f}"})
    if "Mcdowells" in allowed:
        summary_rows.append({"Haulier": "McDowells", "Base Rate": "No rate" if mb is None else f"£{float(mb):,.2f}",
                             "Fuel Surcharge (%)": f"{float(st.session_state['mcd_pct']):.2f}%",
                             "Final Rate": "N/A" if mf is None else f"£{float(mf):,.2f}"})
    if "Pc Howard" in allowed:
        summary_rows.append({"Haulier": "PC Howard", "Base Rate": "No rate" if pb is None else f"£{float(pb):,.2f}",
                             "Fuel Surcharge (%)": f"{float(st.session_state['pch_pct']):.2f}%",
                             "Final Rate": "N/A" if pf is None else f"£{float(pf):,.2f}"})

    st.subheader("3. Calculated Rates")
    if summary_rows:
        df = pd.DataFrame(summary_rows).set_index("Haulier")
        st.table(df.style.apply(highlight_cheapest_factory(), axis=1))
    else:
        st.info("No hauliers available for this warehouse.")

    st.markdown("---")
    st.subheader("Add to Export Lists")

    _clear_so_on_next_run()

    top1, top2, top3 = st.columns([1, 1, 2], gap="medium")
    with top1:
        st.text_input("SO Number", key="so_number", placeholder="e.g. 020502")
    with top2:
        st.write(f"Warehouse: **{st.session_state['warehouse_name']}**")
    with top3:
        customers_df = load_customers_df()
        q = _norm(st.text_input("Customer search", key="cust_search", placeholder="code / name / postcode…"))
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
        label_map = {row["ID"]: customer_label(row) for _, row in filtered.iterrows()}

        st.selectbox(
            "Consignee (for portal row)",
            options=options,
            key="cust_selected_id",
            format_func=lambda x: "— Select —" if x == "" else label_map.get(x, x),
            disabled=("Mcdowells" not in allowed and "Joda" not in allowed and "Pc Howard" not in allowed),
        )

    btns = st.columns([1, 1, 1, 2])

    # Always show the buttons. Disable unavailable hauliers rather than hiding them,
    # so this section cannot appear to vanish if the warehouse/allowed list is unexpected.
    joda_available = "Joda" in allowed
    mcd_available = "Mcdowells" in allowed
    pch_available = "Pc Howard" in allowed

    with btns[3]:
        st.caption(
            "Available: "
            + (", ".join(display_haulier(x) for x in sorted(allowed)) if allowed else "none")
        )

    if btns[0].button("Add Joda", use_container_width=True, disabled=not joda_available):
        try:
            _add_to_sage_basket(build_export_lines_for_haulier_sage("Joda"))

            so_con = st.session_state.get("_so_consignee", {}) or {}
            if _has_usable_consignee(so_con):
                joda_customer = so_con
            else:
                cid = str(st.session_state.get("cust_selected_id", "")).strip()
                if not cid:
                    raise ValueError("Select a consignee (customer) for the Joda/Qargo portal row.")
                crow = customers_df.loc[customers_df["ID"] == cid]
                if crow.empty:
                    raise ValueError("Selected customer not found in customers.xlsx.")
                joda_customer = crow.iloc[0]

            if st.session_state.get("dual"):
                total_pallets = max(1, int(st.session_state.get("pallets", 1)))
                split_101 = int(st.session_state.get("split1", 0) or 0)
                split_201 = int(st.session_state.get("split2", 0) or 0)
                if split_101 + split_201 != total_pallets:
                    raise ValueError("For Joda split collections, 101 + 201 pallets must equal the total number of pallets.")

                total_weight = st.session_state.get("_so_weight", "")
                try:
                    total_weight = float(total_weight)
                except Exception:
                    total_weight = None

                joda_rows = []
                if split_101 > 0:
                    weight_101 = (total_weight * split_101 / total_pallets) if total_weight is not None else None
                    joda_rows.append(build_portal_row_joda(joda_customer, "101 - Skipton", split_101, weight_101))
                if split_201 > 0:
                    weight_201 = (total_weight * split_201 / total_pallets) if total_weight is not None else None
                    joda_rows.append(build_portal_row_joda(joda_customer, "201 - Skipton 2", split_201, weight_201))
                _add_to_portal_rows_joda(joda_rows)
            else:
                _add_to_portal_rows_joda([build_portal_row_joda(joda_customer)])

            mark_so_done(st.session_state["so_number"])
            st.session_state["_clear_so_next"] = True
            st.success("Added Joda lines (+ Qargo portal row).")
            st.rerun()
        except Exception as e:
            st.error(str(e))

    if btns[1].button("Add McDowells", use_container_width=True, disabled=not mcd_available):
        try:
            _add_to_sage_basket(build_export_lines_for_haulier_sage("Mcdowells"))

            so_con = st.session_state.get("_so_consignee", {}) or {}
            if _has_usable_consignee(so_con):
                mcd_customer = so_con
            else:
                cid = str(st.session_state.get("cust_selected_id", "")).strip()
                if not cid:
                    raise ValueError("Select a consignee (customer) for the portal row.")
                crow = customers_df.loc[customers_df["ID"] == cid]
                if crow.empty:
                    raise ValueError("Selected customer not found in customers.xlsx.")
                mcd_customer = crow.iloc[0]
            _add_to_portal_rows_mcd([build_portal_row_mcd(mcd_customer)])

            mark_so_done(st.session_state["so_number"])
            st.session_state["_clear_so_next"] = True
            st.success("Added McDowells lines (+ portal row).")
            st.rerun()
        except Exception as e:
            st.error(str(e))

    if btns[2].button("Add PC Howard", use_container_width=True, disabled=not pch_available):
        try:
            _add_to_sage_basket(build_export_lines_for_haulier_sage("Pc Howard"))

            so_con = st.session_state.get("_so_consignee", {}) or {}
            if _has_usable_consignee(so_con):
                pch_customer = so_con
            else:
                cid = str(st.session_state.get("cust_selected_id", "")).strip()
                if not cid:
                    raise ValueError("Select a consignee (customer) for the PC Howard portal row.")
                crow = customers_df.loc[customers_df["ID"] == cid]
                if crow.empty:
                    raise ValueError("Selected customer not found in customers.xlsx.")
                pch_customer = crow.iloc[0]
            _add_to_portal_rows_pch([build_portal_row_pch(pch_customer)])

            mark_so_done(st.session_state["so_number"])
            st.session_state["_clear_so_next"] = True
            st.success("Added PC Howard lines (+ portal row).")
            st.rerun()
        except Exception as e:
            st.error(str(e))

# -------------------------
# EXPORT TAB (Download + Clear at TOP)
# -------------------------
if selected_page == "Export":
    st.header("Exports")

    # Daily PO refs
    with st.expander("Daily PO refs (today only)", expanded=True):
        st.caption("These apply to all Sage PO lines. Saving also updates any lines already added below.")
        c1, c2, c3, c4 = st.columns([1, 1, 1, 1.2], gap="medium")
        with c1:
            st.number_input("Joda PO Number", min_value=1, step=1, key="po_ref_joda")
            st.caption("default: 1")
        with c2:
            st.number_input("McDowells PO Number", min_value=1, step=1, key="po_ref_mcd")
            st.caption("default: 3")
        with c3:
            st.number_input("PC Howard PO Number", min_value=1, step=1, key="po_ref_pch")
            st.caption("default: 5")

        with c4:
            if st.button("Save PO refs for today", use_container_width=True):
                save_porefs_for_today(
                    _safe_int(st.session_state.get("po_ref_joda"), 1),
                    _safe_int(st.session_state.get("po_ref_mcd"), 3),
                    _safe_int(st.session_state.get("po_ref_pch"), 5),
                )
                apply_po_refs_to_existing_lines()
                st.success("Saved for today and updated existing Sage lines.")
                st.rerun()

    # Sage PO export
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
            st.caption(f"{len(basket)} line(s) in Sage export." if basket else "No Sage lines yet.")

        st.divider()

        if not basket:
            st.info("No Sage PO lines saved yet. Use the Table tab to add lines.")
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

    # Joda/Qargo portal export
    with st.expander("Portal Export — Joda / Qargo (CSV)", expanded=True):
        rows = st.session_state.get("portal_rows_joda", [])
        _ensure_joda_job_numbers(rows)
        _ensure_joda_refs_and_weights(rows)

        export_joda_df = pd.DataFrame(rows).reindex(columns=JODA_QARGO_COLUMNS) if rows else pd.DataFrame(columns=JODA_QARGO_COLUMNS)
        export_joda_df = export_joda_df.where(pd.notnull(export_joda_df), "")
        joda_bytes = export_joda_df.to_csv(index=False, sep=",", na_rep="", lineterminator="\n", quoting=csvlib.QUOTE_MINIMAL).encode("utf-8")

        top = st.columns([1.4, 1.0, 3.6])
        with top[0]:
            st.download_button(
                label="Download Joda / Qargo Portal CSV",
                data=joda_bytes,
                file_name=f"Joda_Qargo_Portal_{date.today().strftime('%Y%m%d')}.csv",
                mime="text/csv",
                use_container_width=True,
                disabled=(len(rows) == 0),
            )
        with top[1]:
            if st.button("Clear all", use_container_width=True, disabled=(len(rows) == 0), key="clear_joda_qargo"):
                st.session_state["portal_rows_joda"] = []
                st.rerun()
        with top[2]:
            st.caption(f"{len(rows)} row(s) in Joda/Qargo export." if rows else "No Joda/Qargo rows yet.")

        st.caption("Uses Qargo Import Template.xlsx. Job numbers self-allocate as YYMMDD001, YYMMDD002, etc. Tick rows below to combine them onto one job number, or split a job into 101/201 collection rows. Weight comes from SOPOrderReturnLines.LineQuantity × StockItems.Weight × 1000 in the Sage SO import where available. Delivery Date uses the shared portal delivery date selected on the Table tab.")
        st.divider()

        if not rows:
            st.info("No Joda/Qargo portal rows saved yet. Use the Table tab → Add Joda.")
        else:
            # Apply any queued checkbox reset before the checkbox widgets are created.
            _apply_pending_joda_selection_clear()

            h = st.columns([0.55, 1.0, 1.25, 1.2, 2.0, 1.1, 0.75, 0.85, 0.75])
            h[0].markdown("**Pick**")
            h[1].markdown("**Job**")
            h[2].markdown("**Order**")
            h[3].markdown("**Collection**")
            h[4].markdown("**Consignee**")
            h[5].markdown("**Postcode**")
            h[6].markdown("**Pallets**")
            h[7].markdown("**Weight**")
            h[8].markdown("**Remove**")
            st.divider()

            remove_id = None
            for r in rows:
                rid = str(r.get("_row_id", ""))
                cols = st.columns([0.55, 1.0, 1.25, 1.2, 2.0, 1.1, 0.75, 0.85, 0.75])
                cols[0].checkbox("", key=f"sel_joda_{rid}", label_visibility="collapsed")
                cols[1].write(r.get("Job Number", ""))
                cols[2].write(r.get("Job Order Number", ""))
                cols[3].write(r.get("_joda_collection_warehouse", "") or st.session_state.get("warehouse_name", ""))
                cols[4].write(r.get("_consignee_label", "") or r.get("Delivery Name", ""))
                cols[5].write(r.get("Delivery Post Code", ""))
                cols[6].write(r.get("Full", ""))
                cols[7].write(r.get("Weight", ""))
                if cols[8].button("🗑", key=f"rm_portal_joda_{rid}", help="Remove this Joda/Qargo row"):
                    remove_id = rid

            if remove_id:
                st.session_state["portal_rows_joda"] = [x for x in st.session_state["portal_rows_joda"] if x.get("_row_id") != remove_id]
                st.rerun()

            selected_joda_ids = _selected_joda_row_ids()
            action_cols = st.columns([1.1, 1.1, 1.1, 3.0])
            if action_cols[0].button("Combine selected", use_container_width=True, disabled=(len(selected_joda_ids) < 2)):
                try:
                    _combine_selected_joda_rows(selected_joda_ids)
                    st.success("Selected Joda rows combined onto one job number.")
                    st.rerun()
                except Exception as e:
                    st.error(str(e))

            if action_cols[1].button("Split selected", use_container_width=True, disabled=(len(selected_joda_ids) == 0)):
                st.session_state["_joda_split_mode"] = True
                st.session_state["_joda_split_ids"] = selected_joda_ids
                st.rerun()

            if action_cols[2].button("Clear selection", use_container_width=True, disabled=(len(selected_joda_ids) == 0)):
                _clear_joda_selection(selected_joda_ids)
                st.session_state["_joda_split_mode"] = False
                st.session_state["_joda_split_ids"] = []
                st.rerun()

            action_cols[3].caption("Non-split jobs get their own YYMMDD sequence job number automatically. Combine gives selected rows one new shared job number. Split replaces a selected row with separate 101/201 collection rows.")

            if st.session_state.get("_joda_split_mode"):
                split_ids = [x for x in st.session_state.get("_joda_split_ids", []) if x in {str(r.get("_row_id", "")) for r in st.session_state.get("portal_rows_joda", [])}]
                if split_ids:
                    st.markdown("#### Split selected Joda job(s)")
                    st.caption("Enter how many pallets should collect from each site. Each split row will receive its own new job number.")
                    for row in st.session_state.get("portal_rows_joda", []):
                        rid = str(row.get("_row_id", ""))
                        if rid not in split_ids:
                            continue
                        total_p = _safe_int(row.get("Full", 0), 0)
                        sc1, sc2, sc3 = st.columns([1.4, 1.0, 1.0])
                        sc1.write(f"{row.get('Job Order Number', '')} — {row.get('_consignee_label', '') or row.get('Delivery Name', '')} ({total_p} pallets)")
                        sc2.number_input("101 pallets", min_value=0, max_value=max(total_p, 0), value=total_p, step=1, key=f"split_joda_{rid}_101")
                        sc3.number_input("201 pallets", min_value=0, max_value=max(total_p, 0), value=0, step=1, key=f"split_joda_{rid}_201")

                    sp1, sp2, sp3 = st.columns([1.2, 1.2, 3])
                    if sp1.button("Apply split", use_container_width=True):
                        try:
                            _apply_joda_split_from_inputs(split_ids)
                            st.success("Selected Joda row(s) split into separate collection jobs.")
                            st.rerun()
                        except Exception as e:
                            st.error(str(e))
                    if sp2.button("Cancel split", use_container_width=True):
                        st.session_state["_joda_split_mode"] = False
                        st.session_state["_joda_split_ids"] = []
                        st.rerun()
                else:
                    st.session_state["_joda_split_mode"] = False
                    st.session_state["_joda_split_ids"] = []

    # McDowells portal export
    with st.expander("Portal Export — McDowells (CSV)", expanded=True):
        rows = st.session_state.get("portal_rows_mcd", [])

        export_mcd_df = pd.DataFrame(rows).reindex(columns=MCD_PORTAL_COLUMNS) if rows else pd.DataFrame(columns=MCD_PORTAL_COLUMNS)
        export_mcd_df = export_mcd_df.where(pd.notnull(export_mcd_df), "")
        mcd_bytes = export_mcd_df.to_csv(index=False, sep=",", na_rep="", lineterminator="\n", quoting=csvlib.QUOTE_MINIMAL).encode("utf-8")

        top = st.columns([1.4, 1.0, 3.6])
        with top[0]:
            st.download_button(
                label="Download McDowells Portal CSV",
                data=mcd_bytes,
                file_name=f"McDowells_Portal_{date.today().strftime('%Y%m%d')}.csv",
                mime="text/csv",
                use_container_width=True,
                disabled=(len(rows) == 0),
            )
        with top[1]:
            if st.button("Clear all", use_container_width=True, disabled=(len(rows) == 0), key="clear_mcd"):
                st.session_state["portal_rows_mcd"] = []
                st.rerun()
        with top[2]:
            st.caption(f"{len(rows)} row(s) in portal export." if rows else "No portal rows yet.")

        st.caption("Depots fixed to 008. Service codes: Economy=2D, Next Day=ND.")
        st.divider()

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
                if cols[4].button("🗑", key=f"rm_portal_mcd_{rid}", help="Remove this portal row"):
                    remove_id = rid

            if remove_id:
                st.session_state["portal_rows_mcd"] = [x for x in st.session_state["portal_rows_mcd"] if x.get("_row_id") != remove_id]
                st.rerun()


    # PC Howard portal export (placeholder format)
    with st.expander("Portal Export — PC Howard (CSV)", expanded=True):
        rows = st.session_state.get("portal_rows_pch", [])

        export_pch_df = pd.DataFrame(rows).reindex(columns=MCD_PORTAL_COLUMNS) if rows else pd.DataFrame(columns=MCD_PORTAL_COLUMNS)
        export_pch_df = export_pch_df.where(pd.notnull(export_pch_df), "")
        pch_bytes = export_pch_df.to_csv(index=False, sep=",", na_rep="", lineterminator="\n", quoting=csvlib.QUOTE_MINIMAL).encode("utf-8")

        top = st.columns([1.4, 1.0, 3.6])
        with top[0]:
            st.download_button(
                label="Download PC Howard Portal CSV",
                data=pch_bytes,
                file_name=f"PC_Howard_Portal_{date.today().strftime('%Y%m%d')}.csv",
                mime="text/csv",
                use_container_width=True,
                disabled=(len(rows) == 0),
            )
        with top[1]:
            if st.button("Clear all", use_container_width=True, disabled=(len(rows) == 0), key="clear_pch"):
                st.session_state["portal_rows_pch"] = []
                st.rerun()
        with top[2]:
            st.caption(f"{len(rows)} row(s) in PC Howard portal export." if rows else "No PC Howard portal rows yet.")

        st.caption("Placeholder portal export using the same layout as McDowells until PC Howard provide a template.")
        st.divider()

        if not rows:
            st.info("No PC Howard portal rows saved yet. Use the Table tab → Add PC Howard.")
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
                if cols[4].button("🗑", key=f"rm_portal_pch_{rid}", help="Remove this PC Howard portal row"):
                    remove_id = rid

            if remove_id:
                st.session_state["portal_rows_pch"] = [x for x in st.session_state["portal_rows_pch"] if x.get("_row_id") != remove_id]
                st.rerun()

# -------------------------
# CUSTOMERS TAB (generic address book)
# -------------------------
if selected_page == "Customers":
    st.header("Customers (customers.xlsx)")
    st.caption("Edits here write back to customers.xlsx. This will be shared across all future portal exports.")

    customers_df = load_customers_df()

    q = _norm(st.text_input("Search (code / name / postcode)", key="ab_search", placeholder="e.g. A0003 or BD7…"))
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
        if cols[3].button("✏️", key=f"ab_edit_{rid}", help="Edit"):
            edit_id = rid
        if cols[4].button("🗑", key=f"ab_del_{rid}", help="Delete"):
            delete_id = rid

    if delete_id:
        customers_df = customers_df[customers_df["ID"] != delete_id].copy()
        save_customers_df(customers_df)
        if st.session_state.get("ab_selected_id") == delete_id:
            st.session_state["ab_selected_id"] = ""
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
