# app.py
import os
import math
import json
from datetime import date, datetime

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
        V2.0.0    
        Enter a UK postcode area, select a service type, set pallets and surcharges,
        and optionally add AM/PM, Tail Lift or Timed Delivery. Dual Collection splits the load.

    **What’s NEW in the Version 2 Release?**    
    **NEW:** Map View (Beta) is live.    
    **NEW:** History tab.    
    **NEW:** Rate Cards Updated for 2026.    
    Note: From 01/01/26 Joda fuel surcharge does not apply on 1–6 pallet quantities (per group when split). McDowells rates have a £5 charge per pallet applied below 5 pallets.
        """,
        unsafe_allow_html=True
    )

# -------------------------
# Config / constants
# -------------------------
DATA_FILE = "joda_surcharge.json"
RATE_XLSX = "haulier prices 2.xlsx"

# Export template CSV (the example you provided). App will mirror these columns.
EXPORT_TEMPLATE_CSV = "PO Import Example File.csv"  # put alongside app.py in your repo
JODA_ACC = "J040"
MCD_ACC = "M127"
JODA_PO_GROUP = 1
MCD_PO_GROUP = 2

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

    # reset on Wednesdays (weekday() == 2)
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
        if isinstance(c, (int, float)) or (isinstance(c, str) and c.isdigit())
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

    # NEW: export basket + defaults used when adding
    st.session_state.setdefault("export_basket", [])  # list[dict]
    st.session_state.setdefault("warehouse_name", "101 - Skipton")
    st.session_state.setdefault("so_number", "")  # manual input, required to add lines

_ensure_defaults()

# -------------------------
# Load from History (existing)
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
    # £5 per pallet if under 5 pallets, applied to up to 4 pallets
    return (5.0 * min(pallet_count, 4)) if pallet_count < 5 else 0.0

def _ddmmyyyy(d: date) -> str:
    return d.strftime("%d/%m/%Y")

# -------------------------
# Export template columns (mirror the provided CSV)
# -------------------------
@st.cache_data
def load_export_template_columns(path: str) -> list[str]:
    if os.path.exists(path):
        try:
            tmp = pd.read_csv(path)
            return tmp.columns.tolist()
        except Exception:
            pass

    # Fallback (if template not found) — these are the columns from your example file:
    return [
        "Purchase Order Import Type",
        "Purchase Order Number",
        "Purchase Order Supplier Acc Code",
        "Purchase Order Document Date",
        "Purchase Order Header Requested Date",
        "Purchase Order Discount Percent",
        "Purchase Order Supplier Document No.",
        "Purchase Order By Default Supply To",
        "Purchase Order AnalysisCode 1",
        "Purchase Order AnalysisCode 2",
        "Purchase Order AnalysisCode 3",
        "Purchase Order AnalysisCode 4",
        "Purchase Order AnalysisCode 5",
        "Purchase Order AnalysisCode 6",
        "Purchase Order AnalysisCode 7",
        "Purchase Order AnalysisCode 8",
        "Purchase Order AnalysisCode 9",
        "Purchase Order AnalysisCode 10",
        "Purchase Order AnalysisCode 11",
        "Purchase Order AnalysisCode 12",
        "Purchase Order AnalysisCode 13",
        "Purchase Order AnalysisCode 14",
        "Purchase Order AnalysisCode 15",
        "Purchase Order AnalysisCode 16",
        "Purchase Order AnalysisCode 17",
        "Purchase Order AnalysisCode 18",
        "Purchase Order AnalysisCode 19",
        "Purchase Order AnalysisCode 20",
        "Item Code",
        "Warehouse Name",
        "Unit Discount Percent",
        "Purchase Order Line Requested Date",
        "Stock Item Unit",
        "Item Description",
        "Free Text Item Description",
        "Free Text Buying Unit Description",
        "Tax Code",
        "Item Quantity",
        "Unit Buying Price",
        "Nominal Code",
        "Nominal Cost Centre",
        "Nominal Department",
        "Additional Charge Codes",
        "Additional Charge Value",
        "Comment Line Description",
        "Comment Line Buying Unit Description",
        "Order Line Internal Notes",
        "Order Line External Notes",
        "POP Delivery Address Line 1",
        "POP Delivery Address Line 2",
        "POP Delivery Address Line 3",
        "POP Delivery Address Line 4",
        "POP Delivery City",
        "POP Delivery County",
        "POP Delivery Country",
        "POP Delivery Post Code",
        "POP Delivery Contact",
        "POP Delivery Fax No",
        "POP Delivery Email Address",
        "POP Delivery Telephone No",
        "POP Goods Received Number",
    ]

EXPORT_COLUMNS = load_export_template_columns(EXPORT_TEMPLATE_CSV)

def _blank_export_row() -> dict:
    return {c: "" for c in EXPORT_COLUMNS}

def _export_line(
    po_group: int,
    supplier_acc: str,
    warehouse_name: str,
    so_number: str,
    area_code: str,
    service: str,
    label: str,
    qty: float,
    unit_price: float,
    doc_date: date | None = None,
    req_date: date | None = None,
) -> dict:
    """
    Creates one CSV line in the example format.
    Column A is always 1 (Purchase Order Import Type).
    Column B groups by haulier (1 for Joda, 2 for McDowells).
    Column AI (Free Text Item Description) carries human-readable detail.
    Columns AL/AM become Item Quantity / Unit Buying Price.
    """
    doc_date = doc_date or date.today()
    req_date = req_date or date.today()

    r = _blank_export_row()
    r["Purchase Order Import Type"] = 1
    r["Purchase Order Number"] = po_group
    r["Purchase Order Supplier Acc Code"] = supplier_acc
    r["Purchase Order Document Date"] = _ddmmyyyy(doc_date)
    r["Purchase Order Header Requested Date"] = _ddmmyyyy(req_date)
    r["Purchase Order Discount Percent"] = 0
    r["Unit Discount Percent"] = 0
    r["Purchase Order Line Requested Date"] = _ddmmyyyy(req_date)

    r["Warehouse Name"] = warehouse_name

    # Free text detail format similar to your example
    so_suffix = f" - SO{so_number}" if str(so_number).strip() else ""
    svc_suffix = "" if not service else f" ({service})"
    r["Free Text Item Description"] = f"{area_code} {label}{svc_suffix}{so_suffix}".strip()

    # Units
    r["Item Quantity"] = float(qty) if qty is not None else 0.0
    r["Unit Buying Price"] = float(unit_price) if unit_price is not None else 0.0

    return r

def _add_to_basket(entry: dict):
    st.session_state.export_basket.append(entry)

def _remove_from_basket(idx: int):
    try:
        st.session_state.export_basket.pop(idx)
    except Exception:
        pass

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
        save_joda_surcharge(st.session_state.joda_pct)
        st.success(f"Saved Joda surcharge at {st.session_state.joda_pct:.2f}%")

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
st.subheader("3. Add to Export List (manual fields)")
cso, cwh, csp = st.columns([1.2, 1.2, 3])
with cso:
    st.text_input("SO Number (manual)", key="so_number", placeholder="e.g. 020502")
with cwh:
    st.text_input("Warehouse Name", key="warehouse_name", placeholder="e.g. 101 - Skipton")
with csp:
    st.caption(
        "Tip: enter the SO number and warehouse once, then use the Add buttons below on any rate you like."
    )

# -------------------------
# Calculations
# -------------------------
# Joda calculation
joda_base = None
joda_final = None
joda_charge_fixed = (7.5 if st.session_state.ampm else 0) + (20 if st.session_state.timed else 0)

if st.session_state.dual:
    b1 = get_base_rate(rate_df, postcode_area, st.session_state.service, "Joda", st.session_state.split1)
    b2 = get_base_rate(rate_df, postcode_area, st.session_state.service, "Joda", st.session_state.split2)
    if b1 is not None and b2 is not None:
        eff1 = joda_effective_pct(st.session_state.split1, st.session_state.joda_pct)
        eff2 = joda_effective_pct(st.session_state.split2, st.session_state.joda_pct)
        g1 = b1 * (1 + eff1 / 100.0) + joda_charge_fixed
        g2 = b2 * (1 + eff2 / 100.0) + joda_charge_fixed
        joda_base = b1 + b2
        joda_final = g1 + g2
else:
    base = get_base_rate(rate_df, postcode_area, st.session_state.service, "Joda", st.session_state.pallets)
    if base is not None:
        joda_base = base
        eff = joda_effective_pct(st.session_state.pallets, st.session_state.joda_pct)
        joda_final = base * (1 + eff / 100.0) + joda_charge_fixed

# McDowells calculation (small-load extra is part of base)
mcd_base = get_base_rate(rate_df, postcode_area, st.session_state.service, "Mcdowells", st.session_state.pallets)
mcd_final = None

mcd_charge_fixed = (10 if st.session_state.ampm else 0) + (19 if st.session_state.timed else 0)
mcd_tail_lift_per_pallet = 3.90 if st.session_state.tail else 0.0
mcd_tail_lift_total = mcd_tail_lift_per_pallet * st.session_state.pallets
mcd_small_extra = mcd_smallload_extra(st.session_state.pallets)

if mcd_base is not None:
    mcd_base_for_calc = mcd_base + mcd_small_extra
    mcd_final = (
        mcd_base_for_calc * (1 + st.session_state.mcd_pct / 100.0)
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
        "Fuel Surcharge (%)": f"{st.session_state.joda_pct:.2f}%",
        "Delivery Charge": "N/A",
        "Final Rate": "N/A"
    })
else:
    shown_pct = (
        joda_effective_pct(st.session_state.pallets, st.session_state.joda_pct)
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
        "Fuel Surcharge (%)": f"{st.session_state.mcd_pct:.2f}%",
        "Delivery Charge": "N/A",
        "Final Rate": "N/A"
    })
else:
    mcd_base_for_display = mcd_base + mcd_small_extra
    summary_rows.append({
        "Haulier": "McDowells",
        "Base Rate": f"£{mcd_base_for_display:,.2f}",
        "Fuel Surcharge (%)": f"{st.session_state.mcd_pct:.2f}%",
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
# History (existing)
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
        "JodaPct": st.session_state.joda_pct,
        "McdPct": st.session_state.mcd_pct,
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
# Tabs
# -------------------------
st.header("4. Calculated Rates")
tab_table, tab_export, tab_history, tab_map = st.tabs(["Table", "Export List", "History", "Map (Beta)"])

# -------------------------
# Add-to-list logic (build export lines from current selection)
# -------------------------
def build_export_lines_for_haulier(haulier: str) -> list[dict]:
    """
    Creates the row(s) you want to export, splitting each delivery-charge component into separate lines.
    Notes:
    - Base "Delivery" line includes fuel (i.e., base after fuel, per pallet unit if qty>1).
    - AM/Timed are their own lines as fixed qty 1.
    - Tail Lift is its own line as qty=pallets, unit=per-pallet price.
    - McDowells small-load extra is already baked into the base before fuel (as per your latest rule).
    """
    so = str(st.session_state.so_number).strip()
    wh = str(st.session_state.warehouse_name).strip()
    area = str(st.session_state.area).strip().upper()
    svc = str(st.session_state.service).strip()

    if not so:
        raise ValueError("SO Number is required before adding lines.")
    if not wh:
        raise ValueError("Warehouse Name is required before adding lines.")

    out: list[dict] = []

    if haulier.lower() == "joda":
        if joda_base is None or joda_final is None:
            raise ValueError("No Joda rate available to add.")

        # Determine pallets and base(s)
        if st.session_state.dual:
            # Split into 2 deliveries, each with their own base and their own fixed charges
            for n, base_n in [(st.session_state.split1, get_base_rate(rate_df, area, svc, "Joda", st.session_state.split1)),
                              (st.session_state.split2, get_base_rate(rate_df, area, svc, "Joda", st.session_state.split2))]:
                if base_n is None:
                    continue
                eff = joda_effective_pct(n, st.session_state.joda_pct)
                base_after_fuel_total = base_n * (1 + eff / 100.0)

                # Base delivery line: qty = pallets (units), unit price per pallet
                unit = base_after_fuel_total / max(n, 1)
                out.append(_export_line(
                    po_group=JODA_PO_GROUP,
                    supplier_acc=JODA_ACC,
                    warehouse_name=wh,
                    so_number=so,
                    area_code=area,
                    service=svc,
                    label="Delivery",
                    qty=n,
                    unit_price=unit
                ))

                # AM/Timed fixed charges as separate lines
                if st.session_state.ampm:
                    out.append(_export_line(
                        po_group=JODA_PO_GROUP,
                        supplier_acc=JODA_ACC,
                        warehouse_name=wh,
                        so_number=so,
                        area_code=area,
                        service=svc,
                        label="AM Charge",
                        qty=1,
                        unit_price=7.5
                    ))
                if st.session_state.timed:
                    out.append(_export_line(
                        po_group=JODA_PO_GROUP,
                        supplier_acc=JODA_ACC,
                        warehouse_name=wh,
                        so_number=so,
                        area_code=area,
                        service=svc,
                        label="Timed Charge",
                        qty=1,
                        unit_price=20.0
                    ))
        else:
            n = int(st.session_state.pallets)
            eff = joda_effective_pct(n, st.session_state.joda_pct)
            base_after_fuel_total = joda_base * (1 + eff / 100.0)

            unit = base_after_fuel_total / max(n, 1)
            out.append(_export_line(
                po_group=JODA_PO_GROUP,
                supplier_acc=JODA_ACC,
                warehouse_name=wh,
                so_number=so,
                area_code=area,
                service=svc,
                label="Delivery",
                qty=n,
                unit_price=unit
            ))

            if st.session_state.ampm:
                out.append(_export_line(
                    po_group=JODA_PO_GROUP,
                    supplier_acc=JODA_ACC,
                    warehouse_name=wh,
                    so_number=so,
                    area_code=area,
                    service=svc,
                    label="AM Charge",
                    qty=1,
                    unit_price=7.5
                ))
            if st.session_state.timed:
                out.append(_export_line(
                    po_group=JODA_PO_GROUP,
                    supplier_acc=JODA_ACC,
                    warehouse_name=wh,
                    so_number=so,
                    area_code=area,
                    service=svc,
                    label="Timed Charge",
                    qty=1,
                    unit_price=20.0
                ))

        return out

    if haulier.lower() in ["mcdowells", "mcd", "mcdowell"]:
        if mcd_base is None or mcd_final is None:
            raise ValueError("No McDowells rate available to add.")

        n = int(st.session_state.pallets)

        # Base includes small-load extra (as per your latest requirement), then fuel is applied to that base.
        base_total_for_calc = (mcd_base + mcd_small_extra)
        base_after_fuel_total = base_total_for_calc * (1 + st.session_state.mcd_pct / 100.0)

        unit = base_after_fuel_total / max(n, 1)
        out.append(_export_line(
            po_group=MCD_PO_GROUP,
            supplier_acc=MCD_ACC,
            warehouse_name=wh,
            so_number=so,
            area_code=area,
            service=svc,
            label="Delivery",
            qty=n,
            unit_price=unit
        ))

        # AM/Timed fixed charges as separate lines
        if st.session_state.ampm:
            out.append(_export_line(
                po_group=MCD_PO_GROUP,
                supplier_acc=MCD_ACC,
                warehouse_name=wh,
                so_number=so,
                area_code=area,
                service=svc,
                label="AM Charge",
                qty=1,
                unit_price=10.0
            ))
        if st.session_state.timed:
            out.append(_export_line(
                po_group=MCD_PO_GROUP,
                supplier_acc=MCD_ACC,
                warehouse_name=wh,
                so_number=so,
                area_code=area,
                service=svc,
                label="Timed Charge",
                qty=1,
                unit_price=19.0
            ))

        # Tail lift is per pallet for McDowells
        if st.session_state.tail:
            out.append(_export_line(
                po_group=MCD_PO_GROUP,
                supplier_acc=MCD_ACC,
                warehouse_name=wh,
                so_number=so,
                area_code=area,
                service=svc,
                label="Tail Lift",
                qty=n,
                unit_price=3.90
            ))

        return out

    raise ValueError(f"Unknown haulier: {haulier}")

# -------------------------
# Table tab
# -------------------------
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

    # Add buttons (one per haulier)
    bcols = st.columns([1, 1, 2])
    with bcols[0]:
        if st.button("Add Joda to Export List", use_container_width=True):
            try:
                lines = build_export_lines_for_haulier("Joda")
                for line in lines:
                    _add_to_basket(line)
                st.success(f"Added {len(lines)} line(s) for Joda.")
            except Exception as e:
                st.error(str(e))
    with bcols[1]:
        if st.button("Add McDowells to Export List", use_container_width=True):
            try:
                lines = build_export_lines_for_haulier("McDowells")
                for line in lines:
                    _add_to_basket(line)
                st.success(f"Added {len(lines)} line(s) for McDowells.")
            except Exception as e:
                st.error(str(e))
    with bcols[2]:
        st.caption("These add rows into the Export List tab (you can remove lines before exporting).")

    # One pallet fewer / one pallet more (existing, but updated to mirror McD small-load-as-base)
    def lookup_adjacent_rate(
        df, area, service, vendor, pallets,
        surcharge_pct, fixed_charge=0.0, per_pallet_charge=0.0,
        joda_rule=False, mcd_smallload_rule=False
    ):
        out = {"lower": None, "higher": None}

        def eff_pct(n):
            return joda_effective_pct(n, surcharge_pct) if joda_rule else float(surcharge_pct)

        def small_extra(n):
            return mcd_smallload_extra(n) if mcd_smallload_rule else 0.0

        def total_extras(n):
            return fixed_charge + per_pallet_charge * n

        def compute_total(n, base_rate):
            base_for_calc = base_rate + small_extra(n)
            return base_for_calc * (1 + eff_pct(n) / 100.0) + total_extras(n)

        if pallets > 1:
            bl = get_base_rate(df, area, service, vendor, pallets - 1)
            if bl is not None:
                n = pallets - 1
                out["lower"] = (n, compute_total(n, bl))

        bh = get_base_rate(df, area, service, vendor, pallets + 1)
        if bh is not None:
            n = pallets + 1
            out["higher"] = (n, compute_total(n, bh))

        return out

    joda_adj = lookup_adjacent_rate(
        rate_df, postcode_area, st.session_state.service, "Joda",
        st.session_state.pallets, st.session_state.joda_pct,
        fixed_charge=joda_charge_fixed, per_pallet_charge=0.0,
        joda_rule=True, mcd_smallload_rule=False
    )

    mcd_adj = lookup_adjacent_rate(
        rate_df, postcode_area, st.session_state.service, "Mcdowells",
        st.session_state.pallets, st.session_state.mcd_pct,
        fixed_charge=mcd_charge_fixed, per_pallet_charge=mcd_tail_lift_per_pallet,
        joda_rule=False, mcd_smallload_rule=True
    )

    st.subheader("One Pallet Fewer / One Pallet More")
    c1, c2 = st.columns(2)

    with c1:
        st.markdown("<b>Joda Rates</b>", unsafe_allow_html=True)
        lines = []
        if joda_adj["lower"]:
            lp, lr = joda_adj["lower"]
            lines.append(f"&nbsp;&nbsp;• {lp} pallet(s): £{lr:,.2f}")
        else:
            lines.append("&nbsp;&nbsp;• <span style='color:gray;'>N/A for fewer pallets</span>")
        if joda_adj["higher"]:
            hp, hr = joda_adj["higher"]
            lines.append(f"&nbsp;&nbsp;• {hp} pallet(s): £{hr:,.2f}")
        else:
            lines.append("&nbsp;&nbsp;• <span style='color:gray;'>N/A for more pallets</span>")
        st.markdown("<br>".join(lines), unsafe_allow_html=True)

    with c2:
        st.markdown("<b>McDowells Rates</b>", unsafe_allow_html=True)
        lines = []
        if mcd_adj["lower"]:
            lp, lr = mcd_adj["lower"]
            lines.append(f"&nbsp;&nbsp;• {lp} pallet(s): £{lr:,.2f}")
        else:
            lines.append("&nbsp;&nbsp;• <span style='color:gray;'>N/A for fewer pallets</span>")
        if mcd_adj["higher"]:
            hp, hr = mcd_adj["higher"]
            lines.append(f"&nbsp;&nbsp;• {hp} pallet(s): £{hr:,.2f}")
        else:
            lines.append("&nbsp;&nbsp;• <span style='color:gray;'>N/A for more pallets</span>")
        st.markdown("<br>".join(lines), unsafe_allow_html=True)

# -------------------------
# Export List tab (NEW)
# -------------------------
with tab_export:
    st.subheader("Saved Lines")
    basket = st.session_state.get("export_basket", [])
    if not basket:
        st.info("Nothing saved yet. Go to the Table tab and click 'Add Joda…' or 'Add McDowells…'.")
    else:
        # Show a compact view + remove buttons
        view_df = pd.DataFrame(basket)
        # Prefer a readable subset if columns exist
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
        st.dataframe(view_df[show_cols], use_container_width=True, hide_index=True)

        rm_cols = st.columns([1, 1, 6])
        with rm_cols[0]:
            if st.button("Remove last line"):
                _remove_from_basket(len(basket) - 1)
                st.rerun()
        with rm_cols[1]:
            if st.button("Clear all"):
                st.session_state.export_basket = []
                st.rerun()

        # Build CSV for download
        export_df = pd.DataFrame(st.session_state.export_basket)
        # Ensure correct column order
        export_df = export_df.reindex(columns=EXPORT_COLUMNS)

        csv_bytes = export_df.to_csv(index=False).encode("utf-8")

        st.download_button(
            label="Download CSV (PO Import Format)",
            data=csv_bytes,
            file_name=f"PO_Import_Export_{date.today().strftime('%Y%m%d')}.csv",
            mime="text/csv",
            use_container_width=True
        )

        st.caption(
            "This export mirrors your example: Column A is 1, Column B groups by haulier (Joda=1, McDowells=2), "
            "Column AI holds the line detail, and quantity/price are in AL/AM."
        )

# -------------------------
# Map tab (existing)
# -------------------------
with tab_map:
    centroids_path_candidates = [
        "postcode_area_centroids.csv",
        "postcode_area_centroids_filled.csv",
        "/mnt/data/postcode_area_centroids_filled.csv",
    ]

    def _find_col(df: pd.DataFrame, candidates: list[str]) -> str | None:
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
        def calc_for_area(area_code: str):
            # Joda
            jb = None
            jf = None
            if st.session_state.dual:
                b1 = get_base_rate(rate_df, area_code, st.session_state.service, "Joda", st.session_state.split1)
                b2 = get_base_rate(rate_df, area_code, st.session_state.service, "Joda", st.session_state.split2)
                if b1 is not None and b2 is not None:
                    p1 = joda_effective_pct(st.session_state.split1, st.session_state.joda_pct)
                    p2 = joda_effective_pct(st.session_state.split2, st.session_state.joda_pct)
                    jf = b1 * (1 + p1 / 100.0) + joda_charge_fixed
                    jf += b2 * (1 + p2 / 100.0) + joda_charge_fixed
                    jb = b1 + b2
            else:
                jb = get_base_rate(rate_df, area_code, st.session_state.service, "Joda", st.session_state.pallets)
                if jb is not None:
                    ep = joda_effective_pct(st.session_state.pallets, st.session_state.joda_pct)
                    jf = jb * (1 + ep / 100.0) + joda_charge_fixed

            # McDowells (smallload as base)
            mb = get_base_rate(rate_df, area_code, st.session_state.service, "Mcdowells", st.session_state.pallets)
            mf = None
            if mb is not None:
                tl_total = (3.90 if st.session_state.tail else 0.0) * st.session_state.pallets
                small_extra = mcd_smallload_extra(st.session_state.pallets)
                mb_for_calc = mb + small_extra
                mf = (
                    mb_for_calc * (1 + st.session_state.mcd_pct / 100.0)
                    + ((10 if st.session_state.ampm else 0) + (19 if st.session_state.timed else 0))
                    + tl_total
                )

            return jb, jf, mb, mf

        areas = rate_df["PostcodeArea"].unique()
        map_rows = []
        for a in areas:
            jb, jf, mb, mf = calc_for_area(a)
            if jf is None and mf is None:
                continue
            crow = centroid_df.loc[centroid_df["Area"] == a]
            if crow.empty:
                continue
            lat = float(crow.iloc[0]["lat"])
            lon = float(crow.iloc[0]["lon"])
            map_rows.append({
                "Area": a,
                "lat": lat,
                "lon": lon,
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
                        "Tail": h.get("Tail", False),
                        "Dual": h.get("Dual", False),
                        "Split1": h.get("Split1", 1),
                        "Split2": h.get("Split2", 1),
                        "JodaPct": h.get("JodaPct", st.session_state.get("joda_pct", 0.0)),
                        "McdPct": h.get("McdPct", st.session_state.get("mcd_pct", 0.0)),
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
