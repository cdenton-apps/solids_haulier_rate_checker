# ======================================
# File: waste_solidus_haulier_app.py
# ======================================

import streamlit as st
import pandas as pd
import math
from PIL import Image
import json
from datetime import date
import os

# ─────────────────────────────────────────
# (1) STREAMLIT PAGE CONFIGURATION (must be first)
# ─────────────────────────────────────────
st.set_page_config(
    page_title="Solidus Haulier Rate Checker",
    layout="wide"
)

# ─────────────────────────────────────────
# (2) HIDE STREAMLIT MENU & FOOTER (optional)
# ─────────────────────────────────────────
hide_streamlit_style = """
    <style>
      /* Hide top-right menu */
      #MainMenu { visibility: hidden; }
      /* Hide “Made with Streamlit” footer */
      footer { visibility: hidden; }
    </style>
"""
st.markdown(hide_streamlit_style, unsafe_allow_html=True)

# ─────────────────────────────────────────
# (3) DISPLAY SOLIDUS LOGO + HEADER SIDE‑BY‑SIDE
# ─────────────────────────────────────────
col_logo, col_text = st.columns([1, 3], gap="medium")

with col_logo:
    logo_path = "assets/solidus_logo.png"
    try:
        logo_img = Image.open(logo_path)
        st.image(logo_img, width=150)
    except Exception:
        st.warning(f"Could not load logo at '{logo_path}'.")

with col_text:
    st.markdown(
        "<h1 style='color:#0D4B6A; margin-bottom:0.25em;'>"
        "Solidus Haulier Rate Checker</h1>",
        unsafe_allow_html=True
    )
    st.markdown(
        """
        Enter a UK postcode area, select a service type (Economy or Next Day),  
        specify the number of pallets, and apply fuel surcharges and optional extras:

        • **Joda’s surcharge (%)** is stored persistently and must be updated once weekly.  
          You can look this up at https://www.jodafreight.com/fuel-surcharge/  
          On Wednesdays it resets to 0 automatically.  
        • **McDowells’ surcharge (%)** is always entered manually each session.  
        • You may optionally add AM/PM Delivery or Timed Delivery,  
        or perform a Dual Collection (For collections from both Unit 4 and ESL):
        """,
        unsafe_allow_html=True
    )

# ─────────────────────────────────────────
# (4) PERSISTENT STORAGE FOR JODA SURCHARGE
# ─────────────────────────────────────────
DATA_FILE = "joda_surcharge.json"

def load_joda_surcharge():
    today_str = date.today().isoformat()
    if not os.path.exists(DATA_FILE):
        initial = {"surcharge": 0.0, "last_updated": today_str}
        with open(DATA_FILE, "w") as f:
            json.dump(initial, f)
        return 0.0

    try:
        with open(DATA_FILE, "r") as f:
            data = json.load(f)
    except Exception:
        data = {"surcharge": 0.0, "last_updated": today_str}

    last_upd = data.get("last_updated", "")
    if date.today().weekday() == 2 and last_upd != today_str:
        data["surcharge"] = 0.0
        data["last_updated"] = today_str
        with open(DATA_FILE, "w") as f:
            json.dump(data, f)
        return 0.0

    return data.get("surcharge", 0.0)

def save_joda_surcharge(new_pct: float):
    today_str = date.today().isoformat()
    data = {"surcharge": float(new_pct), "last_updated": today_str}
    with open(DATA_FILE, "w") as f:
        json.dump(data, f)

joda_stored_pct = load_joda_surcharge()

# ─────────────────────────────────────────
# (5) LOAD & TRANSFORM THE BUILT‑IN EXCEL DATA
# ─────────────────────────────────────────
@st.cache_data
def load_rate_table(excel_path: str) -> pd.DataFrame:
    raw = pd.read_excel(excel_path, header=1)
    raw = raw.rename(columns={
        raw.columns[0]: "PostcodeArea",
        raw.columns[1]: "Service",
        raw.columns[2]: "Vendor"
    })
    raw["PostcodeArea"] = raw["PostcodeArea"].ffill()
    raw["Service"]      = raw["Service"].ffill()
    raw = raw[raw["Vendor"] != "Vendor"].copy()

    pallet_cols = [
        col for col in raw.columns
        if isinstance(col, (int, float)) or (isinstance(col, str) and col.isdigit())
    ]

    melted = raw.melt(
        id_vars=["PostcodeArea", "Service", "Vendor"],
        value_vars=pallet_cols,
        var_name="Pallets",
        value_name="BaseRate"
    )
    melted["Pallets"] = melted["Pallets"].astype(int)
    melted = melted.dropna(subset=["BaseRate"]).copy()

    for col in ["PostcodeArea", "Service", "Vendor"]:
        melted[col] = (
            melted[col]
            .astype(str)
            .str.strip()
            .apply(lambda x: x.upper() if col=="PostcodeArea" else x.title())
        )

    return melted.reset_index(drop=True)

rate_df = load_rate_table("haulier prices.xlsx")
unique_areas = sorted(rate_df["PostcodeArea"].unique())

# ─────────────────────────────────────────
# (6) USER INPUTS
# ─────────────────────────────────────────
st.header("1. Input Parameters")
col_a, col_b, col_c, col_d, col_e, col_f = st.columns([1,1,1,1,1,1], gap="medium")

with col_a:
    input_area = st.selectbox(
        "Postcode Area (e.g. BB, LA, etc.)",
        options=[""] + unique_areas,
        format_func=lambda x: x if x else "— Select area —",
        index=0
    )
    if input_area == "":
        st.info("🔍 Please select a postcode area to continue.")
        st.stop()

with col_b:
    service_option = st.selectbox("Service Type", ["Economy", "Next Day"], index=0)

with col_c:
    num_pallets = st.number_input("Number of Pallets", 1, 26, 1)

with col_d:
    joda_surcharge_pct = st.number_input(
        "Joda Fuel Surcharge (%)",
        min_value=0.0, max_value=100.0,
        value=round(joda_stored_pct, 2), step=0.1,
        format="%.2f"
    )
    if st.button("Save Joda Surcharge"):
        save_joda_surcharge(joda_surcharge_pct)
        st.success(f"✅ Saved Joda surcharge at {joda_surcharge_pct:.2f}%")

with col_e:
    mcd_surcharge_pct = st.number_input(
        "McDowells Fuel Surcharge (%)",
        min_value=0.0, max_value=100.0,
        value=0.0, step=0.1, format="%.2f"
    )

with col_f:
    st.markdown(" ")

st.markdown("---")
postcode_area = input_area

# ─────────────────────────────────────────
# (7) OPTIONAL EXTRAS
# ─────────────────────────────────────────
st.subheader("2. Optional Extras")
col1, col2, col3 = st.columns(3, gap="large")
with col1: ampm_toggle = st.checkbox("AM/PM Delivery")
with col2: timed_toggle = st.checkbox("Timed Delivery")
with col3: dual_toggle = st.checkbox("Dual Collection")

split1 = split2 = None
if dual_toggle:
    st.markdown("**Specify how to split pallets into two despatches e.g. a split load between ESL and U4.**")
    sp1, sp2 = st.columns(2, gap="large")
    with sp1:
        split1 = st.number_input("First Pallet Group", 1, num_pallets-1, 1)
    with sp2:
        split2 = st.number_input("Second Pallet Group", 1, num_pallets-1, num_pallets-1)
    if split1 + split2 != num_pallets:
        st.error("⚠️ Pallet Split values must add up to total pallets.")
        st.stop()

# ─────────────────────────────────────────
# (8) RATE LOOKUP & CALCULATION (no hard stops)
# ─────────────────────────────────────────
def get_base_rate(df, area, service, vendor, pallets):
    subset = df[(df["PostcodeArea"]==area)
                & (df["Service"]==service)
                & (df["Vendor"]==vendor)
                & (df["Pallets"]==pallets)]
    return None if subset.empty else float(subset["BaseRate"].iloc[0])

# Joda
joda_base = None
joda_final = None
if dual_toggle:
    b1 = get_base_rate(rate_df, postcode_area, service_option, "Joda", split1)
    b2 = get_base_rate(rate_df, postcode_area, service_option, "Joda", split2)
    if b1 is not None and b2 is not None:
        charge = (7 if ampm_toggle else 0) + (19 if timed_toggle else 0)
        g1 = b1 * (1 + joda_surcharge_pct/100.0) + charge
        g2 = b2 * (1 + joda_surcharge_pct/100.0) + charge
        joda_base = b1 + b2
        joda_final = g1 + g2
else:
    base = get_base_rate(rate_df, postcode_area, service_option, "Joda", num_pallets)
    if base is not None:
        charge = (7 if ampm_toggle else 0) + (19 if timed_toggle else 0)
        joda_base = base
        joda_final = base * (1 + joda_surcharge_pct/100.0) + charge

# McDowells
mcd_base = get_base_rate(rate_df, postcode_area, service_option, "Mcdowells", num_pallets)
if mcd_base is not None:
    mcd_charge = (10 if ampm_toggle else 0) + (19 if timed_toggle else 0)
    mcd_final = mcd_base * (1 + mcd_surcharge_pct/100.0) + mcd_charge
else:
    mcd_final = None

# ─────────────────────────────────────────
# (9) BUILD SUMMARY TABLE, WITH “No rate” TEXT FOR MISSING
# ─────────────────────────────────────────
summary_rows = []

# Joda row
if joda_base is None:
    summary_rows.append({
        "Haulier": "Joda",
        "Base Rate":      "No rate",
        "Fuel Surcharge (%)": f"{joda_surcharge_pct:.2f}%",
        "Delivery Charge":    "N/A",
        "Final Rate":        "N/A"
    })
else:
    summary_rows.append({
        "Haulier": "Joda",
        "Base Rate":      f"£{joda_base:,.2f}",
        "Fuel Surcharge (%)": f"{joda_surcharge_pct:.2f}%",
        "Delivery Charge":    f"£{(7 if ampm_toggle else 0) + (19 if timed_toggle else 0):,.2f}" if not dual_toggle else f"£{((7 if ampm_toggle else 0)+(19 if timed_toggle else 0))* (2 if dual_toggle else 1):,.2f}",
        "Final Rate":        f"£{joda_final:,.2f}"
    })

# McDowells row
if mcd_base is None:
    summary_rows.append({
        "Haulier": "McDowells",
        "Base Rate":      "No rate",
        "Fuel Surcharge (%)": f"{mcd_surcharge_pct:.2f}%",
        "Delivery Charge":    "N/A",
        "Final Rate":        "N/A"
    })
else:
    summary_rows.append({
        "Haulier": "McDowells",
        "Base Rate":      f"£{mcd_base:,.2f}",
        "Fuel Surcharge (%)": f"{mcd_surcharge_pct:.2f}%",
        "Delivery Charge":    f"£{(10 if ampm_toggle else 0) + (19 if timed_toggle else 0):,.2f}",
        "Final Rate":        f"£{mcd_final:,.2f}"
    })

# ─────────────────────────────────────────
# (10) SHOW ONE-PALET Fewer / MORE
# ─────────────────────────────────────────
def lookup_adjacent_rate(df, area, service, vendor, pallets, surcharge_pct, delivery_charge):
    out = {"lower":None, "higher":None}
    if pallets>1:
        bl = get_base_rate(df, area, service, vendor, pallets-1)
        if bl is not None:
            out["lower"] = ((pallets-1), bl*(1+surcharge_pct/100.0)+delivery_charge)
    bh = get_base_rate(df, area, service, vendor, pallets+1)
    if bh is not None:
        out["higher"] = ((pallets+1), bh*(1+surcharge_pct/100.0)+delivery_charge)
    return out

joda_adj = lookup_adjacent_rate(rate_df, postcode_area, service_option, "Joda", num_pallets,
                                 joda_surcharge_pct, charge if joda_final is not None else 0)
mcd_adj  = lookup_adjacent_rate(rate_df, postcode_area, service_option, "Mcdowells", num_pallets,
                                 mcd_surcharge_pct, mcd_charge if mcd_final is not None else 0)

st.subheader("One Pallet Fewer / One Pallet More")
adj_cols = st.columns(2)

with adj_cols[0]:
    st.markdown("<b>Joda Rates</b>", unsafe_allow_html=True)
    lines=[]
    if joda_adj["lower"]: lines.append(f"  • {joda_adj['lower'][0]} pallet(s): £{joda_adj['lower'][1]:,.2f}")
    else: lines.append("  • <span style='color:gray;'>N/A for fewer pallets</span>")
    if joda_adj["higher"]: lines.append(f"  • {joda_adj['higher'][0]} pallet(s): £{joda_adj['higher'][1]:,.2f}")
    else: lines.append("  • <span style='color:gray;'>N/A for more pallets</span>")
    st.markdown("<br>".join(lines), unsafe_allow_html=True)

with adj_cols[1]:
    st.markdown("<b>McDowells Rates</b>", unsafe_allow_html=True)
    lines=[]
    if mcd_adj["lower"]: lines.append(f"  • {mcd_adj['lower'][0]} pallet(s): £{mcd_adj['lower'][1]:,.2f}")
    else: lines.append("  • <span style='color:gray;'>N/A for fewer pallets</span>")
    if mcd_adj["higher"]: lines.append(f"  • {mcd_adj['higher'][0]} pallet(s): £{mcd_adj['higher'][1]:,.2f}")
    else: lines.append("  • <span style='color:gray;'>N/A for more pallets</span>")
    st.markdown("<br>".join(lines), unsafe_allow_html=True)

# ─────────────────────────────────────────
# (11) FOOTER NOTES
# ─────────────────────────────────────────
st.markdown("---")
st.markdown(
    """
    <small>
    • Joda’s surcharge is stored and resets each Wednesday.  
    • McDowells’ surcharge is always entered manually.  
    • Delivery charges: Joda – AM/PM £7, Timed £19;  
      McDowells – AM/PM £10, Timed £19.  
    • Dual Collection splits Joda into two shipments.  
    • The cheapest final rate is highlighted in green.  
    </small>
    """,
    unsafe_allow_html=True
)
