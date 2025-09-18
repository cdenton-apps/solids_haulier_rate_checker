# import requirements

import streamlit as st
import pandas as pd
import math
from PIL import Image
import json
from datetime import date
import os

# configure streamlit page
st.set_page_config(
    page_title="Solidus Haulier Rate Checker",
    layout="wide"
)

hide_streamlit_style = """
    <style>
      #MainMenu { visibility: hidden; }
      footer { visibility: hidden; }
    </style>
"""
st.markdown(hide_streamlit_style, unsafe_allow_html=True)

# header section
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

# store Joda surcharge
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

# load and transform excel sheet of rates
@st.cache_data
def load_rate_table(excel_path: str, _mtime: float) -> pd.DataFrame:
    raw = pd.read_excel(excel_path, header=1)  # add sheet_name="Rates" if needed
    raw = raw.rename(columns={
        raw.columns[0]: "PostcodeArea",
        raw.columns[1]: "Service",
        raw.columns[2]: "Vendor"
    })
    raw["PostcodeArea"] = raw["PostcodeArea"].ffill()
    raw["Service"]      = raw["Service"].ffill()
    raw["Vendor"]       = raw["Vendor"].ffill()   # forward-fill vendor too
    raw = raw[raw["Vendor"] != "Vendor"].copy()

    # keep only numeric pallet columns (headers like 1,2,3,...)
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
    melted["Pallets"]   = melted["Pallets"].astype(int)
    melted["BaseRate"]  = pd.to_numeric(melted["BaseRate"], errors="coerce")
    melted = melted.dropna(subset=["BaseRate"]).copy()

    melted["PostcodeArea"] = melted["PostcodeArea"].astype(str).str.strip().str.upper()
    melted["Service"]      = melted["Service"].astype(str).str.strip().str.title()
    melted["Vendor"]       = melted["Vendor"].astype(str).str.strip().str.title()

    return melted.reset_index(drop=True)

# Call it with the file's modification time to bust the cache when the file changes
mtime = os.path.getmtime("haulier prices.xlsx")
rate_df = load_rate_table("haulier prices.xlsx", mtime)
unique_areas = sorted(rate_df["PostcodeArea"].unique())

# user inputs
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
    service_option = st.selectbox(
        "Service Type",
        options=["Economy", "Next Day"],
        index=0
    )

with col_c:
    num_pallets = st.number_input(
        "Number of Pallets",
        min_value=1,
        max_value=26,
        value=1,
        step=1
    )

with col_d:
    joda_surcharge_pct = st.number_input(
        "Joda Fuel Surcharge (%)",
        min_value=0.0,
        max_value=100.0,
        value=round(joda_stored_pct, 2),
        step=0.1,
        format="%.2f"
    )
    if st.button("Save Joda Surcharge"):
        save_joda_surcharge(joda_surcharge_pct)
        st.success(f"✅ Saved Joda surcharge at {joda_surcharge_pct:.2f}%")

with col_e:
    mcd_surcharge_pct = st.number_input(
        "McDowells Fuel Surcharge (%)",
        min_value=0.0,
        max_value=100.0,
        value=0.0,
        step=0.1,
        format="%.2f"
    )

with col_f:
    st.markdown(" ")

st.markdown("---")
postcode_area = input_area

# optional extra inputs
st.subheader("2. Optional Extras")
col1, col2, col3 = st.columns(3, gap="large")

with col1:
    ampm_toggle = st.checkbox("AM/PM Delivery")
with col2:
    timed_toggle = st.checkbox("Timed Delivery")
with col3:
    dual_toggle = st.checkbox("Dual Collection")

# ——— New guard: if user tries Dual with only 1 pallet
if dual_toggle and num_pallets == 1:
    st.error("⚠️ Dual Collection requires at least 2 pallets.")
    st.stop()

split1 = split2 = None
if dual_toggle:
    st.markdown(
        "**Specify how to split pallets into two despatches "
        "e.g. split between ESL and U4**"
    )
    sp1, sp2 = st.columns(2, gap="large")
    with sp1:
        split1 = st.number_input("First Pallet Group", 1, num_pallets-1, 1)
    with sp2:
        split2 = st.number_input("Second Pallet Group", 1, num_pallets-1, num_pallets-1)
    if split1 + split2 != num_pallets:
        st.error("⚠️ Pallet Split values must add up to total pallets.")
        st.stop()

# rate lookup and calculate
def get_base_rate(df, area, service, vendor, pallets):
    subset = df[
        (df["PostcodeArea"] == area) &
        (df["Service"] == service) &
        (df["Vendor"] == vendor) &
        (df["Pallets"] == pallets)
    ]
    return None if subset.empty else float(subset["BaseRate"].iloc[0])

# Joda
joda_base = None
joda_final = None
joda_charge = (7 if ampm_toggle else 0) + (19 if timed_toggle else 0)

if dual_toggle:
    b1 = get_base_rate(rate_df, postcode_area, service_option, "Joda", split1)
    b2 = get_base_rate(rate_df, postcode_area, service_option, "Joda", split2)
    if b1 is not None and b2 is not None:
        g1 = b1 * (1 + joda_surcharge_pct/100.0) + joda_charge
        g2 = b2 * (1 + joda_surcharge_pct/100.0) + joda_charge
        joda_base = b1 + b2
        joda_final = g1 + g2
else:
    base = get_base_rate(rate_df, postcode_area, service_option, "Joda", num_pallets)
    if base is not None:
        joda_base = base
        joda_final = base * (1 + joda_surcharge_pct/100.0) + joda_charge

# McDowells
mcd_base = get_base_rate(rate_df, postcode_area, service_option, "Mcdowells", num_pallets)
mcd_final = None
mcd_charge = (10 if ampm_toggle else 0) + (19 if timed_toggle else 0)

if mcd_base is not None:
    mcd_final = mcd_base * (1 + mcd_surcharge_pct/100.0) + mcd_charge

# build summary table
summary_rows = []

# Joda row
if joda_base is None:
    summary_rows.append({
        "Haulier": "Joda",
        "Base Rate":         "No rate",
        "Fuel Surcharge (%)": f"{joda_surcharge_pct:.2f}%",
        "Delivery Charge":    "N/A",
        "Final Rate":         "N/A"
    })
else:
    summary_rows.append({
        "Haulier": "Joda",
        "Base Rate":         f"£{joda_base:,.2f}",
        "Fuel Surcharge (%)": f"{joda_surcharge_pct:.2f}%",
        "Delivery Charge":    f"£{joda_charge:,.2f}",
        "Final Rate":         f"£{joda_final:,.2f}"
    })

# McDowells row
if mcd_base is None:
    summary_rows.append({
        "Haulier": "McDowells",
        "Base Rate":         "No rate",
        "Fuel Surcharge (%)": f"{mcd_surcharge_pct:.2f}%",
        "Delivery Charge":    "N/A",
        "Final Rate":         "N/A"
    })
else:
    summary_rows.append({
        "Haulier": "McDowells",
        "Base Rate":         f"£{mcd_base:,.2f}",
        "Fuel Surcharge (%)": f"{mcd_surcharge_pct:.2f}%",
        "Delivery Charge":    f"£{mcd_charge:,.2f}",
        "Final Rate":         f"£{mcd_final:,.2f}"
    })

# Display warning if neither haulier has a rate
if all(r["Final Rate"] == "N/A" for r in summary_rows):
    st.warning("No rates found for that area/service/pallet combination.")
else:
    summary_df = pd.DataFrame(summary_rows).set_index("Haulier")

    def highlight_cheapest(row):
        fr = row["Final Rate"]
        if fr.startswith("£"):
            val = float(fr.strip("£").replace(",", ""))
            j_r = round(joda_final, 2) if joda_final is not None else float("inf")
            m_r = round(mcd_final, 2) if mcd_final is not None else float("inf")
            if round(val, 2) == min(j_r, m_r):
                return ["background-color: #b3e6b3"] * len(row)
        return [""] * len(row)

    st.header("3. Calculated Rates")
    st.table(summary_df.style.apply(highlight_cheapest, axis=1))
    st.markdown(
        "<i style='color:gray;'>* Rows in green are the cheapest available *</i>",
        unsafe_allow_html=True
    )

# one pallet fewer/more extra section
def lookup_adjacent_rate(df, area, service, vendor, pallets, surcharge_pct, delivery_charge):
    out = {"lower": None, "higher": None}
    if pallets > 1:
        bl = get_base_rate(df, area, service, vendor, pallets - 1)
        if bl is not None:
            out["lower"] = ((pallets - 1), bl * (1 + surcharge_pct/100.0) + delivery_charge)
    bh = get_base_rate(df, area, service, vendor, pallets + 1)
    if bh is not None:
        out["higher"] = ((pallets + 1), bh * (1 + surcharge_pct/100.0) + delivery_charge)
    return out

joda_adj = lookup_adjacent_rate(
    rate_df, postcode_area, service_option, "Joda", num_pallets, joda_surcharge_pct, joda_charge
)
mcd_adj = lookup_adjacent_rate(
    rate_df, postcode_area, service_option, "Mcdowells", num_pallets, mcd_surcharge_pct, mcd_charge
)

st.subheader("One Pallet Fewer / One Pallet More")
adj_cols = st.columns(2)

with adj_cols[0]:
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

with adj_cols[1]:
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

# footer section
st.markdown("---")
st.markdown(
    """
    <small>
    • Joda’s surcharge resets each Wednesday.  
    • McDowells’ surcharge is entered each session.  
    • Delivery charges: Joda – AM/PM £7, Timed £19;  
      McDowells – AM/PM £10, Timed £19.  
    • Dual Collection splits Joda into two shipments.  
    • Green row indicates the cheapest available rate.  
    </small>
    """,
    unsafe_allow_html=True
)
