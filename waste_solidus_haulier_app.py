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
# (2) HIDE STREAMLIT MENU & FOOTER
# ─────────────────────────────────────────
st.markdown("""
    <style>
      #MainMenu { visibility: hidden; }
      footer { visibility: hidden; }
    </style>
""", unsafe_allow_html=True)

# ─────────────────────────────────────────
# (3) HEADER AND LOGO
# ─────────────────────────────────────────
col_logo, col_text = st.columns([1, 3], gap="medium")
with col_logo:
    try:
        logo = Image.open("assets/solidus_logo.png")
        st.image(logo, width=150)
    except:
        st.warning("Logo not found at assets/solidus_logo.png")
with col_text:
    st.markdown(
        "<h1 style='color:#0D4B6A; margin-bottom:0.25em;'>"
        "Solidus Haulier Rate Checker</h1>",
        unsafe_allow_html=True
    )
    st.markdown(
        "Enter a UK postcode area, choose service and pallets, "
        "set fuel surcharges and extras. "
        "Joda’s surcharge resets each Wednesday and is persisted; "
        "McDowells’ is manual.",
        unsafe_allow_html=True
    )

# ─────────────────────────────────────────
# (4) PERSISTENT JODA SURCHARGE
# ─────────────────────────────────────────
DATA_FILE = "joda_surcharge.json"

def load_joda_surcharge():
    today = date.today().isoformat()
    if not os.path.exists(DATA_FILE):
        json.dump({"surcharge":0.0,"last_updated":today}, open(DATA_FILE,"w"))
        return 0.0
    data = json.load(open(DATA_FILE))
    if date.today().weekday()==2 and data.get("last_updated")!=today:
        data = {"surcharge":0.0,"last_updated":today}
        json.dump(data, open(DATA_FILE,"w"))
        return 0.0
    return data.get("surcharge",0.0)

def save_joda_surcharge(pct):
    today = date.today().isoformat()
    json.dump({"surcharge":float(pct),"last_updated":today}, open(DATA_FILE,"w"))

joda_pct = load_joda_surcharge()

# ─────────────────────────────────────────
# (5) LOAD & TRANSFORM EXCEL RATE TABLE
# ─────────────────────────────────────────
@st.cache_data
def load_rates(path):
    df = pd.read_excel(path, header=1)
    df = df.rename(columns={
        df.columns[0]:"PostcodeArea",
        df.columns[1]:"Service",
        df.columns[2]:"Vendor"
    })
    df["PostcodeArea"] = df["PostcodeArea"].ffill()
    df["Service"]      = df["Service"].ffill()
    df = df[df["Vendor"]!="Vendor"].copy()
    pallet_cols = [c for c in df.columns
                   if isinstance(c,(int,float)) or (isinstance(c,str) and c.isdigit())]
    m = df.melt(
        id_vars=["PostcodeArea","Service","Vendor"],
        value_vars=pallet_cols,
        var_name="Pallets",
        value_name="BaseRate"
    )
    m["Pallets"] = m["Pallets"].astype(int)
    m = m.dropna(subset=["BaseRate"]).copy()
    m["PostcodeArea"] = m["PostcodeArea"].astype(str).str.strip().str.upper()
    m["Service"]      = m["Service"].astype(str).str.strip().str.title()
    m["Vendor"]       = m["Vendor"].astype(str).str.strip().str.title()
    return m.reset_index(drop=True)

rates_df = load_rates("haulier prices.xlsx")
areas = sorted(rates_df["PostcodeArea"].unique())

# ─────────────────────────────────────────
# (6) USER INPUTS
# ─────────────────────────────────────────
st.header("1. Input Parameters")
c1,c2,c3,c4,c5,c6 = st.columns(6, gap="medium")

with c1:
    area = st.selectbox(
        "Postcode Area",
        [""]+areas,
        format_func=lambda x: x or "— Select area —"
    )
    if not area:
        st.stop()

with c2:
    service = st.selectbox("Service Type", ["Economy","Next Day"])

with c3:
    pallets = st.number_input("Number of Pallets", 1, 26, 1)

with c4:
    joda_in = st.number_input(
        "Joda FSC (%)",
        0.0, 100.0,
        value=round(joda_pct,2),
        step=0.1, format="%.2f"
    )
    if st.button("Save Joda FSC"):
        save_joda_surcharge(joda_in)
        st.success("Saved Joda surcharge.")

with c5:
    mcd_in = st.number_input(
        "McDowells FSC (%)",
        0.0, 100.0,
        value=0.0, step=0.1, format="%.2f"
    )

with c6:
    st.markdown("")

st.markdown("---")

# ─────────────────────────────────────────
# (7) OPTIONAL EXTRAS
# ─────────────────────────────────────────
st.subheader("2. Optional Extras")
e1,e2,e3 = st.columns(3, gap="large")
ampm  = e1.checkbox("AM/PM Delivery")
timed = e2.checkbox("Timed Delivery")
dual  = e3.checkbox("Dual Collection")

split1 = split2 = None
if dual:
    st.markdown("**Split pallets into two shipments:**")
    s1,s2 = st.columns(2, gap="large")
    with s1:
        split1 = st.number_input("Group 1", 1, pallets-1, 1)
    with s2:
        split2 = st.number_input("Group 2", 1, pallets-1, pallets-1)
    if split1+split2 != pallets:
        st.error("Split must sum to total pallets.")
        st.stop()

# ─────────────────────────────────────────
# (8) RATE LOOKUP & CALCULATION
# ─────────────────────────────────────────
def get_rate(vendor, qty):
    sub = rates_df[
        (rates_df.PostcodeArea==area) &
        (rates_df.Service==service) &
        (rates_df.Vendor==vendor) &
        (rates_df.Pallets==qty)
    ]
    return None if sub.empty else float(sub.BaseRate.iloc[0])

j_charge = (7 if ampm else 0) + (19 if timed else 0)
m_charge = (10 if ampm else 0) + (19 if timed else 0)

# Joda
j_base = None
j_final = None
if dual:
    b1 = get_rate("Joda", split1)
    b2 = get_rate("Joda", split2)
    if b1 is not None and b2 is not None:
        g1 = b1*(1+joda_in/100)+j_charge
        g2 = b2*(1+joda_in/100)+j_charge
        j_base  = b1 + b2
        j_final = g1 + g2
else:
    b = get_rate("Joda", pallets)
    if b is not None:
        j_base  = b
        j_final = b*(1+joda_in/100) + j_charge

# McDowells
m_base  = get_rate("Mcdowells", pallets)
m_final = None if m_base is None else m_base*(1+mcd_in/100) + m_charge

# ─────────────────────────────────────────
# (9) BUILD SUMMARY WITH "No rate"/"N/A"
# ─────────────────────────────────────────
rows = []

# Joda
if j_base is None:
    rows.append({
        "Haulier":"Joda",
        "Base Rate":"No rate",
        "Fuel Surcharge (%)":f"{joda_in:.2f}%",
        "Delivery Charge":"N/A",
        "Final Rate":"N/A"
    })
else:
    rows.append({
        "Haulier":"Joda",
        "Base Rate":f"£{j_base:,.2f}",
        "Fuel Surcharge (%)":f"{joda_in:.2f}%",
        "Delivery Charge":f"£{j_charge:,.2f}",
        "Final Rate":f"£{j_final:,.2f}"
    })

# McDowells
if m_base is None:
    rows.append({
        "Haulier":"McDowells",
        "Base Rate":"No rate",
        "Fuel Surcharge (%)":f"{mcd_in:.2f}%",
        "Delivery Charge":"N/A",
        "Final Rate":"N/A"
    })
else:
    rows.append({
        "Haulier":"McDowells",
        "Base Rate":f"£{m_base:,.2f}",
        "Fuel Surcharge (%)":f"{mcd_in:.2f}%",
        "Delivery Charge":f"£{m_charge:,.2f}",
        "Final Rate":f"£{m_final:,.2f}"
    })

# Display
if all(r["Final Rate"]=="N/A" for r in rows):
    st.warning("No rates found for that selection.")
else:
    df = pd.DataFrame(rows).set_index("Haulier")
    def highlight(row):
        val = row["Final Rate"]
        if val.startswith("£"):
            num = float(val.strip("£").replace(",",""))
            opts = [(j_final or math.inf), (m_final or math.inf)]
            if abs(num - min(opts))<1e-6:
                return ["background-color: #b3e6b3"]*len(row)
        return [""]*len(row)
    st.header("3. Calculated Rates")
    st.table(df.style.apply(highlight, axis=1))

# ─────────────────────────────────────────
# (10) One Pallet Fewer / More
# ─────────────────────────────────────────
def adj_rate(vendor, qty, pct, charge):
    out = {"lower":None,"higher":None}
    if qty>1:
        lb = get_rate(vendor, qty-1)
        if lb is not None: out["lower"] = (qty-1, lb*(1+pct/100)+charge)
    hb = get_rate(vendor, qty+1)
    if hb is not None: out["higher"] = (qty+1, hb*(1+pct/100)+charge)
    return out

j_adj = adj_rate("Joda",       pallets, joda_in, j_charge)
m_adj = adj_rate("Mcdowells",  pallets, mcd_in,  m_charge)

st.subheader("One Pallet Fewer / One Pallet More")
c1,c2 = st.columns(2, gap="large")
with c1:
    st.markdown("**Joda**")
    lines=[]
    if j_adj["lower"]: lines.append(f"• {j_adj['lower'][0]}: £{j_adj['lower'][1]:,.2f}")
    else:                lines.append("• <span style='color:gray;'>N/A fewer</span>")
    if j_adj["higher"]: lines.append(f"• {j_adj['higher'][0]}: £{j_adj['higher'][1]:,.2f}")
    else:                lines.append("• <span style='color:gray;'>N/A more</span>")
    st.markdown("<br>".join(lines), unsafe_allow_html=True)
with c2:
    st.markdown("**McDowells**")
    lines=[]
    if m_adj["lower"]: lines.append(f"• {m_adj['lower'][0]}: £{m_adj['lower'][1]:,.2f}")
    else:                lines.append("• <span style='color:gray;'>N/A fewer</span>")
    if m_adj["higher"]: lines.append(f"• {m_adj['higher'][0]}: £{m_adj['higher'][1]:,.2f}")
    else:                lines.append("• <span style='color:gray;'>N/A more</span>")
    st.markdown("<br>".join(lines), unsafe_allow_html=True)

# ─────────────────────────────────────────
# (11) FOOTER NOTES
# ─────────────────────────────────────────
st.markdown("---")
st.markdown(
    "<small>"
    "• Joda’s surcharge resets Wednesday.  "
    "• McDowells’ is manual.  "
    "• Delivery: Joda AM/PM £7, Timed £19; McDowells AM/PM £10, Timed £19.  "
    "• Dual splits Joda in two shipments.  "
    "• Green row = cheapest."
    "</small>",
    unsafe_allow_html=True
)
