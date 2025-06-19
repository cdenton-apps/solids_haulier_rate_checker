# waste_solidus_haulier_app.py

import streamlit as st
import pandas as pd
import math
from PIL import Image
import json
from datetime import date
import os

# 1) Page config
st.set_page_config(page_title="Solidus Haulier Rate Checker", layout="wide")

# 2) Hide default menu/footer
st.markdown("""
    <style>
      #MainMenu {visibility: hidden;}
      footer {visibility: hidden;}
    </style>
""", unsafe_allow_html=True)

# 3) Header + logo
col_logo, col_text = st.columns([1,3], gap="medium")
with col_logo:
    try:
        st.image(Image.open("assets/solidus_logo.png"), width=150)
    except:
        st.warning("Logo not found.")
with col_text:
    st.markdown(
        "<h1 style='color:#0D4B6A;'>Solidus Haulier Rate Checker</h1>",
        unsafe_allow_html=True
    )
    st.markdown(
        "Enter area, service, pallets, surcharges and extras. "
        "Joda’s surcharge (resets Wed) is persisted; McDowells’ is manual.",
        unsafe_allow_html=True
    )

# 4) Persist Joda surcharge
DATA_FILE = "joda_surcharge.json"
def load_joda_surcharge():
    today = date.today().isoformat()
    if not os.path.exists(DATA_FILE):
        json.dump({"surcharge":0.0,"last_updated":today}, open(DATA_FILE,"w"))
        return 0.0
    data = json.load(open(DATA_FILE))
    if date.today().weekday()==2 and data.get("last_updated")!=today:
        data.update({"surcharge":0.0,"last_updated":today})
        json.dump(data, open(DATA_FILE,"w"))
        return 0.0
    return data.get("surcharge",0.0)

def save_joda_surcharge(p):
    today = date.today().isoformat()
    json.dump({"surcharge":float(p),"last_updated":today}, open(DATA_FILE,"w"))

joda_pct = load_joda_surcharge()

# 5) Load & reshape rates Excel
@st.cache_data
def load_rates(path):
    df = pd.read_excel(path, header=1)
    df = df.rename(columns={df.columns[0]:"PostcodeArea", df.columns[1]:"Service", df.columns[2]:"Vendor"})
    df["PostcodeArea"]=df["PostcodeArea"].ffill()
    df["Service"]=df["Service"].ffill()
    df=df[df["Vendor"]!="Vendor"].copy()
    pallet_cols=[c for c in df.columns if (isinstance(c,(int,float))) or (isinstance(c,str) and c.isdigit())]
    m = df.melt(["PostcodeArea","Service","Vendor"], pallet_cols, "Pallets","BaseRate")
    m["Pallets"]=m["Pallets"].astype(int)
    m=m.dropna(subset=["BaseRate"])
    m["PostcodeArea"]=m["PostcodeArea"].astype(str).str.upper().str.strip()
    m["Service"]=m["Service"].str.title().str.strip()
    m["Vendor"]=m["Vendor"].str.title().str.strip()
    return m.reset_index(drop=True)

rates = load_rates("haulier prices.xlsx")
areas = sorted(rates["PostcodeArea"].unique())

# 6) Inputs
st.header("1. Input Parameters")
c1,c2,c3,c4,c5,c6 = st.columns(6)
with c1:
    area = st.selectbox("Postcode Area", [""]+areas, format_func=lambda x: x or "— Select —")
    if not area: st.stop()
with c2:
    service = st.selectbox("Service", ["Economy","Next Day"])
with c3:
    pallets = st.number_input("Pallets",1,26,1)
with c4:
    joda_in = st.number_input("Joda FSC (%)",0.0,100.0,value=round(joda_pct,2),step=0.1,format="%.2f")
    if st.button("Save Joda FSC"):
        save_joda_surcharge(joda_in); st.success("Saved.")
with c5:
    mcd_in = st.number_input("McDowells FSC (%)",0.0,100.0,0.0,0.1,format="%.2f")
with c6:
    pass

st.markdown("---")

# 7) Optional extras
st.subheader("2. Optional Extras")
e1,e2,e3 = st.columns(3)
ampm = e1.checkbox("AM/PM Delivery")
timed= e2.checkbox("Timed Delivery")
dual = e3.checkbox("Dual Collection")

split1=split2=None
if dual:
    st.markdown("Split pallets into two shipments:")
    s1,s2=st.columns(2)
    with s1: split1=st.number_input("Group 1",1,pallets-1,1)
    with s2: split2=st.number_input("Group 2",1,pallets-1,pallets-1)
    if split1+split2!=pallets: st.error("Sum mismatch"); st.stop()

# helper to look up base rate or None
def get_rate(vendor, qty):
    df=rates
    sub = df[(df.PostcodeArea==area)
             &(df.Service==service)
             &(df.Vendor==vendor)
             &(df.Pallets==qty)]
    return None if sub.empty else float(sub.BaseRate.iloc[0])

# 8) Calculate
j_charge = (7 if ampm else 0)+(19 if timed else 0)
m_charge = (10 if ampm else 0)+(19 if timed else 0)

# Joda
j_base=None; j_final=None
if dual:
    b1=get_rate("Joda",split1); b2=get_rate("Joda",split2)
    if b1 is not None and b2 is not None:
        g1=b1*(1+joda_in/100)+j_charge
        g2=b2*(1+joda_in/100)+j_charge
        j_base=b1+b2; j_final=g1+g2
else:
    b=get_rate("Joda",pallets)
    if b is not None:
        j_base=b; j_final=b*(1+joda_in/100)+j_charge

# McDowells
m_base=get_rate("Mcdowells",pallets)
m_final = None if m_base is None else m_base*(1+mcd_in/100)+m_charge

# 9) Build summary with "No rate"/"N/A"
rows=[]
# Joda row
if j_base is None:
    rows.append({"Haulier":"Joda","Base Rate":"No rate","Fuel Surcharge (%)":f"{joda_in:.2f}%","Delivery Charge":"N/A","Final Rate":"N/A"})
else:
    rows.append({"Haulier":"Joda","Base Rate":f"£{j_base:,.2f}","Fuel Surcharge (%)":f"{joda_in:.2f}%","Delivery Charge":f"£{j_charge:,.2f}","Final Rate":f"£{j_final:,.2f}"})
# McDowells row
if m_base is None:
    rows.append({"Haulier":"McDowells","Base Rate":"No rate","Fuel Surcharge (%)":f"{mcd_in:.2f}%","Delivery Charge":"N/A","Final Rate":"N/A"})
else:
    rows.append({"Haulier":"McDowells","Base Rate":f"£{m_base:,.2f}","Fuel Surcharge (%)":f"{mcd_in:.2f}%","Delivery Charge":f"£{m_charge:,.2f}","Final Rate":f"£{m_final:,.2f}"})

# show warning if none exist
if all(r["Final Rate"]=="N/A" for r in rows):
    st.warning("No rates found for that combination.")
else:
    df=pd.DataFrame(rows).set_index("Haulier")
    def highlight(x):
        fr=x["Final Rate"]
        if fr.startswith("£"):
            v=float(fr.strip("£").replace(",",""))
            cheapest=min([val for val in [(j_final or 1e9),(m_final or 1e9)]])
            return ["background-color:#b3e6b3"]*len(x) if abs(v-cheapest)<1e-6 else [""]*len(x)
        return [""]*len(x)
    st.header("3. Calculated Rates")
    st.table(df.style.apply(highlight,axis=1))
