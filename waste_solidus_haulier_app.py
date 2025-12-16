# app.py — Solidus Haulier Rate Checker (auto-fetch Joda surcharge)

import streamlit as st
import pandas as pd
import pydeck as pdk
from PIL import Image
import json
from datetime import date
import os
import math
import re
import requests
from bs4 import BeautifulSoup

# ── 1) STREAMLIT PAGE CONFIG (must be first)
st.set_page_config(page_title="Solidus Haulier Rate Checker", layout="wide")

# ── 2) HIDE STREAMLIT MENU & FOOTER
st.markdown(
    """
    <style>
      #MainMenu { visibility: hidden; }
      footer { visibility: hidden; }
      .small-note { color:#6b7280; font-size:0.9em; }
    </style>
    """,
    unsafe_allow_html=True
)

# ── 3) HEADER (logo + title)
col_logo, col_text = st.columns([1, 3], gap="medium")

with col_logo:
    logo_path = "assets/solidus_logo.png"
    try:
        st.image(Image.open(logo_path), width=150)
    except Exception:
        st.warning(f"Could not load logo at '{logo_path}'.")

with col_text:
    st.markdown(
        "<h1 style='color:#0D4B6A; margin-bottom:0.25em;'>Solidus Haulier Rate Checker</h1>",
        unsafe_allow_html=True
    )
    st.markdown(
        """
        Enter a UK postcode area, select a service type (Economy or Next Day),  
        specify the number of pallets, and apply fuel surcharges and optional extras:

        • **Joda’s surcharge (%)** is now **auto-fetched** from Joda’s website (cached 1 hour).  
          You can still override it if needed.  
        • **McDowells’ surcharge (%)** is always entered manually each session.  
        • You may optionally add AM/PM Delivery, Tail Lift or Timed Delivery,  
          or perform a Dual Collection (For collections from both Unit 4 and ESL):
        """,
        unsafe_allow_html=True
    )

# ── 4) AUTO-FETCH JODA SURCHARGE (cached 1 hour)
@st.cache_data(ttl=3600)
def fetch_joda_surcharge() -> float:
    """
    Scrape Joda's fuel surcharge page for 'CURRENT SURCHARGE %'.
    Returns float percentage (e.g., 4.84 for 4.84%) or 0.0 if not found.
    """
    url = "https://www.jodafreight.com/fuel-surcharge/"
    try:
        r = requests.get(url, timeout=10)
        r.raise_for_status()
        soup = BeautifulSoup(r.text, "lxml")
        text = soup.get_text(" ", strip=True)
        # Look for 'current surcharge' nearby a % number e.g. 4,84% or 4.84%
        m = re.search(r"current\s*surcharge\s*%[^0-9]*([0-9]+[.,][0-9]+)\s*%", text, flags=re.I)
        if not m:
            m = re.search(r"([0-9]+[.,][0-9]+)\s*%", text)
        if m:
            raw = m.group(1).replace(",", ".")
            val = float(raw)
            if 0 <= val <= 100:
                return val
    except Exception:
        pass
    return 0.0

# ── 5) LOAD + TRANSFORM RATES (cache-busted by file mtime)
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

mtime = os.path.getmtime("haulier prices.xlsx")
rate_df = load_rate_table("haulier prices.xlsx", mtime)
unique_areas = sorted(rate_df["PostcodeArea"].unique())

# ── 6) USER INPUTS
st.header("1. Input Parameters")
col_a, col_b, col_c, col_d, col_e, col_f = st.columns([1,1,1,1,1,1], gap="medium")

with col_a:
    input_area = st.selectbox(
        "Postcode Area",
        options=[""] + unique_areas,
        index=0,
        format_func=lambda x: x if x else "— Select area —"
    )
    if input_area == "":
        st.info("Please select a postcode area to continue.")
        st.stop()

with col_b:
    service_option = st.selectbox("Service Type", options=["Economy", "Next Day"], index=0)

with col_c:
    num_pallets = st.number_input("Number of Pallets", min_value=1, max_value=26, value=1, step=1)

with col_d:
    fetched = fetch_joda_surcharge()
    st.caption("Joda surcharge auto-fetched (cached 1h).")
    override = st.checkbox("Override Joda surcharge", value=False, help="Use if the auto value looks wrong.")
    if override:
        joda_surcharge_pct = st.number_input(
            "Joda Fuel Surcharge (%)",
            min_value=0.0, max_value=100.0,
            value=fetched, step=0.1, format="%.2f"
        )
    else:
        joda_surcharge_pct = fetched
        st.number_input(
            "Joda Fuel Surcharge (%)",
            min_value=0.0, max_value=100.0,
            value=joda_surcharge_pct, step=0.1, format="%.2f", disabled=True
        )

with col_e:
    mcd_surcharge_pct = st.number_input(
        "McDowells Fuel Surcharge (%)", min_value=0.0, max_value=100.0,
        value=0.0, step=0.1, format="%.2f"
    )

with col_f:
    st.markdown(" ")

st.markdown("---")
postcode_area = input_area

# ── 7) OPTIONAL EXTRAS (includes Tail Lift)
st.subheader("2. Optional Extras")
col1, col2, col3, col4 = st.columns(4, gap="large")

with col1:
    ampm_toggle = st.checkbox("AM/PM Delivery")
with col2:
    tail_lift_toggle = st.checkbox("Tail Lift")  # Joda £0, McD £3.90 per pallet
with col3:
    dual_toggle = st.checkbox("Dual Collection")
with col4:
    timed_toggle = st.checkbox("Timed Delivery")

# Guard: dual requires at least 2 pallets
if dual_toggle and num_pallets == 1:
    st.error("Dual Collection requires at least 2 pallets.")
    st.stop()

split1 = split2 = None
if dual_toggle:
    st.markdown("**Split pallets into two despatches (e.g., ESL & U4).**")
    sp1, sp2 = st.columns(2, gap="large")
    with sp1:
        split1 = st.number_input("First Pallet Group", 1, num_pallets - 1, 1)
    with sp2:
        split2 = st.number_input("Second Pallet Group", 1, num_pallets - 1, num_pallets - 1)
    if split1 + split2 != num_pallets:
        st.error("Pallet Split values must add up to total pallets.")
        st.stop()

# ── 8) RATE LOOKUP + CALC
def get_base_rate(df, area, service, vendor, pallets):
    subset = df[
        (df["PostcodeArea"] == area) &
        (df["Service"] == service) &
        (df["Vendor"] == vendor) &
        (df["Pallets"] == pallets)
    ]
    return None if subset.empty else float(subset["BaseRate"].iloc[0])

def apply_joda_surcharge(base_rate: float, pallets: int, surcharge_pct: float, fixed_charge: float) -> float:
    """
    Apply Joda surcharge unless pallet qty is < 7.
    """
    if pallets < 7:
        # No fuel surcharge for 1–6 pallets (effective 01/01/26)
        return base_rate + fixed_charge
    return base_rate * (1 + surcharge_pct / 100.0) + fixed_charge

# Joda (tail lift = £0)
joda_base = None
joda_final = None
joda_charge_fixed = (7 if ampm_toggle else 0) + (19 if timed_toggle else 0)

if dual_toggle:
    b1 = get_base_rate(rate_df, postcode_area, service_option, "Joda", split1)
    b2 = get_base_rate(rate_df, postcode_area, service_option, "Joda", split2)
    if b1 is not None and b2 is not None:
        g1 = apply_joda_surcharge(b1, split1, joda_surcharge_pct, joda_charge_fixed)
        g2 = apply_joda_surcharge(b2, split2, joda_surcharge_pct, joda_charge_fixed)
        joda_base = b1 + b2
        joda_final = g1 + g2
else:
    base = get_base_rate(rate_df, postcode_area, service_option, "Joda", num_pallets)
    if base is not None:
        joda_base = base
        joda_final = apply_joda_surcharge(base, num_pallets, joda_surcharge_pct, joda_charge_fixed)

# McDowells (tail lift = £3.90 per pallet when toggled)
mcd_base = get_base_rate(rate_df, postcode_area, service_option, "Mcdowells", num_pallets)
mcd_final = None
mcd_charge_fixed = (10 if ampm_toggle else 0) + (19 if timed_toggle else 0)
mcd_tail_lift_per_pallet = 3.90 if tail_lift_toggle else 0.0
mcd_tail_lift_total = mcd_tail_lift_per_pallet * num_pallets

if mcd_base is not None:
    mcd_final = (
        mcd_base * (1 + mcd_surcharge_pct / 100.0)
        + mcd_charge_fixed
        + mcd_tail_lift_total
    )

# ── 9) “Calculated Rates” TABS (Table / Map / History)
st.header("3. Calculated Rates")
tab_table, tab_map, tab_history = st.tabs(["Table", "Map (beta)", "History"])

# ---- A) TABLE TAB
with tab_table:
    summary_rows = []

    if joda_base is None:
        summary_rows.append({
            "Haulier": "Joda",
            "Base Rate": "No rate",
            "Fuel Surcharge (%)": f"{joda_surcharge_pct:.2f}%",
            "Delivery Charge": "N/A",
            "Final Rate": "N/A"
        })
    else:
        summary_rows.append({
            "Haulier": "Joda",
            "Base Rate": f"£{joda_base:,.2f}",
            "Fuel Surcharge (%)": f"{joda_surcharge_pct:.2f}%",
            "Delivery Charge": f"£{joda_charge_fixed:,.2f}",
            "Final Rate": f"£{joda_final:,.2f}"
        })

    if mcd_base is None:
        summary_rows.append({
            "Haulier": "McDowells",
            "Base Rate": "No rate",
            "Fuel Surcharge (%)": f"{mcd_surcharge_pct:.2f}%",
            "Delivery Charge": "N/A",
            "Final Rate": "N/A"
        })
    else:
        summary_rows.append({
            "Haulier": "McDowells",
            "Base Rate": f"£{mcd_base:,.2f}",
            "Fuel Surcharge (%)": f"{mcd_surcharge_pct:.2f}%",
            "Delivery Charge": f"£{(mcd_charge_fixed + mcd_tail_lift_total):,.2f}",
            "Final Rate": f"£{mcd_final:,.2f}"
        })

    if all(r["Final Rate"] == "N/A" for r in summary_rows):
        st.warning("No rates found for that area/service/pallet combination.")
    else:
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

        st.table(summary_df.style.apply(highlight_cheapest, axis=1))
        st.markdown("<div class='small-note'>Rows in green are the cheapest available.</div>", unsafe_allow_html=True)

# ---- B) MAP TAB
with tab_map:
    st.caption("Shows rates by postcode area centroid. Uses file: postcode_area_centroids.csv (PostcodeArea,Lat,Lon).")
    centroid_path = "postcode_area_centroids.csv"
    if os.path.exists(centroid_path):
        try:
            centroid_df = pd.read_csv(centroid_path)
            # normalise area code format
            centroid_df["PostcodeArea"] = centroid_df["PostcodeArea"].astype(str).str.upper().str.strip()

            # Build a DF of rates for the selected area/service across pallet counts,
            # then merge with centroids to display points for *all* areas present in the sheet.
            # We compute current Joda/McD final rate at the selected pallet count for each area.
            areas = rate_df["PostcodeArea"].unique()
            rows_map = []
            for area in areas:
                # Joda
                jb = get_base_rate(rate_df, area, service_option, "Joda", num_pallets)
                jf = None
                if jb is not None:
                    jf = apply_joda_surcharge(jb, num_pallets, joda_surcharge_pct, joda_charge_fixed)
                # McD
                mb = get_base_rate(rate_df, area, service_option, "Mcdowells", num_pallets)
                mf = None
                if mb is not None:
                    mf = mb * (1 + mcd_surcharge_pct/100.0) + mcd_charge_fixed + (mcd_tail_lift_per_pallet * num_pallets)

                rows_map.append({
                    "Area": area,
                    "JodaFinal": jf,
                    "McdFinal": mf
                })
            rate_map_df = pd.DataFrame(rows_map)
            map_df = centroid_df.merge(rate_map_df, on="Area", how="inner").dropna(subset=["Latitude","Longitude"])

            # Colour & radius scale
            # Pink for Joda, Blue for McD; we render both layers.
            joda_layer = pdk.Layer(
                "ScatterplotLayer",
                data=map_df,
                get_position='[Longitude, Latitude]',
                get_radius=12000,      # meters; keeps a sensible size as you zoom
                pickable=True,
                get_fill_color=[255, 20, 147, 120],  # pink-ish
            )
            mcd_layer = pdk.Layer(
                "ScatterplotLayer",
                data=map_df,
                get_position='[Longitude, Latitude]',
                get_radius=10000,
                pickable=True,
                get_fill_color=[70, 130, 180, 120],  # steel blue
            )

            tooltip = {
                "html": "<b>{Area}</b><br/>Joda: {JodaFinal}<br/>McDowells: {McdFinal}",
                "style": {"backgroundColor": "rgba(32,32,32,0.85)", "color": "white"}
            }

            # Format currency strings for tooltip
            map_df["JodaFinal"] = map_df["JodaFinal"].apply(lambda v: f"£{v:,.0f}" if pd.notnull(v) else "N/A")
            map_df["McdFinal"] = map_df["McdFinal"].apply(lambda v: f"£{v:,.0f}" if pd.notnull(v) else "N/A")

            view_state = pdk.ViewState(latitude=53.8, longitude=-2.4, zoom=5.2)
            r = pdk.Deck(layers=[joda_layer, mcd_layer], initial_view_state=view_state, tooltip=tooltip, map_style="dark")
            st.pydeck_chart(r, use_container_width=True)
        except Exception as e:
            st.warning(f"Could not render map: {e}")
    else:
        st.info("No centroid file found (postcode_area_centroids.csv). Map view disabled.")

# ---- C) HISTORY TAB
with tab_history:
    if "history" not in st.session_state:
        st.session_state.history = []
    # record current search if there is at least one valid rate
    if (joda_final is not None) or (mcd_final is not None):
        entry = {
            "Area": postcode_area,
            "Service": service_option,
            "Pallets": num_pallets,
            "Dual": bool(dual_toggle),
            "Split": f"{split1}+{split2}" if dual_toggle else "-",
            "AM/PM": bool(ampm_toggle),
            "Timed": bool(timed_toggle),
            "TailLift": bool(tail_lift_toggle),
            "JodaFinal": f"£{joda_final:,.2f}" if joda_final is not None else "N/A",
            "McDFinal": f"£{mcd_final:,.2f}" if mcd_final is not None else "N/A",
        }
        # push newest to front; keep 10
        if not st.session_state.history or st.session_state.history[0] != entry:
            st.session_state.history.insert(0, entry)
            st.session_state.history = st.session_state.history[:10]

    if st.session_state.history:
        st.dataframe(pd.DataFrame(st.session_state.history))
    else:
        st.caption("No history yet — run a search to populate.")

# ── 10) ONE PALLET FEWER / MORE
def lookup_adjacent_rate(df, area, service, vendor, pallets,
                         surcharge_pct, fixed_charge=0.0, per_pallet_charge=0.0,
                         is_joda=False):
    out = {"lower": None, "higher": None}
    if pallets > 1:
        bl = get_base_rate(df, area, service, vendor, pallets - 1)
        if bl is not None:
            if is_joda:
                val = apply_joda_surcharge(bl, pallets-1, surcharge_pct, fixed_charge)
            else:
                val = bl * (1 + surcharge_pct / 100.0) + fixed_charge + per_pallet_charge * (pallets - 1)
            out["lower"] = ((pallets - 1), val)
    bh = get_base_rate(df, area, service, vendor, pallets + 1)
    if bh is not None:
        if is_joda:
            val = apply_joda_surcharge(bh, pallets+1, surcharge_pct, fixed_charge)
        else:
            val = bh * (1 + surcharge_pct / 100.0) + fixed_charge + per_pallet_charge * (pallets + 1)
        out["higher"] = ((pallets + 1), val)
    return out

joda_adj = lookup_adjacent_rate(
    rate_df, postcode_area, service_option, "Joda",
    num_pallets, joda_surcharge_pct,
    fixed_charge=joda_charge_fixed, per_pallet_charge=0.0, is_joda=True
)
mcd_adj = lookup_adjacent_rate(
    rate_df, postcode_area, service_option, "Mcdowells",
    num_pallets, mcd_surcharge_pct,
    fixed_charge=mcd_charge_fixed, per_pallet_charge=mcd_tail_lift_per_pallet, is_joda=False
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

# ── 11) FOOTER
st.markdown("---")
st.markdown(
    """
    <small>
    <b>What's NEW in V2.0?</b><br/>
    • From 01/01/26 Joda fuel surcharge does not apply on 1–6 pallet quantities (per group when split).<br/>
    • Map View (Beta).<br/>
    • History tab (last 10 searches).<br/><br/>

    <b>Notes</b><br/>
    • Joda surcharge is auto-fetched (cached 1h).<br/>
    • Delivery charges: Joda – AM/PM £7, Timed £19; McDowells – AM/PM £10, Timed £19.<br/>
    • Tail Lift: Joda £0; McDowells £3.90 per pallet.<br/>
    • Dual Collection splits Joda into two shipments; McDowells unaffected.
    </small>
    """,
    unsafe_allow_html=True
)
