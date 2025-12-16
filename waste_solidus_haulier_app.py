# app.py — Solidus Haulier Rate Checker (with robust Joda fetch + fixed Map tab)

import os
import re
import math
import json
from datetime import date

import pandas as pd
import numpy as np
import streamlit as st
from PIL import Image

# For auto-fetch Joda surcharge
import requests
from bs4 import BeautifulSoup

# ---------- 1) STREAMLIT PAGE CONFIG ----------
st.set_page_config(page_title="Solidus Haulier Rate Checker", layout="wide")

st.markdown(
    """
    <style>
      #MainMenu { visibility: hidden; }
      footer { visibility: hidden; }
      .note { color: #6b7280; font-size: 0.9rem; }
    </style>
    """,
    unsafe_allow_html=True
)

# ---------- 2) HEADER ----------
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
        specify the number of pallets, and apply fuel surcharges and optional extras.
        """,
        unsafe_allow_html=True
    )

# ---------- 3) AUTO-FETCH JODA SURCHARGE (1h cache) ----------
@st.cache_data(ttl=3600)
def fetch_joda_surcharge() -> float:
    """
    Scrape Joda's fuel surcharge page for 'CURRENT SURCHARGE %'.
    Returns a float (e.g., 4.84 for 4.84%) or 0.0 if not found.
    """
    url = "https://www.jodafreight.com/fuel-surcharge/"
    headers = {
        "User-Agent": (
            "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
            "AppleWebKit/537.36 (KHTML, like Gecko) "
            "Chrome/123.0.0.0 Safari/537.36"
        )
    }
    try:
        r = requests.get(url, headers=headers, timeout=12)
        r.raise_for_status()
        html = r.text

        # Strategy 1: Find number close to 'CURRENT SURCHARGE %'
        m = re.search(
            r"CURRENT\s*SURCHARGE\s*%.*?([0-9]+[.,][0-9]+)\s*%",
            html,
            flags=re.I | re.S,
        )
        if m:
            return float(m.group(1).replace(",", "."))

        # Strategy 2: Parse with BeautifulSoup; look for any text containing '%' near 'surcharge'
        soup = BeautifulSoup(html, "lxml")
        texts = soup.find_all(string=re.compile(r"%"))
        for t in texts:
            parent_text = t.parent.get_text(" ", strip=True).lower() if hasattr(t, "parent") else ""
            this_text = (t or "").strip().lower()
            if "surcharge" in parent_text or "surcharge" in this_text:
                m2 = re.search(r"([0-9]+[.,][0-9]+)", t)
                if m2:
                    return float(m2.group(1).replace(",", "."))

        # Strategy 3: First percentage on page (sanity bounded)
        m3 = re.search(r"([0-9]+[.,][0-9]+)\s*%", html)
        if m3:
            val = float(m3.group(1).replace(",", "."))
            if 0.0 <= val <= 100.0:
                return val
    except Exception:
        pass

    return 0.0

# Manual refresh button for the fetch cache
st.caption("Joda surcharge auto-fetched (cached 1h).")
ref_col1, ref_col2 = st.columns([1, 5])
with ref_col1:
    if st.button("↻ Refresh", help="Clear 1-hour cache and refetch Joda surcharge"):
        fetch_joda_surcharge.clear()
        st.experimental_rerun()

# ---------- 4) LOAD RATES (cache-busted by file mtime) ----------
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

# ---------- 5) INPUTS ----------
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
    # Auto-fetched Joda value (allow override)
    joda_auto = fetch_joda_surcharge()
    override = st.checkbox("Override Joda surcharge", help="Tick to enter a surcharge manually.")
    joda_surcharge_pct = st.number_input(
        "Joda Fuel Surcharge (%)",
        min_value=0.0, max_value=100.0,
        value=float(round(joda_auto, 2) if not override else 0.0),
        step=0.1, format="%.2f", disabled=not override
    )
    if not override:
        joda_surcharge_pct = float(joda_auto)

with col_e:
    mcd_surcharge_pct = st.number_input(
        "McDowells Fuel Surcharge (%)", min_value=0.0, max_value=100.0,
        value=0.0, step=0.1, format="%.2f"
    )

with col_f:
    st.markdown(" ")

st.markdown("---")
postcode_area = input_area

# ---------- 6) OPTIONAL EXTRAS ----------
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

# ---------- Helpers ----------
def get_base_rate(df, area, service, vendor, pallets):
    subset = df[
        (df["PostcodeArea"] == area) &
        (df["Service"] == service) &
        (df["Vendor"] == vendor) &
        (df["Pallets"] == pallets)
    ]
    return None if subset.empty else float(subset["BaseRate"].iloc[0])

def effective_joda_surcharge(pallets_in_group: int, input_pct: float) -> float:
    """
    Business rule: Joda fuel surcharge does NOT apply for quantities 1–6 pallets
    (per group when split). 7 or more → apply input_pct.
    """
    return 0.0 if pallets_in_group < 7 else float(input_pct)

# ---------- 7) CALCULATE ----------
# Fixed delivery charges
joda_fixed = (7 if ampm_toggle else 0) + (19 if timed_toggle else 0)
mcd_fixed  = (10 if ampm_toggle else 0) + (19 if timed_toggle else 0)

# Joda calc
joda_base = None
joda_final = None

if dual_toggle:
    b1 = get_base_rate(rate_df, postcode_area, service_option, "Joda", split1)
    b2 = get_base_rate(rate_df, postcode_area, service_option, "Joda", split2)
    if b1 is not None and b2 is not None:
        # Apply per-group surcharge rule
        s1 = effective_joda_surcharge(split1, joda_surcharge_pct)
        s2 = effective_joda_surcharge(split2, joda_surcharge_pct)
        g1 = b1 * (1 + s1 / 100.0) + joda_fixed
        g2 = b2 * (1 + s2 / 100.0) + joda_fixed
        joda_base = b1 + b2
        joda_final = g1 + g2
else:
    base = get_base_rate(rate_df, postcode_area, service_option, "Joda", num_pallets)
    if base is not None:
        s_eff = effective_joda_surcharge(num_pallets, joda_surcharge_pct)
        joda_base = base
        joda_final = base * (1 + s_eff / 100.0) + joda_fixed

# McDowells calc (tail lift £3.90 per pallet)
mcd_base = get_base_rate(rate_df, postcode_area, service_option, "Mcdowells", num_pallets)
mcd_final = None
mcd_tail_per = 3.90 if tail_lift_toggle else 0.0
mcd_tail_total = mcd_tail_per * num_pallets

if mcd_base is not None:
    mcd_final = mcd_base * (1 + mcd_surcharge_pct / 100.0) + mcd_fixed + mcd_tail_total

# ---------- 8) RESULTS (Table + Map tabs) ----------
st.header("3. Results")

tab_table, tab_map = st.tabs(["Table", "Map (beta)"])

with tab_table:
    summary_rows = []
    if joda_base is None:
        summary_rows.append({
            "Haulier": "Joda",
            "Base Rate": "No rate",
            "Fuel Surcharge (%)": f"{joda_surcharge_pct:.2f}% (rule: <7 pallets → 0%)",
            "Delivery Charge": "N/A",
            "Final Rate": "N/A"
        })
    else:
        # Show the displayed surcharge column as entered (note: rule applied in pricing)
        summary_rows.append({
            "Haulier": "Joda",
            "Base Rate": f"£{joda_base:,.2f}",
            "Fuel Surcharge (%)": f"{joda_surcharge_pct:.2f}% (rule applied)",
            "Delivery Charge": f"£{joda_fixed:,.2f}",
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
            "Delivery Charge": f"£{(mcd_fixed + mcd_tail_total):,.2f}",
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
        st.markdown("<div class='note'>Rows in green are the cheapest available.</div>", unsafe_allow_html=True)

# ---------- Map helpers ----------
@st.cache_data
def load_centroids(csv_path: str) -> pd.DataFrame:
    """
    Load postcode area centroids from CSV and standardize columns to:
      Area, Latitude, Longitude
    Accepts variations like 'PostcodeArea', 'postcode area', 'lat', 'lng', 'lon', etc.
    """
    df = pd.read_csv(csv_path)

    # Normalize column names
    norm = {c: c.strip().lower().replace(" ", "").replace("_", "") for c in df.columns}
    df.columns = [norm[c] for c in df.columns]

    # Find likely columns
    area_col = None
    lat_col = None
    lon_col = None
    for c in df.columns:
        if c in ("area", "postcodearea", "postcodeareacode", "postcode"):
            area_col = c if area_col is None else area_col
        if c in ("lat", "latitude"):
            lat_col = c
        if c in ("lon", "long", "lng", "longitude"):
            lon_col = c

    # If nothing obvious, raise a controlled error
    if area_col is None or lat_col is None or lon_col is None:
        raise ValueError("Centroid CSV must contain area + latitude + longitude columns.")

    out = df[[area_col, lat_col, lon_col]].copy()
    out.columns = ["Area", "Latitude", "Longitude"]
    out["Area"] = out["Area"].astype(str).str.upper().str.strip()
    return out

def compute_final_for_area(area_code: str, pallets: int) -> tuple[float|None, float|None]:
    """Compute Joda & McD final for a given postcode area at current settings."""
    # Joda
    j_final = None
    base = get_base_rate(rate_df, area_code, service_option, "Joda", pallets)
    if base is not None:
        s_eff = effective_joda_surcharge(pallets, joda_surcharge_pct)
        j_final = base * (1 + s_eff / 100.0) + joda_fixed

    # McD
    m_final = None
    m_base = get_base_rate(rate_df, area_code, service_option, "Mcdowells", pallets)
    if m_base is not None:
        m_final = m_base * (1 + mcd_surcharge_pct / 100.0) + mcd_fixed + mcd_tail_per * pallets

    return j_final, m_final

with tab_map:
    st.caption("Map shows the final rate for each postcode area based on current options.")

    try:
        # Try user-provided CSV first; fallback to the filled one we shipped (if present)
        centroids_path_candidates = [
            "postcode_area_centroids_filled.csv",
            "postcode_area_centroids.csv",
            "/mnt/data/postcode_area_centroids_filled.csv",
            "/mnt/data/postcode_area_centroids.csv",
        ]
        centroid_df = None
        for pth in centroids_path_candidates:
            if os.path.exists(pth):
                centroid_df = load_centroids(pth)
                break
        if centroid_df is None:
            raise FileNotFoundError("No centroid CSV found.")

        # Build rate layer for current settings
        records = []
        for area_code in unique_areas:
            j_final, m_final = compute_final_for_area(area_code, num_pallets)
            # Keep rows where we have at least one value
            if (j_final is not None) or (m_final is not None):
                records.append({"Area": area_code, "JodaFinal": j_final, "McDFinal": m_final})

        rates_df = pd.DataFrame(records)
        if rates_df.empty:
            st.info("No mappable rates for the current filters.")
        else:
            merged = centroid_df.merge(rates_df, on="Area", how="inner")
            merged = merged.dropna(subset=["Latitude", "Longitude"])

            # Prepare tooltip label with both values
            def _fmt(v):
                return f"£{v:,.0f}" if pd.notna(v) else "N/A"

            merged["label"] = (
                merged["Area"]
                + "<br>Joda: " + merged["JodaFinal"].apply(_fmt)
                + "<br>McDowells: " + merged["McDFinal"].apply(_fmt)
            )

            # Use st.map via pydeck: two layers with scaling based on zoom is not supported
            # directly, so we choose a reasonable radius. (Streamlit auto-scales a bit.)
            import pydeck as pdk

            view_state = pdk.ViewState(
                latitude=float(merged["Latitude"].mean()),
                longitude=float(merged["Longitude"].mean()),
                zoom=5.2,
                pitch=0
            )

            # Joda layer (pink)
            j_layer = pdk.Layer(
                "ScatterplotLayer",
                data=merged.dropna(subset=["JodaFinal"]),
                get_position=["Longitude", "Latitude"],
                get_radius=25000,  # radius in meters (fixed)
                get_fill_color=[255, 105, 180, 120],  # RGBA
                pickable=True,
                tooltip=True
            )
            # McD layer (blue)
            m_layer = pdk.Layer(
                "ScatterplotLayer",
                data=merged.dropna(subset=["McDFinal"]),
                get_position=["Longitude", "Latitude"],
                get_radius=25000,
                get_fill_color=[78, 121, 237, 120],
                pickable=True,
                tooltip=True
            )

            tooltip = {"text": "{label}"}
            st.pydeck_chart(pdk.Deck(layers=[j_layer, m_layer], initial_view_state=view_state, tooltip=tooltip))
    except Exception as e:
        st.warning(f"Could not render map: {e}")

# ---------- 9) FOOTER ----------
st.markdown("---")
st.markdown(
    """
    <small>
    <b>What's NEW in V2.0?</b><br>
    • NEW: Auto-fetch Joda surcharge (cached 1h).<br>
    • NEW: Map View (Beta) with both rates in one tooltip.<br>
    • Rule: From 01/01/26 Joda fuel surcharge does not apply on 1–6 pallet quantities (per group when split).<br>
    <br>
    <b>Notes</b><br>
    • Joda: AM/PM £7, Timed £19; McDowells: AM/PM £10, Timed £19.<br>
    • Tail Lift: Joda £0; McDowells £3.90 per pallet.<br>
    • Dual Collection splits Joda into two shipments.<br>
    </small>
    """,
    unsafe_allow_html=True
)
