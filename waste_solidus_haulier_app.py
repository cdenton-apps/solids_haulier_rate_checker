# app.py — Solidus Haulier Rate Checker (table+map tabs, pixel-scaling markers)

import os
import math
import json
from datetime import date

import pandas as pd
import streamlit as st
from PIL import Image


# ── 1) STREAMLIT PAGE CONFIG (must be first)
st.set_page_config(page_title="Solidus Haulier Rate Checker", layout="wide")

# ── 2) HIDE STREAMLIT MENU & FOOTER
st.markdown(
    """
    <style>
      #MainMenu { visibility: hidden; }
      footer { visibility: hidden; }
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

        **Joda’s surcharge (%)** is stored persistently and must be updated once weekly  
        (see: https://www.jodafreight.com/fuel-surcharge/).  
        It automatically **resets to 0 on Wednesdays**.

        **McDowells’ surcharge (%)** is always entered manually each session.  
        You may optionally add **AM/PM Delivery**, **Tail Lift** or **Timed Delivery**,  
        and optionally perform a **Dual Collection** (split load between ESL & U4).
        """,
        unsafe_allow_html=True
    )

# ── 4) PERSIST JODA SURCHARGE (resets each Wednesday)
DATA_FILE = "joda_surcharge.json"

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

    # Wednesday reset (weekday() == 2)
    if date.today().weekday() == 2 and data.get("last_updated") != today_str:
        data = {"surcharge": 0.0, "last_updated": today_str}
        with open(DATA_FILE, "w") as f:
            json.dump(data, f)
        return 0.0
    return float(data.get("surcharge", 0.0))

def save_joda_surcharge(new_pct: float):
    today_str = date.today().isoformat()
    with open(DATA_FILE, "w") as f:
        json.dump({"surcharge": float(new_pct), "last_updated": today_str}, f)

joda_stored_pct = load_joda_surcharge()

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
    joda_surcharge_pct = st.number_input(
        "Joda Fuel Surcharge (%)", min_value=0.0, max_value=100.0,
        value=round(joda_stored_pct, 2), step=0.1, format="%.2f"
    )
    if st.button("Save Joda Surcharge"):
        save_joda_surcharge(joda_surcharge_pct)
        st.success(f"Saved Joda surcharge at {joda_surcharge_pct:.2f}%")

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

# Joda rule: fuel surcharge **does not apply** if pallet count (group or total) is < 7
def joda_effective_pct(pallet_count: int, input_pct: float) -> float:
    return 0.0 if pallet_count < 7 else float(input_pct)

# Joda (tail lift = £0)
joda_base = None
joda_final = None
joda_charge_fixed = (7 if ampm_toggle else 0) + (19 if timed_toggle else 0)

if dual_toggle:
    b1 = get_base_rate(rate_df, postcode_area, service_option, "Joda", split1)
    b2 = get_base_rate(rate_df, postcode_area, service_option, "Joda", split2)
    if b1 is not None and b2 is not None:
        eff1 = joda_effective_pct(split1, joda_surcharge_pct)
        eff2 = joda_effective_pct(split2, joda_surcharge_pct)
        g1 = b1 * (1 + eff1 / 100.0) + joda_charge_fixed
        g2 = b2 * (1 + eff2 / 100.0) + joda_charge_fixed
        joda_base = b1 + b2
        joda_final = g1 + g2
else:
    base = get_base_rate(rate_df, postcode_area, service_option, "Joda", num_pallets)
    if base is not None:
        joda_base = base
        eff = joda_effective_pct(num_pallets, joda_surcharge_pct)
        joda_final = base * (1 + eff / 100.0) + joda_charge_fixed

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

# ── 9) SUMMARY TABLE (with “No rate” handling)
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
    shown_pct = (
        joda_effective_pct(num_pallets, joda_surcharge_pct)
        if not dual_toggle else float(joda_surcharge_pct)
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

st.header("3. Calculated Rates")
tab_table, tab_map = st.tabs(["Table", "Map"])

with tab_table:
    if all(r["Final Rate"] == "N/A" for r in summary_rows):
        st.warning("No rates found for that area/service/pallet combination.")
    else:
        st.table(summary_df.style.apply(highlight_cheapest, axis=1))
        st.markdown(
            "<i style='color:gray;'>Rows in green are the cheapest available. "
            "Joda fuel surcharge is **waived** for pallet counts below 7 (per group when split).</i>",
            unsafe_allow_html=True
        )

# ── 10) ONE PALLET FEWER / MORE (respects Joda <7 rule, and McD per-pallet tail lift)
def lookup_adjacent_rate(df, area, service, vendor, pallets,
                         surcharge_pct, fixed_charge=0.0, per_pallet_charge=0.0,
                         joda_rule=False):
    out = {"lower": None, "higher": None}

    def eff_pct(n):
        if joda_rule:
            return joda_effective_pct(n, surcharge_pct)
        return surcharge_pct

    if pallets > 1:
        bl = get_base_rate(df, area, service, vendor, pallets - 1)
        if bl is not None:
            out["lower"] = (
                (pallets - 1),
                bl * (1 + eff_pct(pallets - 1) / 100.0) + fixed_charge + per_pallet_charge * (pallets - 1)
            )
    bh = get_base_rate(df, area, service, vendor, pallets + 1)
    if bh is not None:
        out["higher"] = (
            (pallets + 1),
            bh * (1 + eff_pct(pallets + 1) / 100.0) + fixed_charge + per_pallet_charge * (pallets + 1)
        )
    return out

joda_adj = lookup_adjacent_rate(
    rate_df, postcode_area, service_option, "Joda",
    num_pallets, joda_surcharge_pct,
    fixed_charge=joda_charge_fixed, per_pallet_charge=0.0, joda_rule=True
)
mcd_adj = lookup_adjacent_rate(
    rate_df, postcode_area, service_option, "Mcdowells",
    num_pallets, mcd_surcharge_pct,
    fixed_charge=mcd_charge_fixed, per_pallet_charge=mcd_tail_lift_per_pallet, joda_rule=False
)

with tab_table:
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

# ── 11) MAP TAB (inside "3. Calculated Rates") – BOTH rates in tooltip; pixel-scaling markers
with tab_map:
    # Try user file first; fall back to filled centroids if present
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
                lat_col  = _find_col(tmp, ["lat", "latitude", "y"])
                lon_col  = _find_col(tmp, ["lon", "lng", "longitude", "x"])

                if not area_col or not lat_col or not lon_col:
                    continue

                centroid_df = tmp.rename(columns={
                    area_col: "Area",
                    lat_col:  "lat",
                    lon_col:  "lon"
                }).copy()

                centroid_df["Area"] = centroid_df["Area"].astype(str).str.upper().str.strip()
                centroid_df["lat"]  = pd.to_numeric(centroid_df["lat"], errors="coerce")
                centroid_df["lon"]  = pd.to_numeric(centroid_df["lon"], errors="coerce")
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
            # Joda (respect Joda <7 rule and split)
            jb = None
            jf = None
            if dual_toggle:
                b1 = get_base_rate(rate_df, area_code, service_option, "Joda", split1)
                b2 = get_base_rate(rate_df, area_code, service_option, "Joda", split2)
                if b1 is not None and b2 is not None:
                    p1 = joda_effective_pct(split1, joda_surcharge_pct)
                    p2 = joda_effective_pct(split2, joda_surcharge_pct)
                    jf = b1 * (1 + p1/100.0) + joda_charge_fixed
                    jf += b2 * (1 + p2/100.0) + joda_charge_fixed
                    jb = b1 + b2
            else:
                jb = get_base_rate(rate_df, area_code, service_option, "Joda", num_pallets)
                if jb is not None:
                    ep = joda_effective_pct(num_pallets, joda_surcharge_pct)
                    jf = jb * (1 + ep/100.0) + joda_charge_fixed

            # McDowells
            mb = get_base_rate(rate_df, area_code, service_option, "Mcdowells", num_pallets)
            mf = None
            if mb is not None:
                mf = mb * (1 + mcd_surcharge_pct/100.0) + mcd_charge_fixed + (3.90 if tail_lift_toggle else 0.0)*num_pallets

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
            # Use pixel-based radius so marker size stays sensible while zooming
            # (tweak this if you want larger/smaller dots)
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
                get_radius="size",                # pixels
                radius_units="pixels",            # ← keeps size consistent as you zoom
                pickable=True,
                get_fill_color="""
                    [cheaper == 'JodaFinal' ? 255 : 90,
                     cheaper == 'JodaFinal' ? 64  : 90,
                     cheaper == 'JodaFinal' ? 160 : 255, 200]
                """,
            )

            view_state = pdk.ViewState(latitude=54.5, longitude=-2.5, zoom=4.8)
            st.pydeck_chart(pdk.Deck(layers=[layer], initial_view_state=view_state, tooltip=tooltip))

# ── 12) FOOTER
st.markdown("---")
st.markdown(
    """
    <small>
    • Joda surcharge resets each Wednesday; McDowells is entered per session.  
    • Delivery charges: Joda – AM/PM £7, Timed £19; McDowells – AM/PM £10, Timed £19.  
    • Tail Lift: Joda £0; McDowells £3.90 per pallet.  
    • Dual Collection splits Joda into two shipments; McDowells unaffected.  
    • <b>Joda fuel surcharge is not applied for pallet counts below 7</b> (per group when split).
    </small>
    """,
    unsafe_allow_html=True
)
