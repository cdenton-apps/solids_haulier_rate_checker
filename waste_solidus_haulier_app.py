# ======================================
# File: waste_solidus_haulier_app.py
# ======================================

import streamlit as st
import pandas as pd
import math
from PIL import Image

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
# (3) DISPLAY SOLIDUS LOGO + HEADER SIDE‐BY‐SIDE
# ─────────────────────────────────────────
col_logo, col_text = st.columns([1, 3], gap="medium")

with col_logo:
    logo_path = "assets/solidus_logo.png"
    try:
        logo_img = Image.open(logo_path)
        st.image(logo_img, width=150)
    except Exception:
        st.warning(
            f"⚠️ Could not load logo at '{logo_path}'. Please confirm the file exists and is a valid PNG."
        )

with col_text:
    st.markdown(
        "<h1 style='color:#0D4B6A; margin-bottom:0.25em;'>"
        "Solidus Haulier Rate Checker</h1>",
        unsafe_allow_html=True
    )
    st.markdown(
        """
        Enter a UK postcode, select a service type (Economy or Next Day),  
        specify the number of pallets, and apply fuel surcharges and optional extras:

        • **Joda’s surcharge (%)** must be copied manually from:  
          https://www.jodafreight.com/fuel-surcharge/ (see “CURRENT SURCHARGE %”).  
        • **McDowells’ surcharge (%)** is always entered manually.  
        • You may optionally add AM/PM Delivery or Timed Delivery,  
          or perform a Dual Collection (split your pallets):

          - Joda: AM/PM = £7, Timed = £19  
          - McDowells: AM/PM = £10, Timed = £19

        The app will then:
        1. Calculate the final adjusted rate for both Joda and McDowells.  
        2. Highlight the cheapest option in green.  
        3. Show the price for one fewer and one more pallet (greyed out if unavailable).
        """,
        unsafe_allow_html=True
    )

# ─────────────────────────────────────────
# (4) LOAD & TRANSFORM THE BUILT-IN EXCEL DATA
# ─────────────────────────────────────────
@st.cache_data
def load_rate_table(excel_path: str) -> pd.DataFrame:
    """
    Read 'haulier prices.xlsx' with header=1 (second row as header), forward-fill 
    PostcodeArea & Service, then melt numeric pallet columns into:
      PostcodeArea | Service | Vendor | Pallets | BaseRate
    """
    raw = pd.read_excel(excel_path, header=1)

    # Rename first three columns:
    #   “Unnamed: 0” → “PostcodeArea”
    #   “Unnamed: 1” → “Service”
    #   “Unnamed: 2” → “Vendor”
    raw = raw.rename(columns={
        raw.columns[0]: "PostcodeArea",
        raw.columns[1]: "Service",
        raw.columns[2]: "Vendor"
    })

    # Forward-fill PostcodeArea & Service down through blank cells
    raw["PostcodeArea"] = raw["PostcodeArea"].ffill()
    raw["Service"] = raw["Service"].ffill()

    # Drop any rows where Vendor == "Vendor" (leftover header row)
    raw = raw[raw["Vendor"] != "Vendor"].copy()

    # Identify which columns represent pallet counts (e.g., 1–26)
    pallet_cols = [
        col for col in raw.columns
        if isinstance(col, (int, float)) or (isinstance(col, str) and col.isdigit())
    ]

    # Melt from wide to long: each (PostcodeArea, Service, Vendor, Pallets) → BaseRate
    melted = raw.melt(
        id_vars=["PostcodeArea", "Service", "Vendor"],
        value_vars=pallet_cols,
        var_name="Pallets",
        value_name="BaseRate"
    )
    melted["Pallets"] = melted["Pallets"].astype(int)
    melted = melted.dropna(subset=["BaseRate"]).copy()

    # Standardize text columns
    melted["PostcodeArea"] = melted["PostcodeArea"].astype(str).str.strip().str.upper()
    melted["Service"] = melted["Service"].astype(str).str.strip().str.title()
    melted["Vendor"] = melted["Vendor"].astype(str).str.strip().str.title()

    return melted.reset_index(drop=True)

rate_df = load_rate_table("haulier prices.xlsx")

# Precompute unique postcode areas for autocomplete suggestions
unique_areas = sorted(rate_df["PostcodeArea"].unique())

# ─────────────────────────────────────────
# (5) USER INPUTS
# ─────────────────────────────────────────
st.header("1. Input Parameters")

col_a, col_b, col_c, col_d, col_e = st.columns([1, 1, 1, 1, 1], gap="large")

with col_a:
    input_postcode = st.text_input(
        "Postcode (area) (e.g. BB, LA, etc.)",
        placeholder="Type postcode area…"
    ).strip().upper()

    # Show live suggestions underneath
    if input_postcode:
        matches = [area for area in unique_areas if area.startswith(input_postcode)]
        if matches:
            st.markdown("**Did you mean:** " + ", ".join(matches[:10]) + ("…" if len(matches) > 10 else ""))

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
        max_value=26,  # Excel covers up to 26 pallets
        value=1,
        step=1
    )

with col_d:
    joda_surcharge_pct = st.number_input(
        "Joda Fuel Surcharge (%)",
        min_value=0.00,
        max_value=100.00,
        value=0.00,
        step=0.1,
        format="%.2f",
        help="Copy “CURRENT SURCHARGE %” (e.g. 2.74) from Joda’s site."
    )

with col_e:
    mcd_surcharge_pct = st.number_input(
        "McDowells Fuel Surcharge (%)",
        min_value=0.00,
        max_value=100.00,
        value=0.00,
        step=0.1,
        format="%.2f",
        help="Enter McDowells’ surcharge manually."
    )

st.markdown("---")

# Ensure user entered a postcode area
if not input_postcode:
    st.info("🔍 Please enter a postcode area to continue.")
    st.stop()

# Extract “area” (just use input_postcode as is, since we’re matching areas)
postcode_area = input_postcode

# ─────────────────────────────────────────
# (6) TOGGLE INPUTS (DELIVERY OPTIONS + DUAL)
# ─────────────────────────────────────────
st.subheader("2. Optional Extras")
col1, col2, col3 = st.columns(3, gap="large")

with col1:
    ampm_toggle = st.checkbox("AM/PM Delivery")
with col2:
    timed_toggle = st.checkbox("Timed Delivery")
with col3:
    dual_toggle = st.checkbox("Dual Collection")

# If Dual Collection is enabled, show Pallet Split inputs
split1 = split2 = None
if dual_toggle:
    st.markdown("**Specify how to split pallets into two shipments**")
    sp1, sp2 = st.columns(2, gap="large")
    with sp1:
        split1 = st.number_input(
            "First Pallet Group",
            min_value=1,
            max_value=num_pallets - 1,
            value=1,
            step=1
        )
    with sp2:
        split2 = st.number_input(
            "Second Pallet Group",
            min_value=1,
            max_value=num_pallets - 1,
            value=num_pallets - 1,
            step=1
        )
    # Validate that split1 + split2 == num_pallets
    if split1 + split2 != num_pallets:
        st.error("⚠️ Pallet Split values must add up to total pallets.")
        st.stop()

# ─────────────────────────────────────────
# (7) LOOK UP BASE RATES FOR EACH HAULIER
# ─────────────────────────────────────────
def get_base_rate(df, area, service, vendor, pallets):
    subset = df[
        (df["PostcodeArea"] == area) &
        (df["Service"] == service) &
        (df["Vendor"] == vendor) &
        (df["Pallets"] == pallets)
    ]
    if subset.empty:
        return None
    return float(subset["BaseRate"].iloc[0])

# Joda calculation
if dual_toggle:
    # Two shipments for Joda
    joda_base1 = get_base_rate(rate_df, postcode_area, service_option, "Joda", split1)
    joda_base2 = get_base_rate(rate_df, postcode_area, service_option, "Joda", split2)
    if joda_base1 is None:
        st.error(f"❌ No Joda rate for {split1} pallet(s) (first group).")
        st.stop()
    if joda_base2 is None:
        st.error(f"❌ No Joda rate for {split2} pallet(s) (second group).")
        st.stop()

    # Delivery surcharge per shipment
    joda_delivery_per = 0
    if ampm_toggle:
        joda_delivery_per += 7
    if timed_toggle:
        joda_delivery_per += 19

    # Final for each group
    joda_group1 = joda_base1 * (1 + joda_surcharge_pct / 100.0) + joda_delivery_per
    joda_group2 = joda_base2 * (1 + joda_surcharge_pct / 100.0) + joda_delivery_per

    joda_final = joda_group1 + joda_group2
    joda_base = joda_base1 + joda_base2
    # Total Joda delivery charge = two shipments
    joda_delivery_charge = joda_delivery_per * 2

else:
    # Single‐shipment for Joda
    joda_base = get_base_rate(rate_df, postcode_area, service_option, "Joda", num_pallets)
    if joda_base is None:
        st.error(
            f"❌ No Joda rate for area '{postcode_area}', service '{service_option}', "
            f"{num_pallets} pallet(s)."
        )
        st.stop()
    joda_delivery_charge = 0
    if ampm_toggle:
        joda_delivery_charge += 7
    if timed_toggle:
        joda_delivery_charge += 19

    joda_final = joda_base * (1 + joda_surcharge_pct / 100.0) + joda_delivery_charge

# McDowells calculation
mcd_base = get_base_rate(rate_df, postcode_area, service_option, "Mcdowells", num_pallets)
if mcd_base is None:
    st.error(
        f"❌ No McDowells rate for area '{postcode_area}', service '{service_option}', "
        f"{num_pallets} pallet(s)."
    )
    st.stop()

mcd_delivery_charge = 0
if ampm_toggle:
    mcd_delivery_charge += 10
if timed_toggle:
    mcd_delivery_charge += 19

mcd_final = mcd_base * (1 + mcd_surcharge_pct / 100.0) + mcd_delivery_charge

# ─────────────────────────────────────────
# (8) LOOK UP “ONE PALLET FEWER” & “ONE PALLET MORE” (for display)
# ─────────────────────────────────────────
def lookup_adjacent_rate(df, area, service, vendor, pallets, surcharge_pct, delivery_charge):
    """
    Returns dict {"lower": (count, rate), "higher": (count, rate)}.
    """
    out = {"lower": None, "higher": None}
    if pallets > 1:
        base_lower = get_base_rate(df, area, service, vendor, pallets - 1)
        if base_lower is not None:
            rate_lower = base_lower * (1 + surcharge_pct / 100.0) + delivery_charge
            out["lower"] = ((pallets - 1), rate_lower)
    base_higher = get_base_rate(df, area, service, vendor, pallets + 1)
    if base_higher is not None:
        rate_higher = base_higher * (1 + surcharge_pct / 100.0) + delivery_charge
        out["higher"] = ((pallets + 1), rate_higher)
    return out

# For display, use single‐shipment delivery_charge for adjacent rows
joda_adj = lookup_adjacent_rate(
    rate_df, postcode_area, service_option, "Joda",
    num_pallets, joda_surcharge_pct,
    joda_delivery_charge if not dual_toggle else ((7 if ampm_toggle else 0) + (19 if timed_toggle else 0))
)
mcd_adj = lookup_adjacent_rate(
    rate_df, postcode_area, service_option, "Mcdowells",
    num_pallets, mcd_surcharge_pct, mcd_delivery_charge
)

# ─────────────────────────────────────────
# (9) DISPLAY RESULTS
# ─────────────────────────────────────────
st.header("3. Calculated Rates")

summary_data = [
    {
        "Haulier": "Joda",
        "Base Rate": f"£{joda_base:,.2f}",
        "Fuel Surcharge (%)": f"{joda_surcharge_pct:.2f}%",
        "Delivery Charge": f"£{joda_delivery_charge:,.2f}",
        "Final Rate": f"£{joda_final:,.2f}"
    },
    {
        "Haulier": "McDowells",
        "Base Rate": f"£{mcd_base:,.2f}",
        "Fuel Surcharge (%)": f"{mcd_surcharge_pct:.2f}%",
        "Delivery Charge": f"£{mcd_delivery_charge:,.2f}",
        "Final Rate": f"£{mcd_final:,.2f}"
    }
]
summary_df = pd.DataFrame(summary_data).set_index("Haulier")

def highlight_cheapest(row):
    val = float(row["Final Rate"].strip("£").replace(",", ""))
    cheapest = min(joda_final, mcd_final)
    if math.isclose(val, cheapest, rel_tol=1e-9):
        return ["background-color: #b3e6b3"] * len(row)
    return [""] * len(row)

st.table(summary_df.style.apply(highlight_cheapest, axis=1))
st.markdown(
    "<i style='color:gray;'>* Row highlighted in green indicates the cheapest option *</i>",
    unsafe_allow_html=True
)

# ─────────────────────────────────────────
# (10) SHOW “ONE PALLET FEWER” & “ONE PALLET MORE” (GREYED OUT)
# ─────────────────────────────────────────
st.subheader("4. One Pallet Fewer / One Pallet More (Greyed Out)")

adj_cols = st.columns(2)

with adj_cols[0]:
    st.markdown("<b>Joda Rates</b>", unsafe_allow_html=True)
    lines = []
    if joda_adj["lower"] is not None:
        lp, lr = joda_adj["lower"]
        lines.append(f"&nbsp;&nbsp;• {lp} pallet(s): £{lr:,.2f}")
    else:
        lines.append("&nbsp;&nbsp;• <span style='color:gray;'>N/A for fewer pallets</span>")

    if joda_adj["higher"] is not None:
        hp, hr = joda_adj["higher"]
        lines.append(f"&nbsp;&nbsp;• {hp} pallet(s): £{hr:,.2f}")
    else:
        lines.append("&nbsp;&nbsp;• <span style='color:gray;'>N/A for more pallets</span>")

    st.markdown("<br>".join(lines), unsafe_allow_html=True)

with adj_cols[1]:
    st.markdown("<b>McDowells Rates</b>", unsafe_allow_html=True)
    lines = []
    if mcd_adj["lower"] is not None:
        lp, lr = mcd_adj["lower"]
        lines.append(f"&nbsp;&nbsp;• {lp} pallet(s): £{lr:,.2f}")
    else:
        lines.append("&nbsp;&nbsp;• <span style='color:gray;'>N/A for fewer pallets</span>")

    if mcd_adj["higher"] is not None:
        hp, hr = mcd_adj["higher"]
        lines.append(f"&nbsp;&nbsp;• {hp} pallet(s): £{hr:,.2f}")
    else:
        lines.append("&nbsp;&nbsp;• <span style='color:gray;'>N/A for more pallets</span>")

    st.markdown("<br>".join(lines), unsafe_allow_html=True)

# ─────────────────────────────────────────
# (11) FOOTER / NOTES
# ─────────────────────────────────────────
st.markdown("---")
st.markdown(
    """
    <small>
    • Joda’s surcharge must be copied from Joda’s website and typed above.  
    • McDowells’ surcharge is always entered manually.  
    • Delivery charges: Joda – AM/PM £7, Timed £19;  
      McDowells – AM/PM £10, Timed £19.  
    • Dual Collection splits Joda into two shipments.  
    • The cheapest final rate is highlighted in green.  
    </small>
    """,
    unsafe_allow_html=True
)
