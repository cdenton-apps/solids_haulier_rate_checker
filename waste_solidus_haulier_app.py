# ======================================
# File: waste_solidus_haulier_app.py
# ======================================

import streamlit as st
import pandas as pd
import io
import math
import re
import requests
from bs4 import BeautifulSoup
from PIL import Image

# ─────────────────────────────────────────
# (1) STREAMLIT PAGE CONFIGURATION FIRST
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
# (3) DISPLAY SOLIDUS LOGO + HEADER SIDE-BY-SIDE
# ─────────────────────────────────────────
col_logo, col_text = st.columns([1, 3], gap="medium")
with col_logo:
    logo_path = "assets/solidus_logo.png"
    try:
        logo_img = Image.open(logo_path)
        st.image(logo_img, width=150)
    except Exception:
        st.warning(
            f"⚠️ Could not load logo at '{logo_path}'.\n"
            "Please confirm that the file exists and is a valid PNG."
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
        specify the number of pallets, and apply fuel surcharges (Joda fetched automatically, McDowells entered manually).  
        The app will display the final adjusted rate for both Joda and McDowells, highlight the cheapest option,  
        and also show the price for one fewer and one more pallet (greyed out).
        """,
        unsafe_allow_html=True
    )

# ─────────────────────────────────────────
# (4) FETCH JODA FUEL SURCHARGE VIA REGEX & USER-AGENT
# ─────────────────────────────────────────
@st.cache_data(show_spinner=False)
def fetch_joda_surcharge() -> float:
    """
    Fetch the current Joda Freight surcharge percentage from their website by
    scraping the visible text. We look for the phrase “CURRENT SURCHARGE” followed
    by any digits (with optional decimal) ending in “%”. If found, return that
    number as a float (e.g. 2.74). If anything goes wrong, return None.
    """
    url = "https://www.jodafreight.com/fuel-surcharge/"
    headers = {
        "User-Agent": (
            "Mozilla/5.0 "
            "(Windows NT 10.0; Win64; x64) "
            "AppleWebKit/537.36 (KHTML, like Gecko) "
            "Chrome/115.0.0.0 Safari/537.36"
        )
    }
    try:
        resp = requests.get(url, headers=headers, timeout=10)
        resp.raise_for_status()

        soup = BeautifulSoup(resp.text, "html.parser")
        page_text = soup.get_text(separator=" ", strip=True)

        # More forgiving regex: look for "CURRENT SURCHARGE" then capture "digits(.digits)?%"
        # for example: "CURRENT SURCHARGE % 2.74%"
        match = re.search(
            r"CURRENT\s+SURCHARGE[^0-9]*([0-9]+(?:\.[0-9]+)?)\s*%",
            page_text,
            re.IGNORECASE
        )
        if match:
            return float(match.group(1))
        else:
            return None
    except Exception:
        return None

# Fetch Joda surcharge once at app startup
joda_surcharge_pct = fetch_joda_surcharge()
if joda_surcharge_pct is None:
    st.warning(
        "⚠️ Could not fetch Joda’s current fuel surcharge automatically.\n"
        "We will proceed with 0% surcharge for Joda."
    )
    joda_surcharge_pct = 0.0

# ─────────────────────────────────────────
# (5) LOAD & TRANSFORM THE BUILT-IN EXCEL DATA
# ─────────────────────────────────────────
@st.cache_data
def load_rate_table(excel_path: str) -> pd.DataFrame:
    """
    Reads the 'haulier prices.xlsx' file, reshapes it, and returns a DataFrame with columns:
    ['PostcodeArea', 'Service', 'Vendor', 'Pallets', 'BaseRate'].
    """
    # Read starting at row 2 (header=1)
    raw = pd.read_excel(excel_path, header=1)
    # Rename the first three unnamed columns
    raw = raw.rename(columns={
        raw.columns[0]: "PostcodeArea",
        raw.columns[1]: "Service",
        raw.columns[2]: "Vendor"
    })
    # Forward-fill PostcodeArea and Service down the column
    raw["PostcodeArea"] = raw["PostcodeArea"].ffill()
    raw["Service"] = raw["Service"].ffill()

    # Drop any rows where Vendor == 'Vendor' (the header junk row)
    raw = raw[raw["Vendor"] != "Vendor"].copy()

    # Identify which columns are pallet-count columns (numeric column names)
    pallet_cols = [
        col for col in raw.columns
        if isinstance(col, (int, float)) or (isinstance(col, str) and col.isdigit())
    ]

    # Melt from wide→long form
    melted = raw.melt(
        id_vars=["PostcodeArea", "Service", "Vendor"],
        value_vars=pallet_cols,
        var_name="Pallets",
        value_name="BaseRate"
    )

    # Clean up: Pallets as integer, drop rows where BaseRate is NaN
    melted["Pallets"] = melted["Pallets"].astype(int)
    melted = melted.dropna(subset=["BaseRate"]).copy()

    melted["PostcodeArea"] = melted["PostcodeArea"].astype(str).str.strip().str.upper()
    melted["Service"] = melted["Service"].astype(str).str.strip().str.title()
    melted["Vendor"] = melted["Vendor"].astype(str).str.strip().str.title()

    return melted.reset_index(drop=True)

rate_df = load_rate_table("haulier prices.xlsx")

# ─────────────────────────────────────────
# (6) USER INPUTS
# ─────────────────────────────────────────
st.header("1. Input Parameters")

col_a, col_b, col_c, col_d = st.columns(4, gap="large")
with col_a:
    input_postcode = st.text_input(
        "Postcode (e.g. BB10 1AB)",
        placeholder="Enter 1 or 2 letter area + rest"
    ).strip().upper()

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
        max_value=26,  # Based on columns up to '26' in the file
        value=1,
        step=1
    )

with col_d:
    mcd_surcharge_pct = st.number_input(
        "McDowells Fuel Surcharge (%)",
        min_value=0.0,
        max_value=100.0,
        value=0.0,
        step=0.5,
        format="%.1f"
    )

st.markdown("---")

# Check that postcode_area is at least 1 character
if not input_postcode:
    st.info("🔍 Please enter a postcode to continue.")
    st.stop()

# Extract the postcode area (first 2 letters, e.g. 'BB' from 'BB10 1AB')
postcode_area = input_postcode.split()[0][:2]

# ─────────────────────────────────────────
# (7) LOOK UP BASE RATES FOR BOTH HAULIERS
# ─────────────────────────────────────────
def get_base_rate(
    df: pd.DataFrame,
    area: str,
    service: str,
    vendor: str,
    pallets: int
) -> float:
    """
    Return the base rate for the given vendor, postcode area, service, and pallet count.
    If no matching row exists, return None.
    """
    subset = df[
        (df["PostcodeArea"] == area) &
        (df["Service"] == service) &
        (df["Vendor"] == vendor) &
        (df["Pallets"] == pallets)
    ]
    if subset.empty:
        return None
    return float(subset["BaseRate"].iloc[0])

# Look up Joda base rate
joda_base = get_base_rate(rate_df, postcode_area, service_option, "Joda", num_pallets)
if joda_base is None:
    st.error(
        f"❌ No Joda rate found for area '{postcode_area}', service '{service_option}', "
        f"and {num_pallets} pallet(s)."
    )
    st.stop()

# Look up McDowells base rate
mcd_base = get_base_rate(rate_df, postcode_area, service_option, "Mcdowells", num_pallets)
if mcd_base is None:
    st.error(
        f"❌ No McDowells rate found for area '{postcode_area}', service '{service_option}', "
        f"and {num_pallets} pallet(s)."
    )
    st.stop()

# ─────────────────────────────────────────
# (8) CALCULATE FINAL RATES (WITH SURCHARGES)
# ─────────────────────────────────────────
joda_final = joda_base * (1 + joda_surcharge_pct / 100)
mcd_final = mcd_base * (1 + mcd_surcharge_pct / 100)

# ─────────────────────────────────────────
# (9) LOOK UP “ONE PALLET FEWER” AND “ONE PALLET MORE”
# ─────────────────────────────────────────
def lookup_adjacent_rate(
    df: pd.DataFrame,
    area: str,
    service: str,
    vendor: str,
    pallets: int
):
    """
    For a given vendor/area/service and pallet count:
    - Find base rates for (pallets - 1) and (pallets + 1), if they exist.
    - Apply the same surcharge that we’ll apply to the main count.
    Return a dict: { 'lower': (pallet_count, rate_with_surcharge), 'higher': (...) }
    If the lower/higher pallet count does not exist in the table, return None for that key.
    """
    result = {"lower": None, "higher": None}

    # Lower = pallets - 1 (only if >= 1)
    if pallets > 1:
        lower_base = get_base_rate(df, area, service, vendor, pallets - 1)
        if lower_base is not None:
            if vendor.lower() == "joda":
                lower_rate = lower_base * (1 + joda_surcharge_pct / 100)
            else:
                lower_rate = lower_base * (1 + mcd_surcharge_pct / 100)
            result["lower"] = ((pallets - 1), lower_rate)

    # Higher = pallets + 1 (only if exists)
    higher_base = get_base_rate(df, area, service, vendor, pallets + 1)
    if higher_base is not None:
        if vendor.lower() == "joda":
            higher_rate = higher_base * (1 + joda_surcharge_pct / 100)
        else:
            higher_rate = higher_base * (1 + mcd_surcharge_pct / 100)
        result["higher"] = ((pallets + 1), higher_rate)

    return result

joda_adj = lookup_adjacent_rate(rate_df, postcode_area, service_option, "Joda", num_pallets)
mcd_adj = lookup_adjacent_rate(rate_df, postcode_area, service_option, "Mcdowells", num_pallets)

# ─────────────────────────────────────────
# (10) DISPLAY RESULTS
# ─────────────────────────────────────────
st.header("2. Calculated Rates")

# Prepare a small DataFrame to show in-app summary
summary_data = [
    {
        "Haulier": "Joda",
        "Base Rate": f"£{joda_base:,.2f}",
        "Fuel Surcharge (%)": f"{joda_surcharge_pct:.1f}%",
        "Final Rate": f"£{joda_final:,.2f}"
    },
    {
        "Haulier": "McDowells",
        "Base Rate": f"£{mcd_base:,.2f}",
        "Fuel Surcharge (%)": f"{mcd_surcharge_pct:.1f}%",
        "Final Rate": f"£{mcd_final:,.2f}"
    }
]
summary_df = pd.DataFrame(summary_data).set_index("Haulier")

# Highlight the cheapest final rate row in green
def highlight_cheapest(row):
    val = float(row["Final Rate"].strip("£").replace(",", ""))
    cheapest = min(joda_final, mcd_final)
    if math.isclose(val, cheapest, rel_tol=1e-9):
        return ["background-color: #b3e6b3"] * len(row)
    else:
        return [""] * len(row)

st.table(summary_df.style.apply(highlight_cheapest, axis=1))
st.markdown(
    "<i style='color:gray;'>* Row highlighted in green indicates the cheapest option *</i>",
    unsafe_allow_html=True
)

# ─────────────────────────────────────────
# (11) SHOW ADJACENT PALLET RATES (GREYED OUT)
# ─────────────────────────────────────────
st.subheader("3. One Pallet Fewer / One Pallet More (Greyed Out)")

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
# (12) FOOTER / NOTES
# ─────────────────────────────────────────
st.markdown("---")
st.markdown(
    """
    <small>
    • Joda Freight surcharge is fetched automatically from the Joda website (using a regex).  
    • McDowells surcharge must be entered manually.  
    • If a given pallet count is not offered by a vendor, the app shows “N/A.”  
    • The cheapest final rate is highlighted in green above.  
    </small>
    """,
    unsafe_allow_html=True
)
