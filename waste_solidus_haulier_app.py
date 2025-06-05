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
      /* Hide top-right menu (hamburger) */
      #MainMenu {visibility: hidden;}
      /* Hide “Made with Streamlit” footer */
      footer {visibility: hidden;}
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
        specify the number of pallets, and apply fuel surcharges (Joda fetched automatically if possible, McDowells entered manually).  
        The app will display the final adjusted rate for both Joda and McDowells, highlight the cheapest option,  
        and also show the price for one fewer and one more pallet (greyed out).
        """,
        unsafe_allow_html=True
    )

# ─────────────────────────────────────────
# (4) ATTEMPT TO FETCH JODA FUEL SURCHARGE VIA REGEX & USER-AGENT
# ─────────────────────────────────────────
@st.cache_data(show_spinner=False)
def fetch_joda_surcharge() -> float:
    """
    Try to fetch the “CURRENT SURCHARGE %” number from Joda’s website.
    We grab the page text and use a regex to find digits (possibly with decimals)
    before a “%” sign, immediately following the phrase “CURRENT SURCHARGE”.
    If successful, return a float like 2.74. Otherwise return None.
    """
    url = "https://www.jodafreight.com/fuel-surcharge/"
    headers = {
        "User-Agent": (
            "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
            "AppleWebKit/537.36 (KHTML, like Gecko) "
            "Chrome/115.0.0.0 Safari/537.36"
        )
    }
    try:
        resp = requests.get(url, headers=headers, timeout=10)
        resp.raise_for_status()
        soup = BeautifulSoup(resp.text, "html.parser")
        page_text = soup.get_text(separator=" ", strip=True)

        # Regex: look for “CURRENT SURCHARGE” then any non-digit chars, then capture digits(.digits)? before “%”
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

# Attempt auto-fetch once at app startup
joda_auto_pct = fetch_joda_surcharge()

# ─────────────────────────────────────────
# (5) LOAD & TRANSFORM THE BUILT-IN EXCEL DATA
# ─────────────────────────────────────────
@st.cache_data
def load_rate_table(excel_path: str) -> pd.DataFrame:
    """
    Read 'haulier prices.xlsx' starting at row 2 (header=1), then forward-fill PostcodeArea & Service.
    Melt the numeric pallet columns into long form so we get columns:
      [PostcodeArea, Service, Vendor, Pallets, BaseRate].
    """
    raw = pd.read_excel(excel_path, header=1)
    raw = raw.rename(columns={
        raw.columns[0]: "PostcodeArea",
        raw.columns[1]: "Service",
        raw.columns[2]: "Vendor"
    })
    raw["PostcodeArea"] = raw["PostcodeArea"].ffill()
    raw["Service"] = raw["Service"].ffill()

    # Drop any “Vendor” rows from that extra header line
    raw = raw[raw["Vendor"] != "Vendor"].copy()

    # Identify numeric pallet columns
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

    melted["PostcodeArea"] = melted["PostcodeArea"].astype(str).str.strip().str.upper()
    melted["Service"] = melted["Service"].astype(str).str.strip().str.title()
    melted["Vendor"] = melted["Vendor"].astype(str).str.strip().str.title()

    return melted.reset_index(drop=True)

rate_df = load_rate_table("haulier prices.xlsx")

# ─────────────────────────────────────────
# (6) USER INPUTS
# ─────────────────────────────────────────
st.header("1. Input Parameters")

col_a, col_b, col_c, col_d, col_e = st.columns([1, 1, 1, 1, 1], gap="large")

with col_a:
    input_postcode = st.text_input(
        "Postcode (e.g. BB10 1AB)",
        placeholder="Enter at least 1 or 2 letters"
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
        max_value=26,  # Our Excel has columns up to “26”
        value=1,
        step=1
    )

with col_d:
    # If successfully auto-fetched, prefill; else allow manual
    if joda_auto_pct is not None:
        joda_surcharge_pct = st.number_input(
            "Joda Fuel Surcharge (%)",
            min_value=0.0,
            max_value=100.0,
            value=round(joda_auto_pct, 2),
            step=0.1,
            format="%.2f",
            help="Prefilled from Joda’s site, but you may override."
        )
    else:
        joda_surcharge_pct = st.number_input(
            "Joda Fuel Surcharge (%)",
            min_value=0.0,
            max_value=100.0,
            value=0.00,
            step=0.1,
            format="%.2f",
            help="Could not auto-fetch; please type Joda’s surcharge manually."
        )
        st.warning(
            "⚠️ Auto-fetch failed. Enter Joda’s surcharge manually (e.g. 2.74)."
        )

with col_e:
    mcd_surcharge_pct = st.number_input(
        "McDowells Fuel Surcharge (%)",
        min_value=0.0,
        max_value=100.0,
        value=0.00,
        step=0.1,
        format="%.2f"
    )

st.markdown("---")

# Ensure user entered a postcode
if not input_postcode:
    st.info("🔍 Please enter a postcode to continue.")
    st.stop()

# Extract “area” = the first 2 letters (e.g. “BB” from “BB10 1AB”)
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
    Return the base rate for vendor/area/service/pallets.  If no row, return None.
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

joda_base = get_base_rate(rate_df, postcode_area, service_option, "Joda", num_pallets)
if joda_base is None:
    st.error(
        f"❌ No Joda rate found for area '{postcode_area}', service '{service_option}', "
        f"and {num_pallets} pallet(s)."
    )
    st.stop()

mcd_base = get_base_rate(rate_df, postcode_area, service_option, "Mcdowells", num_pallets)
if mcd_base is None:
    st.error(
        f"❌ No McDowells rate found for area '{postcode_area}', service '{service_option}', "
        f"and {num_pallets} pallet(s)."
    )
    st.stop()

# ─────────────────────────────────────────
# (8) CALCULATE FINAL RATES (APPLY SURCHARGES)
# ─────────────────────────────────────────
joda_final = joda_base * (1 + joda_surcharge_pct / 100.0)
mcd_final = mcd_base * (1 + mcd_surcharge_pct / 100.0)

# ─────────────────────────────────────────
# (9) LOOK UP “ONE PALLET FEWER” & “ONE PALLET MORE”
# ─────────────────────────────────────────
def lookup_adjacent_rate(
    df: pd.DataFrame,
    area: str,
    service: str,
    vendor: str,
    pallets: int
):
    """
    For vendor/area/service and pallet count, find rates for (pallets-1) & (pallets+1)
    (if they exist). Apply the same surcharge. Return a dict:
      {'lower': (pallet_count, rate_with_surcharge), 'higher': (...) }
    If lower/higher doesn’t exist, that key stays None.
    """
    result = {"lower": None, "higher": None}

    # Lower = pallets - 1 if >= 1
    if pallets > 1:
        lower_base = get_base_rate(df, area, service, vendor, pallets - 1)
        if lower_base is not None:
            if vendor.lower() == "joda":
                lr = lower_base * (1 + joda_surcharge_pct / 100.0)
            else:
                lr = lower_base * (1 + mcd_surcharge_pct / 100.0)
            result["lower"] = ((pallets - 1), lr)

    # Higher = pallets + 1 if available
    higher_base = get_base_rate(df, area, service, vendor, pallets + 1)
    if higher_base is not None:
        if vendor.lower() == "joda":
            hr = higher_base * (1 + joda_surcharge_pct / 100.0)
        else:
            hr = higher_base * (1 + mcd_surcharge_pct / 100.0)
        result["higher"] = ((pallets + 1), hr)

    return result

joda_adj = lookup_adjacent_rate(rate_df, postcode_area, service_option, "Joda", num_pallets)
mcd_adj = lookup_adjacent_rate(rate_df, postcode_area, service_option, "Mcdowells", num_pallets)

# ─────────────────────────────────────────
# (10) DISPLAY RESULTS
# ─────────────────────────────────────────
st.header("2. Calculated Rates")

summary_data = [
    {
        "Haulier": "Joda",
        "Base Rate": f"£{joda_base:,.2f}",
        "Fuel Surcharge (%)": f"{joda_surcharge_pct:.2f}%",
        "Final Rate": f"£{joda_final:,.2f}"
    },
    {
        "Haulier": "McDowells",
        "Base Rate": f"£{mcd_base:,.2f}",
        "Fuel Surcharge (%)": f"{mcd_surcharge_pct:.2f}%",
        "Final Rate": f"£{mcd_final:,.2f}"
    }
]
summary_df = pd.DataFrame(summary_data).set_index("Haulier")

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
# (11) SHOW “ONE PALLET FEWER” & “ONE PALLET MORE” (GREYED OUT)
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
    • If Joda’s surcharge auto-fetch succeeded, its value was pre-filled above. Otherwise, you entered it manually.  
    • McDowells surcharge must always be entered manually.  
    • If a given pallet count is not offered by either vendor, that vendor shows “N/A.”  
    • The cheapest final rate is highlighted in green.  
    </small>
    """,
    unsafe_allow_html=True
)
