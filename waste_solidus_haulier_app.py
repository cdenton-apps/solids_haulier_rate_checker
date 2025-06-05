# ======================================
# File: waste_solidus_haulier_app.py
# ======================================

import streamlit as st
import pandas as pd
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
        specify the number of pallets, and apply fuel surcharges:
        - We first attempt to fetch Joda’s “CURRENT SURCHARGE %” from the first row of their table.  
        - If that fails (common on Streamlit Cloud because the table is rendered via JavaScript),  
          we show a warning and let you type Joda’s surcharge manually.  
        - McDowells’ surcharge is always entered manually.  

        The app will then:
        1. Calculate the final adjusted rate for both Joda and McDowells.  
        2. Highlight the cheapest option in green.  
        3. Show the price for one fewer and one more pallet (greyed out if unavailable).
        """,
        unsafe_allow_html=True
    )

# ─────────────────────────────────────────
# (4) TRY TO SCRAPE JODA’S “CURRENT SURCHARGE %” FROM THE FIRST ROW OF THEIR TABLE
# ─────────────────────────────────────────
@st.cache_data(show_spinner=False)
def try_fetch_joda_surcharge_from_table() -> float:
    """
    Attempt to GET Joda’s fuel-surcharge page and parse the first row of the
    first <table> to find the “CURRENT SURCHARGE %” column. Return that number
    as a float (e.g. 2.74). If anything fails, return None.
    """
    try:
        resp = requests.get("https://www.jodafreight.com/fuel-surcharge/", timeout=8)
        resp.raise_for_status()
        soup = BeautifulSoup(resp.text, "html.parser")

        # Find the first <table> element on the page
        tbl = soup.find("table")
        if not tbl:
            return None

        # Collect all <tr> rows
        rows = tbl.find_all("tr")
        if len(rows) < 2:
            return None

        # The first <tr> is typically the header row (with <th> or <td> cells)
        header_cells = rows[0].find_all(["th", "td"])
        # Find the index of any cell whose text contains “CURRENT SURCHARGE” (case‐insensitive)
        surcharge_col_index = None
        for idx, cell in enumerate(header_cells):
            header_text = cell.get_text(strip=True).upper()
            if "CURRENT SURCHARGE" in header_text:
                surcharge_col_index = idx
                break

        if surcharge_col_index is None:
            return None

        # The second row (rows[1]) is the first data row
        first_data_cells = rows[1].find_all("td")
        if surcharge_col_index >= len(first_data_cells):
            return None

        # Extract the cell text (e.g. “2.74%”)
        raw_text = first_data_cells[surcharge_col_index].get_text(strip=True)
        # Use a simple regex to grab numeric portion before “%”
        match = re.search(r"([0-9]+(?:\.[0-9]+)?)\s*%", raw_text)
        if not match:
            return None

        return float(match.group(1))

    except Exception:
        return None


# Attempt to auto-fetch Joda’s surcharge by table parsing
joda_auto_pct = try_fetch_joda_surcharge_from_table()

# ─────────────────────────────────────────
# (5) LOAD & TRANSFORM THE BUILT-IN EXCEL DATA
# ─────────────────────────────────────────
@st.cache_data
def load_rate_table(excel_path: str) -> pd.DataFrame:
    """
    Read 'haulier prices.xlsx' starting at row 2 (header=1).  Forward-fill
    PostcodeArea & Service, then melt numeric pallet columns into:
        PostcodeArea | Service | Vendor | Pallets | BaseRate
    """
    raw = pd.read_excel(excel_path, header=1)
    raw = raw.rename(columns={
        raw.columns[0]: "PostcodeArea",
        raw.columns[1]: "Service",
        raw.columns[2]: "Vendor"
    })
    raw["PostcodeArea"] = raw["PostcodeArea"].ffill()
    raw["Service"] = raw["Service"].ffill()

    # Drop any “Vendor” rows from the redundant header line
    raw = raw[raw["Vendor"] != "Vendor"].copy()

    # Identify numeric columns for pallet counts
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
        max_value=26,  # Your Excel covers up to 26 pallets
        value=1,
        step=1
    )

with col_d:
    if joda_auto_pct is not None:
        joda_surcharge_pct = st.number_input(
            "Joda Fuel Surcharge (%)",
            min_value=0.0,
            max_value=100.0,
            value=round(joda_auto_pct, 2),
            step=0.1,
            format="%.2f",
            help="Prefilled from the first row of Joda’s table; override if needed."
        )
    else:
        joda_surcharge_pct = st.number_input(
            "Joda Fuel Surcharge (%)",
            min_value=0.0,
            max_value=100.0,
            value=0.00,
            step=0.1,
            format="%.2f",
            help="Auto-fetch failed. Please enter Joda’s surcharge manually."
        )
        st.warning(
            "⚠️ Could not auto-fetch Joda’s surcharge. Please type it manually (e.g. 2.74)."
        )

with col_e:
    mcd_surcharge_pct = st.number_input(
        "McDowells Fuel Surcharge (%)",
        min_value=0.0,
        max_value=100.0,
        value=0.00,
        step=0.1,
        format="%.2f",
        help="Enter McDowells’ surcharge manually."
    )

st.markdown("---")

# Ensure user entered something for postcode
if not input_postcode:
    st.info("🔍 Please enter a postcode to continue.")
    st.stop()

# Extract only the “area” (first two letters) from postcode: “BB10 1AB” → “BB”
postcode_area = input_postcode.split()[0][:2]

# ─────────────────────────────────────────
# (7) LOOK UP BASE RATES FOR BOTH HAULIERS
# ─────────────────────────────────────────
def get_base_rate(df, area, service, vendor, pallets):
    """
    Filter rate_df for matching PostcodeArea, Service, Vendor, Pallets.
    Return BaseRate as float, or None if no match.
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
def lookup_adjacent_rate(df, area, service, vendor, pallets):
    """
    For a given vendor/area/service/pallets:
    - If pallets > 1, look up (pallets - 1) base and apply same surcharge.
    - Look up (pallets + 1) base and apply same surcharge.
    Return {"lower": (count, rate), "higher": (count, rate)} or None if missing.
    """
    result = {"lower": None, "higher": None}

    # Lower = pallets - 1 (only if >= 1)
    if pallets > 1:
        lower_base = get_base_rate(df, area, service, vendor, pallets - 1)
        if lower_base is not None:
            if vendor.lower() == "joda":
                lr = lower_base * (1 + joda_surcharge_pct / 100.0)
            else:
                lr = lower_base * (1 + mcd_surcharge_pct / 100.0)
            result["lower"] = ((pallets - 1), lr)

    # Higher = pallets + 1 (if exists)
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
    • If we successfully read Joda’s first table row, you saw its surcharge above. Otherwise, you typed it manually.  
    • McDowells surcharge is always entered manually.  
    • If a vendor does not offer that pallet count, “N/A” is shown.  
    • The cheapest final rate is highlighted in green.  
    </small>
    """,
    unsafe_allow_html=True
)
