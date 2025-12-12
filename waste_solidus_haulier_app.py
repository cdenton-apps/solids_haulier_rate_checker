import streamlit as st
import pandas as pd
from PIL import Image
import json
from datetime import date
import os
import math

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
        **Joda’s surcharge (%)** is stored persistently and must be updated once weekly.  
          You can look this up at https://www.jodafreight.com/fuel-surcharge/  
          On Wednesdays it resets to 0 automatically.
          **McDowells’ surcharge (%)** is always entered manually each session.  
          You may optionally add AM/PM Delivery, Tail Lift or Timed Delivery,  
          or perform a Dual Collection (For collections from both Unit 4 and ESL):
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

    # Wednesday reset
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
    # Set sheet_name="Rates" here if needed for your workbook
    raw = pd.read_excel(excel_path, header=1)
    raw = raw.rename(columns={
        raw.columns[0]: "PostcodeArea",
        raw.columns[1]: "Service",
        raw.columns[2]: "Vendor"
    })

    # forward-fill merged headers
    raw["PostcodeArea"] = raw["PostcodeArea"].ffill()
    raw["Service"]      = raw["Service"].ffill()
    raw["Vendor"]       = raw["Vendor"].ffill()
    # drop accidental header repeats
    raw = raw[raw["Vendor"] != "Vendor"].copy()

    # pallet columns must be purely numeric (1,2,3,...)
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

# Joda (tail lift = £0)
joda_base = None
joda_final = None
joda_charge_fixed = (7 if ampm_toggle else 0) + (19 if timed_toggle else 0)

# >>> Joda surcharge rule helper: apply surcharge only if pallets for that leg >= 7
def apply_joda(price_base: float, pallets_count: int, pct: float, fixed: float) -> float:
    factor = (1 + pct / 100.0) if pallets_count >= 7 else 1.0
    return price_base * factor + fixed

if dual_toggle:
    b1 = get_base_rate(rate_df, postcode_area, service_option, "Joda", split1)
    b2 = get_base_rate(rate_df, postcode_area, service_option, "Joda", split2)
    if b1 is not None and b2 is not None:
        g1 = apply_joda(b1, split1, joda_surcharge_pct, joda_charge_fixed)
        g2 = apply_joda(b2, split2, joda_surcharge_pct, joda_charge_fixed)
        joda_base = b1 + b2
        joda_final = g1 + g2
else:
    base = get_base_rate(rate_df, postcode_area, service_option, "Joda", num_pallets)
    if base is not None:
        joda_base = base
        joda_final = apply_joda(base, num_pallets, joda_surcharge_pct, joda_charge_fixed)

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

# Compute an effective display value for Joda surcharge
def joda_effective_pct_label() -> str:
    if joda_base is None:
        return f"{joda_surcharge_pct:.2f}%"
    if not dual_toggle:
        return "0.00%" if num_pallets < 7 else f"{joda_surcharge_pct:.2f}%"
    # dual: per-leg rule
    both_under = (split1 < 7 and split2 < 7)
    both_over  = (split1 >= 7 and split2 >= 7)
    if both_under:
        return "0.00%"
    if both_over:
        return f"{joda_surcharge_pct:.2f}%"
    return f"{joda_surcharge_pct:.2f}% (partial)"

if joda_base is None:
    summary_rows.append({
        "Haulier": "Joda",
        "Base Rate": "No rate",
        "Fuel Surcharge (%)": joda_effective_pct_label(),
        "Delivery Charge": "N/A",
        "Final Rate": "N/A"
    })
else:
    summary_rows.append({
        "Haulier": "Joda",
        "Base Rate": f"£{joda_base:,.2f}",
        "Fuel Surcharge (%)": joda_effective_pct_label(),
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

    st.header("3. Calculated Rates")
    st.table(summary_df.style.apply(highlight_cheapest, axis=1))
    st.markdown("<i style='color:gray;'>Rows in green are the cheapest available.</i>", unsafe_allow_html=True)

# ── 10) ONE PALLET FEWER / MORE (supports per-pallet charge for McD)
def lookup_adjacent_rate(df, area, service, vendor, pallets,
                         surcharge_pct, fixed_charge=0.0, per_pallet_charge=0.0):
    out = {"lower": None, "higher": None}

    # Helper to apply Joda's <7 pallet rule inside the preview too
    def eff_total(vendor_name: str, pallet_count: int, base_rate: float) -> float:
        if vendor_name == "Joda":
            factor = (1 + surcharge_pct / 100.0) if pallet_count >= 7 else 1.0
            return base_rate * factor + fixed_charge
        else:
            return base_rate * (1 + surcharge_pct / 100.0) + fixed_charge + per_pallet_charge * pallet_count

    if pallets > 1:
        bl = get_base_rate(df, area, service, vendor, pallets - 1)
        if bl is not None:
            out["lower"] = ((pallets - 1), eff_total(vendor, pallets - 1, bl))
    bh = get_base_rate(df, area, service, vendor, pallets + 1)
    if bh is not None:
        out["higher"] = ((pallets + 1), eff_total(vendor, pallets + 1, bh))
    return out

joda_adj = lookup_adjacent_rate(
    rate_df, postcode_area, service_option, "Joda",
    num_pallets, joda_surcharge_pct,
    fixed_charge=joda_charge_fixed, per_pallet_charge=0.0
)
mcd_adj = lookup_adjacent_rate(
    rate_df, postcode_area, service_option, "Mcdowells",
    num_pallets, mcd_surcharge_pct,
    fixed_charge=mcd_charge_fixed, per_pallet_charge=mcd_tail_lift_per_pallet
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
    • Joda surcharge resets each Wednesday; McDowells is entered per session.  
    • Delivery charges: Joda – AM/PM £7, Timed £19; McDowells – AM/PM £10, Timed £19.  
    • Tail Lift: Joda £0; McDowells £3.90 per pallet.  
    • Dual Collection splits Joda into two shipments; McDowells unaffected. 
    TEST
    </small>
    """,
    unsafe_allow_html=True
)
