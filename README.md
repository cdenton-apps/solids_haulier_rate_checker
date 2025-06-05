# Solidus Haulier Rate Checker

This Streamlit app allows you to compare Joda Freight and McDowells shipping rates for UK postcodes.

## Contents

- `haulier prices.xlsx`: Built-in rate table (Joda & McDowells).
- `waste_solidus_haulier_app.py`: Main Streamlit application.
- `assets/solidus_logo.png`: Solidus logo (must be provided).
- `.streamlit/config.toml`: Streamlit theme configuration.
- `requirements.txt`: Python dependencies.
- `.gitignore`: Files to ignore in Git.

## Setup

1. **Clone the repository**:
   ```bash
   git clone <your-repo-url>
   cd solidus-haulier-rate-checker
   ```

2. **Install dependencies**:
   ```bash
   pip install -r requirements.txt
   ```

3. **Ensure the Solidus logo** is saved as `assets/solidus_logo.png`.

4. **Run the app**:
   ```bash
   streamlit run waste_solidus_haulier_app.py
   ```

## Usage

1. Enter a UK postcode (e.g., `BB10 1AB`).
2. Select service type: Economy or Next Day.
3. Specify the number of pallets.
4. Enter McDowells fuel surcharge percentage (Joda’s surcharge is fetched automatically).
5. View final rates for Joda and McDowells, with the cheapest highlighted.
6. See adjacent pallet pricing (±1 pallet) in grey.

***
