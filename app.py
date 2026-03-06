"""
VAT Validator — Mammoet Data Migration
"""

import time
import re
import io
from datetime import datetime

import streamlit as st
import pandas as pd
import requests

st.set_page_config(
    page_title="VAT Validator | Mammoet",
    page_icon="",
    layout="wide",
    initial_sidebar_state="collapsed"
)

st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Barlow:wght@400;500;600;700;800&family=Barlow+Condensed:wght@600;700;800&display=swap');

    html, body, [class*="css"] { font-family: 'Barlow', sans-serif; }
    .stApp { background-color: #0f0f0f; color: #d0d0d0; }
    .block-container { padding: 0 2.5rem !important; max-width: 100% !important; }

    .hero {
        border-bottom: 3px solid #E30613;
        padding: 2rem 0 1.5rem;
        margin-bottom: 2rem;
    }
    .hero-title {
        font-family: 'Barlow Condensed', sans-serif;
        font-size: 2.6rem; font-weight: 800; color: #ffffff;
        letter-spacing: -1px; margin: 0; line-height: 1;
    }
    .hero-title span { color: #E30613; }
    .hero-sub {
        color: #aaaaaa; font-size: 1rem; margin-top: 0.5rem;
    }

    textarea {
        background: #1a1a1a !important; color: #d0d0d0 !important;
        border: 1px solid #2a2a2a !important; border-radius: 8px !important;
    }

    .stButton > button {
        background: #E30613 !important; color: #ffffff !important;
        border: none !important; font-weight: 700 !important;
        padding: 0.7rem 2rem !important; border-radius: 5px !important;
    }

    .stButton > button:hover { background: #c0050f !important; }

    .stDownloadButton > button {
        background: #1a1a1a !important; color: #d0d0d0 !important;
        border: 1px solid #333333 !important;
        border-radius: 5px !important;
    }

    .stTabs [data-baseweb="tab"] {
        font-weight: 600 !important;
    }

    .stTabs [data-baseweb="tab"]:nth-child(2) {
        color: #4a9d6f !important;
    }

    .stTabs [data-baseweb="tab"]:nth-child(3) {
        color: #E30613 !important;
    }

    .stTabs [data-baseweb="tab"]:nth-child(4) {
        color: #c97d20 !important;
    }

    /* Uploaded filename visibility */
    [data-testid="stFileUploaderFileName"] {
        color: #ffffff !important;
        font-weight: 600;
    }

    #MainMenu, footer, header { visibility: hidden; }
</style>
""", unsafe_allow_html=True)

VIES_URL = "https://ec.europa.eu/taxation_customs/vies/rest-api/ms/{country_code}/vat/{vat_number}"

EU_COUNTRY_CODES = {
"AT","BE","BG","CY","CZ","DE","DK","EE","EL","ES",
"FI","FR","HR","HU","IE","IT","LT","LU","LV","MT",
"NL","PL","PT","RO","SE","SI","SK","XI"
}

TIMEOUT = 10
MAX_RETRIES = 3


def parse_vat(raw):
    cleaned = re.sub(r"[\s.\-]", "", str(raw).upper())
    if len(cleaned) < 3:
        return "", ""
    return cleaned[:2], cleaned[2:]


def check_vat(country_code, vat_number):

    if country_code not in EU_COUNTRY_CODES:
        return {"status":"error","company_name":"—","company_address":"—",
                "message":f"Country code '{country_code}' not supported"}

    url = VIES_URL.format(country_code=country_code, vat_number=vat_number)

    try:
        r = requests.get(url, timeout=TIMEOUT)

        if r.status_code == 200:
            data = r.json()
            valid = data.get("isValid", False)

            return {
                "status":"valid" if valid else "invalid",
                "company_name":data.get("name") or "—",
                "company_address":data.get("address") or "—",
                "message":"Valid" if valid else "Invalid according to VIES"
            }

        elif r.status_code == 404:
            return {"status":"invalid","company_name":"—","company_address":"—","message":"Not found in VIES"}

        else:
            return {"status":"error","company_name":"—","company_address":"—","message":f"HTTP {r.status_code}"}

    except Exception as e:
        return {"status":"error","company_name":"—","company_address":"—","message":str(e)}


def detect_vat_column(df):

    for col in df.columns:
        if "vat" in col.lower():
            return col

    return df.columns[0]


def highlight_row(row):

    if row["Status"] == "invalid":
        return ["background-color:#2a0d0f"] * len(row)

    if row["Status"] == "error":
        return ["background-color:#2a1a0a"] * len(row)

    return [""] * len(row)


def run_validation(df_input, vat_col):

    progress_bar = st.progress(0)

    results = []
    total = len(df_input)

    for i, (_, row) in enumerate(df_input.iterrows()):

        raw = str(row[vat_col]).strip()

        country_code, vat_number = parse_vat(raw)

        if not country_code:
            result = {"status":"invalid","company_name":"—","company_address":"—","message":"Invalid format"}
        else:
            result = check_vat(country_code, vat_number)

        results.append({
            "VAT Input": raw,
            "Country": country_code,
            "VAT Number": vat_number,
            "Status": result["status"],
            "Company (VIES)":result["company_name"],
            "Address (VIES)":result["company_address"],
            "Message":result["message"]
        })

        progress_bar.progress((i+1)/total)

        time.sleep(0.4)

    st.session_state.results = pd.DataFrame(results)

    st.rerun()


st.markdown("""
<div class="hero">
    <div class="hero-title">VAT <span>VALIDATOR</span></div>
    <div class="hero-sub">VAT numbers are validated via the official VIES API.</div>
</div>
""", unsafe_allow_html=True)


if "results" not in st.session_state:
    st.session_state.results = None


col_left, col_divider, col_right = st.columns([5,1,6])


df_input = None
vat_col = None


with col_left:

    uploaded = st.file_uploader("", type=["csv","xlsx","xls"])

    if uploaded:

        if uploaded.name.endswith(".csv"):
            df_input = pd.read_csv(uploaded, dtype=str)

        else:
            df_input = pd.read_excel(uploaded, dtype=str)

        vat_col = detect_vat_column(df_input)

        df_input = df_input.fillna("")

        st.write(f"{len(df_input)} rows loaded · VAT column: {vat_col}")

    if df_input is not None:

        label = f"Validate {len(df_input):,} VAT numbers"

        if st.button(label):

            run_validation(df_input, vat_col)


with col_right:

    if st.session_state.results is not None:

        df_res = st.session_state.results

        valid_df = df_res[df_res["Status"]=="valid"]
        invalid_df = df_res[df_res["Status"]=="invalid"]
        error_df = df_res[df_res["Status"]=="error"]

        total = len(df_res)

        t1, t2, t3, t4 = st.tabs([
            f"All results ({total})",
            f"✓ Valid ({len(valid_df)})",
            f"✗ Invalid ({len(invalid_df)})",
            f"! Error ({len(error_df)})"
        ])

        with t1:
            st.dataframe(
                df_res.style.apply(highlight_row, axis=1),
                use_container_width=True,
                hide_index=True
            )

        with t2:
            st.dataframe(valid_df, use_container_width=True, hide_index=True)

        with t3:
            st.dataframe(invalid_df, use_container_width=True, hide_index=True)

        with t4:
            st.dataframe(error_df, use_container_width=True, hide_index=True)
