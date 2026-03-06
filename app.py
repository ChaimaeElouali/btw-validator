"""
BTW Validator — Mammoet Data Migration
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
    page_icon="🔍",
    layout="wide",
    initial_sidebar_state="collapsed"
)

st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Barlow:wght@400;500;600;700;800&family=Barlow+Condensed:wght@600;700;800&display=swap');

    html, body, [class*="css"] { font-family: 'Barlow', sans-serif; }
    .stApp { background-color: #0f0f0f; color: #e0e0e0; }
    .block-container { padding-top: 0 !important; max-width: 860px; margin: 0 auto; }

    /* Hero */
    .hero {
        border-bottom: 3px solid #E30613;
        padding: 2.2rem 0 1.6rem;
        margin-bottom: 2.5rem;
    }
    .hero-title {
        font-family: 'Barlow Condensed', sans-serif;
        font-size: 3rem; font-weight: 800; color: #fff;
        letter-spacing: -1px; margin: 0; line-height: 1;
    }
    .hero-title span { color: #E30613; }
    .hero-sub {
        color: #666; font-size: 0.78rem; margin-top: 0.5rem;
        font-weight: 600; letter-spacing: 2px; text-transform: uppercase;
    }

    /* Section label */
    .section-label {
        font-size: 0.72rem; font-weight: 700; text-transform: uppercase;
        letter-spacing: 2px; color: #E30613; margin-bottom: 0.6rem;
        display: block;
    }
    .section-hint {
        font-size: 0.85rem; color: #aaa; margin-bottom: 0.75rem;
        display: block; font-weight: 400;
    }

    /* Or divider */
    .or-col {
        display: flex; flex-direction: column;
        align-items: center; justify-content: center;
        height: 100%; padding: 0;
    }
    .or-line { flex: 1; width: 1px; background: #333; min-height: 40px; }
    .or-text {
        font-size: 0.75rem; font-weight: 700; letter-spacing: 2px;
        color: #666; padding: 10px 0; text-transform: uppercase;
    }

    /* Upload zone */
    [data-testid="stFileUploader"] { background: transparent !important; }
    [data-testid="stFileUploader"] > div {
        background: #1a1a1a !important;
        border: 2px dashed #333 !important;
        border-radius: 8px !important;
        transition: border-color 0.2s;
    }
    [data-testid="stFileUploader"] > div:hover { border-color: #E30613 !important; }
    [data-testid="stFileUploader"] section { background: #1a1a1a !important; border: none !important; }
    [data-testid="stFileUploader"] button {
        background: #2a2a2a !important; color: #e0e0e0 !important;
        border: 1px solid #444 !important; border-radius: 4px !important;
        font-family: 'Barlow', sans-serif !important; font-weight: 600 !important;
    }
    [data-testid="stFileUploader"] p { color: #888 !important; font-size: 0.85rem !important; }
    [data-testid="stFileUploader"] small { color: #666 !important; }
    [data-testid="stFileUploaderDropzoneInstructions"] { color: #888 !important; }
    [data-testid="stFileUploaderDropzoneInstructions"] div span { color: #888 !important; }

    /* Textarea */
    textarea {
        background: #1a1a1a !important; color: #e0e0e0 !important;
        border: 2px solid #333 !important; border-radius: 8px !important;
        font-family: 'Barlow', sans-serif !important; font-size: 0.9rem !important;
    }
    textarea:focus { border-color: #E30613 !important; box-shadow: none !important; }
    textarea::placeholder { color: #555 !important; }

    /* Buttons */
    .stButton > button {
        background: #E30613 !important; color: #fff !important;
        border: none !important; font-family: 'Barlow', sans-serif !important;
        font-weight: 700 !important; font-size: 1rem !important;
        padding: 0.7rem 2rem !important; border-radius: 6px !important;
    }
    .stButton > button:hover { background: #c0050f !important; }
    .stDownloadButton > button {
        background: #1a1a1a !important; color: #e0e0e0 !important;
        border: 1px solid #333 !important; font-family: 'Barlow', sans-serif !important;
        font-weight: 700 !important; border-radius: 6px !important;
    }

    /* Stat cards */
    .stat-row { display: grid; grid-template-columns: repeat(4,1fr); gap: 12px; margin-bottom: 1.5rem; }
    .stat-card { background: #1a1a1a; border-radius: 8px; padding: 1.25rem 1rem; text-align: center; border: 1px solid #2a2a2a; }
    .stat-card.total   { border-top: 3px solid #555; }
    .stat-card.valid   { border-top: 3px solid #22c55e; }
    .stat-card.invalid { border-top: 3px solid #E30613; }
    .stat-card.error   { border-top: 3px solid #f59e0b; }
    .stat-num { font-family: 'Barlow Condensed', sans-serif; font-size: 2.5rem; font-weight: 800; line-height: 1; }
    .stat-lbl { font-size: 0.7rem; font-weight: 700; text-transform: uppercase; letter-spacing: 1px; color: #666; margin-top: 4px; }
    .total   .stat-num { color: #888; }
    .valid   .stat-num { color: #22c55e; }
    .invalid .stat-num { color: #E30613; }
    .error   .stat-num { color: #f59e0b; }

    /* Tabs */
    .stTabs [data-baseweb="tab-list"] { background: #1a1a1a; border-bottom: 1px solid #2a2a2a; }
    .stTabs [data-baseweb="tab"] {
        color: #666 !important; font-family: 'Barlow', sans-serif !important;
        font-weight: 600 !important; padding: 0.7rem 1.5rem !important;
    }
    .stTabs [aria-selected="true"] {
        color: #fff !important; border-bottom: 3px solid #E30613 !important;
        background: transparent !important;
    }

    .stProgress > div > div { background-color: #E30613 !important; }
    hr { border-color: #2a2a2a !important; }
    #MainMenu, footer, header { visibility: hidden; }
</style>
""", unsafe_allow_html=True)

# ─── VIES API ───────────────────────────────
VIES_URL = "https://ec.europa.eu/taxation_customs/vies/rest-api/ms/{country_code}/vat/{vat_number}"
TIMEOUT = 10
MAX_RETRIES = 3
EU_COUNTRY_CODES = {
    "AT","BE","BG","CY","CZ","DE","DK","EE","EL","ES",
    "FI","FR","HR","HU","IE","IT","LT","LU","LV","MT",
    "NL","PL","PT","RO","SE","SI","SK","XI"
}

def parse_vat(raw):
    if not isinstance(raw, str): raw = str(raw)
    cleaned = re.sub(r"[\s.\-]", "", raw.strip().upper())
    if len(cleaned) < 3: return "", ""
    return cleaned[:2], cleaned[2:]

def check_vat(country_code, vat_number):
    if country_code not in EU_COUNTRY_CODES:
        return {"status":"error","company_name":"—","company_address":"—",
                "message":f"Country code '{country_code}' not supported by VIES"}
    url = VIES_URL.format(country_code=country_code, vat_number=vat_number)
    for attempt in range(1, MAX_RETRIES+1):
        try:
            r = requests.get(url, timeout=TIMEOUT)
            if r.status_code == 200:
                data = r.json()
                valid = data.get("isValid", False)
                return {"status":"valid" if valid else "invalid",
                        "company_name": data.get("name") or "—",
                        "company_address": data.get("address") or "—",
                        "message":"Valid" if valid else "Invalid according to VIES"}
            elif r.status_code == 404:
                return {"status":"invalid","company_name":"—","company_address":"—","message":"Not found in VIES"}
            elif r.status_code in (429,503,504):
                time.sleep(2**attempt); continue
            else:
                return {"status":"error","company_name":"—","company_address":"—","message":f"HTTP {r.status_code}"}
        except requests.exceptions.Timeout:
            if attempt < MAX_RETRIES: time.sleep(2); continue
            return {"status":"error","company_name":"—","company_address":"—","message":"Timeout"}
        except Exception as e:
            return {"status":"error","company_name":"—","company_address":"—","message":str(e)}
    return {"status":"error","company_name":"—","company_address":"—","message":"VIES unavailable"}

def detect_vat_column(df):
    for col in df.columns:
        if any(kw in col.lower() for kw in ["vat","btw","tax"]): return col
    return df.columns[0]

def to_excel_bytes(df):
    output = io.BytesIO()
    valid_df   = df[df["Status"]=="valid"]
    invalid_df = df[df["Status"]=="invalid"]
    error_df   = df[~df["Status"].isin(["valid","invalid"])]
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        pd.DataFrame({
            "Category":["Valid","Invalid","Error/Unknown","Total"],
            "Count":[len(valid_df),len(invalid_df),len(error_df),len(df)],
            "Percentage":[
                f"{len(valid_df)/len(df)*100:.1f}%" if len(df) else "0%",
                f"{len(invalid_df)/len(df)*100:.1f}%" if len(df) else "0%",
                f"{len(error_df)/len(df)*100:.1f}%" if len(df) else "0%","100%"]
        }).to_excel(writer, sheet_name="Summary", index=False)
        ws = writer.sheets["Summary"]
        ws.cell(row=7,column=1,value=f"Report generated: {datetime.now().strftime('%d-%m-%Y %H:%M')}")
        ws.cell(row=8,column=1,value="Source: VIES API — European Commission")
        df.to_excel(writer, sheet_name="All results", index=False)
        if len(valid_df):   valid_df.to_excel(writer, sheet_name="Valid", index=False)
        if len(invalid_df): invalid_df.to_excel(writer, sheet_name="Invalid", index=False)
        if len(error_df):   error_df.to_excel(writer, sheet_name="Errors", index=False)
    return output.getvalue()

def run_validation(df_input, vat_col):
    st.markdown("---")
    st.markdown("**Validating...**")
    progress_bar = st.progress(0)
    status_text  = st.empty()
    results = []
    total = len(df_input)
    for i, (_, row) in enumerate(df_input.iterrows()):
        raw = str(row[vat_col]).strip()
        country_code, vat_number = parse_vat(raw)
        status_text.markdown(
            f"<small style='color:#666'>{i+1}/{total} — <code style='color:#999'>{raw}</code></small>",
            unsafe_allow_html=True)
        result = {"status":"invalid","company_name":"—","company_address":"—","message":"Invalid format"} \
            if not country_code or not vat_number else check_vat(country_code, vat_number)
        row_data = {"VAT Input": raw, "Country": country_code, "VAT Number": vat_number}
        for c in df_input.columns:
            if c != vat_col: row_data[c] = row[c]
        row_data.update({"Status":result["status"],"Company (VIES)":result["company_name"],
                         "Address (VIES)":result["company_address"],"Message":result["message"]})
        results.append(row_data)
        progress_bar.progress((i+1)/total)
        time.sleep(0.4)
    status_text.empty()
    st.session_state.results = pd.DataFrame(results)
    st.rerun()

# ─── HEADER ─────────────────────────────────
st.markdown("""
<div class="hero">
    <div class="hero-title">VAT <span>VALIDATOR</span></div>
    <div class="hero-sub">Mammoet Data Migration &nbsp;·&nbsp; SAP ECC → S/4HANA &nbsp;·&nbsp; VIES API</div>
</div>
""", unsafe_allow_html=True)

if "results" not in st.session_state:
    st.session_state.results = None

# ─── INPUT ──────────────────────────────────
df_input = None
vat_col  = "vat_number"

col_left, col_or, col_right = st.columns([10, 1, 10], gap="small")

with col_left:
    st.markdown('<span class="section-label">📁 Upload a file</span>', unsafe_allow_html=True)
    st.markdown('<span class="section-hint">CSV or Excel file with a column containing VAT numbers</span>', unsafe_allow_html=True)
    uploaded = st.file_uploader("upload", type=["csv","xlsx","xls"], label_visibility="collapsed")
    if uploaded:
        try:
            df_input = pd.read_csv(uploaded, dtype=str) if uploaded.name.endswith(".csv") else pd.read_excel(uploaded, dtype=str)
            vat_col = detect_vat_column(df_input)
            df_input = df_input.fillna("")
            st.markdown(f"<small style='color:#aaa'>✅ {len(df_input)} rows loaded · column: <code style='color:#ccc'>{vat_col}</code></small>", unsafe_allow_html=True)
        except Exception as e:
            st.error(f"Error reading file: {e}")
            df_input = None

with col_or:
    st.markdown("""
<div class="or-col">
  <div class="or-line"></div>
  <div class="or-text">or</div>
  <div class="or-line"></div>
</div>
""", unsafe_allow_html=True)

with col_right:
    st.markdown('<span class="section-label">📋 Paste VAT numbers</span>', unsafe_allow_html=True)
    st.markdown('<span class="section-hint">One VAT number per line — e.g. NL800336808B01, DE811115329</span>', unsafe_allow_html=True)
    pasted = st.text_area("paste", height=158,
        placeholder="NL800336808B01\nDE811115329\nBE0999999999\n...",
        key="paste_input", label_visibility="collapsed")
    if not uploaded and pasted and pasted.strip():
        lines = [l.strip() for l in pasted.splitlines() if l.strip()]
        if lines:
            df_input = pd.DataFrame({"vat_number": lines})
            vat_col = "vat_number"
            st.markdown(f"<small style='color:#aaa'>✅ {len(lines)} VAT number(s) loaded</small>", unsafe_allow_html=True)

# ─── VALIDATE BUTTON ────────────────────────
st.markdown("<br>", unsafe_allow_html=True)
if df_input is not None and len(df_input) > 0:
    col_btn, _ = st.columns([2, 5])
    with col_btn:
        if st.button(f"Validate {len(df_input)} VAT numbers →", use_container_width=True):
            run_validation(df_input, vat_col)
else:
    st.markdown("<small style='color:#555'>Upload a file or paste VAT numbers to get started.</small>", unsafe_allow_html=True)

# ─── RESULTS ────────────────────────────────
if st.session_state.results is not None:
    df_res = st.session_state.results
    valid_df   = df_res[df_res["Status"]=="valid"]
    invalid_df = df_res[df_res["Status"]=="invalid"]
    error_df   = df_res[~df_res["Status"].isin(["valid","invalid"])]
    total = len(df_res)

    st.markdown("---")

    st.markdown(f"""
<div class="stat-row">
  <div class="stat-card total"><div class="stat-num">{total}</div><div class="stat-lbl">Total</div></div>
  <div class="stat-card valid"><div class="stat-num">{len(valid_df)}</div><div class="stat-lbl">Valid</div></div>
  <div class="stat-card invalid"><div class="stat-num">{len(invalid_df)}</div><div class="stat-lbl">Invalid</div></div>
  <div class="stat-card error"><div class="stat-num">{len(error_df)}</div><div class="stat-lbl">Error</div></div>
</div>
""", unsafe_allow_html=True)

    col_dl, col_reset, _ = st.columns([3, 1, 4])
    with col_dl:
        st.download_button(
            label="📥 Download report (.xlsx)",
            data=to_excel_bytes(df_res),
            file_name=f"vat_validation_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )
    with col_reset:
        if st.button("Reset", use_container_width=True):
            st.session_state.results = None
            st.rerun()

    st.markdown("<br>", unsafe_allow_html=True)

    def style_status(val):
        return {"valid":"color:#22c55e;font-weight:700","invalid":"color:#E30613;font-weight:700"}.get(val,"color:#f59e0b;font-weight:700")

    t1, t2, t3, t4 = st.tabs([
        f"All results ({total})",
        f"Valid ({len(valid_df)})",
        f"Invalid ({len(invalid_df)})",
        f"Error ({len(error_df)})"
    ])
    with t1:
        st.dataframe(df_res.style.applymap(style_status, subset=["Status"]), use_container_width=True, hide_index=True, height=400)
    with t2:
        if len(valid_df) > 0:
            st.dataframe(valid_df, use_container_width=True, hide_index=True, height=400)
        else:
            st.info("No valid VAT numbers found.")
    with t3:
        if len(invalid_df) > 0:
            st.dataframe(invalid_df, use_container_width=True, hide_index=True, height=400)
        else:
            st.success("No invalid VAT numbers found!")
    with t4:
        if len(error_df) > 0:
            st.dataframe(error_df, use_container_width=True, hide_index=True, height=400)
        else:
            st.success("No errors found.")

st.markdown("---")
st.markdown("<p style='color:#2a2a2a;font-size:0.75rem;text-align:center;'>Mammoet Data Migration Team · VIES API (European Commission) · SAP ECC → S/4HANA</p>", unsafe_allow_html=True)
