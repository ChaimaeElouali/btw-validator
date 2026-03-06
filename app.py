"""
BTW Validator — Mammoet Data Migration
Streamlit webapplicatie voor validatie van BTW-nummers via VIES API.

Installatie:
    pip install streamlit pandas openpyxl requests

Starten:
    streamlit run app.py
"""

import time
import re
import io
from datetime import datetime

import streamlit as st
import pandas as pd
import requests

st.set_page_config(
    page_title="BTW Validator | Mammoet",
    page_icon="🔍",
    layout="centered",
    initial_sidebar_state="collapsed"
)

st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Barlow:wght@400;500;600;700;800&family=Barlow+Condensed:wght@600;700;800&display=swap');

    html, body, [class*="css"] { font-family: 'Barlow', sans-serif; }
    .stApp { background-color: #111111; color: #f0f0f0; }

    .hero {
        background: #111111;
        border-bottom: 4px solid #E30613;
        padding: 2.2rem 0 1.8rem;
        margin-bottom: 2.5rem;
    }
    .hero-title {
        font-family: 'Barlow Condensed', sans-serif;
        font-size: 2.8rem; font-weight: 800; color: #fff;
        letter-spacing: -0.5px; margin: 0; line-height: 1;
    }
    .hero-title span { color: #E30613; }
    .hero-sub {
        color: #666; font-size: 0.8rem; margin-top: 0.5rem;
        font-weight: 600; letter-spacing: 1.5px; text-transform: uppercase;
    }

    .section-label {
        font-size: 0.75rem; font-weight: 700; text-transform: uppercase;
        letter-spacing: 1.5px; color: #888; margin-bottom: 0.5rem;
    }

    .or-bar {
        display: flex; align-items: center; gap: 1rem;
        margin: 1.2rem 0; color: #444;
        font-size: 0.75rem; font-weight: 700; letter-spacing: 2px; text-transform: uppercase;
    }
    .or-bar::before, .or-bar::after {
        content: ''; flex: 1; height: 1px; background: #2a2a2a;
    }

    .validate-hint {
        color: #555; font-size: 0.85rem; font-style: italic; margin-top: 0.5rem;
    }

    .stat-card {
        background: #1a1a1a; border-radius: 8px; padding: 1.25rem;
        text-align: center; border: 1px solid #2a2a2a;
    }
    .stat-card.valid   { border-top: 4px solid #22c55e; }
    .stat-card.invalid { border-top: 4px solid #E30613; }
    .stat-card.error   { border-top: 4px solid #f59e0b; }
    .stat-card.total   { border-top: 4px solid #555; }
    .stat-number {
        font-family: 'Barlow Condensed', sans-serif;
        font-size: 2.8rem; font-weight: 800; line-height: 1;
    }
    .stat-label {
        font-size: 0.7rem; font-weight: 700; text-transform: uppercase;
        letter-spacing: 1px; color: #666; margin-top: 0.25rem;
    }
    .valid   .stat-number { color: #22c55e; }
    .invalid .stat-number { color: #E30613; }
    .error   .stat-number { color: #f59e0b; }
    .total   .stat-number { color: #aaa; }

    /* Buttons */
    .stButton > button {
        background: #E30613 !important; color: white !important;
        border: none !important; font-family: 'Barlow', sans-serif !important;
        font-weight: 700 !important; font-size: 1rem !important;
        padding: 0.7rem 2rem !important; border-radius: 6px !important;
        width: 100% !important;
    }
    .stButton > button:hover { background: #c0050f !important; }

    .stDownloadButton > button {
        background: #1a1a1a !important; color: #f0f0f0 !important;
        border: 1px solid #333 !important; font-family: 'Barlow', sans-serif !important;
        font-weight: 700 !important; border-radius: 6px !important;
        width: 100% !important;
    }

    /* File uploader */
    [data-testid="stFileUploader"] {
        background: #1a1a1a !important;
        border: 2px dashed #2a2a2a !important;
        border-radius: 8px;
    }
    [data-testid="stFileUploader"]:hover {
        border-color: #E30613 !important;
    }

    /* Textarea */
    textarea {
        background: #1a1a1a !important;
        color: #f0f0f0 !important;
        border: 2px solid #2a2a2a !important;
        border-radius: 8px !important;
        font-family: 'Barlow', sans-serif !important;
        font-size: 0.9rem !important;
    }
    textarea:focus {
        border-color: #E30613 !important;
        box-shadow: none !important;
    }

    /* Tabs */
    .stTabs [data-baseweb="tab-list"] {
        background: #1a1a1a;
        border-bottom: 1px solid #2a2a2a;
        border-radius: 8px 8px 0 0;
    }
    .stTabs [data-baseweb="tab"] {
        color: #555 !important; font-family: 'Barlow', sans-serif !important;
        font-weight: 600 !important; padding: 0.75rem 1.5rem !important;
    }
    .stTabs [aria-selected="true"] {
        color: #fff !important;
        border-bottom: 3px solid #E30613 !important;
        background: transparent !important;
    }

    .stProgress > div > div { background-color: #E30613 !important; }
    hr { border-color: #2a2a2a !important; }
    #MainMenu, footer, header { visibility: hidden; }
    .block-container { padding-top: 0 !important; max-width: 720px; }
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
                "message":f"Landcode '{country_code}' niet ondersteund"}
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
                        "message":"Geldig" if valid else "Ongeldig volgens VIES"}
            elif r.status_code == 404:
                return {"status":"invalid","company_name":"—","company_address":"—","message":"Niet gevonden in VIES"}
            elif r.status_code in (429,503,504):
                time.sleep(2**attempt); continue
            else:
                return {"status":"error","company_name":"—","company_address":"—","message":f"HTTP {r.status_code}"}
        except requests.exceptions.Timeout:
            if attempt < MAX_RETRIES: time.sleep(2); continue
            return {"status":"error","company_name":"—","company_address":"—","message":"Timeout"}
        except Exception as e:
            return {"status":"error","company_name":"—","company_address":"—","message":str(e)}
    return {"status":"error","company_name":"—","company_address":"—","message":"VIES niet beschikbaar"}

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
            "Categorie":["Geldig","Ongeldig","Fout/Onbekend","Totaal"],
            "Aantal":[len(valid_df),len(invalid_df),len(error_df),len(df)],
            "Percentage":[
                f"{len(valid_df)/len(df)*100:.1f}%" if len(df) else "0%",
                f"{len(invalid_df)/len(df)*100:.1f}%" if len(df) else "0%",
                f"{len(error_df)/len(df)*100:.1f}%" if len(df) else "0%","100%"]
        }).to_excel(writer, sheet_name="Samenvatting", index=False)
        ws = writer.sheets["Samenvatting"]
        ws.cell(row=7,column=1,value=f"Rapport: {datetime.now().strftime('%d-%m-%Y %H:%M')}")
        ws.cell(row=8,column=1,value="Bron: VIES API — European Commission")
        df.to_excel(writer, sheet_name="Alle resultaten", index=False)
        if len(valid_df):   valid_df.to_excel(writer, sheet_name="Geldig", index=False)
        if len(invalid_df): invalid_df.to_excel(writer, sheet_name="Ongeldig", index=False)
        if len(error_df):   error_df.to_excel(writer, sheet_name="Fouten", index=False)
    return output.getvalue()

def run_validation(df_input, vat_col):
    st.markdown("---")
    st.markdown("**Validatie bezig...**")
    progress_bar = st.progress(0)
    status_text  = st.empty()
    results = []
    total = len(df_input)
    for i, (_, row) in enumerate(df_input.iterrows()):
        raw = str(row[vat_col]).strip()
        country_code, vat_number = parse_vat(raw)
        status_text.markdown(
            f"<small style='color:#666'>Verwerken {i+1}/{total}: <code style='color:#aaa'>{raw}</code></small>",
            unsafe_allow_html=True)
        result = {"status":"invalid","company_name":"—","company_address":"—","message":"Ongeldig formaat"} \
            if not country_code or not vat_number else check_vat(country_code, vat_number)
        row_data = {"VAT Input": raw, "Land": country_code, "BTW Nummer": vat_number}
        for c in df_input.columns:
            if c != vat_col: row_data[c] = row[c]
        row_data.update({"Status":result["status"],"Bedrijfsnaam (VIES)":result["company_name"],
                         "Adres (VIES)":result["company_address"],"Melding":result["message"]})
        results.append(row_data)
        progress_bar.progress((i+1)/total)
        time.sleep(0.4)
    status_text.empty()
    st.session_state.results = pd.DataFrame(results)
    st.rerun()

# ─── HEADER ─────────────────────────────────
st.markdown("""
<div class="hero">
    <div class="hero-title">BTW <span>VALIDATOR</span></div>
    <div class="hero-sub">Mammoet Data Migration &nbsp;·&nbsp; SAP ECC → S/4HANA &nbsp;·&nbsp; VIES API</div>
</div>
""", unsafe_allow_html=True)

if "results" not in st.session_state:
    st.session_state.results = None

# ─── INPUT ──────────────────────────────────
df_input = None
vat_col  = "vat_number"

st.markdown('<div class="section-label">Bestand uploaden — CSV of Excel</div>', unsafe_allow_html=True)
uploaded = st.file_uploader("upload", type=["csv","xlsx","xls"], label_visibility="collapsed")

if uploaded:
    try:
        df_input = pd.read_csv(uploaded, dtype=str) if uploaded.name.endswith(".csv") else pd.read_excel(uploaded, dtype=str)
        vat_col = detect_vat_column(df_input)
        df_input = df_input.fillna("")
        st.markdown(f"<small style='color:#555'>✅ {len(df_input)} rijen geladen · BTW-kolom: <code style='color:#888'>{vat_col}</code></small>", unsafe_allow_html=True)
    except Exception as e:
        st.error(f"Fout bij inlezen: {e}")
        df_input = None

st.markdown('<div class="or-bar">of</div>', unsafe_allow_html=True)

st.markdown('<div class="section-label">BTW-nummers plakken — één per regel</div>', unsafe_allow_html=True)
pasted = st.text_area("paste", height=160,
    placeholder="NL800336808B01\nDE811115329\nBE0999999999\n...",
    key="paste_input", label_visibility="collapsed")

if not uploaded and pasted and pasted.strip():
    lines = [l.strip() for l in pasted.splitlines() if l.strip()]
    if lines:
        df_input = pd.DataFrame({"vat_number": lines})
        vat_col = "vat_number"
        st.markdown(f"<small style='color:#555'>✅ {len(lines)} BTW-nummer(s) geladen</small>", unsafe_allow_html=True)

st.markdown("<br>", unsafe_allow_html=True)

if df_input is not None and len(df_input) > 0:
    if st.button(f"Valideer {len(df_input)} BTW-nummers →"):
        run_validation(df_input, vat_col)
else:
    st.markdown('<div class="validate-hint">Upload een bestand of plak BTW-nummers om te beginnen.</div>', unsafe_allow_html=True)

# ─── RESULTATEN ─────────────────────────────
if st.session_state.results is not None:
    df_res = st.session_state.results
    valid_df   = df_res[df_res["Status"]=="valid"]
    invalid_df = df_res[df_res["Status"]=="invalid"]
    error_df   = df_res[~df_res["Status"].isin(["valid","invalid"])]
    total = len(df_res)

    st.markdown("---")

    c1, c2, c3, c4 = st.columns(4)
    with c1: st.markdown(f'<div class="stat-card total"><div class="stat-number">{total}</div><div class="stat-label">Totaal</div></div>', unsafe_allow_html=True)
    with c2: st.markdown(f'<div class="stat-card valid"><div class="stat-number">{len(valid_df)}</div><div class="stat-label">Geldig</div></div>', unsafe_allow_html=True)
    with c3: st.markdown(f'<div class="stat-card invalid"><div class="stat-number">{len(invalid_df)}</div><div class="stat-label">Ongeldig</div></div>', unsafe_allow_html=True)
    with c4: st.markdown(f'<div class="stat-card error"><div class="stat-number">{len(error_df)}</div><div class="stat-label">Fout</div></div>', unsafe_allow_html=True)

    st.markdown("<br>", unsafe_allow_html=True)

    col_dl, col_reset = st.columns([3,1])
    with col_dl:
        st.download_button(
            label="📥 Download rapport (.xlsx)",
            data=to_excel_bytes(df_res),
            file_name=f"btw_validatie_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )
    with col_reset:
        if st.button("Opnieuw", use_container_width=True):
            st.session_state.results = None
            st.rerun()

    st.markdown("<br>", unsafe_allow_html=True)

    def style_status(val):
        return {"valid":"color:#22c55e;font-weight:700","invalid":"color:#E30613;font-weight:700"}.get(val,"color:#f59e0b;font-weight:700")

    t1, t2, t3, t4 = st.tabs([
        f"Alle resultaten ({total})",
        f"Geldig ({len(valid_df)})",
        f"Ongeldig ({len(invalid_df)})",
        f"Fout ({len(error_df)})"
    ])
    with t1: st.dataframe(df_res.style.applymap(style_status, subset=["Status"]), use_container_width=True, hide_index=True, height=400)
    with t2: st.dataframe(valid_df, use_container_width=True, hide_index=True, height=400) if len(valid_df) else st.info("Geen geldige BTW-nummers.")
    with t3: st.dataframe(invalid_df, use_container_width=True, hide_index=True, height=400) if len(invalid_df) else st.success("Geen ongeldige BTW-nummers!")
    with t4: st.dataframe(error_df, use_container_width=True, hide_index=True, height=400) if len(error_df) else st.success("Geen fouten.")

st.markdown("---")
st.markdown("<p style='color:#333;font-size:0.75rem;text-align:center;'>Mammoet Data Migration Team · VIES API (European Commission) · SAP ECC → S/4HANA</p>", unsafe_allow_html=True)
