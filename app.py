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
    page_title="BTW Validator | Mammoet",
    page_icon="🔍",
    layout="wide",
    initial_sidebar_state="collapsed"
)

st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Barlow:wght@400;500;600;700;800&family=Barlow+Condensed:wght@600;700;800&display=swap');

    html, body, [class*="css"] { font-family: 'Barlow', sans-serif; }
    .stApp { background-color: #0f0f0f; color: #f0f0f0; }
    .block-container { padding-top: 0 !important; max-width: 900px; margin: 0 auto; }

    /* ── Hero ── */
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
        color: #555; font-size: 0.78rem; margin-top: 0.5rem;
        font-weight: 600; letter-spacing: 2px; text-transform: uppercase;
    }

    /* ── Input cards ── */
    .input-grid {
        display: grid;
        grid-template-columns: 1fr 40px 1fr;
        gap: 0;
        align-items: stretch;
        margin-bottom: 1.5rem;
    }
    .input-card {
        background: #1a1a1a;
        border: 1px solid #2a2a2a;
        border-radius: 10px;
        padding: 1.75rem;
        display: flex;
        flex-direction: column;
        gap: 0.75rem;
    }
    .input-card:hover { border-color: #3a3a3a; }
    .card-label {
        font-size: 0.7rem; font-weight: 700; text-transform: uppercase;
        letter-spacing: 2px; color: #E30613; margin-bottom: 0.25rem;
    }
    .card-desc { font-size: 0.85rem; color: #666; margin: 0; }
    .or-col {
        display: flex;
        flex-direction: column;
        align-items: center;
        justify-content: center;
        gap: 8px;
        padding: 1rem 0;
    }
    .or-line { flex: 1; width: 1px; background: #2a2a2a; }
    .or-text {
        font-size: 0.7rem; font-weight: 700; letter-spacing: 2px;
        text-transform: uppercase; color: #3a3a3a;
        padding: 4px 0;
    }

    /* ── File uploader override ── */
    [data-testid="stFileUploader"] {
        background: #111 !important;
    }
    [data-testid="stFileUploader"] > div {
        background: #111 !important;
        border: 2px dashed #2a2a2a !important;
        border-radius: 8px !important;
        padding: 1.5rem !important;
        transition: border-color 0.2s;
    }
    [data-testid="stFileUploader"] > div:hover {
        border-color: #E30613 !important;
    }
    [data-testid="stFileUploader"] button {
        background: #2a2a2a !important;
        color: #f0f0f0 !important;
        border: none !important;
        border-radius: 4px !important;
    }
    [data-testid="stFileUploader"] p,
    [data-testid="stFileUploader"] small { color: #555 !important; }

    /* ── Textarea ── */
    textarea {
        background: #111 !important;
        color: #f0f0f0 !important;
        border: 2px solid #2a2a2a !important;
        border-radius: 8px !important;
        font-family: 'Barlow', sans-serif !important;
        font-size: 0.9rem !important;
        transition: border-color 0.2s;
    }
    textarea:focus { border-color: #E30613 !important; box-shadow: none !important; }
    textarea::placeholder { color: #444 !important; }

    /* ── Validate button ── */
    .stButton > button {
        background: #E30613 !important;
        color: white !important;
        border: none !important;
        font-family: 'Barlow', sans-serif !important;
        font-weight: 700 !important;
        font-size: 1rem !important;
        padding: 0.75rem 2.5rem !important;
        border-radius: 6px !important;
        letter-spacing: 0.5px;
        transition: background 0.2s !important;
    }
    .stButton > button:hover { background: #c0050f !important; }

    .stDownloadButton > button {
        background: #1a1a1a !important;
        color: #f0f0f0 !important;
        border: 1px solid #333 !important;
        font-family: 'Barlow', sans-serif !important;
        font-weight: 700 !important;
        border-radius: 6px !important;
    }

    /* ── Stat cards ── */
    .stat-row { display: grid; grid-template-columns: repeat(4,1fr); gap: 12px; margin-bottom: 1.5rem; }
    .stat-card {
        background: #1a1a1a; border-radius: 8px; padding: 1.25rem 1rem;
        text-align: center; border: 1px solid #2a2a2a;
    }
    .stat-card.total   { border-top: 3px solid #555; }
    .stat-card.valid   { border-top: 3px solid #22c55e; }
    .stat-card.invalid { border-top: 3px solid #E30613; }
    .stat-card.error   { border-top: 3px solid #f59e0b; }
    .stat-num {
        font-family: 'Barlow Condensed', sans-serif;
        font-size: 2.5rem; font-weight: 800; line-height: 1;
    }
    .stat-lbl { font-size: 0.7rem; font-weight: 700; text-transform: uppercase; letter-spacing: 1px; color: #555; margin-top: 4px; }
    .total   .stat-num { color: #888; }
    .valid   .stat-num { color: #22c55e; }
    .invalid .stat-num { color: #E30613; }
    .error   .stat-num { color: #f59e0b; }

    /* ── Tabs ── */
    .stTabs [data-baseweb="tab-list"] { background: #1a1a1a; border-bottom: 1px solid #2a2a2a; }
    .stTabs [data-baseweb="tab"] {
        color: #555 !important; font-family: 'Barlow', sans-serif !important;
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
                return {"status":"invalid","company_name":"—","company_address":"—","message":"Niet gevonden"}
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
            f"<small style='color:#555'>{i+1}/{total} — <code style='color:#888'>{raw}</code></small>",
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

# ─── INPUT: twee kaarten naast elkaar ───────
col_left, col_or, col_right = st.columns([10, 1, 10], gap="small")

df_input = None
vat_col  = "vat_number"

with col_left:
    st.markdown("""
<div class="input-card">
  <div class="card-label">📁 Bestand uploaden</div>
  <p class="card-desc">CSV of Excel met een kolom BTW-nummers</p>
</div>
""", unsafe_allow_html=True)
    uploaded = st.file_uploader("upload", type=["csv","xlsx","xls"], label_visibility="collapsed")
    if uploaded:
        try:
            df_input = pd.read_csv(uploaded, dtype=str) if uploaded.name.endswith(".csv") else pd.read_excel(uploaded, dtype=str)
            vat_col = detect_vat_column(df_input)
            df_input = df_input.fillna("")
            st.markdown(f"<small style='color:#555'>✅ {len(df_input)} rijen · kolom: <code style='color:#888'>{vat_col}</code></small>", unsafe_allow_html=True)
        except Exception as e:
            st.error(f"Fout: {e}")
            df_input = None

with col_or:
    st.markdown("""
<div class="or-col">
  <div class="or-line"></div>
  <div class="or-text">of</div>
  <div class="or-line"></div>
</div>
""", unsafe_allow_html=True)

with col_right:
    st.markdown("""
<div class="input-card">
  <div class="card-label">📋 BTW-nummers plakken</div>
  <p class="card-desc">Plak nummers uit SAP of Excel, één per regel</p>
</div>
""", unsafe_allow_html=True)
    pasted = st.text_area("paste", height=150,
        placeholder="NL800336808B01\nDE811115329\nBE0999999999\n...",
        key="paste_input", label_visibility="collapsed")
    if not uploaded and pasted and pasted.strip():
        lines = [l.strip() for l in pasted.splitlines() if l.strip()]
        if lines:
            df_input = pd.DataFrame({"vat_number": lines})
            vat_col = "vat_number"
            st.markdown(f"<small style='color:#555'>✅ {len(lines)} BTW-nummer(s)</small>", unsafe_allow_html=True)

# ─── VALIDEER KNOP ──────────────────────────
st.markdown("<br>", unsafe_allow_html=True)
if df_input is not None and len(df_input) > 0:
    col_btn, col_empty = st.columns([1, 3])
    with col_btn:
        if st.button(f"Valideer {len(df_input)} nummers →", use_container_width=True):
            run_validation(df_input, vat_col)
else:
    st.markdown("<small style='color:#333'>Upload een bestand of plak BTW-nummers om te beginnen.</small>", unsafe_allow_html=True)

# ─── RESULTATEN ─────────────────────────────
if st.session_state.results is not None:
    df_res = st.session_state.results
    valid_df   = df_res[df_res["Status"]=="valid"]
    invalid_df = df_res[df_res["Status"]=="invalid"]
    error_df   = df_res[~df_res["Status"].isin(["valid","invalid"])]
    total = len(df_res)

    st.markdown("---")

    st.markdown(f"""
<div class="stat-row">
  <div class="stat-card total"><div class="stat-num">{total}</div><div class="stat-lbl">Totaal</div></div>
  <div class="stat-card valid"><div class="stat-num">{len(valid_df)}</div><div class="stat-lbl">Geldig</div></div>
  <div class="stat-card invalid"><div class="stat-num">{len(invalid_df)}</div><div class="stat-lbl">Ongeldig</div></div>
  <div class="stat-card error"><div class="stat-num">{len(error_df)}</div><div class="stat-lbl">Fout</div></div>
</div>
""", unsafe_allow_html=True)

    col_dl, col_reset, col_empty = st.columns([3, 1, 4])
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
st.markdown("<p style='color:#2a2a2a;font-size:0.75rem;text-align:center;'>Mammoet Data Migration Team · VIES API (European Commission) · SAP ECC → S/4HANA</p>", unsafe_allow_html=True)
