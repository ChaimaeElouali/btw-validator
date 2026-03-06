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

# ─────────────────────────────────────────────
# Pagina configuratie
# ─────────────────────────────────────────────
st.set_page_config(
    page_title="BTW Validator | Mammoet",
    page_icon="🔍",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# ─────────────────────────────────────────────
# Styling
# ─────────────────────────────────────────────
st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Barlow:wght@400;500;600;700;800&family=Barlow+Condensed:wght@600;700;800&display=swap');

    html, body, [class*="css"] { font-family: 'Barlow', sans-serif; }
    .stApp { background-color: #0a0a0a; color: #f0f0f0; }

    .hero {
        background: linear-gradient(135deg, #1a1a1a 0%, #0d0d0d 100%);
        border-bottom: 3px solid #E30613;
        padding: 2.5rem 3rem 2rem;
        margin: -1rem -1rem 2rem;
    }
    .hero-title {
        font-family: 'Barlow Condensed', sans-serif;
        font-size: 3rem; font-weight: 800; color: #ffffff;
        letter-spacing: -0.5px; margin: 0; line-height: 1;
    }
    .hero-title span { color: #E30613; }
    .hero-sub {
        color: #888; font-size: 0.95rem; margin-top: 0.4rem;
        font-weight: 500; letter-spacing: 0.5px; text-transform: uppercase;
    }
    .stat-card {
        background: #161616; border: 1px solid #2a2a2a;
        border-radius: 8px; padding: 1.5rem; text-align: center;
    }
    .stat-card.valid   { border-top: 3px solid #22c55e; }
    .stat-card.invalid { border-top: 3px solid #E30613; }
    .stat-card.error   { border-top: 3px solid #f59e0b; }
    .stat-card.total   { border-top: 3px solid #3b82f6; }
    .stat-number {
        font-family: 'Barlow Condensed', sans-serif;
        font-size: 3rem; font-weight: 800; line-height: 1;
    }
    .stat-label { font-size: 0.8rem; font-weight: 600; text-transform: uppercase; letter-spacing: 1px; color: #888; margin-top: 0.3rem; }
    .valid   .stat-number { color: #22c55e; }
    .invalid .stat-number { color: #E30613; }
    .error   .stat-number { color: #f59e0b; }
    .total   .stat-number { color: #3b82f6; }
    .stProgress > div > div { background-color: #E30613 !important; }
    .stButton > button {
        background: #E30613 !important; color: white !important;
        border: none !important; font-family: 'Barlow', sans-serif !important;
        font-weight: 700 !important; font-size: 1rem !important;
        padding: 0.75rem 2rem !important; border-radius: 6px !important;
    }
    .stButton > button:hover { opacity: 0.85 !important; }
    .stDownloadButton > button {
        background: #1a1a1a !important; color: #f0f0f0 !important;
        border: 1px solid #333 !important; font-family: 'Barlow', sans-serif !important;
        font-weight: 600 !important;
    }
    .stTabs [data-baseweb="tab-list"] { background: #111; border-bottom: 1px solid #2a2a2a; }
    .stTabs [data-baseweb="tab"] {
        color: #888 !important; font-family: 'Barlow', sans-serif !important;
        font-weight: 600 !important; font-size: 0.9rem !important;
        padding: 0.75rem 1.5rem !important;
    }
    .stTabs [aria-selected="true"] {
        color: #ffffff !important; border-bottom: 2px solid #E30613 !important;
        background: transparent !important;
    }
    [data-testid="stFileUploader"] {
        background: #111; border: 1px solid #2a2a2a; border-radius: 8px; padding: 1rem;
    }
    hr { border-color: #2a2a2a !important; }
    .info-box {
        background: #161616; border: 1px solid #2a2a2a;
        border-left: 3px solid #3b82f6; border-radius: 6px;
        padding: 1rem 1.25rem; font-size: 0.9rem; color: #bbb; margin-bottom: 1rem;
    }

    /* SAP Window styling */
    .sap-window {
        background: #f0f0f0; border: 2px solid #808080; border-radius: 3px;
        font-family: 'Courier New', monospace; color: #000;
        overflow: hidden; box-shadow: 3px 3px 8px rgba(0,0,0,0.5);
    }
    .sap-titlebar {
        background: linear-gradient(180deg, #0050a0 0%, #003878 100%);
        color: white; padding: 4px 8px; font-size: 12px; font-weight: bold;
        font-family: 'Courier New', monospace; display: flex; align-items: center;
    }
    .sap-titlebar-btns { margin-left: auto; display: flex; gap: 2px; }
    .sap-titlebar-btn {
        background: #c0c0c0; color: #000; border: 1px solid #808080;
        padding: 0 5px; font-size: 10px; cursor: default; border-radius: 1px;
    }
    .sap-tabs { background: #d4d0c8; border-bottom: 2px solid #808080; display: flex; }
    .sap-tab {
        padding: 3px 10px; font-size: 11px; font-family: 'Courier New', monospace;
        border-right: 1px solid #808080; cursor: default; color: #000;
    }
    .sap-tab.active {
        background: #f0f0f0; border-bottom: 2px solid #f0f0f0;
        margin-bottom: -2px; font-weight: bold;
    }
    .sap-grid-header {
        background: #c0c0c0; border-bottom: 1px solid #808080;
        padding: 3px 8px; font-size: 11px; font-weight: bold;
        color: #000; font-family: 'Courier New', monospace;
    }
    .sap-grid-body { background: white; }
    .sap-grid-row { display: flex; border-bottom: 1px solid #e0e0e0; }
    .sap-grid-row.active-row { background: #c5d9f1; }
    .sap-row-num {
        width: 30px; min-width: 30px; background: #e0e0e0;
        border-right: 1px solid #ccc; text-align: center;
        font-size: 11px; color: #555; padding: 2px 0; line-height: 20px;
        font-family: 'Courier New', monospace;
    }
    .sap-row-val {
        flex: 1; padding: 2px 8px; font-size: 12px;
        font-family: 'Courier New', monospace; color: #000080; line-height: 20px;
    }
    .sap-row-val.empty { color: #bbb; }
    .sap-toolbar {
        background: #d4d0c8; border-top: 1px solid #808080;
        padding: 3px 6px; display: flex; gap: 4px; flex-wrap: wrap;
    }
    .sap-btn {
        background: #d4d0c8; border: 1px solid #808080; padding: 1px 8px;
        font-size: 11px; font-family: 'Courier New', monospace;
        cursor: default; color: #000; border-radius: 1px;
    }
    .sap-statusbar {
        background: #d4d0c8; border-top: 1px solid #808080;
        padding: 2px 8px; font-size: 11px; color: #000;
        font-family: 'Courier New', monospace;
    }

    #MainMenu, footer, header { visibility: hidden; }
    .block-container { padding-top: 0 !important; }
</style>
""", unsafe_allow_html=True)

# ─────────────────────────────────────────────
# VIES API functies
# ─────────────────────────────────────────────
VIES_URL = "https://ec.europa.eu/taxation_customs/vies/rest-api/ms/{country_code}/vat/{vat_number}"
TIMEOUT = 10
MAX_RETRIES = 3

EU_COUNTRY_CODES = {
    "AT", "BE", "BG", "CY", "CZ", "DE", "DK", "EE", "EL", "ES",
    "FI", "FR", "HR", "HU", "IE", "IT", "LT", "LU", "LV", "MT",
    "NL", "PL", "PT", "RO", "SE", "SI", "SK", "XI"
}


def parse_vat(raw: str):
    if not isinstance(raw, str):
        raw = str(raw)
    cleaned = re.sub(r"[\s.\-]", "", raw.strip().upper())
    if len(cleaned) < 3:
        return "", ""
    return cleaned[:2], cleaned[2:]


def check_vat(country_code: str, vat_number: str) -> dict:
    if country_code not in EU_COUNTRY_CODES:
        return {"status": "error", "company_name": "—", "company_address": "—",
                "message": f"Landcode '{country_code}' niet ondersteund door VIES"}
    url = VIES_URL.format(country_code=country_code, vat_number=vat_number)
    for attempt in range(1, MAX_RETRIES + 1):
        try:
            r = requests.get(url, timeout=TIMEOUT)
            if r.status_code == 200:
                data = r.json()
                valid = data.get("isValid", False)
                return {
                    "status": "valid" if valid else "invalid",
                    "company_name": data.get("name") or "—",
                    "company_address": data.get("address") or "—",
                    "message": "Geldig" if valid else "Ongeldig volgens VIES"
                }
            elif r.status_code == 404:
                return {"status": "invalid", "company_name": "—", "company_address": "—",
                        "message": "Niet gevonden in VIES"}
            elif r.status_code in (429, 503, 504):
                time.sleep(2 ** attempt)
                continue
            else:
                return {"status": "error", "company_name": "—", "company_address": "—",
                        "message": f"HTTP {r.status_code}"}
        except requests.exceptions.Timeout:
            if attempt < MAX_RETRIES:
                time.sleep(2)
                continue
            return {"status": "error", "company_name": "—", "company_address": "—",
                    "message": "Timeout — VIES niet bereikbaar"}
        except Exception as e:
            return {"status": "error", "company_name": "—", "company_address": "—",
                    "message": str(e)}
    return {"status": "error", "company_name": "—", "company_address": "—",
            "message": "VIES niet beschikbaar"}


def detect_vat_column(df: pd.DataFrame) -> str:
    keywords = ["vat", "btw", "tax"]
    for col in df.columns:
        if any(kw in col.lower() for kw in keywords):
            return col
    return df.columns[0]


def to_excel_bytes(df: pd.DataFrame) -> bytes:
    output = io.BytesIO()
    valid_df   = df[df["Status"] == "valid"]
    invalid_df = df[df["Status"] == "invalid"]
    error_df   = df[~df["Status"].isin(["valid", "invalid"])]
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        summary = pd.DataFrame({
            "Categorie": ["Geldig", "Ongeldig", "Fout/Onbekend", "Totaal"],
            "Aantal": [len(valid_df), len(invalid_df), len(error_df), len(df)],
            "Percentage": [
                f"{len(valid_df)/len(df)*100:.1f}%" if len(df) else "0%",
                f"{len(invalid_df)/len(df)*100:.1f}%" if len(df) else "0%",
                f"{len(error_df)/len(df)*100:.1f}%" if len(df) else "0%",
                "100%"
            ]
        })
        summary.to_excel(writer, sheet_name="Samenvatting", index=False)
        ws = writer.sheets["Samenvatting"]
        ws.cell(row=7, column=1, value=f"Rapport: {datetime.now().strftime('%d-%m-%Y %H:%M')}")
        ws.cell(row=8, column=1, value="Bron: VIES API — European Commission")
        df.to_excel(writer, sheet_name="Alle resultaten", index=False)
        if len(valid_df):   valid_df.to_excel(writer, sheet_name="Geldig", index=False)
        if len(invalid_df): invalid_df.to_excel(writer, sheet_name="Ongeldig", index=False)
        if len(error_df):   error_df.to_excel(writer, sheet_name="Fouten", index=False)
    return output.getvalue()


def run_validation(df_input: pd.DataFrame, vat_col: str):
    st.markdown("---")
    st.markdown("#### Validatie bezig...")
    progress_bar = st.progress(0)
    status_text  = st.empty()
    results = []
    total = len(df_input)
    for i, (_, row) in enumerate(df_input.iterrows()):
        raw = str(row[vat_col]).strip()
        country_code, vat_number = parse_vat(raw)
        status_text.markdown(
            f"<small style='color:#888'>Verwerken {i+1}/{total}: <code>{raw}</code></small>",
            unsafe_allow_html=True
        )
        if not country_code or not vat_number:
            result = {"status": "invalid", "company_name": "—",
                      "company_address": "—", "message": "Ongeldig formaat"}
        else:
            result = check_vat(country_code, vat_number)
        row_data = {"VAT Input": raw, "Land": country_code, "BTW Nummer": vat_number}
        for c in df_input.columns:
            if c != vat_col:
                row_data[c] = row[c]
        row_data.update({
            "Status": result["status"],
            "Bedrijfsnaam (VIES)": result["company_name"],
            "Adres (VIES)": result["company_address"],
            "Melding": result["message"]
        })
        results.append(row_data)
        progress_bar.progress((i + 1) / total)
        time.sleep(0.4)
    status_text.empty()
    st.session_state.results = pd.DataFrame(results)
    st.rerun()


# ─────────────────────────────────────────────
# Header
# ─────────────────────────────────────────────
st.markdown("""
<div class="hero">
    <div class="hero-title">BTW <span>VALIDATOR</span></div>
    <div class="hero-sub">Mammoet Data Migration · SAP ECC → S/4HANA · VIES API</div>
</div>
""", unsafe_allow_html=True)

if "results" not in st.session_state:
    st.session_state.results = None

# ─────────────────────────────────────────────
# Input tabbladen
# ─────────────────────────────────────────────
tab_file, tab_paste = st.tabs(["📁 Bestand uploaden", "📋 BTW-nummers plakken (SAP-stijl)"])

# ── Tab 1: Bestand uploaden ──
with tab_file:
    col_up, col_info = st.columns([2, 1], gap="large")
    with col_up:
        st.markdown("#### Bestand uploaden")
        uploaded = st.file_uploader(
            "CSV of Excel",
            type=["csv", "xlsx", "xls"],
            label_visibility="collapsed"
        )
    with col_info:
        st.markdown("#### Vereisten")
        st.markdown("""
<div class="info-box">
Minimaal één kolom met BTW-nummers.<br><br>
<strong>Herkende kolomnamen:</strong><br>
<code>vat_number</code>, <code>btw</code>, <code>tax</code><br><br>
<strong>Formaten:</strong> CSV, XLSX<br>
<strong>Voorbeeld:</strong> NL800336808B01
</div>""", unsafe_allow_html=True)

    if uploaded:
        try:
            if uploaded.name.endswith(".csv"):
                df_file = pd.read_csv(uploaded, dtype=str)
            else:
                df_file = pd.read_excel(uploaded, dtype=str)
            vat_col = detect_vat_column(df_file)
            df_file = df_file.fillna("")
            st.markdown("---")
            col_prev, col_btn = st.columns([3, 1], gap="large")
            with col_prev:
                st.markdown(f"**Preview** — {len(df_file)} rijen · BTW-kolom: `{vat_col}`")
                st.dataframe(df_file.head(5), use_container_width=True, hide_index=True)
            with col_btn:
                st.markdown("<br><br>", unsafe_allow_html=True)
                if st.button(f"🔍 Valideer {len(df_file)} nummers", use_container_width=True, key="btn_file"):
                    run_validation(df_file, vat_col)
        except Exception as e:
            st.error(f"Fout bij het inlezen: {e}")

# ── Tab 2: SAP-stijl plakken ──
with tab_paste:

    # Style: plak tekstvak IN de SAP look
    st.markdown("""
<style>
  div[data-testid="stTextArea"] textarea {
      background: white !important;
      border: none !important;
      border-radius: 0 !important;
      font-family: Courier New, monospace !important;
      font-size: 12px !important;
      color: #000080 !important;
      padding: 4px 6px !important;
      resize: none !important;
      box-shadow: none !important;
  }
  div[data-testid="stTextArea"] > label { display: none; }
  div[data-testid="stTextArea"] > div { border: none !important; padding: 0 !important; }
</style>
""", unsafe_allow_html=True)

    # SAP tabbladen + kolomkop
    st.markdown("""
<div style="background:#dedad0; border:2px solid #aaa; border-bottom:none; max-width:520px;">
  <div style="display:flex; gap:2px; padding:4px 4px 0; background:#ccc8b8;">
    <div style="padding:3px 12px; font-size:11px; font-family:Arial; font-weight:bold;
                background:#dedad0; border:1px solid #999; border-bottom:2px solid #dedad0;
                border-radius:3px 3px 0 0; margin-bottom:-1px; cursor:default;">Select Single Values</div>
    <div style="padding:3px 12px; font-size:11px; font-family:Arial;
                background:#c4c0b0; border:1px solid #999; border-radius:3px 3px 0 0; cursor:default;">Select Ranges</div>
    <div style="padding:3px 12px; font-size:11px; font-family:Arial;
                background:#c4c0b0; border:1px solid #999; border-radius:3px 3px 0 0; cursor:default;">Exclude Single Values</div>
    <div style="padding:3px 12px; font-size:11px; font-family:Arial;
                background:#c4c0b0; border:1px solid #999; border-radius:3px 3px 0 0; cursor:default;">Exclude Ranges</div>
  </div>
  <div style="background:#ccc8b8; border-top:1px solid #999; padding:2px 6px;
              font-size:11px; font-weight:bold; font-family:Arial;
              display:flex; align-items:center; gap:8px;">
    <span style="color:#333;">&#11044;</span>
    <span>Single value</span>
  </div>
</div>
""", unsafe_allow_html=True)

    # Tekstvak — zit visueel direct onder de SAP-header
    pasted = st.text_area(
        "input",
        height=200,
        placeholder="Plak hier BTW-nummers, één per regel...",
        key="paste_input",
        label_visibility="collapsed"
    )

    lines_val = [l.strip() for l in (pasted or "").splitlines() if l.strip()]

    # Onderbalk met Paste All / Delete All
    st.markdown(f"""
<div style="background:#dedad0; border:2px solid #aaa; border-top:none; max-width:520px;
            padding:4px 6px; display:flex; gap:6px; align-items:center; font-family:Arial; font-size:11px;">
  <div style="background:#ccc8b8; border:1px solid #888; padding:2px 10px; cursor:default; border-radius:2px;">
    Paste All
  </div>
  <div style="background:#ccc8b8; border:1px solid #888; padding:2px 10px; cursor:default; border-radius:2px;">
    Delete All
  </div>
  <span style="margin-left:auto; color:#555;">Rows: <b>{len(lines_val)}</b></span>
</div>
""", unsafe_allow_html=True)

    # Delete All werkt via Streamlit knop (verborgen achter de HTML-look)
    col_del, col_val_btn = st.columns([1, 3])
    with col_del:
        if st.button("Delete All", key="delete_all"):
            st.session_state["paste_input"] = ""
            st.rerun()
    with col_val_btn:
        if lines_val:
            if st.button(f"Valideer {len(lines_val)} BTW-nummers", key="btn_paste"):
                import pandas as pd
                df_paste = pd.DataFrame({"vat_number": lines_val})
                run_validation(df_paste, "vat_number")

# ─────────────────────────────────────────────
# Resultaten
# ─────────────────────────────────────────────
if st.session_state.results is not None:
    df_res = st.session_state.results
    valid_df   = df_res[df_res["Status"] == "valid"]
    invalid_df = df_res[df_res["Status"] == "invalid"]
    error_df   = df_res[~df_res["Status"].isin(["valid", "invalid"])]
    total      = len(df_res)

    st.markdown("---")
    st.markdown("#### 📊 Resultaten")

    c1, c2, c3, c4 = st.columns(4)
    with c1:
        st.markdown(f'<div class="stat-card total"><div class="stat-number">{total}</div><div class="stat-label">Totaal verwerkt</div></div>', unsafe_allow_html=True)
    with c2:
        st.markdown(f'<div class="stat-card valid"><div class="stat-number">{len(valid_df)}</div><div class="stat-label">Geldig</div></div>', unsafe_allow_html=True)
    with c3:
        st.markdown(f'<div class="stat-card invalid"><div class="stat-number">{len(invalid_df)}</div><div class="stat-label">Ongeldig</div></div>', unsafe_allow_html=True)
    with c4:
        st.markdown(f'<div class="stat-card error"><div class="stat-number">{len(error_df)}</div><div class="stat-label">Fout/Onbekend</div></div>', unsafe_allow_html=True)

    st.markdown("<br>", unsafe_allow_html=True)

    col_dl, col_reset = st.columns([2, 1])
    with col_dl:
        excel_bytes = to_excel_bytes(df_res)
        timestamp = datetime.now().strftime("%Y%m%d_%H%M")
        st.download_button(
            label="📥 Download rapport (.xlsx)",
            data=excel_bytes,
            file_name=f"btw_validatie_{timestamp}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )
    with col_reset:
        if st.button("🔄 Opnieuw beginnen", use_container_width=True):
            st.session_state.results = None
            st.rerun()

    st.markdown("<br>", unsafe_allow_html=True)

    res_tab1, res_tab2, res_tab3, res_tab4 = st.tabs([
        f"📋 Alle resultaten ({total})",
        f"✅ Geldig ({len(valid_df)})",
        f"❌ Ongeldig ({len(invalid_df)})",
        f"⚠️ Fout/Onbekend ({len(error_df)})"
    ])

    def style_status(val):
        return {
            "valid":   "color: #22c55e; font-weight: 700;",
            "invalid": "color: #E30613; font-weight: 700;",
        }.get(val, "color: #f59e0b; font-weight: 700;")

    with res_tab1:
        st.dataframe(df_res.style.applymap(style_status, subset=["Status"]),
                     use_container_width=True, hide_index=True, height=400)
    with res_tab2:
        st.dataframe(valid_df, use_container_width=True, hide_index=True, height=400) if len(valid_df) else st.info("Geen geldige BTW-nummers.")
    with res_tab3:
        st.dataframe(invalid_df, use_container_width=True, hide_index=True, height=400) if len(invalid_df) else st.success("Geen ongeldige BTW-nummers!")
    with res_tab4:
        st.dataframe(error_df, use_container_width=True, hide_index=True, height=400) if len(error_df) else st.success("Geen fouten.")

# Footer
st.markdown("---")
st.markdown(
    "<p style='color:#444; font-size:0.8rem; text-align:center;'>"
    "Mammoet Data Migration Team · Validatie via VIES API (European Commission) · "
    "Alleen geldige BTW-nummers migreren naar SAP S/4HANA</p>",
    unsafe_allow_html=True
)
