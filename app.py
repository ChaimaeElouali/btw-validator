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
    .stApp { background-color: #f0f0f0 !important; }
    /* Remove white card/border around page */
    .stApp > div { background: transparent !important; }
    .block-container {
        padding: 0 2.5rem 1.5rem 2.5rem !important;
        max-width: 100% !important;
    }

    /* ── Hero: full width dark ── */
    .hero {
        background: #0f0f0f;
        border-bottom: 3px solid #E30613;
        padding: 2rem 2.5rem 1.6rem;
        margin: -1.5rem -2.5rem 1.5rem -2.5rem;
    }
    .hero-title {
        font-family: 'Barlow Condensed', sans-serif;
        font-size: 2.8rem; font-weight: 800; color: #ffffff;
        letter-spacing: -1px; margin: 0; line-height: 1;
    }
    .hero-title span { color: #E30613; }
    .hero-sub {
        color: #aaaaaa; font-size: 1rem; margin-top: 0.5rem;
        font-weight: 400; white-space: nowrap;
    }

    /* ── Content padding ── */
    .content-pad { padding: 1.5rem 2.5rem; }

    /* ── Section labels ── */
    .section-label {
        font-size: 0.75rem; font-weight: 700; text-transform: uppercase;
        letter-spacing: 2px; color: #555555;
        margin-bottom: 0.4rem; display: block;
    }
    .section-hint {
        font-size: 0.95rem; color: #444444; margin-bottom: 0.9rem;
        display: block; font-weight: 400; line-height: 1.5;
    }

    /* ── Upload zone ── */
    [data-testid="stFileUploader"] { background: transparent !important; }
    [data-testid="stFileUploader"] > div {
        background: #ffffff !important;
        border: 2px dashed #cccccc !important;
        border-radius: 8px !important;
        transition: border-color 0.2s;
    }
    [data-testid="stFileUploader"] > div:hover { border-color: #E30613 !important; }
    [data-testid="stFileUploader"] section { background: #ffffff !important; border: none !important; }
    [data-testid="stFileUploader"] button {
        background: #e8e8e8 !important; color: #222222 !important;
        border: 1px solid #cccccc !important; border-radius: 4px !important;
        font-family: 'Barlow', sans-serif !important; font-weight: 600 !important;
        font-size: 0.95rem !important;
    }
    [data-testid="stFileUploader"] p,
    [data-testid="stFileUploaderDropzoneInstructions"] div span {
        color: #666666 !important; font-size: 0.95rem !important;
    }
    [data-testid="stFileUploader"] small { color: #888888 !important; }
    [data-testid="stFileUploader"] [data-testid="stMarkdownContainer"] p,
    [data-testid="stFileUploader"] span {
        color: #111111 !important; font-size: 1rem !important; font-weight: 600 !important;
    }

    /* ── All buttons: red ── */
    .stButton > button {
        background: #E30613 !important; color: #ffffff !important;
        border: none !important; font-family: 'Barlow', sans-serif !important;
        font-weight: 700 !important; font-size: 0.9rem !important;
        padding: 0.5rem 1.2rem !important; border-radius: 5px !important;
        white-space: nowrap !important;
    }
    .stButton > button:hover { background: #c0050f !important; }
    .stButton > button:disabled { background: #cccccc !important; color: #888888 !important; }

    /* ── Download buttons: small, outlined ── */
    .stDownloadButton > button {
        background: transparent !important; color: #222222 !important;
        border: 1px solid #aaaaaa !important;
        font-family: 'Barlow', sans-serif !important;
        font-weight: 600 !important; border-radius: 5px !important;
        font-size: 0.75rem !important; padding: 0.3rem 0.7rem !important;
    }
    .stDownloadButton > button:hover {
        background: #222222 !important; color: #ffffff !important;
        border-color: #222222 !important;
    }

    /* ── Divider ── */
    .vert-line {
        width: 1px; background: #cccccc;
        margin: 0 auto; height: 100%; min-height: 400px;
    }

    /* ── Tabs: color per type ── */
    .stTabs [data-baseweb="tab-list"] {
        background: #f0f0f0;
        border-bottom: 2px solid #dddddd;
    }
    .stTabs [data-baseweb="tab"] {
        color: #888888 !important; font-family: 'Barlow', sans-serif !important;
        font-weight: 600 !important; padding: 0.7rem 1.2rem !important;
        font-size: 0.95rem !important;
    }
    /* Tab 1: default */
    .stTabs [data-baseweb="tab"]:nth-child(1) { color: #333333 !important; }
    /* Tab 2: Valid — green */
    .stTabs [data-baseweb="tab"]:nth-child(2) { color: #2d7a4f !important; }
    /* Tab 3: Invalid — red */
    .stTabs [data-baseweb="tab"]:nth-child(3) { color: #E30613 !important; }
    /* Tab 4: Error — orange */
    .stTabs [data-baseweb="tab"]:nth-child(4) { color: #b05e00 !important; }
    .stTabs [aria-selected="true"] {
        font-weight: 700 !important;
        border-bottom: 3px solid #E30613 !important;
        background: transparent !important;
    }

    /* ── Feedback text ── */
    .feedback-text {
        font-size: 1rem; color: #222222; font-weight: 500;
        margin-top: 0.5rem; display: block;
    }

    .stProgress > div > div { background-color: #E30613 !important; }

    @media (max-width: 768px) {
        .hero { padding: 1.2rem 1rem 1rem; }
        .hero-title { font-size: 2rem; }
        .hero-sub { font-size: 0.85rem; white-space: normal; }
        .content-pad { padding: 1rem; }
    }

    hr { border-color: #dddddd !important; }
    #MainMenu, footer, header { visibility: hidden; }
    .stAppHeader, [data-testid="stHeader"] { display: none !important; }
    .stMainBlockContainer { padding-top: 0 !important; }
    .appview-container .main .block-container { padding-top: 0 !important; }
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
            "Category":["Valid","Invalid","Error / Unknown","Total"],
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

def to_xml_bytes(df):
    lines = []
    lines.append('<?xml version="1.0" encoding="UTF-8"?>')
    lines.append('<?mso-application progid="Excel.Sheet"?>')
    lines.append('<Workbook xmlns="urn:schemas-microsoft-com:office:spreadsheet"')
    lines.append(' xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet">')
    lines.append('<Worksheet ss:Name="All results"><Table>')
    lines.append('<Row>')
    for col in df.columns:
        lines.append(f'<Cell><Data ss:Type="String">{col}</Data></Cell>')
    lines.append('</Row>')
    for _, row in df.iterrows():
        lines.append('<Row>')
        for val in row:
            safe = str(val).replace('&','&amp;').replace('<','&lt;').replace('>','&gt;')
            lines.append(f'<Cell><Data ss:Type="String">{safe}</Data></Cell>')
        lines.append('</Row>')
    lines.append('</Table></Worksheet></Workbook>')
    return '\n'.join(lines).encode('utf-8')

def run_validation(df_input, vat_col):
    st.markdown("**Validating VAT numbers...**")
    progress_bar = st.progress(0)
    status_text  = st.empty()
    results = []
    total = len(df_input)
    for i, (_, row) in enumerate(df_input.iterrows()):
        raw = str(row[vat_col]).strip()
        country_code, vat_number = parse_vat(raw)
        status_text.markdown(
            f"<span style='color:#555555; font-size:0.95rem;'>Processing {i+1} of {total}: {raw}</span>",
            unsafe_allow_html=True)
        result = {"status":"invalid","company_name":"—","company_address":"—","message":"Invalid format"} \
            if not country_code or not vat_number else check_vat(country_code, vat_number)
        status_label = {"valid":"✓ valid","invalid":"✗ invalid","error":"! error"}.get(result["status"],"! error")
        row_data = {"VAT Input": raw, "Country": country_code, "VAT Number": vat_number}
        for c in df_input.columns:
            if c != vat_col: row_data[c] = row[c]
        row_data.update({
            "Status": result["status"], "Status Label": status_label,
            "Company (VIES)": result["company_name"],
            "Address (VIES)": result["company_address"],
            "Message": result["message"]
        })
        results.append(row_data)
        progress_bar.progress((i+1)/total)
        time.sleep(0.4)
    status_text.empty()
    st.session_state.results = pd.DataFrame(results)
    st.rerun()

# ─── HEADER (dark) ──────────────────────────
st.markdown("""
<div class="hero">
    <div class="hero-title">VAT <span>VALIDATOR</span></div>
    <div class="hero-sub">VAT numbers are validated in real-time via the official VIES API of the European Commission.</div>
</div>
""", unsafe_allow_html=True)

if "results" not in st.session_state:
    st.session_state.results = None

# ─── MAIN AREA ──────────────────────────────
st.markdown('<div class="content-pad">', unsafe_allow_html=True)

df_input = None
vat_col  = "vat_number"

# Responsive: on mobile stack vertically, on desktop side by side
col_left, col_mid, col_right = st.columns([5, 1, 6], gap="small")

# ── LEFT: Upload ────────────────────────────
with col_left:
    st.markdown('<span class="section-label">Upload a file</span>', unsafe_allow_html=True)
    st.markdown('<span class="section-hint">CSV, Excel (.xlsx / .xls) or Excel XML 2003 (.xml)<br>The file must contain a column with VAT numbers.</span>', unsafe_allow_html=True)

    uploaded = st.file_uploader("upload", type=["csv","xlsx","xls","xml"], label_visibility="collapsed")

    if uploaded:
        try:
            if uploaded.name.endswith(".csv"):
                df_input = pd.read_csv(uploaded, dtype=str)
            elif uploaded.name.endswith(".xml"):
                import xml.etree.ElementTree as ET
                import re as _re
                content_bytes = uploaded.read()
                root = ET.fromstring(content_bytes)
                ns = {}
                for elem in root.iter():
                    m = _re.match(r'\{([^}]+)\}', elem.tag)
                    if m:
                        ns['ss'] = m.group(1)
                        break
                def tag(name):
                    return f"{{{ns['ss']}}}{name}" if ns else name
                rows_data = []
                headers = None
                for worksheet in root.iter(tag('Worksheet')):
                    for table in worksheet.iter(tag('Table')):
                        for i, row in enumerate(table.iter(tag('Row'))):
                            cells = []
                            for cell in row.iter(tag('Cell')):
                                data = cell.find(tag('Data'))
                                cells.append(data.text if data is not None and data.text else "")
                            if i == 0:
                                headers = cells
                            else:
                                rows_data.append(cells)
                    break
                if headers and rows_data:
                    rows_data = [r + [""] * (len(headers) - len(r)) for r in rows_data]
                    df_input = pd.DataFrame(rows_data, columns=headers).astype(str)
                else:
                    raise ValueError("Could not read any data from XML file.")
            else:
                df_input = pd.read_excel(uploaded, dtype=str)
            vat_col = detect_vat_column(df_input)
            df_input = df_input.fillna("")
            st.markdown(
                f"<span class='feedback-text'>✅ <strong>{uploaded.name}</strong> — {len(df_input)} rows loaded · VAT column: <strong>{vat_col}</strong></span>",
                unsafe_allow_html=True)
        except Exception as e:
            st.error(f"Error reading file: {e}")
            df_input = None

    st.markdown("<br>", unsafe_allow_html=True)
    has_input = df_input is not None and len(df_input) > 0
    label = f"Validate {len(df_input)} VAT numbers" if has_input else "Validate VAT numbers"
    if st.button(label, disabled=not has_input):
        run_validation(df_input, vat_col)
    if not has_input:
        st.markdown("<span style='color:#888888; font-size:0.95rem;'>Upload a file to get started.</span>", unsafe_allow_html=True)

# ── DIVIDER ─────────────────────────────────
with col_mid:
    st.markdown('<div class="vert-line"></div>', unsafe_allow_html=True)

# ── RIGHT: Results ──────────────────────────
with col_right:
    if st.session_state.results is not None:
        df_res = st.session_state.results
        valid_df   = df_res[df_res["Status"]=="valid"]
        invalid_df = df_res[df_res["Status"]=="invalid"]
        error_df   = df_res[~df_res["Status"].isin(["valid","invalid"])]
        total = len(df_res)

        st.markdown('<span class="section-label">Results</span>', unsafe_allow_html=True)

        col_dl1, col_dl2, col_rst = st.columns([3, 3, 1])
        with col_dl1:
            st.download_button("Download (.xlsx)", data=to_excel_bytes(df_res),
                file_name=f"vat_validation_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True)
        with col_dl2:
            st.download_button("Download (.xml)", data=to_xml_bytes(df_res),
                file_name=f"vat_validation_{datetime.now().strftime('%Y%m%d_%H%M')}.xml",
                mime="application/xml", use_container_width=True)
        with col_rst:
            if st.button("Reset", use_container_width=True):
                st.session_state.results = None
                st.rerun()

        st.markdown("<br>", unsafe_allow_html=True)

        def style_status(val):
            return {
                "valid":   "color: #2d7a4f; font-weight: 700;",
                "invalid": "color: #E30613; font-weight: 700;",
            }.get(val, "color: #b05e00; font-weight: 700;")

        display_cols = [c for c in df_res.columns if c != "Status"]

        t1, t2, t3, t4 = st.tabs([
            f"All results ({total})",
            f"✓ Valid ({len(valid_df)})",
            f"✗ Invalid ({len(invalid_df)})",
            f"! Error ({len(error_df)})"
        ])
        with t1:
            st.dataframe(df_res[display_cols].style.applymap(style_status, subset=["Status Label"]),
                         use_container_width=True, hide_index=True, height=420)
        with t2:
            if len(valid_df) > 0:
                st.dataframe(valid_df[display_cols], use_container_width=True, hide_index=True, height=420)
            else:
                st.info("No valid VAT numbers found.")
        with t3:
            if len(invalid_df) > 0:
                st.dataframe(invalid_df[display_cols], use_container_width=True, hide_index=True, height=420)
            else:
                st.info("No invalid VAT numbers found.")
        with t4:
            if len(error_df) > 0:
                st.dataframe(error_df[display_cols], use_container_width=True, hide_index=True, height=420)
            else:
                st.info("No errors found.")
    else:
        st.markdown("""
<div style="height:300px; display:flex; align-items:center; justify-content:center;">
  <span style="color:#aaaaaa; font-size:1rem;">Results will appear here after validation.</span>
</div>
""", unsafe_allow_html=True)

st.markdown('</div>', unsafe_allow_html=True)
