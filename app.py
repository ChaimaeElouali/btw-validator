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

    /* ── Hero ── */
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
        font-weight: 400; white-space: nowrap; line-height: 1.5;
    }

    /* ── Section labels ── */
    .section-label {
        font-size: 0.8rem; font-weight: 700; text-transform: uppercase;
        letter-spacing: 1.5px; color: #E30613;
        margin-bottom: 0.4rem; display: block;
    }
    .section-hint {
        font-size: 0.9rem; color: #aaaaaa; margin-bottom: 0.75rem;
        display: block; font-weight: 400; line-height: 1.4;
    }

    /* ── Or divider ── */
    .or-col {
        display: flex; flex-direction: column;
        align-items: center; justify-content: center;
        height: 200px;
    }
    .or-line { flex: 1; width: 1px; background: #2a2a2a; }
    .or-text {
        font-size: 0.75rem; font-weight: 700; letter-spacing: 2px;
        color: #555555; padding: 10px 0; text-transform: uppercase;
    }

    /* ── Upload zone ── */
    [data-testid="stFileUploader"] { background: transparent !important; }
    [data-testid="stFileUploader"] > div {
        background: #1a1a1a !important;
        border: 1px solid #2a2a2a !important;
        border-radius: 8px !important;
        transition: border-color 0.2s;
    }
    [data-testid="stFileUploader"] > div:hover { border-color: #E30613 !important; }
    [data-testid="stFileUploader"] section { background: #1a1a1a !important; border: none !important; }
    [data-testid="stFileUploader"] button {
        background: #2a2a2a !important; color: #d0d0d0 !important;
        border: 1px solid #3a3a3a !important; border-radius: 4px !important;
        font-family: 'Barlow', sans-serif !important; font-weight: 600 !important;
        font-size: 0.9rem !important;
    }
    [data-testid="stFileUploader"] p { color: #888888 !important; font-size: 0.9rem !important; }
    [data-testid="stFileUploader"] small { color: #666666 !important; font-size: 0.85rem !important; }
    [data-testid="stFileUploaderDropzoneInstructions"] div span { color: #888888 !important; font-size: 0.9rem !important; }

    /* ── Textarea ── */
    textarea {
        background: #1a1a1a !important; color: #d0d0d0 !important;
        border: 1px solid #2a2a2a !important; border-radius: 8px !important;
        font-family: 'Barlow', sans-serif !important; font-size: 0.95rem !important;
        line-height: 1.6 !important;
    }
    textarea:focus { border-color: #E30613 !important; box-shadow: none !important; }
    textarea::placeholder { color: #444444 !important; }

    /* ── Validate button ── */
    .stButton > button {
        background: #E30613 !important; color: #ffffff !important;
        border: none !important; font-family: 'Barlow', sans-serif !important;
        font-weight: 700 !important; font-size: 1rem !important;
        padding: 0.7rem 2rem !important; border-radius: 5px !important;
        white-space: nowrap !important; letter-spacing: 0.3px;
    }
    .stButton > button:hover { background: #c0050f !important; }

    /* ── Download button ── */
    .stDownloadButton > button {
        background: #1a1a1a !important; color: #d0d0d0 !important;
        border: 1px solid #333333 !important;
        font-family: 'Barlow', sans-serif !important;
        font-weight: 600 !important; border-radius: 5px !important;
        font-size: 0.95rem !important; padding: 0.6rem 1.5rem !important;
    }
    .stDownloadButton > button:hover { border-color: #555555 !important; color: #ffffff !important; }

    /* ── Stat cards ── */
    .stat-row {
        display: grid; grid-template-columns: repeat(4, 1fr);
        gap: 12px; margin-bottom: 1.5rem;
    }
    .stat-card {
        background: #1a1a1a; border-radius: 6px;
        padding: 1.2rem 1rem; text-align: center;
        border: 1px solid #222222;
    }
    .stat-card.total   { border-top: 3px solid #555555; }
    .stat-card.valid   { border-top: 3px solid #4a9d6f; }
    .stat-card.invalid { border-top: 3px solid #E30613; }
    .stat-card.error   { border-top: 3px solid #c97d20; }
    .stat-num {
        font-family: 'Barlow Condensed', sans-serif;
        font-size: 2.2rem; font-weight: 800; line-height: 1;
    }
    .stat-lbl {
        font-size: 0.75rem; font-weight: 600; text-transform: uppercase;
        letter-spacing: 1.5px; color: #666666; margin-top: 6px;
    }
    .stat-icon { font-size: 0.9rem; margin-bottom: 4px; color: #555555; }
    .total   .stat-num { color: #888888; }
    .valid   .stat-num { color: #4a9d6f; }
    .invalid .stat-num { color: #E30613; }
    .error   .stat-num { color: #c97d20; }

    /* ── Tabs ── */
    .stTabs [data-baseweb="tab-list"] {
        background: #1a1a1a; border-bottom: 1px solid #2a2a2a;
    }
    .stTabs [data-baseweb="tab"] {
        color: #666666 !important; font-family: 'Barlow', sans-serif !important;
        font-weight: 600 !important; padding: 0.7rem 1.5rem !important;
        font-size: 0.9rem !important;
    }
    .stTabs [aria-selected="true"] {
        color: #ffffff !important;
        border-bottom: 2px solid #E30613 !important;
        background: transparent !important;
    }

    /* ── Misc ── */
    code { background: transparent !important; color: #aaaaaa !important; border: none !important; font-size: 0.85rem !important; }
    .stProgress > div > div { background-color: #E30613 !important; }
    hr { border-color: #1e1e1e !important; }
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
    """Generate Excel XML 2003 format."""
    lines = []
    lines.append('<?xml version="1.0" encoding="UTF-8"?>')
    lines.append('<?mso-application progid="Excel.Sheet"?>')
    lines.append('<Workbook xmlns="urn:schemas-microsoft-com:office:spreadsheet"')
    lines.append(' xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet">')
    lines.append('<Worksheet ss:Name="All results">')
    lines.append('<Table>')
    # Header row
    lines.append('<Row>')
    for col in df.columns:
        lines.append(f'<Cell><Data ss:Type="String">{col}</Data></Cell>')
    lines.append('</Row>')
    # Data rows
    for _, row in df.iterrows():
        lines.append('<Row>')
        for val in row:
            safe = str(val).replace('&','&amp;').replace('<','&lt;').replace('>','&gt;')
            lines.append(f'<Cell><Data ss:Type="String">{safe}</Data></Cell>')
        lines.append('</Row>')
    lines.append('</Table>')
    lines.append('</Worksheet>')
    lines.append('</Workbook>')
    return '\n'.join(lines).encode('utf-8')

def run_validation(df_input, vat_col):
    st.markdown("---")
    st.markdown("**Validating VAT numbers...**")
    progress_bar = st.progress(0)
    status_text  = st.empty()
    results = []
    total = len(df_input)
    for i, (_, row) in enumerate(df_input.iterrows()):
        raw = str(row[vat_col]).strip()
        country_code, vat_number = parse_vat(raw)
        status_text.markdown(
            f"<span style='color:#777; font-size:0.9rem;'>Processing {i+1} of {total}: {raw}</span>",
            unsafe_allow_html=True)
        result = {"status":"invalid","company_name":"—","company_address":"—","message":"Invalid format"} \
            if not country_code or not vat_number else check_vat(country_code, vat_number)

        # Status label with accessibility symbol
        status_label = {
            "valid":   "✓ valid",
            "invalid": "✗ invalid",
            "error":   "! error"
        }.get(result["status"], "! error")

        row_data = {"VAT Input": raw, "Country": country_code, "VAT Number": vat_number}
        for c in df_input.columns:
            if c != vat_col: row_data[c] = row[c]
        row_data.update({
            "Status":        result["status"],
            "Status Label":  status_label,
            "Company (VIES)":result["company_name"],
            "Address (VIES)":result["company_address"],
            "Message":       result["message"]
        })
        results.append(row_data)
        progress_bar.progress((i+1)/total)
        time.sleep(0.4)
    status_text.empty()
    st.session_state.results = pd.DataFrame(results)
    st.session_state.scroll_to_results = True
    st.rerun()

# ─── HEADER ─────────────────────────────────
st.markdown("""
<div class="hero">
    <div class="hero-title">VAT <span>VALIDATOR</span></div>
    <div class="hero-sub">VAT numbers are validated in real-time via the official VIES API of the European Commission.</div>
</div>
""", unsafe_allow_html=True)

if "results" not in st.session_state:
    st.session_state.results = None
if "scroll_to_results" not in st.session_state:
    st.session_state.scroll_to_results = False

# ─── MAIN LAYOUT: left = upload, right = results ───
col_left, col_divider, col_right = st.columns([5, 1, 6], gap="small")

df_input = None
vat_col  = "vat_number"

# ── LEFT: Upload ──
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
            st.markdown(f"<span style='color:#aaaaaa; font-size:0.9rem;'>{len(df_input)} rows loaded &nbsp;&middot;&nbsp; VAT column: {vat_col}</span>", unsafe_allow_html=True)
        except Exception as e:
            st.error(f"Error reading file: {e}")
            df_input = None

    st.markdown("<br>", unsafe_allow_html=True)
    has_input = df_input is not None and len(df_input) > 0
    label = f"Validate {len(df_input)} VAT numbers" if has_input else "Validate VAT numbers"
    if st.button(label, use_container_width=False, disabled=not has_input):
        run_validation(df_input, vat_col)

    if not has_input:
        st.markdown("<span style='color:#555555; font-size:0.9rem;'>Upload a file to get started.</span>", unsafe_allow_html=True)

# ── DIVIDER ──
with col_divider:
    st.markdown("""
<div style="display:flex; flex-direction:column; align-items:center; height:500px; padding-top:2rem;">
  <div style="flex:1; width:1px; background:#2a2a2a;"></div>
</div>
""", unsafe_allow_html=True)

# ── RIGHT: Results ──
with col_right:
    if st.session_state.results is not None:
        df_res = st.session_state.results
        valid_df   = df_res[df_res["Status"]=="valid"]
        invalid_df = df_res[df_res["Status"]=="invalid"]
        error_df   = df_res[~df_res["Status"].isin(["valid","invalid"])]
        total = len(df_res)

        st.markdown('<span class="section-label">Results</span>', unsafe_allow_html=True)

        st.markdown(
            f'<div style="display:flex;align-items:center;gap:1.5rem;background:#1a1a1a;border:1px solid #2a2a2a;'
            'border-radius:6px;padding:0.6rem 1.2rem;margin-bottom:1rem;">'
            f'<span style="color:#888;font-size:0.85rem;font-weight:600;">{total} total</span>'
            f'<span style="color:#3a3a3a;margin:0 0.2rem;">|</span>'
            f'<span style="color:#4a9d6f;font-size:0.85rem;font-weight:700;">&#10003; {len(valid_df)} valid</span>'
            f'<span style="color:#3a3a3a;margin:0 0.2rem;">|</span>'
            f'<span style="color:#E30613;font-size:0.85rem;font-weight:700;">&#10007; {len(invalid_df)} invalid</span>'
            f'<span style="color:#3a3a3a;margin:0 0.2rem;">|</span>'
            f'<span style="color:#c97d20;font-size:0.85rem;font-weight:700;">! {len(error_df)} error</span>'
            '</div>',
            unsafe_allow_html=True
        )


        col_dl_xlsx, col_dl_xml, col_reset = st.columns([3, 3, 1])
        with col_dl_xlsx:
            st.download_button(
                label="Download (.xlsx)",
                data=to_excel_bytes(df_res),
                file_name=f"vat_validation_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )
        with col_dl_xml:
            st.download_button(
                label="Download (.xml)",
                data=to_xml_bytes(df_res),
                file_name=f"vat_validation_{datetime.now().strftime('%Y%m%d_%H%M')}.xml",
                mime="application/xml",
                use_container_width=True
            )
        with col_reset:
            if st.button("Reset", use_container_width=True):
                st.session_state.results = None
                st.rerun()

        st.markdown("<br>", unsafe_allow_html=True)

        def style_status(val):
            return {
                "valid":   "color: #4a9d6f; font-weight: 600;",
                "invalid": "color: #E30613; font-weight: 600;",
            }.get(val, "color: #c97d20; font-weight: 600;")

        display_cols = [c for c in df_res.columns if c != "Status"]

        t1, t2, t3, t4 = st.tabs([
            f"All results ({total})",
            f"Valid ({len(valid_df)})",
            f"Invalid ({len(invalid_df)})",
            f"Error ({len(error_df)})"
        ])
        with t1:
            st.dataframe(df_res[display_cols].style.applymap(style_status, subset=["Status Label"]),
                         use_container_width=True, hide_index=True, height=400)
        with t2:
            if len(valid_df) > 0:
                st.dataframe(valid_df[display_cols], use_container_width=True, hide_index=True, height=400)
            else:
                st.info("No valid VAT numbers found.")
        with t3:
            if len(invalid_df) > 0:
                st.dataframe(invalid_df[display_cols], use_container_width=True, hide_index=True, height=400)
            else:
                st.info("No invalid VAT numbers found.")
        with t4:
            if len(error_df) > 0:
                st.dataframe(error_df[display_cols], use_container_width=True, hide_index=True, height=400)
            else:
                st.info("No errors found.")
    else:
        st.markdown("""
<div style="height:300px; display:flex; align-items:center; justify-content:center;">
  <span style="color:#333333; font-size:0.9rem;">Results will appear here after validation.</span>
</div>
""", unsafe_allow_html=True)

st.markdown("---")
st.markdown("<p style='color:#333333; font-size:0.8rem; text-align:center;'>Mammoet Data Migration Team &nbsp;&middot;&nbsp; VIES API (European Commission) &nbsp;&middot;&nbsp; SAP ECC &rarr; S/4HANA</p>", unsafe_allow_html=True)
