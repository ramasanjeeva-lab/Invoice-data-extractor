"""
Invoice Data Extraction Tool — AI VERSION (Google Gemini, free tier)
KSCAA TECH RRC / Group 4

Why this version:
  - Reads the WHOLE invoice with AI, so it understands context
    (no more "Invoice No -> No" mistakes, finds HSN however it's worded)
  - Fills the hard fields: recipient name/address, taxable value,
    CGST/SGST/IGST split, tax amounts, etc.
  - Handles BOTH digital PDFs AND scanned/photographed invoices
    (jpg/png), because Gemini reads images directly.
  - Exports the same Excel file: "Invoice Details" + "Line Items".

Uses Google Gemini's FREE tier (no credit card needed).

HOW TO RUN (in the VS Code terminal):
  python -m pip install -r requirements.txt
  python -m streamlit run app.py
Then paste your free Gemini API key into the box in the app.
"""

import io
import json
import time

import streamlit as st
import pandas as pd
from google import genai
from google.genai import types


# We try these free models in order. If one is busy (503) or unavailable
# for this account (404), the app automatically moves to the next one.
MODELS = [
    "gemini-flash-latest",
    "gemini-3.5-flash",
    "gemini-flash-lite-latest",
    "gemini-2.5-flash-lite",
]

PROMPT = """You are a GST invoice data extraction engine.
Read the attached invoice (it may be a PDF or a photo) and return ONLY valid
JSON, no explanation and no markdown fences. Use this exact structure:

{
  "supplier_name_address": string,
  "supplier_gstin": string,
  "invoice_number": string,
  "date_of_issue": string,
  "nature": string,
  "recipient_name_address": string,
  "recipient_gstin": string,
  "state_name_and_code": string,
  "place_of_supply": string,
  "reverse_charge": string,
  "signature_present": string,
  "total_value": string,
  "taxable_value": string,
  "tax_rate": string,
  "tax_amount": string,
  "line_items": [
    {
      "description": string,
      "hsn_sac": string,
      "quantity": string,
      "unit": string,
      "rate": string,
      "taxable_value": string,
      "amount": string
    }
  ]
}

Rules:
- If a field is not present on the invoice, use an empty string "".
- Do not guess or invent values. Only use what is actually on the invoice.
- For "nature", always decide whether the invoice is for the purchase of
  goods, services, or both, based mainly on the DESCRIPTION of what is being
  sold (e.g. physical products = Goods; consulting, licenses, repairs,
  subscriptions, professional work = Services). Answer exactly "Goods",
  "Services", or "Both". Do this even when no HSN or SAC code is printed.
  If a code IS present, use it only as an extra hint (HSN usually = goods,
  SAC usually = services). Only leave it blank if you truly cannot tell.
- For signature_present, answer "Yes" or "No" based on whether a signature
  or digital signature appears.
"""


def extract_with_ai(api_key, file_bytes, mime_type, status=None):
    client = genai.Client(api_key=api_key)
    file_part = types.Part.from_bytes(data=file_bytes, mime_type=mime_type)

    last_error = None
    for model_name in MODELS:              # try each model in turn
        for attempt in range(3):           # up to 3 tries per model
            try:
                if status:
                    status.write(f"Trying model: {model_name} (attempt {attempt + 1})")
                response = client.models.generate_content(
                    model=model_name,
                    contents=[file_part, PROMPT],
                )
                text = response.text or ""
                cleaned = text.replace("```json", "").replace("```", "").strip()
                return json.loads(cleaned)
            except Exception as e:
                last_error = e
                msg = str(e)
                # Busy / rate-limited -> wait and retry the SAME model
                if ("503" in msg or "UNAVAILABLE" in msg
                        or "429" in msg or "RESOURCE_EXHAUSTED" in msg):
                    time.sleep(2 * (attempt + 1))  # 2s, 4s, 6s
                    continue
                # Model not available for this account -> jump to next model
                if "404" in msg or "NOT_FOUND" in msg:
                    break
                raise  # any other error: show it immediately
        # exhausted this model's retries (or it was 404) -> next model

    raise last_error  # every model failed


def build_excel(details_df, items_df):
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        details_df.to_excel(writer, sheet_name="Invoices", index=False)
        items_df.to_excel(writer, sheet_name="Line Items", index=False)
    buffer.seek(0)
    return buffer


st.set_page_config(page_title="GST Invoice Extractor (AI)", page_icon="\U0001F9FE", layout="wide")

# ---- Colorful theme (injected CSS) ----
st.markdown("""
<style>
/* App background: soft gradient */
.stApp {
    background: linear-gradient(135deg, #f5f7ff 0%, #eef6ff 50%, #f3fff9 100%);
}
/* Gradient header banner */
.hero {
    background: linear-gradient(120deg, #6a11cb 0%, #2575fc 100%);
    padding: 28px 32px;
    border-radius: 18px;
    color: white;
    box-shadow: 0 10px 30px rgba(37,117,252,0.25);
    margin-bottom: 22px;
}
.hero h1 { color: #fff; margin: 0; font-size: 2.1rem; }
.hero p  { color: #e8ecff; margin: 6px 0 0 0; font-size: 1rem; }
/* Section subheaders */
h2, h3 { color: #2b2d6e !important; }
/* Buttons: gradient + rounded */
.stButton>button, .stDownloadButton>button {
    background: linear-gradient(120deg, #6a11cb 0%, #2575fc 100%);
    color: #fff; border: none; border-radius: 10px;
    padding: 0.55rem 1.1rem; font-weight: 600;
    box-shadow: 0 4px 14px rgba(37,117,252,0.3);
}
.stButton>button:hover, .stDownloadButton>button:hover {
    filter: brightness(1.08); color: #fff;
}
/* Sidebar tint */
section[data-testid="stSidebar"] {
    background: linear-gradient(180deg, #ede7ff 0%, #e3f0ff 100%);
}
/* Stat cards */
.stat-card {
    background: #fff; border-radius: 14px; padding: 16px 18px;
    box-shadow: 0 4px 16px rgba(0,0,0,0.06); text-align: center;
    border-top: 4px solid #2575fc;
}
.stat-num { font-size: 1.8rem; font-weight: 700; color: #2575fc; }
.stat-lbl { font-size: 0.85rem; color: #555; }
</style>
""", unsafe_allow_html=True)

st.markdown("""
<div class="hero">
    <h1>\U0001F9FE GST Invoice Data Extractor</h1>
    <p>Reads digital PDFs and scanned/photo invoices with AI, and exports everything to Excel.</p>
</div>
""", unsafe_allow_html=True)

if "api_key" not in st.session_state:
    st.session_state.api_key = ""

# When deployed, the app owner can store a key in Streamlit "Secrets".
# If present, visitors don't need their own key.
try:
    owner_key = st.secrets.get("GEMINI_API_KEY", "")
except Exception:
    owner_key = ""

with st.sidebar:
    st.header("\U0001F511 Setup")
    if owner_key:
        st.success("A shared key is already set up \u2713\nJust upload invoices below.")
    else:
        st.write("Paste your free Gemini API key below (starts with `AQ.`).")
        st.write("Get one free at aistudio.google.com -> Get API key.")
        with st.form("key_form"):
            key_input = st.text_input("Gemini API key", type="password", value="")
            saved = st.form_submit_button("\U0001F4BE Save key")
            if saved:
                st.session_state.api_key = key_input.strip()
        if st.session_state.api_key:
            st.success(f"Key saved \u2713 (ends with ...{st.session_state.api_key[-4:]})")
        else:
            st.info("Paste your key, then click **Save key**.")

# owner key (deployed) takes priority; otherwise use the pasted one
api_key = owner_key or st.session_state.api_key

uploaded_files = st.file_uploader(
    "\U0001F4C1 Upload one or more invoices (PDF, JPG, or PNG)",
    type=["pdf", "jpg", "jpeg", "png"],
    accept_multiple_files=True,   # <-- bulk upload
)

if uploaded_files and not api_key:
    st.warning("Please paste your Gemini API key in the sidebar on the left first.")


def mime_for(name):
    name = name.lower()
    if name.endswith(".pdf"):
        return "application/pdf"
    if name.endswith((".jpg", ".jpeg")):
        return "image/jpeg"
    return "image/png"


def flatten(file_name, data):
    """Turn one invoice's AI result into a single flat row (for the table)."""
    return {
        "Source file": file_name,
        "Supplier name/address": data.get("supplier_name_address", ""),
        "Supplier GSTIN": data.get("supplier_gstin", ""),
        "Invoice number": data.get("invoice_number", ""),
        "Date": data.get("date_of_issue", ""),
        "Nature (Goods/Services)": data.get("nature", ""),
        "Recipient name/address": data.get("recipient_name_address", ""),
        "Recipient GSTIN": data.get("recipient_gstin", ""),
        "State & code": data.get("state_name_and_code", ""),
        "Place of supply": data.get("place_of_supply", ""),
        "Total value": data.get("total_value", ""),
        "Taxable value": data.get("taxable_value", ""),
        "Tax rate": data.get("tax_rate", ""),
        "Tax amount": data.get("tax_amount", ""),
        "Reverse charge": data.get("reverse_charge", ""),
        "Signature present": data.get("signature_present", ""),
    }


if uploaded_files and api_key:
    invoice_rows = []   # one row per invoice
    all_items = []      # every line item, tagged by source
    failures = []       # files that could not be read

    progress = st.progress(0.0)
    status = st.empty()
    total = len(uploaded_files)

    for i, f in enumerate(uploaded_files, start=1):
        status.write(f"Processing {i} of {total}: {f.name}")
        try:
            data = extract_with_ai(api_key, f.getvalue(), mime_for(f.name))
            invoice_rows.append(flatten(f.name, data))
            for item in (data.get("line_items") or []):
                item = dict(item)
                item["Source file"] = f.name
                item["Invoice number"] = data.get("invoice_number", "")
                all_items.append(item)
        except Exception as e:
            failures.append((f.name, str(e)))
        progress.progress(i / total)
        time.sleep(1)  # small gap so the free tier doesn't rate-limit us

    status.empty()
    progress.empty()

    if invoice_rows:
        invoices_df = pd.DataFrame(invoice_rows)
        items_df = pd.DataFrame(all_items) if all_items else pd.DataFrame(
            columns=["Source file", "Invoice number", "description",
                     "hsn_sac", "quantity", "unit", "rate",
                     "taxable_value", "amount"]
        )

        st.success(f"Done! Extracted {len(invoice_rows)} of {total} invoice(s).")

        # Colorful summary cards
        c1, c2, c3 = st.columns(3)
        c1.markdown(f'<div class="stat-card"><div style="font-size:1.6rem">\U0001F4C4</div>'
                    f'<div class="stat-num">{len(invoice_rows)}</div>'
                    f'<div class="stat-lbl">Invoices extracted</div></div>', unsafe_allow_html=True)
        c2.markdown(f'<div class="stat-card"><div style="font-size:1.6rem">\U0001F9FE</div>'
                    f'<div class="stat-num">{len(all_items)}</div>'
                    f'<div class="stat-lbl">Line items found</div></div>', unsafe_allow_html=True)
        c3.markdown(f'<div class="stat-card"><div style="font-size:1.6rem">\u26A0\uFE0F</div>'
                    f'<div class="stat-num">{len(failures)}</div>'
                    f'<div class="stat-lbl">Files skipped</div></div>', unsafe_allow_html=True)
        st.write("")  # spacer

        st.subheader("\U0001F4CB All invoices (one row each)")
        st.dataframe(invoices_df, use_container_width=True)
        st.subheader("\U0001F4E6 All line items")
        st.dataframe(items_df, use_container_width=True)

        excel_file = build_excel(invoices_df, items_df)
        st.download_button(
            "\u2b07\uFE0F Download all as Excel (.xlsx)",
            data=excel_file,
            file_name="invoices_extracted.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
    else:
        st.error("None of the files could be read. See details below.")

    if failures:
        st.warning(f"{len(failures)} file(s) could not be read:")
        for name, err in failures:
            st.write(f"- **{name}** \u2014 {err}")
