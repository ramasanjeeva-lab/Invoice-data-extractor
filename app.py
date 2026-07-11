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
from datetime import datetime

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
  "is_invoice": true or false,
  "document_type": string,
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
  "cgst": string,
  "sgst": string,
  "igst": string,
  "cess": string,
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
- FIRST decide what this document actually is. Set "document_type" to a short
  plain description (e.g. "Tax invoice", "Bank statement", "Receipt",
  "Purchase order", "Delivery challan", "Salary slip", "Resume", "Photograph
  of a person", "Letter", "Unknown document").
- Set "is_invoice" to true ONLY if it is genuinely an invoice or bill
  (tax invoice, purchase invoice, supplier bill). Otherwise set it to false.
- If "is_invoice" is false, leave ALL other fields empty and return
  "line_items": []. Do not attempt to extract invoice details.
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


def clean_number(value):
    """Tally/Zoho need plain numbers: no commas, no currency symbols."""
    if value is None:
        return ""
    s = str(value)
    for junk in ["\u20b9", "Rs.", "Rs", "INR", ",", " "]:
        s = s.replace(junk, "")
    try:
        return float(s)
    except ValueError:
        return ""


def clean_date(value):
    """Normalise to DD-MM-YYYY, which Tally requires."""
    if not value:
        return ""
    s = str(value).strip()
    formats = ["%d-%m-%Y", "%d/%m/%Y", "%d.%m.%Y", "%Y-%m-%d",
               "%d-%b-%Y", "%d %b %Y", "%d-%B-%Y", "%d %B %Y",
               "%d-%m-%y", "%d/%m/%y"]
    for fmt in formats:
        try:
            return datetime.strptime(s, fmt).strftime("%d-%m-%Y")
        except ValueError:
            continue
    return s  # give back the original if we can't parse it


def build_tally_rows(file_name, data):
    """
    TallyPrime purchase-voucher layout: ONE ROW PER LINE ITEM, with the
    voucher header fields repeated on every row of the same invoice.
    Amounts are plain numbers; dates are DD-MM-YYYY.
    """
    rows = []
    items = data.get("line_items") or [{}]
    for item in items:
        rows.append({
            "Voucher Date": clean_date(data.get("date_of_issue", "")),
            "Voucher Type": "Purchase",
            "Voucher Number": data.get("invoice_number", ""),
            "Party Ledger Name": data.get("supplier_name_address", "").split(",")[0].strip(),
            "Party GSTIN": data.get("supplier_gstin", ""),
            "Item / Ledger Name": item.get("description", ""),
            "HSN/SAC": item.get("hsn_sac", ""),
            "Quantity": clean_number(item.get("quantity", "")),
            "Rate": clean_number(item.get("rate", "")),
            "Taxable Value": clean_number(item.get("taxable_value", "")),
            "Amount": clean_number(item.get("amount", "")),
            "CGST": clean_number(data.get("cgst", "")),
            "SGST": clean_number(data.get("sgst", "")),
            "IGST": clean_number(data.get("igst", "")),
            "Place of Supply": data.get("place_of_supply", ""),
            "Source File": file_name,
        })
    return rows


def build_zoho_rows(file_name, data):
    """
    Zoho Books "Bills" import layout: one row per line item, bill header
    repeated. Zoho lets you map columns on import, so headers use its names.
    """
    rows = []
    items = data.get("line_items") or [{}]
    for item in items:
        rows.append({
            "Bill Number": data.get("invoice_number", ""),
            "Bill Date": clean_date(data.get("date_of_issue", "")),
            "Vendor Name": data.get("supplier_name_address", "").split(",")[0].strip(),
            "GST Identification Number (GSTIN)": data.get("supplier_gstin", ""),
            "Place of Supply": data.get("place_of_supply", ""),
            "Item Name": item.get("description", ""),
            "HSN/SAC": item.get("hsn_sac", ""),
            "Quantity": clean_number(item.get("quantity", "")),
            "Rate": clean_number(item.get("rate", "")),
            "Item Total": clean_number(item.get("amount", "")),
            "Bill Total": clean_number(data.get("total_value", "")),
            "Is Inclusive Tax": "false",
            "Source File": file_name,
        })
    return rows


def build_excel(details_df, items_df, tally_df=None, zoho_df=None):
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        details_df.to_excel(writer, sheet_name="Invoices", index=False)
        items_df.to_excel(writer, sheet_name="Line Items", index=False)
        if tally_df is not None:
            tally_df.to_excel(writer, sheet_name="Tally Import", index=False)
        if zoho_df is not None:
            zoho_df.to_excel(writer, sheet_name="Zoho Books Import", index=False)
    buffer.seek(0)
    return buffer


st.set_page_config(page_title="GST Invoice Extractor (AI)", page_icon="\U0001F9FE", layout="wide")

# ---- Theme (injected CSS) ----
st.markdown("""
<style>
.stApp { background: #f4f7fb; }
.block-container { padding-top: 2rem; max-width: 1100px; }

/* HERO ------------------------------------------------------------ */
.hero {
    background: linear-gradient(120deg, #4c1d95 0%, #2563eb 100%);
    padding: 40px 32px 34px;
    border-radius: 20px 20px 0 0;
    text-align: center;
    color: #fff;
}
.hero h1 {
    color: #fff; margin: 0; font-size: 2.4rem; font-weight: 800;
    letter-spacing: -0.5px;
}
.hero h1 .accent { color: #7dd3fc; }
.hero p { color: #dbeafe; margin: 10px 0 0; font-size: 1.05rem; }

/* DROP ZONE WRAPPER ----------------------------------------------- */
.dropwrap {
    background: linear-gradient(120deg, #4c1d95 0%, #2563eb 100%);
    padding: 0 32px 38px;
    border-radius: 0 0 20px 20px;
    margin-bottom: 26px;
}
/* Make Streamlit's uploader look like a big dashed drop zone */
[data-testid="stFileUploaderDropzone"] {
    background: rgba(255,255,255,0.08) !important;
    border: 3px dashed rgba(255,255,255,0.65) !important;
    border-radius: 14px !important;
    padding: 42px 20px !important;
    min-height: 170px;
}
[data-testid="stFileUploaderDropzone"]:hover {
    background: rgba(255,255,255,0.16) !important;
    border-color: #fff !important;
}
[data-testid="stFileUploaderDropzone"] * { color: #fff !important; }
[data-testid="stFileUploaderDropzone"] button {
    background: #f43f5e !important; color: #fff !important;
    border: none !important; border-radius: 8px !important;
    font-weight: 700 !important; padding: 0.6rem 1.6rem !important;
}
[data-testid="stFileUploaderDropzone"] button:hover { background: #e11d48 !important; }

/* FEATURE CARDS --------------------------------------------------- */
.feat {
    background: #fff; border-radius: 14px; padding: 22px 18px;
    text-align: center; box-shadow: 0 2px 10px rgba(15,23,42,0.07);
    height: 100%;
}
.feat .ico { font-size: 2rem; margin-bottom: 8px; }
.feat .txt { color: #475569; font-size: 0.92rem; line-height: 1.45; }

/* STAT CARDS ------------------------------------------------------ */
.stat-card {
    background: #fff; border-radius: 14px; padding: 16px 18px;
    box-shadow: 0 2px 10px rgba(15,23,42,0.07); text-align: center;
    border-top: 4px solid #2563eb;
}
.stat-num { font-size: 1.9rem; font-weight: 800; color: #1e3a8a; }
.stat-lbl { font-size: 0.82rem; color: #64748b; }

/* BUTTONS --------------------------------------------------------- */
.stDownloadButton>button {
    background: #16a34a !important; color: #fff !important; border: none !important;
    border-radius: 10px !important; padding: 0.75rem 1.6rem !important;
    font-weight: 700 !important; font-size: 1.02rem !important; width: 100%;
    box-shadow: 0 4px 14px rgba(22,163,74,0.3);
}
.stDownloadButton>button:hover { background: #15803d !important; }
.stButton>button {
    background: #fff !important; color: #334155 !important;
    border: 2px solid #cbd5e1 !important; border-radius: 10px !important;
    padding: 0.7rem 1.6rem !important; font-weight: 600 !important; width: 100%;
}
.stButton>button:hover { border-color: #f43f5e !important; color: #f43f5e !important; }

/* SIDEBAR --------------------------------------------------------- */
section[data-testid="stSidebar"] { background: #eef2ff; }
h2, h3 { color: #1e293b !important; }
</style>
""", unsafe_allow_html=True)

st.markdown("""
<div class="hero">
    <h1>CONVERT INVOICES TO <span class="accent">EXCEL</span></h1>
    <p>Any invoice \u2014 scanned, photographed, or digital PDF \u2014 read by AI and exported for Tally &amp; Zoho Books.</p>
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

# A changing key lets the Clear button fully reset the uploader
if "uploader_key" not in st.session_state:
    st.session_state.uploader_key = 0

st.markdown('<div class="dropwrap">', unsafe_allow_html=True)
uploaded_files = st.file_uploader(
    "Drag & drop your invoices here \u2014 or click Browse",
    type=["pdf", "jpg", "jpeg", "png"],
    accept_multiple_files=True,
    key=f"uploader_{st.session_state.uploader_key}",
    label_visibility="collapsed",
)
st.markdown('</div>', unsafe_allow_html=True)

# Reassurance cards (only before any files are added, to keep results clean)
if not uploaded_files:
    f1, f2, f3 = st.columns(3)
    f1.markdown('<div class="feat"><div class="ico">\u26A1</div>'
                '<div class="txt">Upload and extraction starts automatically.<br>'
                'No settings to configure.</div></div>', unsafe_allow_html=True)
    f2.markdown('<div class="feat"><div class="ico">\U0001F4F7</div>'
                '<div class="txt">Works with scanned copies and phone photos,<br>'
                'not just digital PDFs.</div></div>', unsafe_allow_html=True)
    f3.markdown('<div class="feat"><div class="ico">\U0001F4D2</div>'
                '<div class="txt">Excel output ready to import into<br>'
                'TallyPrime and Zoho Books.</div></div>', unsafe_allow_html=True)

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
        "CGST": data.get("cgst", ""),
        "SGST": data.get("sgst", ""),
        "IGST": data.get("igst", ""),
        "Cess": data.get("cess", ""),
        "Reverse charge": data.get("reverse_charge", ""),
        "Signature present": data.get("signature_present", ""),
    }


if uploaded_files and api_key:
    invoice_rows = []   # one row per invoice
    all_items = []      # every line item, tagged by source
    tally_rows = []     # Tally purchase-voucher rows
    zoho_rows = []      # Zoho Books bill rows
    not_invoices = []   # (filename, what it actually is)
    failures = []       # files that could not be read at all

    progress = st.progress(0.0)
    status = st.empty()
    total = len(uploaded_files)

    for i, f in enumerate(uploaded_files, start=1):
        status.write(f"Processing {i} of {total}: {f.name}")
        try:
            data = extract_with_ai(api_key, f.getvalue(), mime_for(f.name))

            # --- Is this actually an invoice? ---
            if not data.get("is_invoice", False):
                doc_type = data.get("document_type") or "Not recognised"
                not_invoices.append((f.name, doc_type))
                progress.progress(i / total)
                time.sleep(1)
                continue  # skip extraction entirely

            invoice_rows.append(flatten(f.name, data))
            for item in (data.get("line_items") or []):
                item = dict(item)
                item["Source file"] = f.name
                item["Invoice number"] = data.get("invoice_number", "")
                all_items.append(item)
            tally_rows.extend(build_tally_rows(f.name, data))
            zoho_rows.extend(build_zoho_rows(f.name, data))
        except Exception as e:
            failures.append((f.name, str(e)))
        progress.progress(i / total)
        time.sleep(1)  # small gap so the free tier doesn't rate-limit us

    status.empty()
    progress.empty()

    # --- Report any non-invoices FIRST, clearly ---
    if not_invoices:
        st.warning(f"\U0001F6AB {len(not_invoices)} file(s) are not invoices \u2014 skipped, no details extracted:")
        for name, doc_type in not_invoices:
            st.write(f"- **{name}** \u2014 this looks like a **{doc_type}**, not an invoice.")

    if invoice_rows:
        invoices_df = pd.DataFrame(invoice_rows)
        items_df = pd.DataFrame(all_items) if all_items else pd.DataFrame(
            columns=["Source file", "Invoice number", "description",
                     "hsn_sac", "quantity", "unit", "rate",
                     "taxable_value", "amount"]
        )
        tally_df = pd.DataFrame(tally_rows)
        zoho_df = pd.DataFrame(zoho_rows)

        st.success(f"Done! Extracted {len(invoice_rows)} invoice(s) out of {total} file(s).")

        # Colorful summary cards
        c1, c2, c3, c4 = st.columns(4)
        c1.markdown(f'<div class="stat-card"><div style="font-size:1.6rem">\U0001F4C4</div>'
                    f'<div class="stat-num">{len(invoice_rows)}</div>'
                    f'<div class="stat-lbl">Invoices extracted</div></div>', unsafe_allow_html=True)
        c2.markdown(f'<div class="stat-card"><div style="font-size:1.6rem">\U0001F9FE</div>'
                    f'<div class="stat-num">{len(all_items)}</div>'
                    f'<div class="stat-lbl">Line items found</div></div>', unsafe_allow_html=True)
        c3.markdown(f'<div class="stat-card"><div style="font-size:1.6rem">\U0001F6AB</div>'
                    f'<div class="stat-num">{len(not_invoices)}</div>'
                    f'<div class="stat-lbl">Not invoices</div></div>', unsafe_allow_html=True)
        c4.markdown(f'<div class="stat-card"><div style="font-size:1.6rem">\u26A0\uFE0F</div>'
                    f'<div class="stat-num">{len(failures)}</div>'
                    f'<div class="stat-lbl">Failed to read</div></div>', unsafe_allow_html=True)
        st.write("")  # spacer

        tab1, tab2, tab3, tab4 = st.tabs([
            "\U0001F4CB Invoices", "\U0001F4E6 Line Items",
            "\U0001F4D2 Tally Import", "\U0001F4D7 Zoho Books Import",
        ])
        with tab1:
            st.dataframe(invoices_df, use_container_width=True)
        with tab2:
            st.dataframe(items_df, use_container_width=True)
        with tab3:
            st.caption("Purchase vouchers \u2014 one row per line item, dates as DD-MM-YYYY, "
                       "plain numbers. In TallyPrime: Gateway > Import Data > Vouchers.")
            st.dataframe(tally_df, use_container_width=True)
        with tab4:
            st.caption("Bills \u2014 one row per line item. In Zoho Books: "
                       "Purchases > Bills > More > Import Bills, then map the columns.")
            st.dataframe(zoho_df, use_container_width=True)

        excel_file = build_excel(invoices_df, items_df, tally_df, zoho_df)

        st.write("")
        d1, d2 = st.columns([3, 1])
        with d1:
            st.download_button(
                "\u2b07\uFE0F  Download Excel (Invoices \u2022 Line Items \u2022 Tally \u2022 Zoho)",
                data=excel_file,
                file_name="invoices_extracted.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
        with d2:
            if st.button("\U0001F5D1\uFE0F  Clear all"):
                st.session_state.uploader_key += 1   # new key = empty uploader
                st.rerun()
        st.caption("Tip: download your Excel first \u2014 **Clear all** removes the uploaded files "
                   "and results so you can start a fresh batch.")
    elif not not_invoices:
        st.error("None of the files could be read. See details below.")

    if failures:
        st.warning(f"{len(failures)} file(s) could not be read:")
        for name, err in failures:
            st.write(f"- **{name}** \u2014 {err}")

    # If nothing was extracted, still let the user clear and try again
    if not invoice_rows and (not_invoices or failures):
        if st.button("\U0001F5D1\uFE0F  Clear all and start over"):
            st.session_state.uploader_key += 1
            st.rerun()
