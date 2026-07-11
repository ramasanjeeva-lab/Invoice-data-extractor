"""
Invoice Data Extraction Tool
KSCAA TECH RRC / Group 4

- Reads digital PDFs, scanned PDFs and phone photos (JPG/PNG) using AI
- Detects non-invoices and skips them (says what they are)
- Bulk upload with progress + live timer
- Exports Excel: Invoices | Line Items | Tally Import | Zoho Books Import
"""

import io
import json
import time
from datetime import datetime

import streamlit as st
import pandas as pd
from google import genai
from google.genai import types

MODELS = [
    "gemini-flash-latest",
    "gemini-3.5-flash",
    "gemini-flash-lite-latest",
    "gemini-2.5-flash-lite",
]

PROMPT = """You are a GST invoice data extraction engine.
Read the attached document (PDF or photo) and return ONLY valid JSON,
no explanation and no markdown fences. Use this exact structure:

{
  "is_invoice": true or false,
  "document_type": string,
  "supplier_name": string,
  "supplier_address": string,
  "supplier_gstin": string,
  "vendor_type": string,
  "invoice_number": string,
  "date_of_issue": string,
  "nature": string,
  "recipient_name": string,
  "recipient_address": string,
  "recipient_gstin": string,
  "state_name_and_code": string,
  "place_of_supply": string,
  "reverse_charge": string,
  "signature_present": string,
  "currency": string,
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
  "Purchase order", "Delivery challan", "Salary slip", "Resume",
  "Photograph", "Unknown document").
- Set "is_invoice" to true ONLY if it is genuinely an invoice or bill.
  Otherwise set it to false, leave all other fields empty, and return
  "line_items": []. Do not extract anything from a non-invoice.
- DATES: always output "date_of_issue" in DD-MM-YYYY format. Convert whatever
  format appears on the invoice (e.g. "30-May-2025" -> "30-05-2025",
  "2025/05/30" -> "30-05-2025", "May 30, 2025" -> "30-05-2025").
- SUPPLIER: put ONLY the company/person name in "supplier_name" and the full
  postal address in "supplier_address". Never mix them. Same for recipient.
- VENDOR TYPE: set "vendor_type" to "Local" if the supplier is based in India
  (Indian address, or has an Indian GSTIN). Set it to "Foreign" if the supplier
  is outside India (foreign address, or invoice is in a foreign currency such
  as USD/EUR/GBP/AED). Always answer "Local" or "Foreign".
- TAX AMOUNTS: for cgst, sgst, igst, cess, tax_amount, if the tax is absent,
  nil, N/A, or a dash, return "0" (the digit zero). Never return "NIL" or "-".
- NATURE: decide "Goods", "Services", or "Both" based mainly on WHAT is being
  sold (physical products = Goods; consulting, licences, repairs, subscriptions,
  professional work = Services). Do this even when no HSN/SAC code is printed.
  If a code is present use it as a hint (HSN usually goods, SAC usually services).
- For signature_present, answer "Yes" or "No".
- Do not guess or invent values. Use only what is on the document.
"""


# ------------------------------------------------------------------ #
# AI call: try each model, retry on busy / rate limit
# ------------------------------------------------------------------ #
def extract_with_ai(api_key, file_bytes, mime_type):
    client = genai.Client(api_key=api_key)
    file_part = types.Part.from_bytes(data=file_bytes, mime_type=mime_type)

    last_error = None
    for model_name in MODELS:
        for attempt in range(3):
            try:
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
                if ("503" in msg or "UNAVAILABLE" in msg
                        or "429" in msg or "RESOURCE_EXHAUSTED" in msg):
                    time.sleep(2 * (attempt + 1))
                    continue
                if "404" in msg or "NOT_FOUND" in msg:
                    break
                raise
    raise last_error


# ------------------------------------------------------------------ #
# Cleaning helpers
# ------------------------------------------------------------------ #
def clean_number(value, blank_as_zero=False):
    """Plain numbers for Tally/Zoho: no commas, no currency symbols."""
    if value is None or str(value).strip() == "":
        return 0 if blank_as_zero else ""
    s = str(value).strip()
    if s.upper() in ("NIL", "N/A", "NA", "-", "--", "NONE"):
        return 0 if blank_as_zero else ""
    for junk in ["\u20b9", "Rs.", "Rs", "INR", "$", "\u20ac", "\u00a3", ",", " "]:
        s = s.replace(junk, "")
    try:
        return float(s)
    except ValueError:
        return 0 if blank_as_zero else ""


def tax_value(value):
    """Tax fields: missing/NIL becomes 0."""
    return clean_number(value, blank_as_zero=True)


def clean_date(value):
    """Force DD-MM-YYYY."""
    if not value:
        return ""
    s = str(value).strip()
    formats = ["%d-%m-%Y", "%d/%m/%Y", "%d.%m.%Y", "%Y-%m-%d", "%Y/%m/%d",
               "%d-%b-%Y", "%d %b %Y", "%d-%B-%Y", "%d %B %Y",
               "%b %d, %Y", "%B %d, %Y", "%d-%m-%y", "%d/%m/%y"]
    for fmt in formats:
        try:
            return datetime.strptime(s, fmt).strftime("%d-%m-%Y")
        except ValueError:
            continue
    return s


# ------------------------------------------------------------------ #
# Sheet builders
# ------------------------------------------------------------------ #
def flatten(file_name, data):
    return {
        "Source file": file_name,
        "Supplier name": data.get("supplier_name", ""),
        "Supplier address": data.get("supplier_address", ""),
        "Supplier GSTIN": data.get("supplier_gstin", ""),
        "Vendor type": data.get("vendor_type", ""),
        "Invoice number": data.get("invoice_number", ""),
        "Date": clean_date(data.get("date_of_issue", "")),
        "Nature (Goods/Services)": data.get("nature", ""),
        "Recipient name": data.get("recipient_name", ""),
        "Recipient address": data.get("recipient_address", ""),
        "Recipient GSTIN": data.get("recipient_gstin", ""),
        "State & code": data.get("state_name_and_code", ""),
        "Place of supply": data.get("place_of_supply", ""),
        "Currency": data.get("currency", ""),
        "Total value": clean_number(data.get("total_value", "")),
        "Taxable value": clean_number(data.get("taxable_value", "")),
        "Tax rate": data.get("tax_rate", ""),
        "Tax amount": tax_value(data.get("tax_amount", "")),
        "CGST": tax_value(data.get("cgst", "")),
        "SGST": tax_value(data.get("sgst", "")),
        "IGST": tax_value(data.get("igst", "")),
        "Cess": tax_value(data.get("cess", "")),
        "Reverse charge": data.get("reverse_charge", ""),
        "Signature present": data.get("signature_present", ""),
    }


def build_tally_rows(file_name, data):
    """TallyPrime purchase voucher: one row per line item, header repeated."""
    rows = []
    for item in (data.get("line_items") or [{}]):
        rows.append({
            "Voucher Date": clean_date(data.get("date_of_issue", "")),
            "Voucher Type": "Purchase",
            "Voucher Number": data.get("invoice_number", ""),
            "Party Ledger Name": data.get("supplier_name", ""),
            "Party Address": data.get("supplier_address", ""),
            "Party GSTIN": data.get("supplier_gstin", ""),
            "Vendor Type": data.get("vendor_type", ""),
            "Item / Ledger Name": item.get("description", ""),
            "HSN/SAC": item.get("hsn_sac", ""),
            "Quantity": clean_number(item.get("quantity", "")),
            "Rate": clean_number(item.get("rate", "")),
            "Taxable Value": clean_number(item.get("taxable_value", "")),
            "Amount": clean_number(item.get("amount", "")),
            "CGST": tax_value(data.get("cgst", "")),
            "SGST": tax_value(data.get("sgst", "")),
            "IGST": tax_value(data.get("igst", "")),
            "Cess": tax_value(data.get("cess", "")),
            "Place of Supply": data.get("place_of_supply", ""),
            "Source File": file_name,
        })
    return rows


def build_zoho_rows(file_name, data):
    """Zoho Books Bills import: one row per line item, header repeated."""
    rows = []
    for item in (data.get("line_items") or [{}]):
        rows.append({
            "Bill Number": data.get("invoice_number", ""),
            "Bill Date": clean_date(data.get("date_of_issue", "")),
            "Vendor Name": data.get("supplier_name", ""),
            "Vendor Address": data.get("supplier_address", ""),
            "GST Identification Number (GSTIN)": data.get("supplier_gstin", ""),
            "Vendor Type": data.get("vendor_type", ""),
            "Currency Code": data.get("currency", "") or "INR",
            "Place of Supply": data.get("place_of_supply", ""),
            "Item Name": item.get("description", ""),
            "HSN/SAC": item.get("hsn_sac", ""),
            "Quantity": clean_number(item.get("quantity", "")),
            "Rate": clean_number(item.get("rate", "")),
            "Item Total": clean_number(item.get("amount", "")),
            "CGST": tax_value(data.get("cgst", "")),
            "SGST": tax_value(data.get("sgst", "")),
            "IGST": tax_value(data.get("igst", "")),
            "Bill Total": clean_number(data.get("total_value", "")),
            "Is Inclusive Tax": "false",
            "Source File": file_name,
        })
    return rows


def build_excel(inv_df, items_df, tally_df, zoho_df):
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        inv_df.to_excel(writer, sheet_name="Invoices", index=False)
        items_df.to_excel(writer, sheet_name="Line Items", index=False)
        tally_df.to_excel(writer, sheet_name="Tally Import", index=False)
        zoho_df.to_excel(writer, sheet_name="Zoho Books Import", index=False)
    buffer.seek(0)
    return buffer


def mime_for(name):
    n = name.lower()
    if n.endswith(".pdf"):
        return "application/pdf"
    if n.endswith((".jpg", ".jpeg")):
        return "image/jpeg"
    return "image/png"


# ------------------------------------------------------------------ #
# UI
# ------------------------------------------------------------------ #
st.set_page_config(page_title="Invoice Data Extraction Tool",
                   page_icon="\U0001F4CA", layout="wide")

st.markdown("""
<style>
.stApp { background: #f4f7fb; }
.block-container { padding-top: 1.5rem; max-width: 1150px; }

.hero {
    background: linear-gradient(120deg, #4c1d95 0%, #2563eb 100%);
    padding: 24px 30px; border-radius: 16px;
    color: #fff; margin-bottom: 20px;
}
.hero-row { display:flex; align-items:center; gap:18px; }
.hero-logo { flex:0 0 auto; display:flex; align-items:center;
    filter: drop-shadow(0 4px 10px rgba(0,0,0,.25)); }
.hero-txt { flex:1 1 auto; }
.hero h1 { color:#fff; margin:0; font-size:1.9rem; font-weight:800; letter-spacing:-.3px; }
.hero p  { color:#dbeafe; margin:5px 0 0; font-size:0.95rem; }

/* Panel headings */
.panel-h { font-weight:700; color:#1e293b; margin:0 0 8px; font-size:1rem; }

/* Compact dashed drop zone, LIGHT so file names stay readable */
[data-testid="stFileUploaderDropzone"] {
    background:#ffffff !important;
    border:2px dashed #94a3b8 !important;
    border-radius:12px !important;
    padding:18px !important;
    min-height:120px;
}
[data-testid="stFileUploaderDropzone"]:hover { border-color:#2563eb !important; }

/* Cards */
.stat-card { background:#fff; border-radius:12px; padding:14px;
    box-shadow:0 2px 8px rgba(15,23,42,.07); text-align:center;
    border-top:4px solid #2563eb; }
.stat-num { font-size:1.6rem; font-weight:800; color:#1e3a8a; }
.stat-lbl { font-size:.78rem; color:#64748b; }

/* Buttons */
.stDownloadButton>button { background:#16a34a !important; color:#fff !important;
    border:none !important; border-radius:10px !important; font-weight:700 !important;
    padding:.7rem 1rem !important; width:100%; }
.stDownloadButton>button:hover { background:#15803d !important; }
.stButton>button { background:#fff !important; color:#334155 !important;
    border:2px solid #cbd5e1 !important; border-radius:10px !important;
    font-weight:600 !important; padding:.65rem 1rem !important; width:100%; }
.stButton>button:hover { border-color:#f43f5e !important; color:#f43f5e !important; }

section[data-testid="stSidebar"] { background:#eef2ff; }
</style>
""", unsafe_allow_html=True)

LOGO_SVG = """
<svg width="58" height="58" viewBox="0 0 64 64" fill="none" xmlns="http://www.w3.org/2000/svg">
  <defs>
    <linearGradient id="pg" x1="0" y1="0" x2="1" y2="1">
      <stop offset="0%" stop-color="#ffffff"/><stop offset="100%" stop-color="#e0e7ff"/>
    </linearGradient>
    <linearGradient id="cg" x1="0" y1="0" x2="1" y2="1">
      <stop offset="0%" stop-color="#34d399"/><stop offset="100%" stop-color="#059669"/>
    </linearGradient>
  </defs>
  <!-- page -->
  <rect x="10" y="6" width="38" height="50" rx="5" fill="url(#pg)"/>
  <!-- folded corner -->
  <path d="M40 6 L48 14 L40 14 Z" fill="#c7d2fe"/>
  <!-- header lines -->
  <rect x="16" y="14" width="18" height="3" rx="1.5" fill="#94a3b8"/>
  <rect x="16" y="21" width="24" height="2.5" rx="1.25" fill="#cbd5e1"/>
  <!-- colourful data bars -->
  <rect x="16" y="29" width="6"  height="14" rx="2" fill="#f59e0b"/>
  <rect x="25" y="33" width="6"  height="10" rx="2" fill="#3b82f6"/>
  <rect x="34" y="26" width="6"  height="17" rx="2" fill="#a855f7"/>
  <!-- baseline -->
  <rect x="16" y="45" width="24" height="2" rx="1" fill="#e2e8f0"/>
  <!-- green check badge -->
  <circle cx="47" cy="45" r="12" fill="url(#cg)" stroke="#ffffff" stroke-width="2.5"/>
  <path d="M41.5 45.5 L45.5 49.5 L53 41.5" stroke="#ffffff" stroke-width="3.2"
        stroke-linecap="round" stroke-linejoin="round" fill="none"/>
</svg>
"""

st.markdown(f"""
<div class="hero">
  <div class="hero-row">
    <div class="hero-logo">{LOGO_SVG}</div>
    <div class="hero-txt">
      <h1>Invoice Data Extraction Tool</h1>
      <p>Scanned, photographed or digital invoices &mdash; read by AI, exported for Tally &amp; Zoho Books.</p>
    </div>
  </div>
</div>
""", unsafe_allow_html=True)

if "api_key" not in st.session_state:
    st.session_state.api_key = ""
if "uploader_key" not in st.session_state:
    st.session_state.uploader_key = 0
if "results" not in st.session_state:
    st.session_state.results = None

try:
    owner_key = st.secrets.get("GEMINI_API_KEY", "")
except Exception:
    owner_key = ""

with st.sidebar:
    st.header("\U0001F511 Setup")
    if owner_key:
        st.success("A shared key is already set up \u2713")
    else:
        with st.form("key_form"):
            key_input = st.text_input("Gemini API key", type="password")
            if st.form_submit_button("\U0001F4BE Save key"):
                st.session_state.api_key = key_input.strip()
        if st.session_state.api_key:
            st.success(f"Key saved \u2713 (...{st.session_state.api_key[-4:]})")
        else:
            st.info("Paste your key, then click Save key.")

api_key = owner_key or st.session_state.api_key

# ---------- SIDE BY SIDE: upload (left) | actions (right) ----------
left, right = st.columns([2, 1])

with left:
    st.markdown('<div class="panel-h">\U0001F4C1 Upload invoices</div>', unsafe_allow_html=True)
    uploaded_files = st.file_uploader(
        "Drag and drop invoices here, or click Browse (PDF, JPG, PNG)",
        type=["pdf", "jpg", "jpeg", "png"],
        accept_multiple_files=True,
        key=f"uploader_{st.session_state.uploader_key}",
    )

with right:
    st.markdown('<div class="panel-h">\u2699\uFE0F Actions</div>', unsafe_allow_html=True)
    action_box = st.container()

if uploaded_files and not api_key:
    st.warning("Please add your Gemini API key in the sidebar first.")

# ---------- Processing ----------
if uploaded_files and api_key:
    invoice_rows, all_items, tally_rows, zoho_rows = [], [], [], []
    not_invoices, failures = [], []

    progress = st.progress(0.0)
    status = st.empty()
    total = len(uploaded_files)
    batch_start = time.time()

    for i, f in enumerate(uploaded_files, start=1):
        file_start = time.time()
        elapsed = int(time.time() - batch_start)
        status.info(f"\u23F3 Processing {i} of {total}: **{f.name}** \u2014 elapsed {elapsed}s")
        try:
            data = extract_with_ai(api_key, f.getvalue(), mime_for(f.name))
            if not data.get("is_invoice", False):
                not_invoices.append((f.name, data.get("document_type") or "Not recognised"))
            else:
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
        time.sleep(1)

    took = int(time.time() - batch_start)
    status.empty()
    progress.empty()

    st.session_state.results = {
        "invoices": invoice_rows, "items": all_items,
        "tally": tally_rows, "zoho": zoho_rows,
        "not_invoices": not_invoices, "failures": failures,
        "total": total, "took": took,
    }

# ---------- Results ----------
res = st.session_state.results
if res:
    took, total = res["took"], res["total"]
    inv, items = res["invoices"], res["items"]
    not_inv, fails = res["not_invoices"], res["failures"]

    st.caption(f"\u23F1\uFE0F Finished in **{took} seconds** ({total} file(s)).")

    if not_inv:
        st.warning(f"\U0001F6AB {len(not_inv)} file(s) are not invoices \u2014 skipped:")
        for name, dtype in not_inv:
            st.write(f"- **{name}** \u2014 looks like a **{dtype}**, not an invoice.")

    if fails:
        st.error(f"{len(fails)} file(s) could not be read:")
        for name, err in fails:
            st.write(f"- **{name}** \u2014 {err}")

    if inv:
        inv_df = pd.DataFrame(inv)
        items_df = pd.DataFrame(items) if items else pd.DataFrame()
        tally_df = pd.DataFrame(res["tally"])
        zoho_df = pd.DataFrame(res["zoho"])
        excel_file = build_excel(inv_df, items_df, tally_df, zoho_df)

        with action_box:
            st.download_button(
                "\u2b07\uFE0F  Download Excel",
                data=excel_file,
                file_name="invoices_extracted.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
            if st.button("\U0001F5D1\uFE0F  Clear all"):
                st.session_state.uploader_key += 1
                st.session_state.results = None
                st.rerun()
            st.caption("4 sheets: Invoices, Line Items, Tally, Zoho.")

        c1, c2, c3, c4 = st.columns(4)
        for col, icon, num, lbl in [
            (c1, "\U0001F4C4", len(inv), "Invoices"),
            (c2, "\U0001F9FE", len(items), "Line items"),
            (c3, "\U0001F6AB", len(not_inv), "Not invoices"),
            (c4, "\u26A0\uFE0F", len(fails), "Failed"),
        ]:
            col.markdown(f'<div class="stat-card"><div style="font-size:1.4rem">{icon}</div>'
                         f'<div class="stat-num">{num}</div>'
                         f'<div class="stat-lbl">{lbl}</div></div>', unsafe_allow_html=True)
        st.write("")

        t1, t2, t3, t4 = st.tabs(["\U0001F4CB Invoices", "\U0001F4E6 Line Items",
                                  "\U0001F4D2 Tally Import", "\U0001F4D7 Zoho Books Import"])
        with t1:
            st.dataframe(inv_df, use_container_width=True)
        with t2:
            st.dataframe(items_df, use_container_width=True)
        with t3:
            st.caption("Purchase vouchers \u2014 one row per line item, DD-MM-YYYY dates, plain numbers. "
                       "TallyPrime: Gateway > Import Data > Vouchers.")
            st.dataframe(tally_df, use_container_width=True)
        with t4:
            st.caption("Bills \u2014 one row per line item. "
                       "Zoho Books: Purchases > Bills > More > Import Bills.")
            st.dataframe(zoho_df, use_container_width=True)
    else:
        with action_box:
            if st.button("\U0001F5D1\uFE0F  Clear all"):
                st.session_state.uploader_key += 1
                st.session_state.results = None
                st.rerun()
