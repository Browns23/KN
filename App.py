#!/usr/bin/env python3
# app.py – Streamlit UI for Invoice Watcher logic
#
# Upload one or more invoice PDFs → parse → preview → download Excel.

import io
import os
import re
import traceback
from datetime import datetime
from typing import Dict, Any, Optional, List

import pdfplumber
import pandas as pd
import streamlit as st
from openpyxl import Workbook

# ─────────────────────────────────────────────────────────────
# UNIVERSAL INVOICE NUMBER EXTRACTOR
# ─────────────────────────────────────────────────────────────
def extract_invoice_id(filename: str) -> str:
    """
    Extracts a cleaner invoice ID from the filename.
    """
    filename = filename.upper()

    # 1. Try SY pattern first
    m = re.search(r"(SY\d+[A-Z]?)", filename)
    if m:
        return m.group(1)

    # 2. Extract first number + optional last letter (ex: 42308, 55671A)
    m = re.search(r"(\d+[A-Z]?)", filename)
    if m:
        return m.group(1)

    # 3. Fallback: return filename without extension
    return os.path.splitext(filename)[0]


# ─────────────────────────────────────────────────────────────
# Fixed header row (Updated to include Freight_amount)
# ─────────────────────────────────────────────────────────────
HEADERS: List[str] = [
    "Timestamp", "Filename", "Invoice_Date", "Currency", "Shipper",
    "Weight_KG", "Volume_M3", "Chargeable_KG", "Chargeable_CBM",
    "Pieces", "Subtotal", "Freight_Mode", "Freight_amount",
]

# ─────────────────────────────────────────────────────────────
# Helper functions + regex (Modified for the new invoice structure and robustness)
# ─────────────────────────────────────────────────────────────

_f = lambda s: float(s.replace(",", "")) if s else None
_to_kg = lambda v, u: v if u.lower().startswith("kg") else v * 0.453592

# NEW: Pattern to capture Invoice No and Date in one go (e.g., "2400046912 14.10.2025")
INV_NO_DATE_PAT = re.compile(
    r"(?:INVOICE NO\.? / DATE|INVOICE NO)\s*\"?,?\"\s*(.+?)\s*(\d{1,2}\.\d{1,2}\.\d{2,4})", re.I
)

# Pattern to capture Shipper (assuming it's followed by a detailed address block)
SHIPPER_PAT = re.compile(
    r"SHIPPER\s*[:\n]\s*(.+?)(?:\n\s*\d|\n{2,})", re.I | re.S
)

# Pattern for pieces, gross weight (GW) and volume (VOL)
ROW_PIECES_GW_VOL_PAT = re.compile(
    r"(?:Gross Weight|G\.?\s*W\.?\s*K?\.?)\s*[:\-]?\s*([\d,.]+)\s*KGS?"
    r"\s+Volume\s*[:\-]?\s*([\d,.]+)\s+M3"
    r"(?:\s+Pieces\s*[:\-]?\s*(\d+))?", re.I
)

# MODIFIED: Pattern for chargeable weight (making 'CHARGEABLE' optional)
CHARGEABLE_KG_PAT = re.compile(
    r"CH(?:ARGEABLE)?\.?\s*W(?:EIGHT)?\s*[:\-]?\s*([\d,.]+)\s*(KG|KGS?|LB|M3|CBM)", re.I
)

# MODIFIED: Highly specific pattern for the primary Airfreight charge amount
AIRFREIGHT_CHARGE_LINE = re.compile(
    r"AIRFREIGHT\s+CHARGE.*?([\d,]+\.\d{2})", re.I
)
# Fallback for general freight/transportation (Less specific)
GENERAL_FRT_AMOUNT_PAT = re.compile(
    r"(?:BASIC\s+FREIGHT|TRANSPORTATION|FREIGHT).*?\s+([\d,]+\.\d{2})", re.I
)
SEA_FRT_AMOUNT_PAT = re.compile(
    r"(?:SEA|OCEAN)\s*FREIGHT.*?\s+([\d,]+\.\d{2})", re.I
)

# Pattern for Subtotal and Currency (Updated to use the currency group from the subtotal line)
SUBTOTAL_PAT = re.compile(
    r"Sub-?Total\s*[:\-]?\s*([A-Z]{3})\s+([\d,]+\.\d{2})", re.I
)

CURRENCY_ANY = re.compile(r"\b(CAD|USD|EUR|GBP|AUD)\b", re.I)


# ─────────────────────────────────────────────────────────────
# PARSE ONE PDF (Updated to handle new invoice fields)
# ─────────────────────────────────────────────────────────────
def parse_invoice_pdf_bytes(data: bytes, filename: str) -> Optional[Dict[str, Any]]:
    """
    Parses a single invoice PDF's text content to extract relevant financial and freight data.
    """
    try:
        with pdfplumber.open(io.BytesIO(data)) as pdf:
            # Extract text from all pages
            text = "\n".join(p.extract_text() or "" for p in pdf.pages)

        # --- 1. Invoice Date and Number (Primary source from the new format) ---
        inv_date = None
        invoice_no_from_content = None
        m = INV_NO_DATE_PAT.search(text)
        if m:
            invoice_no_from_content = m.group(1).strip()
            inv_date = m.group(2).strip()
            # If we found a better invoice number in the content, use it.
            if invoice_no_from_content:
                filename = invoice_no_from_content

        # --- 2. Currency and Subtotal (This is the total charge including all line items) ---
        currency = "USD"
        subtotal = None
        m = SUBTOTAL_PAT.search(text)
        if m:
            currency = m.group(1).upper()
            subtotal = _f(m.group(2))
        else:
            m = CURRENCY_ANY.search(text)
            if m:
                currency = m.group(1).upper()

        # --- 3. Shipper ---
        shipper = None
        m = SHIPPER_PAT.search(text)
        if m:
            # Clean up the shipper name (remove address lines, just keep the name)
            shipper_block = m.group(1).strip().split('\n')[0]
            shipper = re.sub(r"\s+", " ", shipper_block)

        # --- 4. Weight, Volume, Pieces ---
        w_kg = v_m3 = pieces = None
        m = ROW_PIECES_GW_VOL_PAT.search(text)
        if m:
            w, v, p = m.groups()
            w_kg = _f(w)
            v_m3 = _f(v)
            if p:
                pieces = int(p)

        # A fallback pattern for Pieces (if not in the main row)
        if pieces is None:
            m = re.search(r"Pieces\s*[:\-]?\s*(\d+)", text, re.I)
            if m:
                pieces = int(m.group(1))

        # --- 5. Chargeable Weight (More robust capture) ---
        c_kg = c_cbm = None
        m = CHARGEABLE_KG_PAT.search(text)
        if m:
            val, unit = m.groups()
            val = _f(val)
            if unit.lower().startswith(("kg", "lb")):
                c_kg = _to_kg(val, unit)
            elif unit.lower().startswith(("m3", "cbm")):
                c_cbm = val


        # --- 6. Freight Mode and Amount (Priority given to AIRFREIGHT CHARGE) ---
        # This captures the *main* freight line item (879.20 USD) for the Freight_amount column.
        f_mode = f_amount = None

        m_air_charge = AIRFREIGHT_CHARGE_LINE.search(text)
        if m_air_charge:
            f_mode, f_amount = "Air", _f(m_air_charge.group(1))
        else:
            m_air_general = GENERAL_FRT_AMOUNT_PAT.search(text)
            if m_air_general:
                f_mode, f_amount = "Air", _f(m_air_general.group(1))
            else:
                m_sea = SEA_FRT_AMOUNT_PAT.search(text)
                if m_sea:
                    f_mode, f_amount = "Sea", _f(m_sea.group(1))
        
        # --- 7. Final Return ---
        return {
            "Timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "Filename": filename,
            "Invoice_Date": inv_date,
            "Currency": currency,
            "Shipper": shipper,
            "Weight_KG": w_kg,
            "Volume_M3": v_m3,
            "Chargeable_KG": c_kg,
            "Chargeable_CBM": c_cbm,
            "Pieces": pieces,
            "Subtotal": subtotal, # This should be the final amount (1,183.96 USD)
            "Freight_Mode": f_mode,
            "Freight_amount": f_amount, # This should be the main line item amount (879.20 USD)
        }
    except Exception:
        # Print stack trace to Streamlit console for debugging
        traceback.print_exc()
        return None


# ─────────────────────────────────────────────────────────────
# STREAMLIT UI (Unchanged)
# ─────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Silog Invoice Processor – A→Z (Streamlit)",
    page_icon="📄",
    layout="wide",
)

st.title("📄 Silog Invoice Processor – A→Z")
st.caption(
    "Upload freight invoices → Extract Invoice Date, Shipper, Weight, Volume, "
    "Chargeable, Subtotal, Freight Mode & Amount → Download Excel summary."
)

uploads = st.file_uploader(
    "Upload PDF invoice files",
    type=["pdf"],
    accept_multiple_files=True,
    help="Drag & drop or browse invoices.",
)

MAX_MB = 25
too_big = False
if uploads:
    for f in uploads:
        if f.size > MAX_MB * 1024 * 1024:
            st.error(f"❌ {f.name} is larger than {MAX_MB} MB")
            too_big = True

extract_btn = st.button("Extract Invoices", type="primary", disabled=(not uploads or too_big))

if extract_btn and uploads and not too_big:
    rows: List[Dict[str, Any]] = []
    progress = st.progress(0)
    status = st.empty()

    total = len(uploads)
    for i, f in enumerate(uploads, start=1):
        status.write(f"Parsing: **{f.name}**")

        data = f.read()

        # 🔥 APPLY INVOICE ID CLEANING HERE
        invoice_id = extract_invoice_id(f.name)

        # Pass the extracted ID, which might be overwritten by a better ID found in content
        row = parse_invoice_pdf_bytes(data, filename=invoice_id)

        if row:
            rows.append(row)
        else:
            st.warning(f"⚠️ Could not extract data from {f.name}")

        progress.progress(i / total)

    if not rows:
        st.error("❌ No data extracted.")
    else:
        df = pd.DataFrame(rows)
        # Ensure all required columns are present, even if empty
        for col in HEADERS:
            if col not in df.columns:
                df[col] = None
        # Reorder columns according to HEADERS
        df = df[HEADERS]

        st.subheader("Preview Results")
        st.dataframe(df, use_container_width=True)

        # Build Excel in memory
        output = io.BytesIO()
        wb = Workbook()
        ws = wb.active
        ws.append(HEADERS)
        for _, r in df.iterrows():
            ws.append([r[h] for h in HEADERS])
        wb.save(output)
        output.seek(0)

        st.success(f"✅ Extraction complete! {len(rows)} invoices processed.")

        st.download_button(
            "⬇️ Download Excel",
            data=output,
            file_name="Invoice_Summary.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
