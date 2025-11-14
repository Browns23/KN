#!/usr/bin/env python3
# app.py – Universal Streamlit UI for Invoice Extraction

import io
from datetime import datetime
from typing import Dict, Any, List

import pandas as pd
import streamlit as st
from openpyxl import Workbook

from parse_logic import parse_invoice_pdf_bytes, extract_invoice_id, HEADERS


# ──────────────────────────────────────────────
# STREAMLIT PAGE
# ──────────────────────────────────────────────
st.set_page_config(
    page_title="Universal Invoice Processor – A→Z",
    page_icon="📄",
    layout="wide",
)

st.title("📄 Universal Freight Invoice Processor – A→Z")
st.caption("Upload ANY freight invoice → Auto-extract → Preview → Download Excel")

uploads = st.file_uploader(
    "Upload PDF invoice files",
    type=["pdf"],
    accept_multiple_files=True
)

MAX_MB = 25
too_big = False

if uploads:
    for f in uploads:
        if f.size > MAX_MB * 1024 * 1024:
            st.error(f"❌ {f.name} is larger than {MAX_MB} MB (limit: {MAX_MB}MB)")
            too_big = True

extract_btn = st.button("Extract Invoices", type="primary",
                        disabled=(not uploads or too_big))

# ──────────────────────────────────────────────
# EXTRACTION
# ──────────────────────────────────────────────
if extract_btn and uploads and not too_big:
    rows: List[Dict[str, Any]] = []
    progress = st.progress(0)
    msg = st.empty()

    total = len(uploads)
    for i, f in enumerate(uploads, start=1):
        msg.write(f"Processing **{f.name}**...")
        data = f.read()

        invoice_id = extract_invoice_id(f.name)
        row = parse_invoice_pdf_bytes(data, filename=invoice_id)

        if row:
            rows.append(row)
        else:
            st.warning(f"⚠️ Unable to extract data from {f.name}")

        progress.progress(i / total)

    if not rows:
        st.error("❌ No valid invoice data extracted.")
    else:
        df = pd.DataFrame(rows)

        # Ensure all required columns exist
        for col in HEADERS:
            if col not in df.columns:
                df[col] = None

        df = df[HEADERS]

        # Show table
        st.subheader("Preview Extracted Data")
        st.dataframe(df, use_container_width=True)

        # ───────── Excel export ─────────
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
            file_name=f"Invoice_Summary_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
