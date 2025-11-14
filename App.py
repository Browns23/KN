import io
import streamlit as st
import pandas as pd
from openpyxl import Workbook
from datetime import datetime

from parse_logic import parse_invoice_pdf_bytes, HEADERS

st.title("KN Invoice Extractor (OCR Enabled)")

uploads = st.file_uploader("Upload KN Invoice PDFs", type=["pdf"], accept_multiple_files=True)

if st.button("Extract") and uploads:
    rows = []
    for f in uploads:
        data = f.read()
        row = parse_invoice_pdf_bytes(data, filename=f.name)
        if row:
            rows.append(row)
        else:
            st.warning(f"Could not extract {f.name}")

    df = pd.DataFrame(rows)
    df = df[HEADERS]

    st.dataframe(df)

    out = io.BytesIO()
    wb = Workbook()
    ws = wb.active
    ws.append(HEADERS)
    for _, r in df.iterrows():
        ws.append([r[h] for h in HEADERS])
    wb.save(out)
    out.seek(0)

    st.download_button(
        "Download Excel",
        out,
        file_name=f"KN_Invoice_Summary_{datetime.now().strftime('%Y%m%d')}.xlsx"
    )
