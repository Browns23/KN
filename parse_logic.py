# parse_logic.py – KN-only OCR invoice parser

import io
import re
from datetime import datetime
from typing import Optional, Dict, Any

from pdf2image import convert_from_bytes
import easyocr

# OCR reader (english)
reader = easyocr.Reader(['en'], gpu=False)

HEADERS = [
    "Timestamp", "Filename", "Invoice_Date", "Currency", "Shipper",
    "Weight_KG", "Volume_M3", "Chargeable_KG", "Chargeable_CBM",
    "Pieces", "Subtotal", "Freight_Mode", "Freight_amount",
]

def extract_text_ocr(data: bytes) -> str:
    """Extract text from scanned KN invoices using OCR."""
    pages = convert_from_bytes(data)
    full = ""
    for p in pages:
        txt = reader.readtext(p, detail=0, paragraph=True)
        full += "\n".join(txt) + "\n"
    return full


def parse_invoice_pdf_bytes(data: bytes, filename: str) -> Optional[Dict[str, Any]]:
    """Parse a scanned KN invoice using OCR only (no pdfplumber)."""
    text = extract_text_ocr(data)

    # Normalize text for easier regex
    t = text.replace("\n", " ").upper()

    # Invoice Date
    m = re.search(r"INVOICE NO.? / DATE\s*(\d{1,2}\.\d{1,2}\.\d{4})", t)
    invoice_date = m.group(1) if m else None

    # Invoice Number
    m = re.search(r"INVOICE NO.? / DATE\s*(\d+)", t)
    inv_no = m.group(1) if m else filename

    # Shipper
    m = re.search(r"SHIPPER\s+([A-Z0-9 .,&/-]+?)\s+CONSIGNEE", t)
    shipper = m.group(1).strip() if m else None

    # Weight KG
    m = re.search(r"(\d+\.\d+)\s*KG", t)
    weight = float(m.group(1)) if m else None

    # Volume CBM
    m = re.search(r"(\d+\.\d+)\s*(CBM|M3)", t)
    volume = float(m.group(1)) if m else None

    # Pieces
    m = re.search(r"(\d+)\s+ELEGANT SHOES", t)
    pieces = int(m.group(1)) if m else None

    # Chargeable weight
    m = re.search(r"CHG\.? WT\.?\s*(\d+\.\d+)", t)
    chargeable_kg = float(m.group(1)) if m else None

    # Freight (AIRFREIGHT … USD)
    m = re.search(r"AIRFREIGHT.*?(\d+\.\d+)\s*USD", t)
    freight_amount = float(m.group(1)) if m else None

    # Subtotal USD
    m = re.search(r"SUBTOTAL\s*USD\s*(\d+\.\d+)", t)
    subtotal = float(m.group(1)) if m else None

    return {
        "Timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "Filename": inv_no,
        "Invoice_Date": invoice_date,
        "Currency": "USD",
        "Shipper": shipper,
        "Weight_KG": weight,
        "Volume_M3": volume,
        "Chargeable_KG": chargeable_kg,
        "Chargeable_CBM": None,
        "Pieces": pieces,
        "Subtotal": subtotal,
        "Freight_Mode": "Air",
        "Freight_amount": freight_amount,
    }
