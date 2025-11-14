# parse_logic.py – UNIVERSAL FREIGHT INVOICE PARSER
import io
import re
from datetime import datetime
from typing import Dict, Any, Optional, List

import pdfplumber


# ───────────────────────────────────────────────
# FIXED HEADERS (NO CHANGES PER USER REQUEST)
# ───────────────────────────────────────────────
HEADERS: List[str] = [
    "Timestamp", "Filename", "Invoice_Date", "Currency", "Shipper",
    "Weight_KG", "Volume_M3", "Chargeable_KG", "Chargeable_CBM",
    "Pieces", "Subtotal", "Freight_Mode", "Freight_Rate",
]

# ───────────────────────────────────────────────
# HELPERS
# ───────────────────────────────────────────────
_f = lambda s: float(s.replace(",", "")) if s else None
_to_kg = lambda val, unit: val if unit.lower().startswith("kg") else (val * 0.453592)

# STANDARD CURRENCY DETECTOR
CURRENCY_ANY = re.compile(r"\b(USD|CAD|EUR|GBP|AUD)\b", re.I)

# EXCHANGE RATE DETECTOR
# Examples captured:
# "1.16993", "Rate: 0.71736526", "1 EUR = 1.16993 USD"
EX_RATE_PAT = re.compile(
    r"(\d+\.\d{4,6})\s*(?:USD|CAD|EUR|GBP|AUD)?\b", re.I
)

# SHIPPER SECTION (multiple patterns)
SHIPPER_PAT = re.compile(
    r"(?:SHIPPER|EXPORTER|VENDOR|SUPPLIER)\s*[:\-]?\s*(.+?)(?:CONSIGNEE|NOTIFY|DELIVERY|BILL TO|$)",
    re.I | re.S
)

# INVOICE DATE PATTERNS
DATE_PATS = [
    re.compile(r"INVOICE\s*NO\.?.*?(\d{1,2}[./-]\d{1,2}[./-]\d{2,4})", re.I | re.S),
    re.compile(r"DATE[:\s]+(\d{1,2}[./-]\d{1,2}[./-]\d{2,4})", re.I),
    re.compile(r"\b(\d{1,2}\s*[A-Za-z]{3,}\s*\d{2,4})", re.I),
]

# INVOICE NUMBER FROM FILENAME
def extract_invoice_id(filename: str):
    fn = filename.upper()
    m = re.search(r"(SY\d+[A-Z]?)", fn)
    if m: return m.group(1)
    m = re.search(r"(\d{4,8}[A-Z]?)", fn)
    if m: return m.group(1)
    return filename

# FREIGHT MODE + CHARGE
AIR_FRT = re.compile(r"AIR(?:\s*FREIGHT|\s*CHARGE|\s*RATE)?.*?([\d,]+\.\d{2})", re.I)
SEA_FRT = re.compile(r"(?:SEA|OCEAN).*?FREIGHT.*?([\d,]+\.\d{2})", re.I)
ROAD_FRT = re.compile(r"(?:TRUCK|ROAD|LTL|FTL).*?([\d,]+\.\d{2})", re.I)

# WEIGHT / VOLUME / PIECES PATTERNS
GW_PAT = re.compile(r"GROSS\s*W(?:EIGHT)?\s*[:\-]?\s*([\d,.]+)\s*KG", re.I)
VOL_PAT = re.compile(r"(?:VOLUME|CBM|M3)\s*[:\-]?\s*([\d,.]+)", re.I)
CH_W_PAT = re.compile(r"(?:CHARGEABLE|CH\.?W)\s*[:\-]?\s*([\d,.]+)\s*(KG|LB|M3|CBM)", re.I)
PIECES_PAT = re.compile(r"(?:PIECES|PCS|CTN|CARTONS)\s*[:\-]?\s*([\d,]+)", re.I)

# SUBTOTAL
SUBTOTAL_PAT = re.compile(r"(?:SUBTOTAL|TOTAL\s*USD|TOTAL)\s*[:\-]?\s*([\d,]+\.\d{2})", re.I)


# ───────────────────────────────────────────────
# MAIN PARSER
# ───────────────────────────────────────────────
def parse_invoice_pdf_bytes(data: bytes, filename: str) -> Optional[Dict[str, Any]]:
    try:
        with pdfplumber.open(io.BytesIO(data)) as pdf:
            text = "\n".join(page.extract_text() or "" for page in pdf.pages)

        # ───────────── INVOICE DATE ─────────────
        inv_date = None
        for pat in DATE_PATS:
            m = pat.search(text)
            if m:
                inv_date = m.group(1).strip()
                break

        # ───────────── SHIPPER ─────────────
        shipper = None
        m = SHIPPER_PAT.search(text)
        if m:
            shipper = re.sub(r"\s+", " ", m.group(1).strip())

        # ───────────── Pieces / Weight / Volume / Chargeable ─────────────
        pieces = w_kg = v_m3 = c_kg = c_cbm = None

        m = PIECES_PAT.search(text)
        if m: pieces = int(m.group(1).replace(",", ""))

        m = GW_PAT.search(text)
        if m: w_kg = _f(m.group(1))

        m = VOL_PAT.search(text)
        if m: v_m3 = _f(m.group(1))

        m = CH_W_PAT.search(text)
        if m:
            val = _f(m.group(1))
            unit = m.group(2)
            if unit.lower().startswith(("kg", "lb")):
                c_kg = _to_kg(val, unit)
            else:
                c_cbm = val

        # ───────────── Currency ─────────────
        currency = "USD"
        m = CURRENCY_ANY.search(text)
        if m:
            currency = m.group(1).upper()

        # ───────────── Freight Mode + Freight Rate ─────────────
        f_mode = f_rate = None

        if AIR_FRT.search(text):
            f_mode = "Air"
            f_rate = _f(AIR_FRT.search(text).group(1))
        elif SEA_FRT.search(text):
            f_mode = "Sea"
            f_rate = _f(SEA_FRT.search(text).group(1))
        elif ROAD_FRT.search(text):
            f_mode = "Road"
            f_rate = _f(ROAD_FRT.search(text).group(1))

        # ───────────── Subtotal (Before Conversion) ─────────────
        subtotal_raw = None
        m = SUBTOTAL_PAT.search(text)
        if m:
            subtotal_raw = _f(m.group(1))

        # ───────────── Exchange Rate Detection ─────────────
        exchange_rate = None
        # KN and many invoices show multiple exchange rates; pick the most realistic (0.5–2.0)
        rates = [float(x.group(1)) for x in EX_RATE_PAT.finditer(text)]
        rates_clean = [r for r in rates if 0.3 < r < 2.5]  # filter random big numbers

        if rates_clean:
            exchange_rate = rates_clean[0]  # first good conversion ratio

        # ───────────── Subtotal Conversion to USD ─────────────
        subtotal = subtotal_raw
        if subtotal_raw and exchange_rate:
            # Convert to USD ONLY if invoice supplies exchange rate
            if currency != "USD":
                subtotal = round(subtotal_raw * exchange_rate, 2)
                currency = "USD"   # Update because converted

        # ───────────── FINAL RETURN ─────────────
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
            "Subtotal": subtotal,
            "Freight_Mode": f_mode,
            "Freight_Rate": f_rate,
        }

    except Exception:
        return None
