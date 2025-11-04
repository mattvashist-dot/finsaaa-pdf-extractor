import io, re, traceback
from datetime import datetime
from typing import List, Dict, Optional

import pandas as pd
import streamlit as st
from pdfminer.high_level import extract_text

# ========================
# App config / constants
# ========================
st.set_page_config(page_title="FINSA PDF → CSV (v9)", layout="centered")
st.title("FINSA PDF → Excel (CSV) – v9")
st.caption("Upload up to 100 FINSA PDFs, extract structured data to your mapping, and download one combined CSV.")

MAX_FILES = 100
MAX_FILE_MB = 25

# Default columns (keeps mapping upload open; order used if no mapping CSV provided)
DEFAULT_COLS: List[str] = [
    "ReferralManager","ReferralEmail","Brand","QuoteNumber","QuoteDate",
    "Company","FirstName","LastName","ContactEmail","ContactPhone",
    "CurrencyCode","ContactName","Country","manufacturer_Name","item_id",
    "item_desc","Quantity","TotalSales","PDF","CustomerNumber",
    "UnitSales","Unit_Cost","sales_cost","cust_type","QuoteComment",
    "Created_By","quote_line_no","DemoQuote"
]

# ========================
# Helpers
# ========================
@st.cache_data(show_spinner=False)
def _extract_text(file_bytes: bytes) -> str:
    try:
        return extract_text(io.BytesIO(file_bytes)) or ""
    except Exception:
        return ""

def _fmt_phone(s: str) -> str:
    s = (s or "").strip()
    digits = re.sub(r"\D", "", s)
    if len(digits) >= 10:
        out = f"{digits[:3]}-{digits[3:6]}-{digits[6:10]}"
        if len(digits) > 10:
            out += f" x{digits[10:]}"
        return out
    return s

def _num_clean(s: str) -> str:
    if not s:
        return ""
    return re.sub(r"[ ,]", "", s)

# ========================
# Regex patterns (use flags arg, avoid inline (?i) in alternations)
# ========================
DATE_RX = re.compile(r"\b([0-3]\d/[01]\d/\d{4})\b")       # dd/mm/yyyy
HEADER_QUOTE_HINT = re.compile(r"\bCotizaci[oó]n\b", re.I)

QNUM_RXS = [
    re.compile(r"(?:\bN[uú]mero(?:\s+de\s+cotizaci[oó]n)?|Numero(?:\s+de\s+cotizacion)?|No\.?|N°)\s*:?\s*\n?(\d{5,8})", re.I),
    re.compile(r"(\d{5,8})\s*\nN[uú]mero\s*:", re.I),
]
COMPANY_RX = re.compile(r"(?:Cliente\s*:\s*([^\n]+)|Cliente\s*\n([^\n]+))", re.I)

# NEW: Prefer Contacto: <name> on the same line. Fallback to AT'N on next line.
CONTACTO_LINE_RX = re.compile(r"Contacto\s*:\s*([^\n]+)", re.I)
ATN_NEXTLINE_RX  = re.compile(r"AT'N\s*\n([^\n]+)", re.I)

PHONE_RX   = re.compile(r"Tel[eé]fono\s*:\s*([^\n]+)", re.I)
CUR_RX     = re.compile(r"Moneda\s*:\s*([A-Z]{3})", re.I)
SIGN_RX    = re.compile(r"Atentamente\s*\n([A-ZÁÉÍÓÚÑ ]{3,})", re.I)

# Item block: supports either header style (MODELO... or ARTICULO...IMPORTE...)
ITEM_BLOCK_RX = re.compile(
    r"(?is)(?:MODELO.*?CANTIDAD.*?UNIDAD.*?\n|ARTICULO.*?IMPORTE.*?\n)(.*?)(?:\nSub[\s-]?Total|\nSubtotal|\nSub Total)"
)

# Units that may appear near quantity
UNIT_TOKENS = r"(?:PZA|KIT|PZAS|SET|UND|PCS)"

# Money formats like 6,872.00 / 27,488.00 / 288.50 / 28.008,00
MONEY_RX = re.compile(r"\b\d{1,3}(?:[.,]\d{3})*(?:[.,]\d{2})\b")
INT_OR_DEC_RX = re.compile(r"\b\d+(?:\.\d+)?\b")
UNIT_NEAR_QTY_RX = re.compile(rf"\b(\d+(?:\.\d+)?)\s+{UNIT_TOKENS}\b", re.I)

# ========================
# Totals: pick the 3rd monetary line (Subtotal, IVA, Total)
# ========================
def find_total_third_money_line(raw_text: str) -> str:
    """
    FINSA totals block pattern:
      SubTotal
      55496.00
      IVA 16%
      8879.36
      Total
      64375.36
      43   <-- ignore (no decimals)
    Return the 3rd decimal-like monetary value = Total.
    """
    if not raw_text:
        return ""
    lines = [ln.strip() for ln in raw_text.splitlines() if ln.strip()]

    # find a starting anchor
    start_idx = None
    anchors = re.compile(r"(Sub[\s-]?Total|Subtotal|Sub Total)", re.I)
    for i, ln in enumerate(lines):
        if anchors.search(ln):
            start_idx = i
            break
    if start_idx is None:
        for i, ln in enumerate(lines):
            if re.search(r"\bTotal\b", ln, flags=re.I) and not re.search(r"Total\s*de\s*Art", ln, flags=re.I):
                start_idx = i
                break
    if start_idx is None:
        return ""

    window = lines[start_idx:start_idx + 20]

    monies = []
    for ln in window:
        for m in MONEY_RX.findall(ln):
            val = m.replace(" ", "")
            # normalize: 64,375.36 or 64.375,36 -> 64375.36
            if "," in val and "." in val:
                if val.rfind(",") > val.rfind("."):
                    val = val.replace(".", "").replace(",", ".")
                else:
                    val = val.replace(",", "")
            elif "," in val and "." not in val:
                val = val.replace(".", "").replace(",", ".")
            if "." in val:  # ignore plain integers
                monies.append(val)

    if len(monies) >= 3:
        return monies[2]
    elif monies:
        return monies[-1]
    return ""

# ========================
# Quantity parsing — robust for wrapped item rows
# ========================
def split_into_item_rows(block_text: str) -> list:
    """
    Combine wrapped lines into item rows by detecting when a line contains
    >= 2 money tokens (price & importe). We accumulate lines until we see that,
    then flush a completed item row.
    """
    if not block_text:
        return []
    raw_lines = [ln for ln in block_text.splitlines() if ln.strip()]
    rows, cur = [], []
    for ln in raw_lines:
        cur.append(ln)
        monies = MONEY_RX.findall(ln)
        if len(monies) >= 2:
            rows.append(" ".join(cur))
            cur = []
    if cur:
        if rows:
            rows[-1] += " " + " ".join(cur)
        else:
            rows.append(" ".join(cur))
    return rows

def infer_quantity_from_row(row_text: str) -> Optional[float]:
    """
    Given a consolidated row string:
      1) Prefer number immediately before a UNIT token ("7 PZA") → qty=7
      2) Else take the rightmost integer/decimal that occurs before the first money value
         (helps when units split over lines or columns are misread)
    """
    if not row_text:
        return None

    m = UNIT_NEAR_QTY_RX.search(row_text)
    if m:
        try:
            return float(m.group(1))
        except Exception:
            pass

    money_spans = list(MONEY_RX.finditer(row_text))
    if not money_spans:
        return None
    first_money_start = money_spans[0].start()

    candidates = []
    for nm in INT_OR_DEC_RX.finditer(row_text[:first_money_start]):
        tok = nm.group(0)
        try:
            v = float(tok)
            if v >= 0 and v < 100000:
                candidates.append((nm.start(), v))
        except Exception:
            continue
    if candidates:
        return candidates[-1][1]
    return None

def sum_quantities_advanced(block_text: str) -> str:
    rows = split_into_item_rows(block_text)
    total = 0.0
    found_any = False
    for r in rows:
        q = infer_quantity_from_row(r)
        if q is not None:
            total += q
            found_any = True
    if not found_any:
        return ""
    return f"{total:.2f}".rstrip("0").rstrip(".")

# ========================
# Name parsing
# ========================
def parse_first_last_from_string(full: str) -> (str, str, str):
    """
    Normalize capitalization and split first token vs the rest.
    Returns (first, last, full_clean).
    """
    full = (full or "").strip()
    if not full:
        return "", "", ""
    # Preserve common multi-word surnames; simple heuristic via title()
    full_clean = " ".join(full.split())
    full_title = full_clean.title()
    parts = [p for p in full_title.split() if p]
    first = parts[0] if parts else ""
    last = " ".join(parts[1:]) if len(parts) > 1 else ""
    return first, last, full_title

def extract_contact_names(raw_text: str) -> (str, str, str):
    """
    Prefer 'Contacto: <NAME>' on the same line.
    Fallback to AT'N on next line if Contacto is not present.
    Returns (first, last, full_contact_name).
    """
    m_contact = CONTACTO_LINE_RX.search(raw_text or "")
    if m_contact:
        first, last, full = parse_first_last_from_string(m_contact.group(1))
        return first, last, full
    m_atn = ATN_NEXTLINE_RX.search(raw_text or "")
    if m_atn:
        first, last, full = parse_first_last_from_string(m_atn.group(1))
        return first, last, full
    return "", "", ""

# ========================
# PDF parser core (implements your rules)
# ========================
def parse_pdf(file_name: str, data: bytes, out_cols: List[str]) -> Dict[str, str]:
    raw = _extract_text(data)
    lines = [ln for ln in (raw or "").splitlines() if ln.strip()]

    # QuoteNumber: prefer near "Cotización" header; else use "Número:"
    qnum = ""
    for i, ln in enumerate(lines[:25]):
        if HEADER_QUOTE_HINT.search(ln):
            window = lines[max(0, i-3):i+6]
            cand = next((x.strip() for x in window if re.fullmatch(r"\d{5,8}", x.strip())), "")
            if cand:
                qnum = cand
                break
    if not qnum:
        for rx in QNUM_RXS:
            m = rx.search(raw or "")
            if m:
                qnum = (m.group(1) or "").strip()
                if qnum:
                    break

    # QuoteDate: dd/mm/yyyy → mm/dd/yyyy
    qdate = ""
    m_date = DATE_RX.search(raw or "")
    if m_date:
        try:
            qdate = datetime.strptime(m_date.group(1), "%d/%m/%Y").strftime("%m/%d/%Y")
        except Exception:
            qdate = ""

    # Company (keep numeric prefix if present)
    company = ""
    m_co = COMPANY_RX.search(raw or "")
    if m_co:
        company = (m_co.group(1) or m_co.group(2) or "").strip()

    # Contact name (NEW logic: Contacto: preferred; fallback AT'N)
    first_name, last_name, contact_full = extract_contact_names(raw)

    # Phone
    phone = ""
    m_phone = PHONE_RX.search(raw or "")
    if m_phone:
        phone = _fmt_phone(m_phone.group(1))

    # Currency
    currency = ""
    m_cur = CUR_RX.search(raw or "")
    if m_cur:
        currency = (m_cur.group(1) or "").upper()

    # ReferralManager from Atentamente
    referral_mgr = ""
    m_sig = SIGN_RX.search(raw or "")
    if m_sig:
        referral_mgr = (m_sig.group(1) or "").title().strip()

    # TotalSales = 3rd monetary line in totals section
    total = find_total_third_money_line(raw)

    # Items: sum quantities; blank item_id/item_desc if multiple items
    item_id = item_desc = ""
    qty_total = ""
    multi = False
    m_block = ITEM_BLOCK_RX.search(raw or "")
    if m_block:
        block = m_block.group(1) or ""
        rows = split_into_item_rows(block)
        multi = len([r for r in rows if r.strip()]) > 1
        qty_total = sum_quantities_advanced(block)

        if not multi and rows:
            first_line = rows[0]
            m_model = re.search(r"([A-Z0-9][A-Z0-9\-]{5,})", first_line)
            if m_model:
                cand = m_model.group(1)
                if cand not in {"KIT","PZA","SET","UND","PCS"} and len(cand) >= 6:
                    item_id = cand
            desc = first_line
            if item_id:
                desc = desc.replace(item_id, "").strip()
            desc = re.sub(r"\s+\d+(?:\.\d{2})?\s*(?:PZA|KIT|PZAS|SET|UND|PCS)?\s*[\d, ]*\.\d{2}.*$", "", desc).strip()
            item_desc = desc
        else:
            item_id = ""
            item_desc = ""

    # PDF name
    pdf_name = f"FINSA_{qnum}.pdf" if qnum else (file_name or "")

    # Build row according to mapping
    row = {col: "" for col in out_cols}
    def setcol(col, val):
        if col in row:
            row[col] = val

    setcol("ReferralManager", referral_mgr)
    setcol("ReferralEmail", "")                      # per rule: always blank
    setcol("Brand", "Finsa")
    setcol("QuoteNumber", qnum)
    setcol("QuoteDate", qdate)
    setcol("Company", company)
    setcol("FirstName", first_name)
    setcol("LastName", last_name)
    setcol("ContactEmail", "")
    setcol("ContactPhone", phone)
    setcol("CurrencyCode", currency)
    setcol("ContactName", contact_full.strip())
    setcol("Country", "Mexico" if currency == "MXN" else "")
    setcol("manufacturer_Name", "FINSA")
    setcol("item_id", item_id)
    setcol("item_desc", item_desc)
    setcol("Quantity", qty_total)
    setcol("TotalSales", total)
    setcol("PDF", pdf_name)
    return row

# ========================
# Streamlit UI
# ========================
st.sidebar.header("Output Mapping")
mapping_file = st.sidebar.file_uploader(
    "Upload mapping CSV (header defines output columns & order) – optional", type=["csv"]
)
if mapping_file is not None:
    try:
        mapping_cols = pd.read_csv(mapping_file, nrows=0).columns.tolist()
    except Exception:
        st.sidebar.error("Could not read mapping header; using defaults.")
        mapping_cols = DEFAULT_COLS
else:
    mapping_cols = DEFAULT_COLS

strict = st.sidebar.checkbox(
    "Strict validation", value=True,
    help="Require QuoteNumber, Company, QuoteDate, TotalSales."
)

files = st.file_uploader("Upload up to 100 FINSA PDF quotes", type=["pdf"], accept_multiple_files=True)
if files and len(files) > MAX_FILES:
    st.warning(f"You selected {len(files)} files. Only the first {MAX_FILES} will be processed.")
    files = files[:MAX_FILES]

if st.button("🔄 Extract to CSV", use_container_width=True):
    if not files:
        st.warning("Please upload at least one PDF.")
        st.stop()

    rows, errors = [], []
    with st.spinner("Parsing PDFs…"):
        for f in files:
            if f.size > MAX_FILE_MB * 1024 * 1024:
                errors.append(f"{f.name}: exceeds {MAX_FILE_MB} MB limit.")
                continue
            try:
                row = parse_pdf(f.name, f.read(), mapping_cols)
                rows.append(row)
            except Exception as e:
                rows.append({"PDF": f.name, "_ERROR": f"{type(e).__name__}: {e}",
                             "_TRACE": traceback.format_exc()[:1500]})
                errors.append(f"{f.name}: {e}")

    if errors:
        st.error("\n".join(errors))

    if rows:
        df = pd.DataFrame(rows, columns=mapping_cols)

        if strict:
            problems = []
            for idx, r in df.iterrows():
                for col in ["QuoteNumber","Company","QuoteDate","TotalSales"]:
                    if col in df.columns and not str(r.get(col, "")).strip():
                        problems.append(f"Row {idx+1}: Missing '{col}'")
            if problems:
                st.error("Validation failed. Please review:")
                st.code("\n".join(problems), language="text")
                st.stop()

        st.success(f"Parsed {len(rows)} file(s) successfully.")
        st.dataframe(df, use_container_width=True)

        csv_bytes = df.to_csv(index=False).encode("utf-8-sig")
        st.download_button("⬇️ Download CSV", data=csv_bytes,
                           file_name="finsa_parsed.csv", mime="text/csv",
                           use_container_width=True)

st.markdown("---")
st.caption("v9: First/Last name now taken from 'Contacto: <NAME>' (preferred), with AT'N as fallback; Quantity parser (multi-line) and TotalSales (3rd monetary line) remain in place.")
