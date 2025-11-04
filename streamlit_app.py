import io, re, traceback
from datetime import datetime
from typing import List, Dict, Optional

import pandas as pd
import streamlit as st
from pdfminer.high_level import extract_text

# ========================
# App config / constants
# ========================
st.set_page_config(page_title="FINSA PDF → CSV (v14)", layout="centered")
st.title("FINSA PDF → Excel (CSV) – v14")
st.caption("Headers preserved; stronger ReferralManager & City; all earlier parsing rules kept.")

MAX_FILES = 100
MAX_FILE_MB = 25

# Your full header list (used only if no mapping CSV is uploaded)
DEFAULT_COLS: List[str] = [
    "ReferralManager","ReferralEmail","Brand","QuoteNumber","QuoteDate","Company",
    "FirstName","LastName","ContactEmail","ContactPhone","Address","County","City",
    "State","ZipCode","Country","manufacturer_Name","item_id","item_desc","Quantity",
    "TotalSales","PDF","CustomerNumber","UnitSales","Unit_Cost","sales_cost","cust_type",
    "QuoteComment","Created_By","quote_line_no","DemoQuote"
]

# Columns we actively fill from PDFs (others remain blank on purpose)
FILLED_COLS: List[str] = [
    "ReferralManager","Brand","QuoteNumber","QuoteDate","Company","FirstName",
    "LastName","ContactPhone","City","Quantity","TotalSales","PDF"
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

def _digits_only(s: str) -> str:
    return re.sub(r"\D", "", s or "")

def _fmt_phone_mx(raw: str) -> str:
    """
    Returns numbers only (no EXT). Does not enforce 10 digits (quotes vary).
    """
    if not raw:
        return ""
    # cut off at EXT or similar
    cut = re.split(r"\bEXT\.?\b", raw, flags=re.I)[0]
    cut = re.split(r"\bEXTENSI[oó]N\b", cut, flags=re.I)[0]
    digits = _digits_only(cut)
    return digits

# ========================
# Regex patterns
# ========================
DATE_RX = re.compile(r"\b([0-3]\d/[01]\d/\d{4})\b")       # dd/mm/yyyy
HEADER_QUOTE_HINT = re.compile(r"\bCotizaci[oó]n\b", re.I)

# 4–8 digits (some quotes like 8724)
QNUM_RXS = [
    re.compile(r"(?:\bN[uú]mero(?:\s+de\s+cotizaci[oó]n)?|Numero(?:\s+de\s+cotizacion)?|No\.?|N°)\s*:?\s*\n?(\d{4,8})", re.I),
    re.compile(r"(\d{4,8})\s*\nN[uú]mero\s*:", re.I),
]

CLIENTE_ANYLINE_RX    = re.compile(r"Cliente\s*:?\s*([^\n]*)", re.I)
CONTACTO_LINE_RX      = re.compile(r"Contacto\s*:\s*([^\n]+)", re.I)
ATN_WORD_RX           = re.compile(r"Atentamente\s*$", re.I)

PHONE_LINE_RXS = [
    re.compile(r"([A-Za-zÁÉÍÓÚÑñ .]+?)\s+TEL\.?\s*[:.]?\s*([^\n]*)", re.I),  # captures City (group1) + rest (group2)
    re.compile(r"([A-Za-zÁÉÍÓÚÑñ .]+?)\s+Tel[eé]fono\s*[:.]?\s*([^\n]*)", re.I),
    re.compile(r"\bTEL\.?\s*[:.]?\s*([^\n]+)", re.I),  # phone only (no city)
    re.compile(r"Tel[eé]fono\s*[:.]?\s*([^\n]+)", re.I),
]

CUR_RX                = re.compile(r"Moneda\s*:\s*([A-Z]{3})", re.I)

# Item block: supports either header style (MODELO... or ARTICULO...IMPORTE...)
ITEM_BLOCK_RX = re.compile(
    r"(?is)(?:MODELO.*?CANTIDAD.*?UNIDAD.*?\n|ARTICULO.*?IMPORTE.*?\n)(.*?)(?:\nSub[\s-]?Total|\nSubtotal|\nSub Total)"
)

UNIT_TOKENS = r"(?:PZA|KIT|PZAS|SET|UND|PCS)"
MONEY_RX    = re.compile(r"\b\d{1,3}(?:[.,]\d{3})*(?:[.,]\d{2})\b")
INT_OR_DEC_RX = re.compile(r"\b\d+(?:\.\d+)?\b")
UNIT_NEAR_QTY_RX = re.compile(rf"\b(\d+(?:\.\d+)?)\s+{UNIT_TOKENS}\b", re.I)

# ========================
# Totals: pick the 3rd monetary line (Subtotal, IVA, Total)
# ========================
def find_total_third_money_line(raw_text: str) -> str:
    if not raw_text:
        return ""
    lines = [ln.strip() for ln in raw_text.splitlines() if ln.strip()]

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
            if "," in val and "." in val:
                if val.rfind(",") > val.rfind("."):
                    val = val.replace(".", "").replace(",", ".")
                else:
                    val = val.replace(",", "")
            elif "," in val and "." not in val:
                val = val.replace(".", "").replace(",", ".")
            if "." in val:
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
            if 0 <= v < 100000:
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
# Name parsing & extraction
# ========================
def clean_name_line(s: str) -> str:
    s = (s or "").strip()
    if not s:
        return ""
    # Skip obvious footer/link noise
    if re.search(r"(visita|www|http|https|\.com)", s, flags=re.I):
        return ""
    # Strip leading/trailing numbers and punctuation
    s = re.sub(r"^[\d\W_]+\s+", "", s, flags=re.UNICODE)
    s = re.sub(r"\s+[\d\W_]+$", "", s, flags=re.UNICODE)
    # Keep words containing letters (incl accents/Ñ)
    tokens = [t for t in s.split() if re.search(r"[A-Za-zÁÉÍÓÚáéíóúÑñ]", t)]
    return " ".join(tokens).title() if tokens else ""

def extract_contact_names(raw_text: str) -> (str, str, str):
    m_contact = CONTACTO_LINE_RX.search(raw_text or "")
    if m_contact:
        full = clean_name_line(m_contact.group(1))
        parts = full.split() if full else []
        first = parts[0] if parts else ""
        last  = " ".join(parts[1:]) if len(parts) > 1 else ""
        if first or last:
            return first, last, full
    # legacy fallback (AT'N next line)
    m_atn_next = re.search(r"AT'N\s*\n([^\n]+)", raw_text or "", flags=re.I)
    if m_atn_next:
        full = clean_name_line(m_atn_next.group(1))
        parts = full.split() if full else []
        return (parts[0] if parts else "", " ".join(parts[1:]) if len(parts) > 1 else "", full)
    return "", "", ""

def extract_referral_manager(raw_text: str) -> str:
    """
    Find the *last* occurrence of 'Atentamente' anywhere (bottom/center aligned).
    Take the next non-empty line as the name, cleaned of numbers and junk.
    Try up to 6 lines below to handle tiny text or line breaks.
    """
    if not raw_text:
        return ""
    lines = [ln.rstrip() for ln in raw_text.splitlines()]
    last_idx = None
    for i, ln in enumerate(lines):
        if ATN_WORD_RX.search(ln or ""):
            last_idx = i
    if last_idx is None:
        return ""
    for j in range(last_idx + 1, min(last_idx + 7, len(lines))):
        candidate = clean_name_line(lines[j])
        if candidate:
            return candidate
    return ""

# ========================
# Quote number extraction (includes 'Cotización N04     8724')
# ========================
def extract_quote_number(lines: list, raw_text: str) -> str:
    for idx, ln in enumerate(lines[:40]):  # header zone
        if HEADER_QUOTE_HINT.search(ln):
            window = [ln]
            for k in range(1, 6):
                if idx + k < len(lines):
                    window.append(lines[idx + k])
            joined = " ".join(window)
            nums = re.findall(r"\b(\d{4,8})\b", joined)
            if nums:
                return nums[-1].lstrip("0") or nums[-1]
    for rx in QNUM_RXS:
        m = rx.search(raw_text or "")
        if m:
            cand = (m.group(1) or "").strip()
            if re.fullmatch(r"\d{4,8}", cand):
                return cand.lstrip("0") or cand
    top = " ".join(lines[:40])
    m = re.search(r"\b(\d{4,8})\b", top)
    if m:
        return (m.group(1) or "").lstrip("0") or m.group(1)
    return ""

# ========================
# Company extraction (robust)
# ========================
def extract_company(lines: list, raw_text: str) -> str:
    for i, ln in enumerate(lines[:80]):
        m = CLIENTE_ANYLINE_RX.search(ln)
        if m:
            same = (m.group(1) or "").strip()
            next_line = lines[i+1].strip() if (not same or len(same) < 3) and i + 1 < len(lines) else ""
            candidate = same or next_line
            candidate = re.sub(r"\s+", " ", candidate).strip()
            if re.search(r"(Contacto|Vendedor|Tel[eé]fono|Moneda|No\.|N°|Fecha|Atentamente)", candidate, re.I):
                continue
            if re.search(r"(visita|www|http|https|\.com)", candidate, re.I):
                continue
            if re.search(r"[A-Za-zÁÉÍÓÚáéíóúÑñ]", candidate):
                return candidate
    m2 = re.search(r"Cliente\s*:?\s*([^\n]+)\n?([^\n]*)", raw_text or "", flags=re.I)
    if m2:
        cand = (m2.group(1) or "").strip()
        if not cand or len(cand) < 3:
            cand = (m2.group(2) or "").strip()
        cand = re.sub(r"\s+", " ", cand).strip()
        if not re.search(r"(visita|www|http|https|\.com)", cand, flags=re.I):
            return cand
    return ""

# ========================
# City + Phone extraction
# ========================
def extract_city_and_phone(raw_text: str) -> (str, str):
    """
    City = text before TEL./Teléfono on the same line (letters & spaces, normalized).
    Phone = digits only from the right side; EXT removed.
    We scan for city+TEL lines first, then fall back to phone-only lines.
    """
    # Try lines that include city before TEL / Teléfono
    for rx in PHONE_LINE_RXS[:2]:  # first two patterns capture city + trailing text
        for m in rx.finditer(raw_text or ""):
            city_raw = (m.group(1) or "").strip()
            right    = (m.group(2) or "").strip()
            city = re.sub(r"\s+", " ", city_raw).strip(" .,:;-")
            # avoid false positives like "Observaciones:" etc.
            if not re.search(r"[A-Za-zÁÉÍÓÚáéíóúÑñ]", city):
                city = ""
            phone = _fmt_phone_mx(right)
            if city or phone:
                return (city.title(), phone)
    # Fallback: phone-only line patterns
    for rx in PHONE_LINE_RXS[2:]:
        for m in rx.finditer(raw_text or ""):
            right = (m.group(1) or "").strip()
            phone = _fmt_phone_mx(right)
            if phone:
                return ("", phone)
    return ("", "")

# ========================
# PDF parser core
# ========================
def parse_pdf(file_name: str, data: bytes, out_cols: List[str]) -> Dict[str, str]:
    raw = _extract_text(data)
    lines = [ln for ln in (raw or "").splitlines() if ln.strip()]

    qnum = extract_quote_number(lines, raw)

    qdate = ""
    m_date = DATE_RX.search(raw or "")
    if m_date:
        try:
            qdate = datetime.strptime(m_date.group(1), "%d/%m/%Y").strftime("%m/%d/%Y")
        except Exception:
            qdate = ""

    company = extract_company(lines, raw)

    first_name, last_name, contact_full = extract_contact_names(raw)

    city, phone = extract_city_and_phone(raw)

    currency = ""
    m_cur = CUR_RX.search(raw or "")
    if m_cur:
        currency = (m_cur.group(1) or "").upper()

    referral_mgr = extract_referral_manager(raw)

    total = find_total_third_money_line(raw)

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

    pdf_name = f"FINSA_{qnum}.pdf" if qnum else (file_name or "")

    # Build a row with ALL columns requested by the mapping, but only fill mapped ones.
    row = {col: "" for col in out_cols}
    def setcol(col, val):
        if col in row:
            row[col] = val

    # Fill mapped columns
    setcol("ReferralManager", referral_mgr)
    setcol("Brand", "Finsa")
    setcol("QuoteNumber", qnum)
    setcol("QuoteDate", qdate)
    setcol("Company", company)
    setcol("FirstName", first_name)
    setcol("LastName", last_name)
    setcol("ContactPhone", phone)
    setcol("City", city)
    setcol("Quantity", qty_total)
    setcol("TotalSales", total)
    setcol("PDF", pdf_name)

    # Per earlier rules
    setcol("ReferralEmail", "")
    setcol("ContactEmail", "")
    setcol("CurrencyCode", currency)
    setcol("Country", "Mexico" if currency == "MXN" else "")
    setcol("manufacturer_Name", "FINSA")

    # Leave other columns blank intentionally
    return row

# ========================
# Streamlit UI
# ========================
st.sidebar.header("Output Mapping")
mapping_file = st.sidebar.file_uploader(
    "Upload mapping CSV (header defines output columns & order) – REQUIRED to preserve exact headers, otherwise default is used.",
    type=["csv"]
)
if mapping_file is not None:
    try:
        # We only need the header for order; rows not needed.
        mapping_cols = pd.read_csv(mapping_file, nrows=0).columns.tolist()
    except Exception:
        st.sidebar.error("Could not read mapping header; falling back to default headers.")
        mapping_cols = DEFAULT_COLS
else:
    mapping_cols = DEFAULT_COLS

# QuoteNumber is review-only (no blocking)
strict = st.sidebar.checkbox(
    "Show validation review (Company, QuoteDate, TotalSales)",
    value=True,
    help="Issues are shown for review but export is always allowed. QuoteNumber is review-only."
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
        # ALWAYS build DataFrame with the mapping header order (export headers preserved).
        df = pd.DataFrame(rows, columns=mapping_cols)

        # List PDFs missing QuoteNumber (review only)
        try:
            missing_q = df[df.get("QuoteNumber", "").astype(str).str.strip() == ""]
            if not missing_q.empty:
                st.warning(
                    "QuoteNumber not found (left blank) in the following file(s):\n"
                    + "\n".join(f"- {r.get('PDF', '')}" for _, r in missing_q.iterrows())
                )
        except Exception:
            pass

        # Show non-blocking validation (optional)
        if strict:
            problems = []
            for idx, r in df.iterrows():
                for col in ["Company","QuoteDate","TotalSales"]:
                    if col in df.columns and not str(r.get(col, "")).strip():
                        problems.append(f"Row {idx+1}: Missing '{col}'")
            if problems:
                st.error("Validation review:")
                st.code("\n".join(problems), language="text")

        st.success(f"Parsed {len(rows)} file(s). Export contains ALL columns in your header order.")
        st.dataframe(df, use_container_width=True)

        csv_bytes = df.to_csv(index=False).encode("utf-8-sig")
        st.download_button("⬇️ Download CSV", data=csv_bytes,
                           file_name="finsa_parsed.csv", mime="text/csv",
                           use_container_width=True)

st.markdown("---")
st.caption("v14: Headers preserved as uploaded; ReferralManager = name after the last 'Atentamente'; City from text before TEL./Teléfono; robust Company/Quote/Quantity/Total/Phone kept.")
