import io, re, traceback
from datetime import datetime
from typing import List, Dict, Optional

import pandas as pd
import streamlit as st
from pdfminer.high_level import extract_text

# ========================
# App config / constants
# ========================
st.set_page_config(page_title="FINSA PDF → CSV (v11)", layout="centered")
st.title("FINSA PDF → Excel (CSV) – v11")
st.caption("Improved QuoteNumber from 'Cotización …' and robust ReferralManager from 'Atentamente' lines.")

MAX_FILES = 100
MAX_FILE_MB = 25

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
# Regex patterns
# ========================
DATE_RX = re.compile(r"\b([0-3]\d/[01]\d/\d{4})\b")       # dd/mm/yyyy
HEADER_QUOTE_HINT = re.compile(r"\bCotizaci[oó]n\b", re.I)

# allow 4–8 digits (some quotes are 4 digits like 8724)
QNUM_RXS = [
    re.compile(r"(?:\bN[uú]mero(?:\s+de\s+cotizaci[oó]n)?|Numero(?:\s+de\s+cotizacion)?|No\.?|N°)\s*:?\s*\n?(\d{4,8})", re.I),
    re.compile(r"(\d{4,8})\s*\nN[uú]mero\s*:", re.I),
]

COMPANY_RX = re.compile(r"(?:Cliente\s*:\s*([^\n]+)|Cliente\s*\n([^\n]+))", re.I)
CONTACTO_LINE_RX = re.compile(r"Contacto\s*:\s*([^\n]+)", re.I)
ATN_LINE_RX      = re.compile(r"Atentamente\s*$", re.I)    # line that says Atentamente
PHONE_RX         = re.compile(r"Tel[eé]fono\s*:\s*([^\n]+)", re.I)
CUR_RX           = re.compile(r"Moneda\s*:\s*([A-Z]{3})", re.I)

# Item block: supports either header style (MODELO... or ARTICULO...IMPORTE...)
ITEM_BLOCK_RX = re.compile(
    r"(?is)(?:MODELO.*?CANTIDAD.*?UNIDAD.*?\n|ARTICULO.*?IMPORTE.*?\n)(.*?)(?:\nSub[\s-]?Total|\nSubtotal|\nSub Total)"
)

UNIT_TOKENS = r"(?:PZA|KIT|PZAS|SET|UND|PCS)"
MONEY_RX = re.compile(r"\b\d{1,3}(?:[.,]\d{3})*(?:[.,]\d{2})\b")
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
# Name parsing (Contacto preferred; ReferralManager from Atentamente)
# ========================
def parse_first_last_from_string(full: str) -> (str, str, str):
    full = (full or "").strip()
    if not full:
        return "", "", ""
    full_clean = " ".join(full.split())
    # Remove leading/trailing numbers (e.g., "43 Roberto Carrera" -> "Roberto Carrera")
    full_clean = re.sub(r"^\d+\s+", "", full_clean)
    full_clean = re.sub(r"\s+\d+\s*$", "", full_clean)
    full_title = full_clean.title()
    parts = [p for p in full_title.split() if p]
    first = parts[0] if parts else ""
    last = " ".join(parts[1:]) if len(parts) > 1 else ""
    return first, last, full_title

def extract_contact_names(raw_text: str) -> (str, str, str):
    m_contact = CONTACTO_LINE_RX.search(raw_text or "")
    if m_contact:
        first, last, full = parse_first_last_from_string(m_contact.group(1))
        return first, last, full
    # Fallback: if "AT'N" layout exists (older quotes), handle it here if needed
    m_atn_next = re.search(r"AT'N\s*\n([^\n]+)", raw_text or "", flags=re.I)
    if m_atn_next:
        first, last, full = parse_first_last_from_string(m_atn_next.group(1))
        return first, last, full
    return "", "", ""

def extract_referral_manager(raw_text: str) -> str:
    """
    Find 'Atentamente' line; take the first non-empty line after it,
    strip any leading numbers, keep only the name (letters/spaces/accents),
    and title-case it.
    """
    if not raw_text:
        return ""
    lines = [ln.rstrip() for ln in raw_text.splitlines()]
    for i, ln in enumerate(lines):
        if ATN_LINE_RX.search(ln or ""):
            # get next non-empty line within a short window
            for j in range(i+1, min(i+4, len(lines))):
                candidate = (lines[j] or "").strip()
                if not candidate:
                    continue
                # remove leading/trailing numbers
                candidate = re.sub(r"^\d+\s+", "", candidate)
                candidate = re.sub(r"\s+\d+\s*$", "", candidate)
                # remove trailing punctuation
                candidate = candidate.strip(" .,:;|-")
                # keep words that are alphabetic (allow accents and Ññ)
                tokens = candidate.split()
                name_tokens = [t for t in tokens if re.search(r"[A-Za-zÁÉÍÓÚáéíóúÑñ]", t)]
                if not name_tokens:
                    continue
                name = " ".join(name_tokens).title()
                if len(name) >= 2:
                    return name
    return ""

# ========================
# Quote number extraction (improved for 'Cotización N04     8724')
# ========================
def extract_quote_number(lines: list, raw_text: str) -> str:
    # 1) Look near any line containing "Cotización"
    for idx, ln in enumerate(lines[:40]):  # restrict to header area
        if HEADER_QUOTE_HINT.search(ln):
            window = [ln]
            # include neighbors (same line and next few lines)
            for k in range(1, 6):
                if idx + k < len(lines):
                    window.append(lines[idx + k])
            joined = " ".join(window)
            # find 4–8 digit groups; take the last one (ignores things like N04)
            nums = re.findall(r"\b(\d{4,8})\b", joined)
            if nums:
                return nums[-1].lstrip("0") or nums[-1]  # keep significant digits
    # 2) Fallback to previous patterns (Número:, etc.)
    for rx in QNUM_RXS:
        m = rx.search(raw_text or "")
        if m:
            cand = (m.group(1) or "").strip()
            if re.fullmatch(r"\d{4,8}", cand):
                return cand.lstrip("0") or cand
    # 3) Final fallback: search top area for standalone 4–8 digits
    top = " ".join(lines[:40])
    m = re.search(r"\b(\d{4,8})\b", top)
    if m:
        return (m.group(1) or "").lstrip("0") or m.group(1)
    return ""

# ========================
# PDF parser core
# ========================
def parse_pdf(file_name: str, data: bytes, out_cols: List[str]) -> Dict[str, str]:
    raw = _extract_text(data)
    lines = [ln for ln in (raw or "").splitlines() if ln.strip()]

    # QuoteNumber improved
    qnum = extract_quote_number(lines, raw)

    # QuoteDate
    qdate = ""
    m_date = DATE_RX.search(raw or "")
    if m_date:
        try:
            qdate = datetime.strptime(m_date.group(1), "%d/%m/%Y").strftime("%m/%d/%Y")
        except Exception:
            qdate = ""

    # Company
    company = ""
    m_co = COMPANY_RX.search(raw or "")
    if m_co:
        company = (m_co.group(1) or m_co.group(2) or "").strip()

    # Contact names
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

    # ReferralManager from Atentamente (robust, strips numbers)
    referral_mgr = extract_referral_manager(raw)

    # TotalSales = 3rd monetary line in totals section
    total = find_total_third_money_line(raw)

    # Items / Quantity
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
    setcol("ContactName", (first_name + " " + last_name).strip() or "")
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

# QuoteNumber is review-only (no blocking)
strict = st.sidebar.checkbox(
    "Strict validation (Company, QuoteDate, TotalSales only)",
    value=True,
    help="If fields are missing, they are listed for review but export is still allowed."
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

        # Validation review (no blocking)
        if strict:
            problems = []
            required_cols = [c for c in ["Company","QuoteDate","TotalSales"] if c in df.columns]
            for idx, r in df.iterrows():
                for col in required_cols:
                    if not str(r.get(col, "")).strip():
                        problems.append(f"Row {idx+1}: Missing '{col}'")
            if problems:
                st.error("Validation review:")
                st.code("\n".join(problems), language="text")

        # Preview & Download
        st.success(f"Parsed {len(rows)} file(s). You can review issues above and still export.")
        st.dataframe(df, use_container_width=True)

        csv_bytes = df.to_csv(index=False).encode("utf-8-sig")
        st.download_button("⬇️ Download CSV", data=csv_bytes,
                           file_name="finsa_parsed.csv", mime="text/csv",
                           use_container_width=True)

st.markdown("---")
st.caption("v11: QuoteNumber pulled from 'Cotización …' line when present; ReferralManager from the line after 'Atentamente' with numbers stripped; robust Quantity & TotalSales logic retained.")
