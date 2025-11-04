import io, re, traceback
from datetime import datetime
from typing import List, Dict

import pandas as pd
import streamlit as st
from pdfminer.high_level import extract_text

st.set_page_config(page_title="FINSA PDF → CSV (v6)", layout="centered")
st.title("FINSA PDF → Excel (CSV) – v6")
st.caption("Upload up to 100 FINSA PDFs, map to your CSV headers, preview, then download one combined CSV.")

MAX_FILES = 100
MAX_FILE_MB = 25

# Default columns = your mapping order (keep this in sync with your mapping CSV header)
DEFAULT_COLS: List[str] = [
    "ReferralManager","ReferralEmail","Brand","QuoteNumber","QuoteDate",
    "Company","FirstName","LastName","ContactEmail","ContactPhone",
    "CurrencyCode","ContactName","Country","manufacturer_Name","item_id",
    "item_desc","Quantity","TotalSales","PDF","CustomerNumber",
    "UnitSales","Unit_Cost","sales_cost","cust_type","QuoteComment",
    "Created_By","quote_line_no","DemoQuote"
]

# ------------- Helpers -------------
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

DATE_RX = re.compile(r"\b([0-3]\d/[01]\d/\d{4})\b")  # dd/mm/yyyy
HEADER_QUOTE_HINT = re.compile(r"(?i)\bCotizaci[oó]n\b")
QNUM_RXS = [
    re.compile(r"(?i)(?:\bN[uú]mero(?:\s+de\s+cotizaci[oó]n)?|Numero(?:\s+de\s+cotizacion)?|No\.?|N°)\s*:?\s*\n?(\d{5,8})"),
    re.compile(r"(?i)(\d{5,8})\s*\nN[uú]mero\s*:")
]
COMPANY_RX = re.compile(r"(?i)Cliente\s*:\s*([^\n]+)|(?i)Cliente\s*\n([^\n]+)")
CONTACT_RX = re.compile(r"(?i)Contacto\s*:\s*([^\n]+)|(?i)AT'N\s*\n([^\n]+)")
PHONE_RX   = re.compile(r"(?i)Tel[eé]fono\s*:\s*([^\n]+)")
CUR_RX     = re.compile(r"(?i)Moneda\s*:\s*([A-Z]{3})")
SIGN_RX    = re.compile(r"(?i)Atentamente\s*\n([A-ZÁÉÍÓÚÑ ]{3,})")

# Table block from header row down to the Subtotal line (handles Sub-Total/Sub Total/Subtotal)
ITEM_BLOCK_RX = re.compile(r"(?is)(?:MODELO.*?CANTIDAD.*?UNIDAD.*?\n)(.*?)(?:\nSub[\s-]?Total|\nSubtotal|\nSub Total)")

def find_amount(raw_text: str, lbl_pattern: str) -> str:
    rx = re.compile(rf"(?i){lbl_pattern}\s*[:\-]?\s*\n?\s*([\d\.,\s]+)")
    matches = list(rx.finditer(raw_text or ""))
    for m in reversed(matches):
        g1 = m.group(1) if m else None
        if g1:
            return _num_clean(g1)
    return ""

def sum_quantities(block: str) -> str:
    if not block:
        return ""
    # quantities appear as: "1.00 KIT" / "7 PZA" etc → capture the numeric
    qtys = re.findall(r"(?:^|\s)(\d+(?:\.\d{1,2})?)\s+(?:PZA|KIT|PZAS|SET|UND|PCS)\b", block, flags=re.I)
    try:
        s = sum(float(q) for q in qtys)
        s_str = f"{s:.2f}"
        return s_str.rstrip("0").rstrip(".")
    except Exception:
        return ""

def parse_pdf(file_name: str, data: bytes, out_cols: List[str]) -> Dict[str, str]:
    raw = _extract_text(data)
    lines = [ln for ln in (raw or "").splitlines() if ln and ln.strip()]

    # QuoteNumber: prefer 5–8 digits near "Cotización" in the header zone; else use "Número:"
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

    # Company: the text after "Cliente" (keep numeric prefix if present)
    company = ""
    m_co = COMPANY_RX.search(raw or "")
    if m_co:
        company = (m_co.group(1) or m_co.group(2) or "").strip()

    # Contact (Contacto or AT'N)
    first_name = last_name = ""
    m_ct = CONTACT_RX.search(raw or "")
    if m_ct:
        full = (m_ct.group(1) or m_ct.group(2) or "").strip().title()
        parts = [p for p in full.split() if p]
        if parts:
            first_name = parts[0]
            last_name = " ".join(parts[1:])

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

    # Total at bottom (avoid "Total de Artículos")
    total = find_amount(raw, r"Total(?!\s*de\s*Art)")

    # Items
    item_id = item_desc = ""
    qty_total = ""
    multi = False
    m_block = ITEM_BLOCK_RX.search(raw or "")
    if m_block:
        block = m_block.group(1) or ""
        nonempty = [ln for ln in block.splitlines() if ln.strip()]
        multi = len(nonempty) > 1
        qty_total = sum_quantities(block)

        # only set item_id/desc if there is a single line item
        if not multi and nonempty:
            first_line = nonempty[0]
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

    # Build output row according to mapping
    row = {col: "" for col in out_cols}
    def setcol(col, val):
        if col in row:
            row[col] = val

    setcol("ReferralManager", referral_mgr)
    setcol("ReferralEmail", "")                     # always blank per spec
    setcol("Brand", "Finsa")
    setcol("QuoteNumber", qnum)
    setcol("QuoteDate", qdate)
    setcol("Company", company)
    setcol("FirstName", first_name)
    setcol("LastName", last_name)
    setcol("ContactEmail", "")
    setcol("ContactPhone", phone)
    setcol("CurrencyCode", currency)
    setcol("ContactName", f"{first_name} {last_name}".strip())
    setcol("Country", "Mexico" if currency == "MXN" else "")
    setcol("manufacturer_Name", "FINSA")
    setcol("item_id", item_id)
    setcol("item_desc", item_desc)
    setcol("Quantity", qty_total)
    setcol("TotalSales", total)
    setcol("PDF", pdf_name)
    # Other columns remain empty unless you later give rules for them
    return row

# ------------- Sidebar -------------
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

# ------------- Main -------------
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
                errors.append(f"{f.name}: exceeds size limit ({MAX_FILE_MB} MB).")
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
                        problems.append(f"Row {idx+1}: Missing required '{col}'")
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
st.caption("v6: exact FINSA mapping; one row per PDF; sum Quantity across items; blank item_id/item_desc on multi-line; ReferralEmail blank; ReferralManager from Atentamente; Country=Mexico when Currency=MXN.")
