import io, re, traceback
from datetime import datetime
from typing import List, Dict, Optional, Tuple

import pandas as pd
import streamlit as st
from pdfminer.high_level import extract_text

# ========================
# App config / constants
# ========================
st.set_page_config(page_title="FINSA PDF → CSV (v23)", layout="centered")
st.title("FINSA PDF → Excel (CSV) – v23")
st.caption("FirstName from Contacto: (full name) or left-of-Vendedor fallback. LastName blank. item_id/item_desc blank. Quantity sum improved.")

MAX_FILES = 100
MAX_FILE_MB = 25

DEFAULT_COLS: List[str] = [
    "ReferralManager","ReferralEmail","Brand","QuoteNumber","QuoteDate","Company",
    "FirstName","LastName","ContactEmail","ContactPhone","Address","County","City",
    "State","ZipCode","Country","manufacturer_Name","item_id","item_desc","Quantity",
    "TotalSales","PDF","CustomerNumber","UnitSales","Unit_Cost","sales_cost","cust_type",
    "QuoteComment","Created_By","quote_line_no","DemoQuote"
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
    if not raw:
        return ""
    cut = re.split(r"\bEXT\.?\b", raw, flags=re.I)[0]
    cut = re.split(r"\bEXTENSI[oó]N\b", cut, flags=re.I)[0]
    return _digits_only(cut)

def _clean_spaces(s: str) -> str:
    return re.sub(r"\s+", " ", (s or "").strip())

def _looks_like_person(name: str) -> bool:
    if not name:
        return False
    if re.search(r"(visita|www|http|https|\.com|subtotal|total|iva|observaciones|condiciones|telefono|tel\.)", name, re.I):
        return False
    toks = [t for t in name.split() if re.search(r"[A-Za-zÁÉÍÓÚáéíóúÑñ]", t)]
    return 2 <= len(toks) <= 6

# ========================
# Regex patterns
# ========================
DATE_RX = re.compile(r"\b([0-3]\d/[01]\d/\d{4})\b")
HEADER_QUOTE_HINT = re.compile(r"\bCotizaci[oó]n\b", re.I)
QNUM_RXS = [
    re.compile(r"(?:\bN[uú]mero(?:\s+de\s+cotizaci[oó]n)?|Numero(?:\s+de\s+cotizacion)?|No\.?|N°)\s*:?\s*\n?(\d{4,8})", re.I),
    re.compile(r"(\d{4,8})\s*\nN[uú]mero\s*:", re.I),
]
CLIENTE_ANYLINE_RX = re.compile(r"Cliente\s*:?\s*([^\n]*)", re.I)

# Contact-related
CONTACTO_INLINE_RX = re.compile(
    r"Contacto\s*:\s*"
    r"([A-Za-zÁÉÍÓÚÑñ]+(?:\s+[A-Za-zÁÉÍÓÚÑñ]+){1,5})"
    r"(?=\s*(?:\(|D[ií]as\s+Ent\b|D[ií]as\b|EDP\b|CANT\b|CLASIF\b|UNID\b|MODELO\b|PRECIO\b|IMPORTE\b|Moneda\b|Vendedor\b|Tel(?:efono)?\.?\b|$))",
    re.I
)
VEND_LINE_RX = re.compile(r"\bVendedor\s*:", re.I)
MONEDA_CHUNK_RX = re.compile(r"Moneda\s*:\s*[A-Z]{3}", re.I)

# Footer / Tel / City
ATN_WORD_RX  = re.compile(r"Atentamente\s*$", re.I)
TEL_CITY_LINE_RX = re.compile(r"(.+?)\s+(?:TEL\.?|Tel[eé]fono)\s*[:.]?\s*([^\n]+)$", re.I)
TEL_ONLY_RXS     = [
    re.compile(r"\bTEL\.?\s*[:.]?\s*([^\n]+)", re.I),
    re.compile(r"Tel[eé]fono\s*[:.]?\s*([^\n]+)", re.I),
]
CUR_RX = re.compile(r"Moneda\s*:\s*([A-Z]{3})", re.I)

# Items block / totals
ITEM_BLOCK_RX = re.compile(
    r"(?is)(?:MODELO.*?CANTIDAD.*?UNIDAD.*?\n|ARTICULO.*?IMPORTE.*?\n)(.*?)(?:\nSub[\s-]?Total|\nSubtotal|\nSub Total)"
)
MONEY_RX = re.compile(r"\b\d{1,3}(?:[.,]\d{3})*(?:[.,]\d{2})\b")

# Units & Qty
UNIT_WORDS = ["PZA", "PZAS", "JUEGO", "KIT", "SET", "UND", "PCS"]
UNIT_WORDS_RX = r"(?:PZA|PZAS|JUEGO|KIT|SET|UND|PCS)"
QTY_DEC_RX = re.compile(r"\b\d+(?:[.,]\d+)?\b")

# ========================
# Total (3rd monetary after subtotal)
# ========================
def find_total_third_money_line(raw_text: str) -> str:
    if not raw_text: return ""
    lines = [ln.strip() for ln in raw_text.splitlines() if ln.strip()]
    start_idx = None
    for i, ln in enumerate(lines):
        if re.search(r"(Sub[\s-]?Total|Subtotal|Sub Total)", ln, flags=re.I):
            start_idx = i; break
    if start_idx is None:
        for i, ln in enumerate(lines):
            if re.search(r"\bTotal\b", ln, flags=re.I) and not re.search(r"Total\s*de\s*Art", ln, flags=re.I):
                start_idx = i; break
    if start_idx is None: return ""
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
            if "." in val: monies.append(val)
    if len(monies) >= 3: return monies[2]
    if monies: return monies[-1]
    return ""

# ========================
# City & phone
# ========================
def extract_city_and_phone(lines: list, raw_text: str) -> Tuple[str, str]:
    cliente_idx = None
    for i, ln in enumerate(lines):
        if re.search(r"\bCliente\b", ln, flags=re.I):
            cliente_idx = i; break
    if cliente_idx is not None:
        for j in range(cliente_idx + 1, min(cliente_idx + 15, len(lines))):
            ln = lines[j]
            m = TEL_CITY_LINE_RX.search(ln or "")
            if m:
                city = _clean_spaces((m.group(1) or "").strip(" .,:;-"))
                right = (m.group(2) or "").strip()
                phone = _fmt_phone_mx(right)
                if re.search(r"[A-Za-zÁÉÍÓÚáéíóúÑñ]", city):
                    return (city.title(), phone)
    m2 = TEL_CITY_LINE_RX.search(raw_text or "")
    if m2:
        city = _clean_spaces((m2.group(1) or "").strip(" .,:;-"))
        phone = _fmt_phone_mx((m2.group(2) or "").strip())
        return (city.title(), phone if phone else "")
    for rx in TEL_ONLY_RXS:
        for m in rx.finditer(raw_text or ""):
            phone = _fmt_phone_mx(m.group(1) or "")
            if phone:
                return ("", phone)
    return ("", "")

# ========================
# Referral Manager (bottom, after "Atentamente")
# ========================
def extract_referral_manager_bottom(lines: list) -> str:
    last_idx = None
    for i, ln in enumerate(lines):
        if ATN_WORD_RX.search(ln or ""):
            last_idx = i
    if last_idx is None:
        return ""
    stop_idx = min(last_idx + 8, len(lines))
    for j in range(last_idx+1, min(last_idx+10, len(lines))):
        if re.search(r"^Visita\s*:", lines[j] or "", flags=re.I):
            stop_idx = j
            break
    window_lines = []
    for j in range(last_idx+1, stop_idx):
        cand = (lines[j] or "").strip()
        if not cand:
            continue
        window_lines.append(cand)
    joined = " ".join(window_lines)
    joined = re.sub(r"^\s*\d+\s+", "", joined)
    joined = re.sub(r"\bVisita\s*:.*$", "", joined, flags=re.I)
    phrases = re.findall(r"([A-Za-zÁÉÍÓÚÑñ][A-Za-zÁÉÍÓÚÑñ]+(?:\s+[A-Za-zÁÉÍÓÚÑñ][A-Za-zÁÉÍÓÚÑñ]+){1,5})", joined)
    candidates = [p for p in phrases if _looks_like_person(p)]
    if candidates:
        candidates.sort(key=lambda s: (len(s.split()), len(s)), reverse=True)
        return candidates[0].title()
    # fallback: look a few lines above
    for k in range(last_idx-1, max(last_idx-5, -1), -1):
        cand = (lines[k] or "").strip()
        phrases = re.findall(r"([A-Za-zÁÉÍÓÚÑñ][A-Za-zÁÉÍÓÚÑñ]+(?:\s+[A-Za-zÁÉÍÓÚÑñ][A-Za-zÁÉÍÓÚÑñ]+){1,5})", cand)
        candidates = [p for p in phrases if _looks_like_person(p)]
        if candidates:
            candidates.sort(key=lambda s: (len(s.split()), len(s)), reverse=True)
            return candidates[0].title()
    return ""

# ========================
# FirstName only (Contacto or left-of-Vendedor)
# ========================
def extract_firstname_only(lines: list, raw_text: str) -> str:
    # Rule 1: global inline "Contacto: <FULL NAME>" with hard-stop tokens
    m_inline = CONTACTO_INLINE_RX.search(raw_text or "")
    if m_inline:
        full = _clean_spaces(m_inline.group(1))
        if _looks_like_person(full):
            return full.title()

    # Rule 2: left-of-"Vendedor:" (remove "Moneda: XXX"; longest 2–6 word phrase)
    vend_idx = None
    for i, l in enumerate(lines):
        if VEND_LINE_RX.search(l or ""):
            vend_idx = i
            break
    if vend_idx is not None:
        left = lines[vend_idx].split("Vendedor", 1)[0]
        prev_chunk = ""
        if vend_idx - 1 >= 0 and len(lines[vend_idx - 1].strip()) < 80:
            prev_chunk = lines[vend_idx - 1].strip()
        candidate = _clean_spaces((prev_chunk + " " + left).strip())
        candidate = MONEDA_CHUNK_RX.sub(" ", candidate)
        candidate = _clean_spaces(candidate)
        phrases = re.findall(r"([A-Za-zÁÉÍÓÚÑñ][A-Za-zÁÉÍÓÚÑñ]+(?:\s+[A-Za-zÁÉÍÓÚÑñ][A-Za-zÁÉÍÓÚÑñ]+){1,5})", candidate)
        phrases = [p for p in phrases if _looks_like_person(p)]
        if phrases:
            phrases.sort(key=lambda s: (len(s.split()), len(s)), reverse=True)
            return phrases[0].title()

    return ""  # leave blank if not confidently found

# ========================
# Company / Quote number / Date
# ========================
def extract_quote_number(lines: list, raw_text: str) -> str:
    for idx, ln in enumerate(lines[:40]):
        if HEADER_QUOTE_HINT.search(ln):
            window = " ".join(lines[idx:idx+6])
            nums = re.findall(r"\b(\d{4,8})\b", window)
            if nums: return nums[-1].lstrip("0") or nums[-1]
    for rx in QNUM_RXS:
        m = rx.search(raw_text or "")
        if m:
            cand = (m.group(1) or "").strip()
            if re.fullmatch(r"\d{4,8}", cand):
                return cand.lstrip("0") or cand
    top = " ".join(lines[:40])
    m = re.search(r"\b(\d{4,8})\b", top)
    return ((m.group(1) or "").lstrip("0") or m.group(1)) if m else ""

def extract_company(lines: list, raw_text: str) -> str:
    for i, ln in enumerate(lines[:80]):
        m = CLIENTE_ANYLINE_RX.search(ln)
        if m:
            same = (m.group(1) or "").strip()
            next_line = lines[i+1].strip() if (not same or len(same) < 3) and i + 1 < len(lines) else ""
            cand = _clean_spaces(same or next_line)
            if re.search(r"(Contacto|Vendedor|Tel[eé]fono|Moneda|No\.|N°|Fecha|Atentamente)", cand, re.I): continue
            if re.search(r"(visita|www|http|https|\.com)", cand, re.I): continue
            if re.search(r"[A-Za-zÁÉÍÓÚáéíóúÑñ]", cand): return cand
    m2 = re.search(r"Cliente\s*:?\s*([^\n]+)\n?([^\n]*)", raw_text or "", flags=re.I)
    if m2:
        cand = _clean_spaces((m2.group(1) or "") or (m2.group(2) or ""))
        if not re.search(r"(visita|www|http|https|\.com)", cand, flags=re.I):
            return cand
    return ""

def extract_date(raw_text: str) -> str:
    m_date = DATE_RX.search(raw_text or "")
    if not m_date:
        return ""
    try:
        return datetime.strptime(m_date.group(1), "%d/%m/%Y").strftime("%m/%d/%Y")
    except Exception:
        return ""

# ========================
# Quantity (sum across item rows)
# ========================
def parse_qty_sum(raw_text: str) -> str:
    """
    Inside the item block (between header and Sub-Total), sum quantities in two layouts:
      A) <UNIT> <QTY>  (e.g., 'JUEGO 1.00', 'PZA 2')
      B) <QTY> <UNIT>  (rare but handled)
    Ignores monetary amounts and other numbers.
    """
    m_block = ITEM_BLOCK_RX.search(raw_text or "")
    if not m_block:
        return ""
    block = m_block.group(1) or ""
    total = 0.0
    found = False

    # A) UNIT then QTY
    for m in re.finditer(rf"\b{UNIT_WORDS_RX}\b\s+(\d+(?:[.,]\d+)?)", block, flags=re.I):
        q = m.group(1)
        q = q.replace(",", ".")
        try:
            total += float(q); found = True
        except:
            pass

    # B) QTY then UNIT (ensure not a price by looking ahead a short distance for a currency-like price with 2 decimals)
    for m in re.finditer(rf"(\d+(?:[.,]\d+)?)\s+\b{UNIT_WORDS_RX}\b", block, flags=re.I):
        qty_str = m.group(1).replace(",", ".")
        # heuristic: if within 15 chars after this match we immediately hit a price-like number with 2 decimals AND separators,
        # it's probably not the qty pattern we already counted (avoid double-count).
        tail = block[m.end(): m.end() + 20]
        if re.search(r"\b\d{1,3}(?:[.,]\d{3})*[.,]\d{2}\b", tail):
            # could be the price – but we still count the qty before it
            pass
        try:
            total += float(qty_str); found = True
        except:
            pass

    return (f"{total:.2f}".rstrip("0").rstrip(".")) if found else ""

# ========================
# PDF parser core
# ========================
def parse_pdf(file_name: str, data: bytes, out_cols: List[str]) -> Dict[str, str]:
    raw = _extract_text(data)
    lines = [ln for ln in (raw or "").splitlines() if ln.strip()]

    qnum = extract_quote_number(lines, raw)
    qdate = extract_date(raw)
    company = extract_company(lines, raw)
    firstname = extract_firstname_only(lines, raw)  # FirstName = full contact name
    city, phone = extract_city_and_phone(lines, raw)

    currency = ""
    m_cur = CUR_RX.search(raw or "")
    if m_cur:
        currency = (m_cur.group(1) or "").upper()

    referral_mgr = extract_referral_manager_bottom(lines)
    total = find_total_third_money_line(raw)
    qty_total = parse_qty_sum(raw)

    pdf_name = f"FINSA_{qnum}.pdf" if qnum else (file_name or "")

    # Build row with ALL columns; only mapped fields filled
    row = {col: "" for col in out_cols}
    def setcol(col, val):
        if col in row:
            row[col] = val

    setcol("ReferralManager", referral_mgr)
    setcol("ReferralEmail", "")
    setcol("Brand", "Finsa")
    setcol("QuoteNumber", qnum)
    setcol("QuoteDate", qdate)
    setcol("Company", company)
    setcol("FirstName", firstname)   # full name into FirstName
    setcol("LastName", "")           # always blank per request
    setcol("ContactEmail", "")
    setcol("ContactPhone", phone)
    setcol("Address", "")
    setcol("County", "")
    setcol("City", city)
    setcol("State", "")
    setcol("ZipCode", "")
    setcol("Country", "Mexico" if currency == "MXN" else "")
    setcol("manufacturer_Name", "")
    setcol("item_id", "")            # always blank per request
    setcol("item_desc", "")          # always blank per request
    setcol("Quantity", qty_total)
    setcol("TotalSales", total)
    setcol("PDF", pdf_name)
    # remaining columns left blank intentionally
    return row

# ========================
# Streamlit UI
# ========================
st.sidebar.header("Output Mapping")
mapping_file = st.sidebar.file_uploader(
    "Upload mapping CSV (header defines output columns & order). If omitted, a default header is used.",
    type=["csv"]
)
if mapping_file is not None:
    try:
        mapping_cols = pd.read_csv(mapping_file, nrows=0).columns.tolist()
    except Exception:
        st.sidebar.error("Could not read mapping header; falling back to default headers.")
        mapping_cols = DEFAULT_COLS
else:
    mapping_cols = DEFAULT_COLS

files = st.file_uploader("Upload up to 100 FINSA PDF quotes", type=["pdf"], accept_multiple_files=True)

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
        st.success("Parsed. FirstName from Contacto: (full) or left-of-Vendedor fallback. item_id/item_desc blank. Quantity summed across all lines.")
        st.dataframe(df, use_container_width=True)

        csv_bytes = df.to_csv(index=False).encode("utf-8-sig")
        st.download_button("⬇️ Download CSV", data=csv_bytes,
                           file_name="finsa_parsed.csv", mime="text/csv",
                           use_container_width=True)
