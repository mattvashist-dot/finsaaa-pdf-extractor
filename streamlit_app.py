import io, re, traceback
from datetime import datetime
from typing import List, Dict, Optional, Tuple

import pandas as pd
import streamlit as st
from pdfminer.high_level import extract_text

# ========================
# App config / constants
# ========================
st.set_page_config(page_title="FINSA PDF → CSV (v18)", layout="centered")
st.title("FINSA PDF → Excel (CSV) – v18")
st.caption("Stronger Contacto→First/Last; bottom Atentamente→ReferralManager (robust merge & name picking). Headers preserved.")

MAX_FILES = 100
MAX_FILE_MB = 25

# Fallback header list (used only if no mapping CSV is uploaded)
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
    """Return digits only; remove any EXT/EXTENSIÓN parts."""
    if not raw:
        return ""
    cut = re.split(r"\bEXT\.?\b", raw, flags=re.I)[0]
    cut = re.split(r"\bEXTENSI[oó]N\b", cut, flags=re.I)[0]
    return _digits_only(cut)

def _clean_inline_name(s: str) -> str:
    """
    Clean a potential name chunk:
      - remove URLs/Visita, leading/trailing punctuation/digits
      - drop parentheses content
      - collapse spaces
      - Title case (keeps all-caps too)
    """
    s = (s or "").strip()
    if not s:
        return ""
    if re.search(r"(visita|www|http|https|\.com)", s, flags=re.I):
        return ""
    s = re.sub(r"\([^)]*\)", " ", s)  # remove (...) like (PAGOS)
    s = re.sub(r"^[\d\W_]+\s+", "", s, flags=re.UNICODE)
    s = re.sub(r"\s+[\d\W_]+$", "", s, flags=re.UNICODE)
    s = re.sub(r"\s+", " ", s).strip()
    return s

def _looks_like_name(s: str) -> bool:
    """
    Heuristic: 2–6 alphabetic-ish tokens; no keywords; tokens contain letters (with accents).
    Accepts UPPERCASE or Titlecase.
    """
    if not s:
        return False
    if re.search(r"(visita|www|http|https|\.com|observaciones|condiciones|total|subtotal|iva)", s, flags=re.I):
        return False
    toks = [t for t in s.split() if re.search(r"[A-Za-zÁÉÍÓÚáéíóúÑñ]", t)]
    return 1 <= len(toks) <= 6

def _best_name_from_text(text: str) -> str:
    """
    From a short joined text, try to extract the longest plausible human name.
    Strategy:
      - split on punctuation to phrases
      - pick longest phrase that looks like a name (2–6 tokens with letters)
      - as fallback, join all tokens and re-check
    """
    text = _clean_inline_name(text)
    if not text:
        return ""
    # remove leading codes like "43 " or "380 "
    text = re.sub(r"^\d+\s+", "", text)
    # split phrases on separators
    parts = re.split(r"[|•·,;:]+", text)
    candidates = []
    for p in parts:
        p = _clean_inline_name(p)
        if _looks_like_name(p):
            candidates.append(p)
    if candidates:
        # choose the one with the most tokens, then longest length
        candidates.sort(key=lambda s: (len(s.split()), len(s)), reverse=True)
        return candidates[0]
    # fallback: take words with letters only
    toks = [t for t in text.split() if re.search(r"[A-Za-zÁÉÍÓÚáéíóúÑñ]", t)]
    if len(toks) >= 1:
        return " ".join(toks)
    return ""

def _split_first_last(full: str) -> Tuple[str, str]:
    full = (full or "").strip()
    if not full:
        return "", ""
    parts = full.split()
    return parts[0], " ".join(parts[1:]) if len(parts) > 1 else ""

# ========================
# Regex patterns
# ========================
DATE_RX = re.compile(r"\b([0-3]\d/[01]\d/\d{4})\b")       # dd/mm/yyyy
HEADER_QUOTE_HINT = re.compile(r"\bCotizaci[oó]n\b", re.I)
QNUM_RXS = [
    re.compile(r"(?:\bN[uú]mero(?:\s+de\s+cotizaci[oó]n)?|Numero(?:\s+de\s+cotizacion)?|No\.?|N°)\s*:?\s*\n?(\d{4,8})", re.I),
    re.compile(r"(\d{4,8})\s*\nN[uú]mero\s*:", re.I),
]
CLIENTE_ANYLINE_RX = re.compile(r"Cliente\s*:?\s*([^\n]*)", re.I)

# Contacto / Vigencia / Atentamente / Tel
CONTACTO_LINE_RX = re.compile(r"Contacto\s*:\s*([^\n]+)", re.I)
VIGENCIA_RX      = re.compile(r"Vigencia\s*:", re.I)
ATN_WORD_RX      = re.compile(r"Atentamente\s*$", re.I)

TEL_CITY_LINE_RX = re.compile(r"(.+?)\s+(?:TEL\.?|Tel[eé]fono)\s*[:.]?\s*([^\n]+)$", re.I)
TEL_ONLY_RXS     = [
    re.compile(r"\bTEL\.?\s*[:.]?\s*([^\n]+)", re.I),
    re.compile(r"Tel[eé]fono\s*[:.]?\s*([^\n]+)", re.I),
]

CUR_RX = re.compile(r"Moneda\s*:\s*([A-Z]{3})", re.I)
ITEM_BLOCK_RX = re.compile(
    r"(?is)(?:MODELO.*?CANTIDAD.*?UNIDAD.*?\n|ARTICULO.*?IMPORTE.*?\n)(.*?)(?:\nSub[\s-]?Total|\nSubtotal|\nSub Total)"
)
UNIT_TOKENS = r"(?:PZA|KIT|PZAS|SET|UND|PCS)"
MONEY_RX = re.compile(r"\b\d{1,3}(?:[.,]\d{3})*(?:[.,]\d{2})\b")
INT_OR_DEC_RX = re.compile(r"\b\d+(?:\.\d+)?\b")
UNIT_NEAR_QTY_RX = re.compile(rf"\b(\d+(?:\.\d+)?)\s+{UNIT_TOKENS}\b", re.I)

# ========================
# Totals: third monetary line
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
# Quantity parsing
# ========================
def split_into_item_rows(block_text: str) -> list:
    if not block_text: return []
    raw_lines = [ln for ln in block_text.splitlines() if ln.strip()]
    rows, cur = [], []
    for ln in raw_lines:
        cur.append(ln)
        if len(MONEY_RX.findall(ln)) >= 2:
            rows.append(" ".join(cur)); cur = []
    if cur:
        if rows: rows[-1] += " " + " ".join(cur)
        else: rows.append(" ".join(cur))
    return rows

def infer_quantity_from_row(row_text: str) -> Optional[float]:
    if not row_text: return None
    m = UNIT_NEAR_QTY_RX.search(row_text)
    if m:
        try: return float(m.group(1))
        except: pass
    money_spans = list(MONEY_RX.finditer(row_text))
    if not money_spans: return None
    first_money_start = money_spans[0].start()
    candidates = []
    for nm in INT_OR_DEC_RX.finditer(row_text[:first_money_start]):
        tok = nm.group(0)
        try:
            v = float(tok)
            if 0 <= v < 100000:
                candidates.append((nm.start(), v))
        except: pass
    if candidates: return candidates[-1][1]
    return None

def sum_quantities_advanced(block_text: str) -> str:
    rows = split_into_item_rows(block_text)
    total = 0.0; found = False
    for r in rows:
        q = infer_quantity_from_row(r)
        if q is not None:
            total += q; found = True
    return (f"{total:.2f}".rstrip("0").rstrip(".")) if found else ""

# ========================
# Core field extractors
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
            cand = re.sub(r"\s+", " ", (same or next_line)).strip()
            if re.search(r"(Contacto|Vendedor|Tel[eé]fono|Moneda|No\.|N°|Fecha|Atentamente)", cand, re.I): continue
            if re.search(r"(visita|www|http|https|\.com)", cand, re.I): continue
            if re.search(r"[A-Za-zÁÉÍÓÚáéíóúÑñ]", cand): return cand
    m2 = re.search(r"Cliente\s*:?\s*([^\n]+)\n?([^\n]*)", raw_text or "", flags=re.I)
    if m2:
        cand = (m2.group(1) or "").strip() or (m2.group(2) or "").strip()
        cand = re.sub(r"\s+", " ", cand).strip()
        if not re.search(r"(visita|www|http|https|\.com)", cand, flags=re.I):
            return cand
    return ""

def extract_first_last_from_vigencia_contacto(lines: list, raw_text: str) -> Tuple[str, str]:
    """
    Strict rule: find 'Vigencia:' line; within the next 8 lines, find 'Contacto:' and parse the name.
    If not found, fallback to any 'Contacto:' in the doc.
    """
    # Preferred: relative to Vigencia
    vig_idx = None
    for i, ln in enumerate(lines):
        if VIGENCIA_RX.search(ln):
            vig_idx = i; break
    if vig_idx is not None:
        for j in range(vig_idx+1, min(vig_idx+9, len(lines))):
            m = re.search(r"Contacto\s*:\s*([^\n]+)", lines[j], flags=re.I)
            if m:
                full = _clean_inline_name(m.group(1))
                # remove parenthetical hints like (PAGOS)
                full = re.sub(r"\([^)]*\)", " ", full)
                full = re.sub(r"\s+", " ", full).strip()
                first, last = _split_first_last(full)
                return first, last
    # Fallback: any 'Contacto:' line in the doc
    m_any = CONTACTO_LINE_RX.search(raw_text or "")
    if m_any:
        full = _clean_inline_name(m_any.group(1))
        full = re.sub(r"\([^)]*\)", " ", full)
        full = re.sub(r"\s+", " ", full).strip()
        first, last = _split_first_last(full)
        return first, last
    return "", ""

def extract_referral_manager_bottom(lines: list) -> str:
    """
    Use the last 'Atentamente' (bottom).
    Build a short window of up to 6 lines after it (until 'Visita:') and
    extract the longest plausible person name after stripping codes & URLs.
    Also handles cases where code+name are on same line (e.g., '43 Roberto Carrera').
    Fallback: look just above 'Atentamente' if nothing found below.
    """
    last_idx = None
    for i, ln in enumerate(lines):
        if ATN_WORD_RX.search(ln or ""):
            last_idx = i
    if last_idx is None:
        return ""
    # Bound by 'Visita:'
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
        # if the numeric code and name are in one line, keep it; we'll strip digits later
        if re.fullmatch(r"\d+", cand):
            window_lines.append(cand)  # retain; strip later
        else:
            window_lines.append(cand)
    joined = " ".join(window_lines).strip()
    joined = re.sub(r"\bVisita\s*:.*$", "", joined, flags=re.I)  # hard stop at footer
    # Strip obvious codes at start (e.g., "43 ", "260 ")
    joined = re.sub(r"^\s*\d+\s+", "", joined)
    # Clean and pick best name
    name = _best_name_from_text(joined)
    if name and _looks_like_name(name):
        return name.title()

    # fallback: look above Atentamente
    above_lines = []
    for k in range(last_idx-1, max(last_idx-5, -1), -1):
        cand = (lines[k] or "").strip()
        if cand:
            above_lines.append(cand)
    joined_above = " ".join(above_lines)
    joined_above = re.sub(r"^\s*\d+\s+", "", joined_above)
    name2 = _best_name_from_text(joined_above)
    return name2.title() if name2 else ""

def extract_city_and_phone(lines: list, raw_text: str) -> Tuple[str, str]:
    """
    Preferred: after first 'Cliente:' line, take the first following line with TEL./Teléfono
    as 'City [before TEL]' and phone digits after TEL (EXT removed).
    Fallback: generic search.
    """
    cliente_idx = None
    for i, ln in enumerate(lines):
        if re.search(r"\bCliente\b", ln, flags=re.I):
            cliente_idx = i; break
    if cliente_idx is not None:
        for j in range(cliente_idx + 1, min(cliente_idx + 15, len(lines))):
            ln = lines[j]
            m = TEL_CITY_LINE_RX.search(ln)
            if m:
                city = re.sub(r"\s+", " ", (m.group(1) or "").strip(" .,:;-"))
                phone = _fmt_phone_mx(m.group(2) or "")
                if re.search(r"[A-Za-zÁÉÍÓÚáéíóúÑñ]", city):
                    return (city.title(), phone)
    # fallback: generic
    m2 = TEL_CITY_LINE_RX.search(raw_text or "")
    if m2:
        city = re.sub(r"\s+", " ", (m2.group(1) or "").strip(" .,:;-"))
        right = (m2.group(2) or "").strip()
        phone = _fmt_phone_mx(right)
        return (city.title(), phone if phone else "")
    for rx in TEL_ONLY_RXS:
        for m in rx.finditer(raw_text or ""):
            phone = _fmt_phone_mx(m.group(1) or "")
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

    first_name, last_name = extract_first_last_from_vigencia_contacto(lines, raw)

    city, phone = extract_city_and_phone(lines, raw)

    currency = ""
    m_cur = CUR_RX.search(raw or "")
    if m_cur:
        currency = (m_cur.group(1) or "").upper()

    referral_mgr = extract_referral_manager_bottom(lines)

    total = find_total_third_money_line(raw)

    # Items / Quantity (kept, but you can leave item_id/desc blank for multi-line)
    item_id = item_desc = ""
    qty_total = ""
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
            item_desc = re.sub(r"\s+\d+(?:\.\d{2})?\s*(?:PZA|KIT|PZAS|SET|UND|PCS)?\s*[\d, ]*\.\d{2}.*$", "", desc).strip()

    pdf_name = f"FINSA_{qnum}.pdf" if qnum else (file_name or "")

    # Build row with ALL columns per mapping header; only mapped fields are filled
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
    setcol("FirstName", first_name)
    setcol("LastName", last_name)
    setcol("ContactEmail", "")
    setcol("ContactPhone", phone)
    setcol("Address", "")
    setcol("County", "")
    setcol("City", city)
    setcol("State", "")
    setcol("ZipCode", "")
    setcol("Country", "Mexico" if currency == "MXN" else "")
    # manufacturer_Name intentionally NOT set (left blank)
    setcol("item_id", item_id if item_id else "")
    setcol("item_desc", item_desc if item_desc else "")
    setcol("Quantity", qty_total)
    setcol("TotalSales", total)
    setcol("PDF", pdf_name)
    # other columns left blank intentionally
    return row

# ========================
# Streamlit UI
# ========================
st.sidebar.header("Output Mapping")
mapping_file = st.sidebar.file_uploader(
    "Upload mapping CSV (header defines output columns & order) – REQUIRED to preserve exact headers; otherwise a default header is used.",
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

# Optional review (non-blocking)
strict = st.sidebar.checkbox(
    "Show validation review (Company, QuoteDate, TotalSales)",
    value=True
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
        # DataFrame strictly in mapping header order (headers preserved)
        df = pd.DataFrame(rows, columns=mapping_cols)

        # Review-only list of missing QuoteNumber
        try:
            missing_q = df[df.get("QuoteNumber", "").astype(str).str.strip() == ""]
            if not missing_q.empty:
                st.warning(
                    "QuoteNumber not found (left blank) in the following file(s):\n"
                    + "\n".join(f"- {r.get('PDF', '')}" for _, r in missing_q.iterrows())
                )
        except Exception:
            pass

        if strict:
            problems = []
            for idx, r in df.iterrows():
                for col in ["Company","QuoteDate","TotalSales"]:
                    if col in df.columns and not str(r.get(col, "")).strip():
                        problems.append(f"Row {idx+1}: Missing '{col}'")
            if problems:
                st.error("Validation review:")
                st.code("\n".join(problems), language="text")

        st.success(f"Parsed {len(rows)} file(s). Export contains ALL columns in your header order; mapped fields filled.")
        st.dataframe(df, use_container_width=True)

        csv_bytes = df.to_csv(index=False).encode("utf-8-sig")
        st.download_button("⬇️ Download CSV", data=csv_bytes,
                           file_name="finsa_parsed.csv", mime="text/csv",
                           use_container_width=True)

st.markdown("---")
st.caption("v18: Contacto→First/Last tuned around 'Vigencia'; Atentamente footer logic merged lines and extracts longest person name; headers preserved.")
