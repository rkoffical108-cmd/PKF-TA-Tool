"""
PKF TA Claim Tool — Streamlit version
Runs on Streamlit Community Cloud (always-on, free)
"""

import re, io, uuid
from datetime import datetime, date
from typing import Optional
from pathlib import Path

import streamlit as st
from PIL import Image
import pytesseract
from pdf2image import convert_from_bytes
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter
from pypdf import PdfWriter, PdfReader

# ── page config ───────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="PKF TA Claim Tool",
    page_icon="🧾",
    layout="wide",
)

# ── custom CSS ────────────────────────────────────────────────────────────────
st.markdown("""
<style>
[data-testid="stAppViewContainer"] { background: #f0f4fa; }
.pkf-header {
    background: #1a3660;
    color: white;
    padding: 18px 28px 14px;
    border-bottom: 4px solid #c9a84c;
    margin: -1rem -1rem 1.5rem -1rem;
}
.pkf-header h1 { font-size: 20px; margin: 0; font-weight: 700; }
.pkf-header .sub { font-size: 12px; color: #aec6e8; margin-bottom: 4px; }
.pkf-logo { font-size: 22px; font-weight: 800; color: #c9a84c; }
.row-card {
    background: white;
    border: 1px solid #c8d4e8;
    border-radius: 6px;
    padding: 16px;
    margin-bottom: 12px;
}
.flag-warn {
    background: #fff3cd;
    border: 1px solid #e6a817;
    color: #7a5c00;
    border-radius: 4px;
    padding: 4px 10px;
    font-size: 12px;
    font-weight: 700;
}
.ocr-ok  { color: #1a6b2e; font-size: 12px; }
.ocr-err { color: #c0392b; font-size: 12px; }
</style>
<div class="pkf-header">
  <div><span class="pkf-logo">PKF</span> <span class="sub">Sridhar &amp; Santhanam LLP</span></div>
  <h1>Travel Allowance Claim Tool</h1>
</div>
""", unsafe_allow_html=True)

# ── accounting heads ──────────────────────────────────────────────────────────
ACCOUNTING_HEADS = [
    {"sl":  1, "acc": 6535, "name": "Air Fare"},
    {"sl":  2, "acc": 6540, "name": "Bus Fare"},
    {"sl":  3, "acc": 6545, "name": "Expenses during travel"},
    {"sl":  4, "acc": 6555, "name": "Train Fare"},
    {"sl":  5, "acc": 6560, "name": "Pick up & Drop"},
    {"sl":  6, "acc": 6565, "name": "Boarding Expenses(Stay & Food)"},
    {"sl":  7, "acc": 6578, "name": "Laundry Expenses"},
    {"sl":  8, "acc": 6590, "name": "Local Conveyance"},
    {"sl":  9, "acc": 6571, "name": "Other Charges"},
    {"sl": 10, "acc": 6575, "name": "Telephone charges - clients"},
    {"sl": 11, "acc": 6580, "name": "Visa Handling Charges"},
    {"sl": 12, "acc": 6585, "name": "Insurance Travel"},
    {"sl": 13, "acc": 6577, "name": "Car Hire Charges -Clients"},
    {"sl": 14, "acc": 6595, "name": "Car Parking Charges"},
]
ACC_OPTIONS = ["— select —"] + [f"{h['acc']} — {h['name']}" for h in ACCOUNTING_HEADS]
ACC_MAP     = {h["acc"]: h["name"] for h in ACCOUNTING_HEADS}


# ── OCR helpers ───────────────────────────────────────────────────────────────
def images_from_bytes(data: bytes, filename: str) -> list:
    if filename.lower().endswith(".pdf"):
        return convert_from_bytes(data, dpi=300)
    return [Image.open(io.BytesIO(data))]


def extract_amount(text: str) -> Optional[float]:
    cleaned = re.sub(r"(?<=\d),(?=\d{3})", "", text)
    CURR = r"(?:₹|£|Rs\.?|INR|R[s5]\.?|%|R\[|F(?=\d))"

    # P1 strong keyword + currency prefix
    strong_kw = (
        r"(?:grand\s*total|total\s*amount|amount\s*payable|net\s*payable"
        r"|net\s*total|bill\s*total|payable\s*amount)"
        r"[^\d]{0,30}" + CURR + r"\s*([0-9]+(?:\.[0-9]{1,2})?)"
    )
    p1 = [float(m.group(1)) for m in re.finditer(strong_kw, cleaned, re.IGNORECASE)
          if 1 <= _safe_float(m.group(1)) <= 999999]
    if p1: return p1[-1]

    # P2 last currency-prefixed amount
    p2 = [float(m.group(1)) for m in re.finditer(
          CURR + r"\s*([0-9]+(?:\.[0-9]{1,2})?)", cleaned, re.IGNORECASE)
          if 1 <= _safe_float(m.group(1)) <= 999999]
    if p2: return p2[-1]

    # P3 last keyword-adjacent
    p3 = [float(m.group(1)) for m in re.finditer(
          r"(?:total|amount|grand\s*total|net\s*amount|paid|payable|fare)"
          r"[^\d]{0,20}([0-9]+(?:\.[0-9]{1,2})?)", cleaned, re.IGNORECASE)
          if 1 <= _safe_float(m.group(1)) <= 999999]
    if p3: return p3[-1]

    # P4 last standalone number
    p4 = [float(m.group(1)) for m in re.finditer(r"\b([0-9]{2,6}(?:\.[0-9]{1,2})?)\b", cleaned)
          if 10 <= _safe_float(m.group(1)) <= 99999]

    # Strip leading misread-prefix digit: Tesseract reads ₹ as 2/7/6/9 etc.
    # e.g. ₹480 → "2480" (4 digits). Strip first digit if result is 50–9999.
    # Only applies to 4+ digit numbers where stripping gives a plausible amount.
    MISREAD = {"2","7","6","8","9","3","4"}
    def _strip(candidates):
        out = []
        for v in candidates:
            s = str(int(v)) if v == int(v) else str(v)
            if len(s) >= 4 and s[0] in MISREAD:
                stripped = float(s[1:])
                if 50 <= stripped <= 9999:
                    out.append(stripped)
                    continue
            out.append(v)
        return out

    if p4: return _strip(p4)[-1]
    return None


def _safe_float(s):
    try: return float(s)
    except: return 0


def extract_date(text: str) -> Optional[date]:
    months = {
        "jan":1,"feb":2,"mar":3,"apr":4,"may":5,"jun":6,
        "jul":7,"aug":8,"sep":9,"oct":10,"nov":11,"dec":12,
        "january":1,"february":2,"march":3,"april":4,
        "june":6,"july":7,"august":8,"september":9,
        "october":10,"november":11,"december":12,
    }
    cy = date.today().year
    patterns = [
        (r"\b(\d{1,2})\s+([A-Za-z]{3,9})\s+(20\d{2})\b",       "dmy"),
        (r"\b([A-Za-z]{3,9})\s+(\d{1,2})[,\s]+(20\d{2})\b",    "mdy"),
        (r"\b(20\d{2})[/\-](\d{2})[/\-](\d{2})\b",              "ymd"),
        (r"\b(\d{2})[/\-\.](\d{2})[/\-\.](20\d{2})\b",          "dmy_num"),
        (r"\b([A-Za-z]{3,9})\s+(\d{1,2})(?:[,\s]|$)",           "md_noyear"),
    ]
    for pat, fmt in patterns:
        for m in re.finditer(pat, text, re.IGNORECASE):
            try:
                g = m.groups()
                if fmt == "dmy":
                    day, mon, year = int(g[0]), months.get(g[1].lower()), int(g[2])
                elif fmt == "mdy":
                    mon, day, year = months.get(g[0].lower()), int(g[1]), int(g[2])
                elif fmt == "ymd":
                    year, mon, day = int(g[0]), int(g[1]), int(g[2])
                elif fmt == "dmy_num":
                    day, mon, year = int(g[0]), int(g[1]), int(g[2])
                elif fmt == "md_noyear":
                    mon, day, year = months.get(g[0].lower()), int(g[1]), cy
                else:
                    continue
                if mon and 1 <= mon <= 12 and 1 <= day <= 31 and 2000 <= year <= 2099:
                    return date(year, mon, day)
            except (ValueError, TypeError):
                pass
    return None


def ocr_bill(data: bytes, filename: str) -> dict:
    imgs = images_from_bytes(data, filename)
    full = ""
    for img in imgs:
        raw = img.convert("RGB")
        for cfg in ["--psm 3 -l eng", "--psm 6 -l eng", "--psm 11 -l eng"]:
            try: full += pytesseract.image_to_string(raw, config=cfg) + "\n"
            except: pass
    return {"amount": extract_amount(full), "date": extract_date(full)}


def ocr_gpay(data: bytes, filename: str) -> Optional[float]:
    """
    GPay/BHIM/UPI screenshot OCR.
    Tries both normal and inverted image (for dark/coloured backgrounds like BHIM green).
    Crops top 40% where amount is displayed large, upscales 3x.
    Falls back to full-image OCR if crop yields nothing.
    """
    imgs = images_from_bytes(data, filename)
    if not imgs: return None
    img = imgs[0].convert("RGB")
    w, h = img.size
    top_crop = img.crop((0, 0, w, int(h * 0.40)))
    top_crop = top_crop.resize((w * 3, int(h * 0.40) * 3), Image.LANCZOS)

    def _scan_for_amount(pil_img):
        grey = pil_img.convert("L")
        # Try normal and inverted (for white text on dark bg)
        for variant in [grey, Image.fromarray(255 - __import__("numpy").array(grey))]:
            try:
                text = pytesseract.image_to_string(variant, config="--psm 3 -l eng")
                for line in text.split("\n"):
                    line = line.strip()
                    # Remove currency symbols and commas
                    line_clean = re.sub(r"[₹£%,Rs\.INR]", "", line).strip()
                    m = re.fullmatch(r"([0-9]+(?:\.[0-9]{1,2})?)", line_clean)
                    if m:
                        v = float(m.group(1))
                        if 50 <= v <= 99999:
                            return v
            except Exception:
                pass
        return None

    result = _scan_for_amount(top_crop)
    if result:
        return result

    # Fallback: full image general OCR
    full = ""
    for img2 in imgs:
        for cfg in ["--psm 3 -l eng", "--psm 6 -l eng"]:
            try: full += pytesseract.image_to_string(img2.convert("RGB"), config=cfg) + "\n"
            except: pass
    return extract_amount(full)


# ── Excel builder ─────────────────────────────────────────────────────────────
def _border():
    s = Side(style="thin")
    return Border(left=s, right=s, top=s, bottom=s)

GREY_FILL = PatternFill("solid", fgColor="D9D9D9")
RED_FILL  = PatternFill("solid", fgColor="FFC7CE")
HDR_FILL  = PatternFill("solid", fgColor="1F4E79")


def build_excel(header: dict, rows: list, thresholds: dict, thr_enabled: bool) -> bytes:
    wb = Workbook()
    ws1 = wb.active
    ws1.title = "TA CON-SHEET "

    head_totals = {h["acc"]: 0.0 for h in ACCOUNTING_HEADS}
    for r in rows:
        acc = r.get("acc", 0)
        if acc in head_totals:
            head_totals[acc] += r.get("total", 0)

    def cf(bold=False, size=10):
        return Font(name="Cambria", bold=bold, size=size)

    # Title
    ws1.merge_cells("A1:H1")
    t = ws1.cell(1, 1, "TRAVELLING EXPENSES STATEMENT")
    t.font = cf(bold=True, size=12)
    t.alignment = Alignment(horizontal="center")

    def lbl(row, col, v): ws1.cell(row, col, v).font = cf(bold=True)
    def val(row, col, v): ws1.cell(row, col, v).font = cf()

    lbl(2,1,"CLIENT NAME :");          val(2,2, header.get("client_name",""))
    lbl(2,4,"Team Members Name ");     val(2,5, header.get("team_members",""))
    lbl(3,4,"Audit Start Date");       val(3,5, header.get("audit_start",""))
    lbl(4,1,"CLIENT LOCATION");        val(4,2, header.get("client_location",""))
    lbl(4,4,"Audit End Date");         val(4,5, header.get("audit_end",""))
    lbl(5,4,"Total Mandays");          val(5,5, header.get("total_mandays",""))
    lbl(6,1,"NATURE OF AUDIT :");      val(6,2, header.get("nature_of_audit",""))
    lbl(7,4,"To & Fro -Tickets Booked"); val(7,5, header.get("tickets_booked",""))
    lbl(8,4,"Onward booked by");       val(8,5, header.get("onward_by","NA"))
    lbl(9,4,"Return booked by");       val(9,5, header.get("return_by","NA"))

    for col, text in enumerate(["S.L.NO","Accounting No.","Accounting Heads","Consolidated Total Expenses"], 1):
        c = ws1.cell(11, col, text)
        c.font = Font(name="Cambria", bold=True, size=10, color="FFFFFF")
        c.fill = HDR_FILL; c.border = _border()
        c.alignment = Alignment(horizontal="center", wrap_text=True)

    restricted = {6571, 6575, 6577, 6580, 6585, 6595}
    for idx, head in enumerate(ACCOUNTING_HEADS):
        row = 12 + idx
        acc = head["acc"]; total = head_totals.get(acc, 0.0)
        ws1.cell(row, 1, head["sl"]).font  = cf()
        ws1.cell(row, 2, acc).font         = cf()
        ws1.cell(row, 3, head["name"]).font= cf()
        tc = ws1.cell(row, 4, total)
        tc.font = cf(); tc.number_format = "#,##0.00"; tc.border = _border()
        fill = GREY_FILL if acc in restricted else None
        thr = thresholds.get(acc)
        if thr_enabled and thr and total > thr:
            tc.fill = RED_FILL
        elif fill:
            for col in range(1, 5): ws1.cell(row, col).fill = fill

    advance = float(header.get("advance_amount", 0))
    for row, l, v in [
        (26,"Total Expenses","=SUM(D12:D25)"),
        (28,"Advance Amount:", advance),
        (29,"Total Expenses","=SUM(D12:D25)"),
        (30,"Balance to repay office", 0),
        (31,"Balance to receive from office","=E28-E29"),
        (32,"Submission Date :", header.get("submission_date","")),
        (33,"Signature :", header.get("signature","")),
    ]:
        ws1.cell(row, 4, l).font = cf(bold=True)
        c = ws1.cell(row, 5, v); c.font = cf()
        if isinstance(v, float): c.number_format = "#,##0.00"

    ws1.column_dimensions["A"].width = 8
    ws1.column_dimensions["B"].width = 14
    ws1.column_dimensions["C"].width = 32
    ws1.column_dimensions["D"].width = 26
    ws1.column_dimensions["E"].width = 20

    # ── Annexure ──────────────────────────────────────────────────────────────
    ws2 = wb.create_sheet("Annexure-Breakup daywise")
    ws2.cell(1,1,"ANNEXURE").font = cf(bold=True, size=12)
    for row, l, v in [
        (2,"Name. of the team members", header.get("team_members","")),
        (3,"Client Location Visited",   header.get("client_location","")),
        (4,"Nature of Audit",           header.get("nature_of_audit","")),
    ]:
        ws2.cell(row,1,l).font = cf(bold=True)
        ws2.cell(row,3,v).font = cf()

    col_hdrs = ["Date","Day","Nature of Expenditure","Bill Amount","Extra",
                "Total","Sup Bills Y/N","Remarks -Mandatory","Mode of travel-Mandatory","Team members name"]
    for col, h in enumerate(col_hdrs, 1):
        c = ws2.cell(5, col, h)
        c.font = Font(name="Cambria", bold=True, size=10, color="FFFFFF")
        c.fill = HDR_FILL; c.border = _border()
        c.alignment = Alignment(horizontal="center", wrap_text=True)

    sorted_rows = sorted(rows, key=lambda r: r.get("date") or date.min)

    for idx, r in enumerate(sorted_rows):
        er = 6 + idx
        acc       = r.get("acc", 0)
        bill_amt  = float(r.get("bill_amount", 0) or 0)
        extra_amt = float(r.get("extra", 0) or 0)
        total_amt = max(bill_amt, extra_amt)
        dt        = r.get("date")

        c1 = ws2.cell(er, 1, dt); c1.number_format = "DD/MM/YYYY"; c1.font = cf()
        ws2.cell(er, 2, f'=TEXT(A{er},"dddd")').font = cf()
        ws2.cell(er, 3, ACC_MAP.get(acc, "")).font    = cf()

        for col, v in [(4, bill_amt),(5, extra_amt)]:
            c = ws2.cell(er, col, v); c.number_format = "#,##0.00"; c.font = cf()

        tc = ws2.cell(er, 6, f"=D{er}+E{er}")
        tc.number_format = "#,##0.00"; tc.font = cf()

        ws2.cell(er, 7,  r.get("sup_bills","Yes")).font    = cf()
        ws2.cell(er, 8,  r.get("remarks","")).font         = cf()
        ws2.cell(er, 9,  r.get("mode_of_travel","")).font  = cf()
        ws2.cell(er, 10, r.get("team_member","")).font     = cf()

        row_fill = None
        if thr_enabled:
            thr = thresholds.get(acc)
            if thr and total_amt > thr:
                row_fill = RED_FILL

        for col in range(1, 11):
            ws2.cell(er, col).border = _border()
            if row_fill: ws2.cell(er, col).fill = row_fill

    total_row = 6 + len(sorted_rows)
    ws2.cell(total_row, 3, "Total").font = cf(bold=True)
    tc = ws2.cell(total_row, 6, f"=SUM(F6:F{total_row-1})")
    tc.number_format = "#,##0.00"; tc.font = cf(bold=True)
    ws1["D19"] = f"='Annexure-Breakup daywise'!F{total_row}"

    widths = [14,12,28,13,10,10,13,28,22,22]
    for i, w in enumerate(widths, 1):
        ws2.column_dimensions[get_column_letter(i)].width = w

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def build_pdf(bill_files: list) -> Optional[bytes]:
    """Merge all bill bytes into one PDF in order."""
    writer = PdfWriter()
    added  = 0
    for data, filename in bill_files:
        if not data: continue
        try:
            if filename.lower().endswith(".pdf"):
                reader = PdfReader(io.BytesIO(data))
                for page in reader.pages:
                    writer.add_page(page)
            else:
                img = Image.open(io.BytesIO(data)).convert("RGB")
                buf = io.BytesIO()
                img.save(buf, format="PDF")
                buf.seek(0)
                reader = PdfReader(buf)
                for page in reader.pages:
                    writer.add_page(page)
            added += 1
        except Exception:
            pass
    if added == 0: return None
    out = io.BytesIO()
    writer.write(out)
    return out.getvalue()


# ── session state initialisation ──────────────────────────────────────────────
if "rows" not in st.session_state:
    st.session_state.rows = [{}]   # start with one empty row

if "row_count" not in st.session_state:
    st.session_state.row_count = 1


# ── STEP 1: Claim header ──────────────────────────────────────────────────────
st.markdown("### 1️⃣ Claim Details")
with st.container():
    c1, c2 = st.columns(2)
    with c1:
        client_name     = st.text_input("Client Name",      placeholder="e.g. Sterling Holidays")
        client_location = st.text_area("Client Location",   placeholder="Full address", height=80)
        nature_of_audit = st.text_input("Nature of Audit",  placeholder="e.g. Internal Audit-Q1_FY_26-27")
        advance_amount  = st.number_input("Advance Amount (₹)", min_value=0.0, step=100.0, format="%.2f")
    with c2:
        team_members    = st.text_input("Team Members",     placeholder="Comma-separated names")
        audit_start     = st.date_input("Audit Start Date", value=None)
        audit_end       = st.date_input("Audit End Date",   value=None)
        tickets_booked  = st.selectbox("Tickets Booked By", ["NA","Office","Client","Self Booking"])
        onward_by       = st.text_input("Onward booked by", value="NA")
        return_by       = st.text_input("Return booked by", value="NA")
        submission_date = st.date_input("Submission Date",  value=date.today())
        signature       = st.text_input("Signature",        placeholder="Name of signatory")

# Mandays auto-compute
total_mandays = ""
if audit_start and audit_end and audit_end >= audit_start:
    total_mandays = (audit_end - audit_start).days + 1
    st.info(f"**Total Mandays: {total_mandays}** (auto-computed)")

st.divider()

# ── STEP 2: Thresholds ────────────────────────────────────────────────────────
st.markdown("### 2️⃣ Per-Person Per-Day Limits *(optional)*")
thr_enabled = st.toggle("Enable threshold limits")
thresholds = {}
if thr_enabled:
    st.caption("Rows exceeding the limit will be flagged in red in the Excel output but still included in totals.")
    tc1, tc2, tc3 = st.columns(3)
    with tc1:
        thresholds[6590] = st.number_input("Local Conveyance (₹)", min_value=0.0, value=300.0, step=50.0)
        thresholds[6565] = st.number_input("Boarding/Food (₹)",    min_value=0.0, value=300.0, step=50.0)
    with tc2:
        thresholds[6535] = st.number_input("Air Fare (₹)",         min_value=0.0, value=0.0, step=500.0)
        thresholds[6555] = st.number_input("Train Fare (₹)",       min_value=0.0, value=0.0, step=100.0)
    with tc3:
        thresholds[6540] = st.number_input("Bus Fare (₹)",         min_value=0.0, value=0.0, step=50.0)
        thresholds[6560] = st.number_input("Pick up & Drop (₹)",   min_value=0.0, value=0.0, step=50.0)

st.divider()

# ── STEP 3: Expense rows ──────────────────────────────────────────────────────
st.markdown("### 3️⃣ Expense Entries")

collected_rows  = []
bill_files_list = []

# Parse team member list for dropdown
team_member_list = [m.strip() for m in team_members.split(",") if m.strip()] if team_members else []

for i in range(st.session_state.row_count):
    key = f"row_{i}"

    # Session state keys for OCR results — persisted across rerenders
    ss_amt  = f"{key}_ocr_amt"
    ss_date = f"{key}_ocr_date"
    ss_file = f"{key}_last_file"   # track which file was last scanned

    with st.expander(f"**Row {i+1}**", expanded=True):

        # ── Bill upload FIRST — OCR runs here, sets session state ─────────────
        # Must be before date/amount widgets so rerun pre-fills them correctly
        bc1, bc2 = st.columns(2)
        with bc1:
            st.markdown("**Bill (Image / PDF)**")
            bill_file = st.file_uploader("Upload bill", type=["jpg","jpeg","png","pdf"],
                                          key=f"{key}_bill", label_visibility="collapsed")
            if bill_file:
                file_id = f"{bill_file.name}_{bill_file.size}"
                if st.session_state.get(ss_file) != file_id:
                    bill_data = bill_file.read()
                    with st.spinner("Scanning bill…"):
                        try:
                            details = ocr_bill(bill_data, bill_file.name)
                            amt_result  = details.get("amount")
                            date_result = details.get("date")
                            st.session_state[ss_amt]  = amt_result
                            st.session_state[ss_date] = date_result
                            st.session_state[ss_file] = file_id
                            st.session_state[f"{key}_bill_bytes"] = (bill_data, bill_file.name)
                            # Inject into widget session state BEFORE widgets render
                            if date_result:
                                st.session_state[f"{key}_date"] = date_result
                            if amt_result:
                                st.session_state[f"{key}_bill_amt"] = float(amt_result)
                        except Exception as e:
                            st.markdown(f'<span class="ocr-err">⚠ OCR error: {e}</span>', unsafe_allow_html=True)
                    st.rerun()   # rerun so widgets render with injected values

                # Show OCR status
                ocr_amt  = st.session_state.get(ss_amt)
                ocr_date_val = st.session_state.get(ss_date)
                if ocr_amt:
                    st.markdown(f'<span class="ocr-ok">✓ Amount: ₹{ocr_amt:.2f}</span>', unsafe_allow_html=True)
                else:
                    st.markdown('<span class="ocr-err">⚠ Amount not detected — enter manually</span>', unsafe_allow_html=True)
                if ocr_date_val:
                    st.markdown(f'<span class="ocr-ok">✓ Date: {ocr_date_val.strftime("%d/%m/%Y")} — filled below</span>', unsafe_allow_html=True)
                else:
                    st.markdown('<span class="ocr-err">⚠ Date not detected — select manually</span>', unsafe_allow_html=True)

                bdata = st.session_state.get(f"{key}_bill_bytes")
                if bdata and bdata[0]:
                    bill_files_list.append(bdata)

        with bc2:
            st.markdown("**GPay Screenshot** *(if extra paid)*")
            gpay_file = st.file_uploader("Upload GPay", type=["jpg","jpeg","png"],
                                          key=f"{key}_gpay", label_visibility="collapsed")
            if gpay_file:
                gpay_data = gpay_file.read()
                with st.spinner("Scanning GPay…"):
                    try:
                        amt = ocr_gpay(gpay_data, gpay_file.name)
                        if amt:
                            st.session_state[f"{key}_gpay_amt"] = float(amt)
                            st.markdown(f'<span class="ocr-ok">✓ Detected ₹{amt:.2f}</span>', unsafe_allow_html=True)
                        else:
                            st.markdown('<span class="ocr-err">⚠ Could not extract — enter manually</span>', unsafe_allow_html=True)
                    except Exception as e:
                        st.markdown(f'<span class="ocr-err">⚠ OCR error: {e}</span>', unsafe_allow_html=True)
                bill_files_list.append((gpay_data, gpay_file.name))

        st.divider()

        # ── Row fields — render AFTER upload so OCR values are in session state ─
        rc1, rc2, rc3 = st.columns(3)
        with rc1:
            row_date = st.date_input("Date", key=f"{key}_date")
        with rc2:
            acc_sel = st.selectbox("Accounting Head", ACC_OPTIONS, key=f"{key}_acc")
        with rc3:
            mode_options = ["Auto", "Cab", "Train", "Bus", "Flight", "Others — specify"]
            mode_sel = st.selectbox("Mode of Travel", mode_options, key=f"{key}_mode_sel")
            if mode_sel == "Others — specify":
                mode = st.text_input("Specify mode", placeholder="e.g. Bike, Ferry…", key=f"{key}_mode_other")
            else:
                mode = mode_sel

        rc4, rc5, rc6 = st.columns(3)
        with rc4:
            rem_options = ["Onward", "Return", "Onward & Return", "Others — specify"]
            rem_sel = st.selectbox("Remarks", rem_options, key=f"{key}_rem_sel")
            if rem_sel == "Others — specify":
                remarks = st.text_input("Specify remarks", placeholder="e.g. Site visit, Client meeting…", key=f"{key}_rem_other")
            else:
                remarks = rem_sel
        with rc5:
            sup_bills = st.selectbox("Sup. Bills?", ["Yes","No"], key=f"{key}_sup")
        with rc6:
            if team_member_list:
                member = st.selectbox("Team Member", team_member_list, key=f"{key}_mem")
            else:
                member = st.text_input("Team Member", placeholder="Add team members in Step 1 first", key=f"{key}_mem")

        # ── Amount fields ─────────────────────────────────────────────────────
        amt1, amt2 = st.columns(2)
        with amt1:
            bill_amt = st.number_input("Bill Amount (₹)", min_value=0.0,
                                        step=0.01, format="%.2f", key=f"{key}_bill_amt")
        with amt2:
            gpay_amt = st.number_input("GPay Value (₹)", min_value=0.0,
                                        step=0.01, format="%.2f", key=f"{key}_gpay_amt")

        # Total = whichever is higher (bill or GPay)
        # GPay paid = actual amount paid when it exceeds the bill
        total_amt = max(bill_amt, gpay_amt)
        st.markdown(f"**Total: ₹ {total_amt:,.2f}**")

        acc_val = 0
        if acc_sel != "— select —":
            acc_val = int(acc_sel.split(" — ")[0])
            if thr_enabled:
                thr = thresholds.get(acc_val, 0)
                if thr and total_amt > thr:
                    st.markdown(f'<span class="flag-warn">⚠ Exceeds ₹{thr:.0f} limit</span>', unsafe_allow_html=True)

        collected_rows.append({
            "date":           row_date,
            "acc":            acc_val,
            "mode_of_travel": mode,
            "remarks":        remarks,
            "sup_bills":      sup_bills,
            "team_member":    member,
            "bill_amount":    bill_amt,
            "extra":          gpay_amt,
            "total":          max(bill_amt, gpay_amt),
        })

col_add, col_remove = st.columns([1, 5])
with col_add:
    if st.button("➕ Add Row"):
        st.session_state.row_count += 1
        st.rerun()
with col_remove:
    if st.session_state.row_count > 1 and st.button("➖ Remove Last Row"):
        st.session_state.row_count -= 1
        st.rerun()

st.divider()

# ── STEP 4: Generate ──────────────────────────────────────────────────────────
st.markdown("### 4️⃣ Generate Output")

if st.button("🚀 Generate Excel & PDF", type="primary"):
    errors = []
    if not client_name:     errors.append("Client Name is required")
    if not nature_of_audit: errors.append("Nature of Audit is required")
    if not team_members:    errors.append("Team Members is required")
    if not signature:       errors.append("Signature is required")
    if not any(r["acc"] for r in collected_rows):
        errors.append("Add at least one expense row with an accounting head")

    if errors:
        for e in errors:
            st.error(e)
    else:
        header = {
            "client_name":    client_name,
            "client_location": client_location,
            "nature_of_audit": nature_of_audit,
            "team_members":   team_members,
            "audit_start":    audit_start.strftime("%d/%m/%Y") if audit_start else "",
            "audit_end":      audit_end.strftime("%d/%m/%Y")   if audit_end   else "",
            "total_mandays":  total_mandays,
            "advance_amount": advance_amount,
            "tickets_booked": tickets_booked,
            "onward_by":      onward_by,
            "return_by":      return_by,
            "submission_date": submission_date.strftime("%d/%m/%Y") if submission_date else "",
            "signature":      signature,
        }

        with st.spinner("Building Excel…"):
            excel_bytes = build_excel(header, collected_rows, thresholds, thr_enabled)

        with st.spinner("Merging bills into PDF…"):
            pdf_bytes = build_pdf(bill_files_list)

        run_id = uuid.uuid4().hex[:6].upper()
        st.success("✅ Files ready for download!")

        dl1, dl2 = st.columns(2)
        with dl1:
            st.download_button(
                label="⬇ Download Excel",
                data=excel_bytes,
                file_name=f"TA_Claim_{client_name.replace(' ','_')}_{run_id}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
        with dl2:
            if pdf_bytes:
                st.download_button(
                    label="⬇ Download Consolidated PDF",
                    data=pdf_bytes,
                    file_name=f"Bills_Consolidated_{run_id}.pdf",
                    mime="application/pdf",
                )
            else:
                st.info("No bill files uploaded — PDF not generated")
