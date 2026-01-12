import os
import uuid
import json
import logging
import re
import math
import time
from copy import copy
from typing import Dict, Any, List, Optional, Tuple


import base64
from pydantic import BaseModel
from pypdf import PdfReader

from fastapi import FastAPI, Body, HTTPException
from fastapi.responses import FileResponse, JSONResponse
from openpyxl import load_workbook
from openpyxl.drawing.image import Image as XLImage
from openpyxl.styles import Protection

logging.basicConfig(level=logging.INFO)

TEMPLATE_PATH = os.environ.get("BOYD_TEMPLATE_PATH", "templates/Blank.xlsx")
SHEET_NAME = os.environ.get("BOYD_SHEET_NAME", "Proposal")
OUTPUT_DIR = os.environ.get("BOYD_OUTPUT_DIR", "/tmp/output")
LOGO_PATH = os.environ.get("BOYD_LOGO_PATH", "assets/logo.png")

os.makedirs(OUTPUT_DIR, exist_ok=True)

app = FastAPI()


# =========================================================
# Cleanup: delete old generated workbooks (older than N minutes)
# =========================================================
def cleanup_old_generated_workbooks(
    directory: str,
    older_than_minutes: int = 30,
    filename_prefix: str = "Boyd_Proposal_",
    allowed_exts: Tuple[str, ...] = (".xlsx", ".xlsm"),
) -> None:
    if not os.path.isdir(directory):
        return

    now = time.time()
    max_age_sec = older_than_minutes * 60

    for fn in os.listdir(directory):
        fn_lower = fn.lower()
        if not fn.startswith(filename_prefix):
            continue
        if not fn_lower.endswith(tuple(ext.lower() for ext in allowed_exts)):
            continue

        path = os.path.join(directory, fn)
        if not os.path.isfile(path):
            continue

        try:
            age_sec = now - os.path.getmtime(path)
            if age_sec > max_age_sec:
                os.remove(path)
                logging.info("Deleted old output workbook: %s (age %.1f min)", fn, age_sec / 60.0)
        except Exception as e:
            logging.warning("Failed deleting %s: %s", fn, e)


# =========================================================
# Basic helpers
# =========================================================
def safe_str(x) -> str:
    return "" if x is None else str(x)

def safe_num(x):
    try:
        if x is None or x == "":
            return None
        return float(x)
    except Exception:
        return None

def round_up_dollars(value):
    if value is None or value == "":
        return None
    try:
        return math.ceil(float(value))
    except Exception:
        return None

def join_address_lines(addr_lines: List[str]) -> str:
    return "\n".join([line for line in addr_lines if line and line.strip()])

def write_cell(ws, cell: str, value):
    ws[cell].value = value

def insert_logo(ws):
    if not os.path.exists(LOGO_PATH):
        logging.warning("Logo not found at %s; skipping insert.", LOGO_PATH)
        return
    img = XLImage(LOGO_PATH)
    ws.add_image(img, "A1")


# =========================================================
# Lock/Unlock helpers
# =========================================================
def lock_cell(ws, cell_ref: str):
    ws[cell_ref].protection = Protection(locked=True)

def unlock_cell(ws, cell_ref: str):
    ws[cell_ref].protection = Protection(locked=False)


# =========================================================
# Footer row height capture/restore
# =========================================================
def capture_row_heights(ws, start_row: int, end_row: int) -> dict:
    heights = {}
    for r in range(start_row, end_row + 1):
        heights[r] = ws.row_dimensions[r].height
    return heights

def restore_row_heights(ws, heights: dict, row_offset: int):
    for original_row, height in heights.items():
        target_row = original_row + row_offset
        ws.row_dimensions[target_row].height = height


# =========================================================
# Sign type + summary split (ROBUST)
# =========================================================
TYPE_DESC_SPLIT_RE = re.compile(r"\s*[-–—]\s*", flags=re.UNICODE)

def looks_like_sign_code(code: str) -> bool:
    if not code:
        return False
    c = code.strip()
    if not c or len(c) > 40:
        return False
    if not re.match(r"^[A-Za-z0-9]", c):
        return False
    if not re.match(r"^[A-Za-z0-9./&_ ,]+$", c):
        return False
    return True

def split_sign_type_and_summary(raw_sign_type: str) -> Tuple[str, str]:
    if not raw_sign_type:
        return "", ""

    s = str(raw_sign_type).strip()
    if not s:
        return "", ""

    parts = TYPE_DESC_SPLIT_RE.split(s, maxsplit=1)
    if len(parts) == 2:
        code = parts[0].strip()
        summary = parts[1].strip()
        if looks_like_sign_code(code):
            return code, summary

    return s, ""


# =========================================================
# Merge shifting helpers (used by adjust body)
# =========================================================
CELL_RE = re.compile(r"^([A-Z]+)(\d+)$")

def shift_cell_ref(cell_ref: str, row_offset: int) -> str:
    m = CELL_RE.match(cell_ref)
    if not m:
        return cell_ref
    col, row = m.group(1), int(m.group(2))
    return f"{col}{row + row_offset}"

def parse_range(a1_range: str) -> Tuple[str, str]:
    if ":" in a1_range:
        a, b = a1_range.split(":")
        return a, b
    return a1_range, a1_range

def split_ref(ref: str):
    m = CELL_RE.match(ref)
    if not m:
        return ref, None
    return m.group(1), int(m.group(2))

def shift_range_overlap_safe(a1_range: str, footer_start_row: int, row_offset: int) -> str:
    a, b = parse_range(a1_range)
    _, a_row = split_ref(a)
    _, b_row = split_ref(b)

    if a_row is None or b_row is None:
        return a1_range

    top = min(a_row, b_row)
    bottom = max(a_row, b_row)

    should_shift = (top >= footer_start_row) or (top < footer_start_row <= bottom)

    def apply_shift(ref):
        col, row = split_ref(ref)
        if row is None:
            return ref
        if should_shift:
            row += row_offset
        return f"{col}{row}"

    new_a = apply_shift(a)
    new_b = apply_shift(b)
    return f"{new_a}:{new_b}" if ":" in a1_range else new_a

def save_merged_ranges(ws) -> List[str]:
    return [str(rng) for rng in ws.merged_cells.ranges]

def unmerge_all(ws, merges: List[str]):
    for rng in merges:
        ws.unmerge_cells(rng)

def restore_merges(ws, merges: List[str], footer_start_row: int, row_offset: int):
    for rng in merges:
        new_rng = shift_range_overlap_safe(rng, footer_start_row, row_offset)
        ws.merge_cells(new_rng)

def copy_row_style(ws, src_row: int, dst_row: int, max_col: int):
    ws.row_dimensions[dst_row].height = ws.row_dimensions[src_row].height
    for col in range(1, max_col + 1):
        src = ws.cell(row=src_row, column=col)
        dst = ws.cell(row=dst_row, column=col)
        if src.has_style:
            dst._style = copy(src._style)
        dst.number_format = src.number_format
        dst.alignment = copy(src.alignment)
        dst.border = copy(src.border)
        dst.fill = copy(src.fill)
        dst.font = copy(src.font)
        dst.protection = copy(src.protection)


# =========================================================
# Body adjust with merge + style preservation
# =========================================================
def adjust_body_rows_preserve_footer(
    ws,
    sign_count: int,
    body_start: int = 28,
    body_end: int = 47,
    extra_blank_rows: int = 3
) -> int:
    base_rows = body_end - body_start + 1
    needed_rows = sign_count + extra_blank_rows
    footer_start = body_end + 1
    diff = needed_rows - base_rows

    if diff == 0:
        return 0

    merges = save_merged_ranges(ws)
    unmerge_all(ws, merges)

    max_col = ws.max_column

    if diff > 0:
        logging.info("Inserting %d row(s) at %d to expand body.", diff, footer_start)
        ws.insert_rows(footer_start, amount=diff)
        for r in range(footer_start, footer_start + diff):
            copy_row_style(ws, src_row=body_end, dst_row=r, max_col=max_col)
    else:
        delete_count = abs(diff)
        delete_start = body_start + needed_rows
        logging.info("Deleting %d row(s) at %d to shrink body.", delete_count, delete_start)
        ws.delete_rows(delete_start, amount=delete_count)

    restore_merges(ws, merges, footer_start, diff)
    return diff


# =========================================================
# Approximate Row Height "AutoFit"
# =========================================================
def approximate_autofit_rows(ws, row_start: int, row_end: int, text_cols: List[str], min_height: float = 15.0):
    CHARS_PER_LINE = 60
    LINE_HEIGHT = 15
    for r in range(row_start, row_end + 1):
        max_lines = 1
        for col in text_cols:
            v = ws[f"{col}{r}"].value
            if not v:
                continue
            text = str(v)
            explicit_lines = text.split("\n")
            line_count = 0
            for ln in explicit_lines:
                if not ln:
                    line_count += 1
                else:
                    wrapped = max(1, (len(ln) // CHARS_PER_LINE) + (1 if len(ln) % CHARS_PER_LINE else 0))
                    line_count += wrapped
            max_lines = max(max_lines, line_count)
        ws.row_dimensions[r].height = max(min_height, max_lines * LINE_HEIGHT)


# =========================================================
# NEW: unlock merged-cell top-lefts that overlap B–E body rows
# =========================================================
def unlock_body_selection(ws, body_row_start: int, body_row_end: int) -> None:
    """
    Excel uses the top-left cell of a merged range to determine locked/unlocked.
    If body rows contain merged cells whose top-left is locked, selection gets blocked.
    This unlocks:
      - all B–E cells in body rows
      - the top-left cell of any merged range overlapping body rows and columns B–E
    """
    # 1) Unlock B–E cells directly
    for r in range(body_row_start, body_row_end + 1):
        for col in ("B", "C", "D", "E"):
            unlock_cell(ws, f"{col}{r}")
        # keep totals locked
        lock_cell(ws, f"F{r}")

    # 2) Unlock top-left of merged ranges that intersect (rows body) and (cols B–E)
    # B=2, E=5
    for rng in ws.merged_cells.ranges:
        min_row, max_row = rng.min_row, rng.max_row
        min_col, max_col = rng.min_col, rng.max_col

        rows_intersect = not (max_row < body_row_start or min_row > body_row_end)
        cols_intersect = not (max_col < 2 or min_col > 5)

        if rows_intersect and cols_intersect:
            top_left = ws.cell(row=min_row, column=min_col)
            top_left.protection = Protection(locked=False)


# =========================================================
# Build proposal workbook from extracted estimate_data dict
# =========================================================
def build_proposal_workbook(estimate_data: Dict[str, Any]) -> Tuple[str, str]:
    """
    Generates proposal workbook from estimate_data dict and returns (out_name, out_path).
    """
    if not isinstance(estimate_data, dict) or not estimate_data:
        raise HTTPException(status_code=400, detail="estimate_data must be a non-empty JSON object.")

    if not os.path.exists(TEMPLATE_PATH):
        raise HTTPException(status_code=500, detail=f"Template not found at {TEMPLATE_PATH}")

    wb = load_workbook(TEMPLATE_PATH)
    if SHEET_NAME not in wb.sheetnames:
        raise HTTPException(status_code=500, detail=f"Sheet '{SHEET_NAME}' not found in workbook.")
    ws = wb[SHEET_NAME]

    # Disable protection temporarily but preserve template settings
    ws.protection.sheet = False
    insert_logo(ws)

    # Capture template footer row heights BEFORE any insertion/deletion
    FOOTER_HEIGHT_START = 48
    FOOTER_HEIGHT_END = 120
    footer_row_heights = capture_row_heights(ws, FOOTER_HEIGHT_START, FOOTER_HEIGHT_END)

    # ---------------- Header mapping ----------------
    write_cell(ws, "E5", safe_str(estimate_data.get("estimate_date")))
    write_cell(ws, "D8", safe_str(estimate_data.get("project_id")))
    write_cell(ws, "C22", safe_str(estimate_data.get("salesperson")))
    write_cell(ws, "C23", safe_str(estimate_data.get("project_manager")))
    write_cell(ws, "C25", safe_str(estimate_data.get("project_description")))

    # ---------------- Sold-to / Ship-to ----------------
    sold_to = estimate_data.get("sold_to", {}) or {}
    ship_to = estimate_data.get("ship_to", {}) or {}

    write_cell(ws, "D11", safe_str(sold_to.get("name")))
    write_cell(ws, "D13", join_address_lines(sold_to.get("address_lines") or []))
    sold_csz = " ".join([p for p in [
        safe_str(sold_to.get("city")),
        safe_str(sold_to.get("state")),
        safe_str(sold_to.get("zip"))
    ] if p.strip()])
    write_cell(ws, "D16", sold_csz)
    write_cell(ws, "D17", safe_str(sold_to.get("phone")))

    write_cell(ws, "C11", safe_str(ship_to.get("name")))
    write_cell(ws, "C13", join_address_lines(ship_to.get("address_lines") or []))
    ship_csz = " ".join([p for p in [
        safe_str(ship_to.get("city")),
        safe_str(ship_to.get("state")),
        safe_str(ship_to.get("zip"))
    ] if p.strip()])
    write_cell(ws, "C16", ship_csz)
    write_cell(ws, "C17", safe_str(ship_to.get("phone")))

    # ---------------- Dynamic body resize ----------------
    sign_types = estimate_data.get("sign_types", []) or []
    sign_count = len(sign_types)

    BODY_START = 28
    BODY_END = 47
    EXTRA_BLANK = 3

    footer_row_offset = adjust_body_rows_preserve_footer(
        ws,
        sign_count=sign_count,
        body_start=BODY_START,
        body_end=BODY_END,
        extra_blank_rows=EXTRA_BLANK
    )

    total_body_rows_needed = sign_count + EXTRA_BLANK
    body_last_row = BODY_START + total_body_rows_needed - 1
    body_change_count = footer_row_offset - 5
    if body_change_count > 0:
        change_text = f"Add {body_change_count}"
    elif body_change_count < 0:
        change_text = f"Subtract {body_change_count}"
    else:
        change_text = "No Change"
    ws["H3"].value = change_text

    # ---------------- Write sign lines ----------------
    COL_ITEM, COL_SIGN_TYPE, COL_DESC, COL_QTY, COL_UNIT, COL_TOTAL = "A", "B", "C", "D", "E", "F"
    current_row = BODY_START
    item_num = 1

    sign_code_counts: Dict[str, int] = {}

    for sign in sign_types:
        ws[f"{COL_ITEM}{current_row}"].value = item_num

        raw_type = safe_str(sign.get("sign_type"))
        clean_type, summary = split_sign_type_and_summary(raw_type)

        sign_code_counts.setdefault(clean_type, 0)
        sign_code_counts[clean_type] += 1
        occurrence = sign_code_counts[clean_type]

        desc_summary = summary.strip() if summary else safe_str(sign.get("description")).strip()

        if occurrence == 1:
            ws[f"{COL_SIGN_TYPE}{current_row}"].value = clean_type
            qty_val = safe_num(sign.get("qty"))
            unit_val = round_nearest_dollar(sign.get("unit_price"))
            ws[f"{COL_QTY}{current_row}"].value = qty_val
            ws[f"{COL_DESC}{current_row}"].value = desc_summary
            ws[f"{COL_UNIT}{current_row}"].value = unit_val
            if qty_val is not None and unit_val is not None:
                ws[f"{COL_TOTAL}{current_row}"].value = f"=D{current_row}*E{current_row}"
            else:
                ws[f"{COL_TOTAL}{current_row}"].value = None
        else:
            ws[f"{COL_SIGN_TYPE}{current_row}"].value = None
            ws[f"{COL_QTY}{current_row}"].value = None
            ws[f"{COL_DESC}{current_row}"].value = f"ALTERNATE {desc_summary}"
            unit_val = round_nearest_dollar(sign.get("unit_price"))
            ws[f"{COL_UNIT}{current_row}"].value = unit_val
            ws[f"{COL_TOTAL}{current_row}"].value = None

        current_row += 1
        item_num += 1

    # ---------------- Totals ----------------
    totals = estimate_data.get("totals", {}) or {}
    grand_total = safe_num(totals.get("total"))
    shipping_total = sum_extended(estimate_data.get("shipping"))
    install_total = sum_extended(estimate_data.get("installation"))

    SUBTOTAL_CELL = "F48"
    SHIPPING_CELL = "F49"
    INSTALL_CELL = "F53"
    TOTAL_CELL = "F54"

    subtotal_cell_ref = shift_cell_ref(SUBTOTAL_CELL, footer_row_offset)
    body_sum_range = f"F{BODY_START}:F{body_last_row}"
    ws[subtotal_cell_ref].value = f"=SUM({body_sum_range})"

    shipping_value = shipping_total if shipping_total else 0.00
    write_cell(ws, shift_cell_ref(SHIPPING_CELL, footer_row_offset), shipping_value)

    install_value = install_total if install_total else 0.00
    write_cell(ws, shift_cell_ref(INSTALL_CELL, footer_row_offset), install_value)

    if grand_total is not None:
        write_cell(ws, shift_cell_ref(TOTAL_CELL, footer_row_offset), grand_total)

    # ---------------- Row height adjustment ----------------
    last_used_row = ws.max_row
    approximate_autofit_rows(ws, row_start=27, row_end=last_used_row, text_cols=["C"], min_height=15.0)

    restore_row_heights(ws, footer_row_heights, footer_row_offset)

    # ---------------- Selection rules ----------------
    apply_sheet_protection_for_selection(ws, body_row_start=BODY_START, body_row_end=body_last_row)

    file_id = uuid.uuid4().hex
    out_name = f"Boyd_Proposal_{file_id}.xlsx"
    out_path = os.path.join(OUTPUT_DIR, out_name)
    wb.save(out_path)

    return out_name, out_path

# =========================================================
# FastAPI endpoints
# =========================================================
@app.post("/generate_proposal")
def generate_proposal(payload: Dict[str, Any] = Body(default=None)):
    logging.info("Incoming request: payload keys = %s", list(payload.keys()) if payload else None)

    if not payload or "payload" not in payload:
        raise HTTPException(status_code=400, detail="Missing required field 'payload' (JSON string).")

    # Delete old outputs (older than 30 minutes)
    cleanup_old_generated_workbooks(OUTPUT_DIR, older_than_minutes=30)

    try:
        estimate_data = json.loads(payload["payload"])
    except Exception as e:
        raise HTTPException(status_code=400, detail=f"Invalid JSON string in 'payload': {str(e)}")

    out_name, _ = build_proposal_workbook(estimate_data)

    base_url = os.environ.get("RAILWAY_PUBLIC_URL", "").rstrip("/")
    if not base_url:
        base_url = "https://fastapi-production-37f6.up.railway.app"

    download_url = f"{base_url}/download/{out_name}"
    return JSONResponse({"download_url": download_url, "filename": out_name})

@app.get("/download/{filename}")
def download_file(filename: str):
    file_path = os.path.join(OUTPUT_DIR, filename)
    if not os.path.exists(file_path):
        raise HTTPException(status_code=404, detail="File not found")

    return FileResponse(
        file_path,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        filename=filename
    )


# =========================================================
# PDF -> estimate_data extraction (best-effort, text-based)
# =========================================================
MONEY_RE = re.compile(r"^\$?\(?([0-9]{1,3}(?:,[0-9]{3})*|[0-9]+)(?:\.[0-9]{2})?\)?$")

def _money_to_float(s: str) -> Optional[float]:
    if s is None:
        return None
    t = str(s).strip()
    if not t:
        return None
    neg = False
    if t.startswith("(") and t.endswith(")"):
        neg = True
        t = t[1:-1].strip()
    t = t.replace("$", "").replace(",", "")
    try:
        val = float(t)
        return -val if neg else val
    except Exception:
        return None

def extract_text_from_pdf(pdf_path: str) -> str:
    reader = PdfReader(pdf_path)
    parts = []
    for page in reader.pages:
        try:
            parts.append(page.extract_text() or "")
        except Exception:
            parts.append("")
    return "\n".join(parts)

def parse_estimate_pdf_text_to_data(text: str) -> Dict[str, Any]:
    """
    Best-effort extraction from text-based PDFs.
    This is intentionally conservative: it will fill sign_types and totals when recognizable,
    and leave other fields blank when not reliably detectable.
    """
    lines = [ln.strip() for ln in text.splitlines() if ln.strip()]
    data: Dict[str, Any] = {
        "estimate_date": "",
        "project_id": "",
        "project_description": "",
        "salesperson": "",
        "project_manager": "",
        "project_engineer": "",
        "sold_to": {"name": "", "address_lines": [], "city": "", "state": "", "zip": "", "phone": ""},
        "ship_to": {"name": "", "address_lines": [], "city": "", "state": "", "zip": "", "phone": ""},
        "sign_types": [],
        "shipping": [],
        "installation": [],
        "totals": {"sub_total": None, "mark_up": None, "total": None},
    }

    # Basic totals heuristics
    for ln in lines:
        lower = ln.lower()
        # try grab total like "Total: $12,345.67"
        if "grand total" in lower or lower.startswith("total"):
            m = re.search(r"\$\s*([0-9,]+(?:\.[0-9]{2})?)", ln)
            if m:
                data["totals"]["total"] = _money_to_float(m.group(1))

    # Line item heuristic: split by 2+ spaces; last 3 cols often qty/unit/ext or qty/ext
    for ln in lines:
        # quick filter for sign-like lines: must contain dash separator and a number
        if "-" not in ln and "–" not in ln and "—" not in ln:
            continue
        cols = re.split(r"\s{2,}", ln)
        if len(cols) < 2:
            continue

        # Try infer qty/unit/ext from right side
        # Look for patterns like: [desc..., qty, unit, ext] or [desc..., qty, ext]
        qty = None
        unit = None
        ext = None

        # Walk from right, collect up to 3 numeric-ish fields
        tail = cols[-3:] if len(cols) >= 3 else cols[-2:]
        # Flatten tokens in tail columns
        tokens = []
        for c in tail:
            tokens.extend(c.split())
        # Take last 3 tokens
        tokens = tokens[-3:]

        if len(tokens) >= 2:
            # last token is often extended total money
            maybe_ext = _money_to_float(tokens[-1])
            if maybe_ext is not None:
                ext = maybe_ext
            # previous token is unit price or qty
            maybe_unit_or_qty = _money_to_float(tokens[-2])
            # qty might be integer token (no $)
            maybe_qty = None
            try:
                maybe_qty = int(tokens[-2].replace(",", "").replace("$", "").replace("(", "").replace(")", ""))
            except Exception:
                maybe_qty = None

            # if we have 3 tokens, middle could be unit and left could be qty
            if len(tokens) == 3:
                try:
                    maybe_qty2 = int(tokens[-3].replace(",", "").replace("$", "").replace("(", "").replace(")", ""))
                except Exception:
                    maybe_qty2 = None
                maybe_unit2 = _money_to_float(tokens[-2])
                if maybe_qty2 is not None and maybe_unit2 is not None and ext is not None:
                    qty, unit = float(maybe_qty2), maybe_unit2

            # if only 2 tokens and second looks like money and first looks like int
            if qty is None and ext is not None:
                if maybe_qty is not None:
                    qty = float(maybe_qty)
                if maybe_unit_or_qty is not None and maybe_qty is None:
                    unit = maybe_unit_or_qty

        # Require at least qty or unit or ext to be found
        if qty is None and unit is None and ext is None:
            continue

        # Left side should contain sign type
        left = cols[0]
        sign_type, desc = split_sign_type_and_summary(left)
        if not sign_type:
            continue

        data["sign_types"].append({
            "sign_type": sign_type,
            "description": desc,
            "qty": qty,
            "unit_price": unit,
            "extended_total": ext,
        })

    return data


# =========================================================
# Chunked upload endpoints (supports BOTH /upload/* and /pdf-upload/*)
# =========================================================
UPLOAD_DIR = os.environ.get("BOYD_UPLOAD_DIR", "/tmp/uploads")
os.makedirs(UPLOAD_DIR, exist_ok=True)

class UploadInitResponseModel(BaseModel):
    upload_id: str

class UploadChunkRequestModel(BaseModel):
    upload_id: str
    index: int
    data_base64: str

class UploadFinishRequestModel(BaseModel):
    upload_id: str
    filename: str

def _upload_dir(upload_id: str) -> str:
    return os.path.join(UPLOAD_DIR, upload_id)

def _decode_base64_strict(s: str) -> bytes:
    s2 = "".join(str(s).split())
    try:
        return base64.b64decode(s2, validate=True)
    except Exception as e:
        raise HTTPException(status_code=400, detail=f"Invalid base64 payload: {e}")

def _ensure_upload_id(upload_id: str) -> str:
    d = _upload_dir(upload_id)
    if not os.path.isdir(d):
        raise HTTPException(status_code=404, detail="upload_id not found. Call init/start first.")
    return d

def _cleanup_upload_id(upload_id: str) -> None:
    d = _upload_dir(upload_id)
    try:
        shutil.rmtree(d, ignore_errors=True)
    except Exception:
        pass

def _stitch_chunks_to_pdf(upload_id: str) -> str:
    d = _ensure_upload_id(upload_id)
    chunk_files = sorted([fn for fn in os.listdir(d) if fn.endswith(".chunk")])
    if not chunk_files:
        raise HTTPException(status_code=400, detail="No chunks uploaded.")
    with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp:
        for fn in chunk_files:
            tmp.write(open(os.path.join(d, fn), "rb").read())
        return tmp.name

# Canonical
@app.post("/pdf-upload/start", response_model=UploadInitResponseModel)
def pdf_upload_start():
    upload_id = uuid.uuid4().hex
    os.makedirs(_upload_dir(upload_id), exist_ok=True)
    return UploadInitResponseModel(upload_id=upload_id)

@app.post("/pdf-upload/chunk")
def pdf_upload_chunk(req: UploadChunkRequestModel):
    d = _ensure_upload_id(req.upload_id)
    chunk_bytes = _decode_base64_strict(req.data_base64)
    path = os.path.join(d, f"{req.index:06d}.chunk")
    with open(path, "wb") as f:
        f.write(chunk_bytes)
    return {"ok": True, "saved_index": req.index, "saved_bytes": len(chunk_bytes)}

@app.post("/pdf-upload/finish")
def pdf_upload_finish(req: UploadFinishRequestModel):
    cleanup_old_generated_workbooks(OUTPUT_DIR, older_than_minutes=30)

    pdf_path = _stitch_chunks_to_pdf(req.upload_id)
    try:
        pdf_bytes = Path(pdf_path).read_bytes()
        if len(pdf_bytes) < 1000 or not pdf_bytes.startswith(b"%PDF"):
            raise HTTPException(status_code=400, detail="Reassembled file does not look like a valid PDF (missing/invalid chunks).")

        text = extract_text_from_pdf(pdf_path)
        estimate_data = parse_estimate_pdf_text_to_data(text)

        out_name, _ = build_proposal_workbook(estimate_data)

        base_url = os.environ.get("RAILWAY_PUBLIC_URL", "").rstrip("/")
        if not base_url:
            base_url = "https://fastapi-production-37f6.up.railway.app"
        download_url = f"{base_url}/download/{out_name}"
        return JSONResponse({"download_url": download_url, "filename": out_name})
    finally:
        try:
            os.unlink(pdf_path)
        except Exception:
            pass
        _cleanup_upload_id(req.upload_id)

# Compatibility aliases expected by your current OpenAPI
@app.post("/upload/init", response_model=UploadInitResponseModel)
def upload_init_alias():
    return pdf_upload_start()

@app.post("/upload/chunk")
def upload_chunk_alias(req: UploadChunkRequestModel):
    return pdf_upload_chunk(req)

@app.post("/upload/finish")
def upload_finish_alias(req: UploadFinishRequestModel):
    return pdf_upload_finish(req)
