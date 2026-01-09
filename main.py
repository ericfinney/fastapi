import os
import uuid
import json
import logging
import re
import math
import time
from copy import copy
from typing import Dict, Any, List, Optional, Tuple

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
    """
    Deletes generated proposal files older than N minutes.
    Safe-guards:
      - Only deletes files starting with filename_prefix
      - Only deletes allowed_exts
      - Only deletes regular files
    """
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
    """
    Always round UP to the next whole dollar.
    12.00 -> 12
    12.01 -> 13
    """
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
    """
    Reinserts logo at A1 every time. Pillow must be installed.
    """
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

def lock_all_cells(ws, max_row: int, max_col: int) -> None:
    """
    Lock all cells in a rectangular region.
    """
    for r in range(1, max_row + 1):
        for c in range(1, max_col + 1):
            ws.cell(row=r, column=c).protection = Protection(locked=True)


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
# Handles:
#   "A1 - Room ID"
#   "A1- Room ID"
#   "A1 -Room ID"
#   "E5.P&P, D/F - Double Sided 12 x 18 DOT, Post w/ Plate Mount"
TYPE_DESC_SPLIT_RE = re.compile(r"\s*[-–—]\s*", flags=re.UNICODE)

def looks_like_sign_code(code: str) -> bool:
    """
    Allow commas/spaces like: "E5.P&P, D/F"
    Keep guardrails to avoid splitting normal sentences.
    """
    if not code:
        return False
    c = code.strip()
    if not c or len(c) > 60:
        return False
    if not re.match(r"^[A-Za-z0-9]", c):
        return False
    if not re.match(r"^[A-Za-z0-9./&_ ,]+$", c):
        return False
    return True

def split_sign_type_and_summary(raw_sign_type: str) -> Tuple[str, str]:
    """
    Splits the FIRST dash separator into:
      (sign_type_code, summary_text)

    Examples:
      "A1 - Room ID" -> ("A1", "Room ID")
      "A1- Room ID"  -> ("A1", "Room ID")
      "A1 -Room ID"  -> ("A1", "Room ID")
      "E5.P&P, D/F - Double Sided 12 x 18 DOT, Post w/ Plate Mount"
        -> ("E5.P&P, D/F", "Double Sided 12 x 18 DOT, Post w/ Plate Mount")
    """
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
        # Don't copy protection - we'll set it explicitly later
        # dst.protection = copy(src.protection)


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
# Totals helpers
# =========================================================
def sum_extended(items: Optional[List[Dict[str, Any]]]) -> Optional[float]:
    if not items:
        return None
    total = 0.0
    found = False
    for it in items:
        val = safe_num(it.get("extended_total"))
        if val is not None:
            total += val
            found = True
    return total if found else None


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
# Selection control: ONLY allow selecting body B–E
# =========================================================
def unlock_body_selection(ws, body_row_start: int, body_row_end: int) -> None:
    """
    Excel decides selection/lock behavior for merged ranges based on the *top-left* cell.
    So we must:
      1) Unlock B–E in body rows
      2) Unlock the top-left of any merged range that intersects B–E and the body rows
      3) Keep F locked (and everything else locked)
    """
    # 1) Unlock B–E directly
    for r in range(body_row_start, body_row_end + 1):
        for col in ("B", "C", "D", "E"):
            unlock_cell(ws, f"{col}{r}")

    # 2) Unlock top-left of merged ranges that intersect (rows body) and (cols B–E)
    # B=2, E=5
    for rng in ws.merged_cells.ranges:
        min_row, max_row = rng.min_row, rng.max_row
        min_col, max_col = rng.min_col, rng.max_col

        rows_intersect = not (max_row < body_row_start or min_row > body_row_end)
        cols_intersect = not (max_col < 2 or min_col > 5)

        if rows_intersect and cols_intersect:
            ws.cell(row=min_row, column=min_col).protection = Protection(locked=False)


def apply_sheet_protection_for_selection(ws, body_row_start: int, body_row_end: int) -> None:
    """
    Locks everything, unlocks only B–E body cells, then enables sheet protection
    so ONLY those unlocked cells are selectable.
    """
    logging.info(f"Applying sheet protection for body rows {body_row_start} to {body_row_end}")
    
    # Lock everything in the used region
    max_row = max(ws.max_row, body_row_end + 5)
    max_col = max(ws.max_column, 6)  # at least A–F
    logging.info(f"Locking all cells up to row {max_row}, col {max_col}")
    lock_all_cells(ws, max_row=max_row, max_col=max_col)

    # Unlock the allowed body selection area (and merged top-lefts)
    logging.info(f"Unlocking body cells B-E in rows {body_row_start}-{body_row_end}")
    unlock_body_selection(ws, body_row_start, body_row_end)
    
    # Verify a few cells are actually unlocked BEFORE enabling protection
    test_cells = [f"B{body_row_start}", f"C{body_row_start}", f"D{body_row_start}", f"E{body_row_start}", f"A{body_row_start}", f"F{body_row_start}"]
    for cell_ref in test_cells:
        locked = ws[cell_ref].protection.locked
        logging.info(f"BEFORE protection - Cell {cell_ref} locked status: {locked}")

    # Prevent selecting locked cells, allow selecting unlocked cells
    ws.protection.selectLockedCells = False
    ws.protection.selectUnlockedCells = True

    # Optional hardening
    ws.protection.formatCells = False
    ws.protection.formatColumns = False
    ws.protection.formatRows = False
    ws.protection.insertRows = False
    ws.protection.deleteRows = False
    ws.protection.insertColumns = False
    ws.protection.deleteColumns = False
    
    # Enable protection with no password using the enable() method
    logging.info("Enabling sheet protection with enable() method")
    ws.protection.enable()
    
    # Verify protection settings
    logging.info(f"Sheet protection enabled: {ws.protection.sheet}")
    logging.info(f"Select locked cells: {ws.protection.selectLockedCells}")
    logging.info(f"Select unlocked cells: {ws.protection.selectUnlockedCells}")
    
    # Verify cells are STILL unlocked after enabling protection
    for cell_ref in test_cells:
        locked = ws[cell_ref].protection.locked
        logging.info(f"AFTER protection - Cell {cell_ref} locked status: {locked}")


# =========================================================
# FastAPI endpoints
# =========================================================
@app.get("/")
def root():
    return {
        "status": "ok",
        "template_exists": os.path.exists(TEMPLATE_PATH),
        "template_path": TEMPLATE_PATH,
        "sheet_name": SHEET_NAME,
        "logo_exists": os.path.exists(LOGO_PATH),
        "logo_path": LOGO_PATH
    }


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

    if not isinstance(estimate_data, dict) or not estimate_data:
        raise HTTPException(status_code=400, detail="Decoded 'payload' must be a non-empty JSON object.")

    if not os.path.exists(TEMPLATE_PATH):
        raise HTTPException(status_code=500, detail=f"Template not found at {TEMPLATE_PATH}")

    try:
        wb = load_workbook(TEMPLATE_PATH)
        if SHEET_NAME not in wb.sheetnames:
            raise HTTPException(status_code=500, detail=f"Sheet '{SHEET_NAME}' not found in workbook.")
        ws = wb[SHEET_NAME]
        
        # Completely remove any existing sheet protection from the template
        # Create a fresh protection object with no password
        from openpyxl.worksheet.protection import SheetProtection
        ws.protection = SheetProtection()

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

        # ---------------- Write sign lines ----------------
        COL_ITEM, COL_SIGN_TYPE, COL_DESC, COL_QTY, COL_UNIT, COL_TOTAL = "A", "B", "C", "D", "E", "F"
        current_row = BODY_START
        item_num = 1

        # Track duplicates by sign code
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
                ws[f"{COL_QTY}{current_row}"].value = safe_num(sign.get("qty"))
                ws[f"{COL_DESC}{current_row}"].value = desc_summary
                ws[f"{COL_TOTAL}{current_row}"].value = round_up_dollars(sign.get("extended_total"))
                ws[f"{COL_UNIT}{current_row}"].value = round_up_dollars(sign.get("unit_price"))
            else:
                ws[f"{COL_SIGN_TYPE}{current_row}"].value = None
                ws[f"{COL_QTY}{current_row}"].value = None
                ws[f"{COL_DESC}{current_row}"].value = f"ALTERNATE {desc_summary}"
                ws[f"{COL_TOTAL}{current_row}"].value = None
                ws[f"{COL_UNIT}{current_row}"].value = round_up_dollars(sign.get("unit_price"))

            current_row += 1
            item_num += 1

        # ---------------- Totals (hard-coded cells shifted) ----------------
        totals = estimate_data.get("totals", {}) or {}
        subtotal = safe_num(totals.get("sub_total"))
        grand_total = safe_num(totals.get("total"))
        shipping_total = sum_extended(estimate_data.get("shipping"))
        install_total = sum_extended(estimate_data.get("installation"))

        SUBTOTAL_CELL = "F48"
        SHIPPING_CELL = "F49"
        INSTALL_CELL = "F53"
        TOTAL_CELL = "F54"

        if subtotal is not None:
            write_cell(ws, shift_cell_ref(SUBTOTAL_CELL, footer_row_offset), subtotal)
        if shipping_total is not None:
            write_cell(ws, shift_cell_ref(SHIPPING_CELL, footer_row_offset), shipping_total)
        if install_total is not None:
            write_cell(ws, shift_cell_ref(INSTALL_CELL, footer_row_offset), install_total)
        if grand_total is not None:
            write_cell(ws, shift_cell_ref(TOTAL_CELL, footer_row_offset), grand_total)

        # ---------------- Row height adjustment below row 26 ----------------
        last_used_row = ws.max_row
        approximate_autofit_rows(
            ws,
            row_start=27,
            row_end=last_used_row,
            text_cols=["C"],
            min_height=15.0
        )

        # Restore footer row heights to template values (shifted)
        restore_row_heights(ws, footer_row_heights, footer_row_offset)

        # ---------------- Selection rules (ONLY allow selecting B–E in body) ----------------
        # CRITICAL: Must be the VERY LAST thing before saving - don't touch any cells after this!
        apply_sheet_protection_for_selection(ws, body_row_start=BODY_START, body_row_end=body_last_row)
        
        logging.info("About to save workbook - no more cell modifications after this point")

        # ---------------- Save output workbook ----------------
        file_id = uuid.uuid4().hex
        out_name = f"Boyd_Proposal_{file_id}.xlsx"
        out_path = os.path.join(OUTPUT_DIR, out_name)
        wb.save(out_path)

    except Exception as e:
        logging.exception("Proposal generation failed")
        raise HTTPException(status_code=500, detail=str(e))

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