import os
import uuid
import json
import logging
import re
from copy import copy
from typing import Dict, Any, List, Optional, Tuple

from fastapi import FastAPI, Body, HTTPException
from fastapi.responses import FileResponse, JSONResponse
from openpyxl import load_workbook
from openpyxl.drawing.image import Image as XLImage

logging.basicConfig(level=logging.INFO)

TEMPLATE_PATH = os.environ.get("BOYD_TEMPLATE_PATH", "templates/Blank.xlsx")
SHEET_NAME = os.environ.get("BOYD_SHEET_NAME", "Proposal")
OUTPUT_DIR = os.environ.get("BOYD_OUTPUT_DIR", "/tmp/output")
LOGO_PATH = os.environ.get("BOYD_LOGO_PATH", "assets/logo.png")

os.makedirs(OUTPUT_DIR, exist_ok=True)

app = FastAPI()


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

def join_address_lines(addr_lines: List[str]) -> str:
    return "\n".join([line for line in addr_lines if line and line.strip()])

def write_cell(ws, cell: str, value):
    ws[cell].value = value

def insert_logo(ws):
    """
    Reinserts logo at A1 every time. Pillow must be installed.
    """
    if not os.path.exists(LOGO_PATH):
        logging.warning(f"Logo not found at {LOGO_PATH}; skipping insert.")
        return
    img = XLImage(LOGO_PATH)
    ws.add_image(img, "A1")


# =========================================================
# Footer row height capture/restore
# =========================================================
def capture_row_heights(ws, start_row: int, end_row: int) -> dict:
    """
    Capture row heights from the template.
    Returns {row_number: height or None}.
    """
    heights = {}
    for r in range(start_row, end_row + 1):
        heights[r] = ws.row_dimensions[r].height
    return heights

def restore_row_heights(ws, heights: dict, row_offset: int):
    """
    Restore captured row heights, shifted by row_offset.
    Example: original row 48 height is applied to row 48 + row_offset.
    """
    for original_row, height in heights.items():
        target_row = original_row + row_offset
        ws.row_dimensions[target_row].height = height


# =========================================================
# Sign type + summary split (ROBUST)
# =========================================================
def split_sign_type_and_summary(raw_sign_type: str):
    """
    Split only on the FIRST dash used as CODE - SUMMARY separator.

    Supports codes like:
      D - Donor Room
      D- Donor Room
      E5.W - 12 x 18 DOT, Wall Mount
      E5.VA.P&P - Something
      A4.X - Exterior Utility Room ID

    Code allowed chars:
      letters, numbers, dots, slashes, ampersands, underscores
    """
    if not raw_sign_type:
        return "", ""

    s = raw_sign_type.strip()

    parts = re.split(r"\s*-\s*", s, maxsplit=1)
    if len(parts) == 2:
        code = parts[0].strip()
        summary = parts[1].strip()

        # Guardrail: only treat as code if it looks like a sign code
        if re.match(r"^[A-Za-z0-9./&_]+$", code):
            return code, summary

    return s, ""


def build_description_one_cell(sign: Dict[str, Any]) -> str:
    """
    Description should show ONLY the summary from sign_type (the part after '-').
    If no dash exists, fallback to the sign['description'].
    """
    raw_sign_type = safe_str(sign.get("sign_type"))
    _, summary = split_sign_type_and_summary(raw_sign_type)

    if summary:
        return summary.strip()

    return safe_str(sign.get("description")).strip()


# =========================================================
# Merge shifting helpers (Critical for Option 2)
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
    """
    Shift merged ranges if they are below the footer boundary OR overlap it.
    Prevents merge corruption when a merge spans across the insertion row.
    """
    a, b = parse_range(a1_range)
    a_col, a_row = split_ref(a)
    b_col, b_row = split_ref(b)

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
    """
    Copy styles + row height from src_row to dst_row.
    Includes borders.
    """
    ws.row_dimensions[dst_row].height = ws.row_dimensions[src_row].height

    for col in range(1, max_col + 1):
        src = ws.cell(row=src_row, column=col)
        dst = ws.cell(row=dst_row, column=col)

        if src.has_style:
            dst._style = copy(src._style)

        dst.number_format = src.number_format
        dst.alignment = copy(src.alignment)
        dst.border = copy(src.border)         # ✅ keeps borders
        dst.fill = copy(src.fill)
        dst.font = copy(src.font)
        dst.protection = copy(src.protection)


# =========================================================
# Body adjust (Option 2) with merge + style preservation
# =========================================================
def adjust_body_rows_preserve_footer(
    ws,
    sign_count: int,
    body_start: int = 28,
    body_end: int = 47,
    extra_blank_rows: int = 3
) -> int:
    """
    Ensures body has sign_count + extra_blank_rows rows.
    Inserts/deletes rows at the footer boundary while preserving merges and styles.

    Returns row_offset applied to the footer:
      positive -> footer moved down
      negative -> footer moved up
    """
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
        logging.info(f"Inserting {diff} row(s) at {footer_start} to expand body.")
        ws.insert_rows(footer_start, amount=diff)

        for r in range(footer_start, footer_start + diff):
            copy_row_style(ws, src_row=body_end, dst_row=r, max_col=max_col)

    else:
        delete_count = abs(diff)
        delete_start = body_start + needed_rows
        logging.info(f"Deleting {delete_count} row(s) at {delete_start} to shrink body.")
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
    """
    Approximates AutoFit:
    - column width ~50 => ~60 characters per line
    - line height = 15
    """
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

        insert_logo(ws)

        # ✅ Capture template footer row heights BEFORE any insertion/deletion
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

        # Clear the body rows we will use
        total_body_rows_needed = sign_count + EXTRA_BLANK
        for r in range(BODY_START, BODY_START + total_body_rows_needed):
            for c in ["A", "B", "C", "D", "E", "F"]:
                ws[f"{c}{r}"].value = None

        # ---------------- Write sign lines ----------------
        COL_ITEM, COL_SIGN_TYPE, COL_DESC, COL_QTY, COL_UNIT, COL_TOTAL = "A", "B", "C", "D", "E", "F"
        current_row = BODY_START
        item_num = 1

        for sign in sign_types:
            ws[f"{COL_ITEM}{current_row}"].value = item_num

            raw_type = safe_str(sign.get("sign_type"))
            clean_type, _ = split_sign_type_and_summary(raw_type)
            ws[f"{COL_SIGN_TYPE}{current_row}"].value = clean_type

            # ✅ Description shows only the summary
            ws[f"{COL_DESC}{current_row}"].value = build_description_one_cell(sign)

            ws[f"{COL_QTY}{current_row}"].value = safe_num(sign.get("qty"))

            unit_price = safe_num(sign.get("unit_price"))
            ws[f"{COL_UNIT}{current_row}"].value = round(unit_price) if unit_price is not None else None

            ws[f"{COL_TOTAL}{current_row}"].value = safe_num(sign.get("extended_total"))

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

        # ✅ Restore footer row heights to template values (shifted)
        restore_row_heights(ws, footer_row_heights, footer_row_offset)

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
